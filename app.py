"""
Payroll Reconciliation App  –  Stonelink Property Management
SOP-compliant automation of the Weekly Payroll & Hours Reporting process.

Outputs a two-tab Excel workbook:
  Tab 1 – Weekly Recap  (Planet Synergy PM Report format, §5.4–5.7)
           One row per TSheets shift entry, with SOP columns A–L,
           highlighted subtotal row per technician.
  Tab 2 – Payroll Hours  (§5.8)
           One row per employee: Regular / OT / PTO / Total
"""

import io
import math
import re
from collections import defaultdict

import pandas as pd
import pdfplumber
import streamlit as st
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

# ─────────────────────────────────────────────────────────────────────────────
# CONSTANTS
# ─────────────────────────────────────────────────────────────────────────────

MELD_RE = re.compile(r'\b(T[A-Z0-9]{5,10})\b')

# Keywords that make a shift entry non-billable (§5.5 / §5.6)
NON_BILLABLE_KW = [
    "non-billable", "not billable", "office", "meeting", "dump", "vertedero",
    "ofisina", "llaves", "keys", "paperwork", "estimate",
    "entrega", "reunión", "reunion",
    "these hours are not being charged",
    "mattress pick up at the office",
    "drop off", "pick up at the office",
    "clock-out",                 # Judith Hernandez entries
    "parking", "car cleaning",
]

# Exact SOP-required phrase for flat billing (§5.5 – zero variation allowed)
FLAT_BILL_PHRASE = "not billable- to be flat bill"

# SOP styling
DARK_BLUE    = PatternFill("solid", fgColor="1F4E79")
MED_BLUE     = PatternFill("solid", fgColor="2E75B6")
LIGHT_BLUE   = PatternFill("solid", fgColor="D6E4F0")
YELLOW_FILL  = PatternFill("solid", fgColor="FFE699")   # Paying > Billable
RED_FILL     = PatternFill("solid", fgColor="FF9999")   # missing meld
ORANGE_FILL  = PatternFill("solid", fgColor="FFB347")   # non-billable with no note (actionable error)
GREEN_FILL   = PatternFill("solid", fgColor="C6EFCE")   # subtotal
PURPLE_FILL  = PatternFill("solid", fgColor="D9B3FF")   # flat-bill entries
WHITE_FILL   = PatternFill("solid", fgColor="FFFFFF")
ALT_FILL     = PatternFill("solid", fgColor="EBF3FB")

WHITE_FONT  = Font(color="FFFFFF", bold=True, size=10)
BOLD_FONT   = Font(bold=True, size=10)
BASE_FONT   = Font(size=10)

_THIN = Side(border_style="thin", color="AAAAAA")
BOX  = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

OT_THRESHOLD = 40.0   # hours before overtime kicks in (§5.8)


# ─────────────────────────────────────────────────────────────────────────────
# HELPERS
# ─────────────────────────────────────────────────────────────────────────────

def extract_meld_ids(text: str) -> list[str]:
    return list(dict.fromkeys(MELD_RE.findall(text.upper())))


def is_non_billable(notes: str) -> bool:
    low = notes.lower()
    return any(kw in low for kw in NON_BILLABLE_KW)


def is_flat_bill(notes: str) -> bool:
    """Return True if notes contain the exact SOP flat-billing phrase (§5.5)."""
    return FLAT_BILL_PHRASE in notes.lower()


def calc_billable(actual: float, non_billable: bool) -> float:
    """
    SOP §5.5:
    • Non-billable → 0
    • Minimum billable = 1 hour
    • Round UP to nearest 0.25 hour (ceiling, not nearest)
    """
    if non_billable:
        return 0.0
    rounded = math.ceil(actual * 4) / 4      # round UP to nearest 0.25
    return max(1.0, rounded)


def non_billable_note(notes: str) -> str:
    """
    Return a Notes-column (col J) string for non-billable entries.
    If the notes already contain a reason, pass it through as-is.
    The exact flat-bill phrase is preserved verbatim per SOP §5.5.
    """
    low = notes.lower()
    # If technician already wrote a proper non-billable note, use it directly
    if "not billable" in low or "non-billable" in low:
        # Extract the first sentence/clause after the non-billable marker
        for phrase in ["not billable-", "not billable –", "not billable -",
                       "non-billable –", "non-billable -"]:
            if phrase in low:
                idx = low.index(phrase)
                return notes[idx:].split("\n")[0].strip()
        return notes.split("\n")[0].strip()
    # Auto-generate a reason
    if "office" in low or "meeting" in low or "reunión" in low or "reunion" in low:
        return "Not billable – office/meeting time"
    if "dump" in low or "vertedero" in low:
        return "Not billable – dump run"
    if "keys" in low or "llaves" in low or "entrega" in low or "drop off" in low:
        return "Not billable – key drop/pickup"
    if "clock-out" in low:
        return "Not billable – clock-out entry"
    if "these hours are not being charged" in low:
        return "Not billable – technician noted no charge"
    if "drove" in low or "drive" in low or "travel" in low:
        return "Not billable – travel time"
    return "Not billable – admin/other"


# ─────────────────────────────────────────────────────────────────────────────
# PDF PARSER  (QB Time / TSheets payroll report)
# ─────────────────────────────────────────────────────────────────────────────

# Matches a line that contains ONLY digits, dots, and spaces (numeric-only lines)
_NUMERIC_LINE_RE = re.compile(r'^[\d\s.]+$')
# Matches a line of exactly 5 space-separated floats (single-line summary format)
_FIVE_FLOATS_RE  = re.compile(
    r'^(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)\s+(\d+\.\d+)$'
)
# Matches summary with labels on one line:
# "Regular 37.34 PTO 0.00 OT 0.00 DT 0.00 Total Hours 37.34"
_LABELED_SUMMARY_RE = re.compile(
    r'Regular\s+(\d+\.\d+)\s+PTO\s+(\d+\.\d+)\s+OT\s+(\d+\.\d+)'
    r'\s+DT\s+(\d+\.\d+)\s+Total\s+Hours?\s+(\d+\.\d+)',
    re.IGNORECASE
)
# Words that are never part of an employee name
_NON_NAME_WORDS = {
    'Regular', 'PTO', 'OT', 'DT', 'Total', 'Hours', 'Time', 'in', 'out',
    'Duration', 'Generated', 'for', 'Stonelink', 'Property', 'Management',
    'Shift',
}


def _is_name_word(w: str) -> bool:
    """Return True if word looks like part of a proper name."""
    return bool(w) and w[0].isupper() and re.match(r'^[A-Za-z]+$', w) \
           and w not in _NON_NAME_WORDS


def _extract_name_words(line: str) -> list[str]:
    """Pull name-like words from a line, stopping at non-name tokens."""
    words = []
    for w in line.split():
        if _is_name_word(w):
            words.append(w)
        else:
            break   # stop at first non-name token (numbers, labels, etc.)
    return words


def _find_name(lines: list[str], date_idx: int) -> str:
    """
    Robustly locate the employee name in the lines near the date-range line.
    Handles:
      • Name on the line immediately before the date range
      • Name split across two lines (e.g. "Leonardo\\nGonzalez")
      • Lines that contain name + summary labels on the same line
      • Numeric-only lines appearing between the name and date range
    """
    candidate_parts: list[str] = []

    for back in range(1, 12):
        idx = date_idx - back
        if idx < 0:
            break
        line = lines[idx].strip()
        if not line:
            continue
        if line.startswith("Generated"):
            break
        if re.match(r'^\d{2}/\d{2}/\d{4}', line):   # another date range → stop
            break
        if _NUMERIC_LINE_RE.match(line):              # skip pure-numeric lines
            continue

        words = _extract_name_words(line)
        if words:
            candidate_parts = words + candidate_parts
            # If we already have 2+ words, that's enough for a full name
            if len(candidate_parts) >= 2:
                break
            # Otherwise keep looking backwards for a first name
            continue
        # If no name words found, stop searching further back
        break

    return " ".join(candidate_parts) if candidate_parts else ""


def parse_qb_pdf(file_obj) -> list[dict]:
    """
    Parse the TSheets Payroll PDF and return a list of employee dicts.

    Each dict:
      name, period_start, period_end,
      regular, pto, ot, dt, total_hours,
      days: {date_label: daily_total},
      shifts: list of shift dicts
    """
    with pdfplumber.open(file_obj) as pdf:
        full_text = "\n".join(p.extract_text() or "" for p in pdf.pages)
    return _parse_lines(full_text)


def _parse_lines(full_text: str) -> list[dict]:
    lines = full_text.splitlines()

    date_range_re = re.compile(r'(\d{2}/\d{2}/\d{4})\s+to\s+(\d{2}/\d{2}/\d{4})')
    month_re = re.compile(
        r'^(January|February|March|April|May|June|July|August|'
        r'September|October|November|December)\s+\d{1,2},\s+\d{4}'
    )
    shift_re = re.compile(
        r'^(\d{1,2}:\d{2}[ap]m(?:\s*\(EDT\))?)\s+'
        r'(\d{1,2}:\d{2}[ap]m(?:\s*\(EDT\))?)\s+'
        r'(\d+\.\d+)\s+Shift Total'
    )
    single_float_re = re.compile(r'^\d+\.\d+$')

    employees: list[dict] = []
    emp: dict | None = None
    shift: dict | None = None
    cur_date = cur_date_iso = ""
    reading_notes = False
    notes_buf = ""
    expect: dict = {}

    def _flush_shift():
        nonlocal shift, reading_notes, notes_buf
        if shift and emp is not None:
            if reading_notes:
                shift["notes"] = notes_buf.strip()
                shift["meld_ids"] = extract_meld_ids(shift["notes"])
                shift["non_billable"] = is_non_billable(shift["notes"])
            emp["shifts"].append(shift)
        shift = None
        reading_notes = False
        notes_buf = ""

    def _flush_emp():
        _flush_shift()
        if emp and emp.get("name"):
            employees.append(emp)

    i = 0
    while i < len(lines):
        line = lines[i].strip()

        # ── New employee block (date-range line) ──────────────────────────
        m = date_range_re.match(line)
        if m:
            name = _find_name(lines, i)

            # Summary location varies by PDF version:
            #   A) Same line as date range: "03/08/2026 to 03/14/2026  37.34 0.00 ..."
            #   B) Line BEFORE date range:  "37.34 0.00 0.00 0.00 37.34" / "03/08/2026 ..."
            #   C) Line AFTER date range
            summary_from_date_line = _FIVE_FLOATS_RE.search(
                line[len(m.group(0)):].strip()
            )

            # Check line BEFORE date range (most common in this PDF version)
            prev_five = None
            if not summary_from_date_line and i > 0:
                prev_five = _FIVE_FLOATS_RE.match(lines[i - 1].strip())

            _flush_emp()
            emp = dict(name=name, period_start=m.group(1), period_end=m.group(2),
                       regular=0.0, pto=0.0, ot=0.0, dt=0.0, total_hours=0.0,
                       days={}, shifts=[])
            expect = {"regular": True}

            def _apply_five(mx):
                emp["regular"]     = float(mx.group(1))
                emp["pto"]         = float(mx.group(2))
                emp["ot"]          = float(mx.group(3))
                emp["dt"]          = float(mx.group(4))
                emp["total_hours"] = float(mx.group(5))
                expect.clear()

            if summary_from_date_line:
                _apply_five(summary_from_date_line)
            elif prev_five:
                _apply_five(prev_five)
            else:
                # Check if NEXT line is a five-float or labeled summary
                if i + 1 < len(lines):
                    nxt = lines[i + 1].strip()
                    m5 = _FIVE_FLOATS_RE.match(nxt)
                    if m5:
                        _apply_five(m5)
                        i += 2; continue
                    ml = _LABELED_SUMMARY_RE.search(nxt)
                    if ml:
                        emp["regular"]     = float(ml.group(1))
                        emp["pto"]         = float(ml.group(2))
                        emp["ot"]          = float(ml.group(3))
                        emp["dt"]          = float(ml.group(4))
                        emp["total_hours"] = float(ml.group(5))
                        expect.clear()
                        i += 2; continue

            i += 1; continue

        if emp is None:
            i += 1; continue

        # ── Summary: per-line format ("Regular\n37.34\nPTO\n0.00 ...") ───
        for lbl, key in [("Regular","regular"), ("PTO","pto"),
                          ("OT","ot"), ("DT","dt"), ("Total Hours","total_hours")]:
            if line == lbl:
                expect[key] = True
                break
        else:
            for key in list(expect):
                if expect.get(key) and single_float_re.match(line):
                    emp[key] = float(line)
                    expect[key] = False
                    break

        # ── Summary: five-float on one line (fallback anywhere in block) ──
        if not emp["total_hours"] and _FIVE_FLOATS_RE.match(line):
            m5 = _FIVE_FLOATS_RE.match(line)
            emp["regular"]     = float(m5.group(1))
            emp["pto"]         = float(m5.group(2))
            emp["ot"]          = float(m5.group(3))
            emp["dt"]          = float(m5.group(4))
            emp["total_hours"] = float(m5.group(5))
            expect = {}
            i += 1; continue

        # ── Summary: labeled single-line format (fallback anywhere in block) ──
        # "Regular 37.34 PTO 0.00 OT 0.00 DT 0.00 Total Hours 37.34"
        if not emp["total_hours"]:
            ml = _LABELED_SUMMARY_RE.search(line)
            if ml:
                emp["regular"]     = float(ml.group(1))
                emp["pto"]         = float(ml.group(2))
                emp["ot"]          = float(ml.group(3))
                emp["dt"]          = float(ml.group(4))
                emp["total_hours"] = float(ml.group(5))
                expect = {}
                i += 1; continue

        # ── Date header (e.g. "March 9, 2026  7.01") ─────────────────────
        if month_re.match(line):
            _flush_shift()
            # Daily total sometimes on the same line: "March 9, 2026 7.01"
            parts = line.split()
            cur_date = " ".join(parts[:3]) if len(parts) >= 3 else line
            cur_date_iso = _date_iso(cur_date)
            # Try same-line daily total
            if len(parts) == 4 and single_float_re.match(parts[3]):
                emp["days"][cur_date] = float(parts[3])
                i += 1; continue
            # Try next-line daily total
            if i + 1 < len(lines) and single_float_re.match(lines[i+1].strip()):
                emp["days"][cur_date] = float(lines[i+1].strip())
                i += 2; continue
            i += 1; continue

        # ── Shift row ─────────────────────────────────────────────────────
        ms = shift_re.match(line)
        if ms:
            _flush_shift()
            shift = dict(
                date=cur_date, date_iso=cur_date_iso,
                time_in=ms.group(1).replace("(EDT)","").strip(),
                time_out=ms.group(2).replace("(EDT)","").strip(),
                duration=float(ms.group(3)),
                notes="", meld_ids=[], non_billable=False,
            )
            i += 1; continue

        # ── Notes ─────────────────────────────────────────────────────────
        if line.startswith("NOTES:") and shift:
            reading_notes = True
            notes_buf = line[6:].strip()
            i += 1; continue

        if reading_notes and shift:
            if line and not month_re.match(line) and not shift_re.match(line) \
                    and not line.startswith("Generated"):
                notes_buf += " " + line
                i += 1; continue
            else:
                shift["notes"] = notes_buf.strip()
                shift["meld_ids"] = extract_meld_ids(shift["notes"])
                shift["non_billable"] = is_non_billable(shift["notes"])
                reading_notes = False
                notes_buf = ""
                # fall through — re-process this line

        if line.startswith("Generated for"):
            _flush_emp()
            emp = None
            i += 1; continue

        i += 1

    _flush_emp()
    return employees


def _date_iso(s: str) -> str:
    MONTHS = {"January":"01","February":"02","March":"03","April":"04",
               "May":"05","June":"06","July":"07","August":"08",
               "September":"09","October":"10","November":"11","December":"12"}
    m = re.match(r'(\w+)\s+(\d+),\s+(\d+)', s)
    if not m: return ""
    return f"{m.group(3)}-{MONTHS.get(m.group(1),'00')}-{m.group(2).zfill(2)}"


# ─────────────────────────────────────────────────────────────────────────────
# CSV / MELD LOADER
# ─────────────────────────────────────────────────────────────────────────────

def load_melds(file_obj) -> pd.DataFrame:
    """
    Accepts either:
      • melds_report CSV  (Meld Number, Property Name, Unit, Work Category, …)
      • Work Log Summary  (Meld, Agent, Unit, Title, Description, Check-In Hours, Address Line 1)
    Returns DataFrame indexed by uppercase Meld Number.
    """
    df = pd.read_csv(file_obj, dtype=str)
    df.columns = df.columns.str.strip()
    # Normalise Meld Number column (two possible names)
    if "Meld Number" in df.columns:
        df = df.rename(columns={"Meld Number": "Meld"})
    if "Meld" not in df.columns:
        raise ValueError("CSV must contain a 'Meld' or 'Meld Number' column.")
    df["Meld"] = df["Meld"].str.strip().str.upper()
    # Numeric hours
    for col in ["Total Labor Hours", "Check-In Hours"]:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    df = df[~df["Meld"].duplicated(keep="first")]
    return df.set_index("Meld")


def meld_lookup(row: pd.Series | None, field: str, default="") -> str:
    if row is None:
        return default
    val = row.get(field, default)
    if isinstance(val, pd.Series):
        val = val.iloc[0] if len(val) > 0 else default
    return str(val if val is not None else "").strip()


# ─────────────────────────────────────────────────────────────────────────────
# BUILD ROW DATA
# ─────────────────────────────────────────────────────────────────────────────

def build_recap_rows(employees: list[dict], melds_df: pd.DataFrame) -> list[dict]:
    """
    Build one dict per TSheets shift entry (SOP §5.4–5.5).
    Columns match the SOP's Weekly Recap (Planet Synergy PM Report) layout.
    """
    rows = []
    for emp in employees:
        for sh in emp["shifts"]:
            primary = sh["meld_ids"][0] if sh["meld_ids"] else ""
            mrow = melds_df.loc[primary] if primary and primary in melds_df.index else None

            # Resolve address: prefer Meld CSV, fall back to notes
            address = (meld_lookup(mrow, "Address Line 1")
                       or meld_lookup(mrow, "Property Name")
                       or "")
            unit     = meld_lookup(mrow, "Unit")
            trade    = (meld_lookup(mrow, "Work Category")
                        or meld_lookup(mrow, "Title")
                        or meld_lookup(mrow, "Work Type")
                        or "")
            desc     = (meld_lookup(mrow, "Description")
                        or meld_lookup(mrow, "Meld Description")
                        or "")
            status   = meld_lookup(mrow, "Meld Status")

            nb       = sh["non_billable"]
            flat_b   = is_flat_bill(sh["notes"])
            actual   = sh["duration"]
            paying   = actual                          # coordinator adjusts in Excel
            billable = calc_billable(actual, nb)

            # Auto-note for column J
            j_note = ""
            if flat_b:
                j_note = "Not billable- to be flat bill"   # exact SOP phrase §5.5
            elif nb:
                j_note = non_billable_note(sh["notes"])
            elif len(sh["meld_ids"]) > 1:
                j_note = "Multi-meld entry: " + ", ".join(sh["meld_ids"][1:])
            elif primary and mrow is None:
                j_note = f"⚠ Meld {primary} not found in CSV – verify"

            # Flag: non-billable entry with no explanatory note (§5.5 actionable error)
            nb_missing_note = nb and not j_note.strip()

            rows.append({
                # SOP column mapping
                "MWO":        primary,                 # A
                "Tech":       emp["name"],             # B
                "Address":    address,                 # C
                "Unit":       unit,                    # D
                "Check-In":   sh["time_in"],           # E
                "Actual":     actual,                  # F  ← from TSheets
                "Paying":     paying,                  # G  ← editable, starts = Actual
                "Billable":   billable,                # H  ← SOP-computed
                # (col I intentionally blank)
                "Notes":      j_note,                  # J
                "Trade":      trade,                   # K
                "Work Performed": desc,                # L
                # Extra metadata (not in SOP layout, used for flags)
                "_employee":       emp["name"],
                "_date":           sh["date"],
                "_date_iso":       sh["date_iso"],
                "_time_out":       sh["time_out"],
                "_meld_ids":       sh["meld_ids"],
                "_non_bill":       nb,
                "_flat_bill":      flat_b,
                "_nb_missing_note":nb_missing_note,
                "_meld_found":     mrow is not None,
                "_raw_notes":      sh["notes"],
                "_status":         status,
            })
    return rows


def build_payroll_rows(employees: list[dict]) -> list[dict]:
    """
    SOP §5.8: Payroll Hours Spreadsheet
    Regular (≤40 h from Paying), OT (>40 h), PTO, Total.
    PTO does NOT count toward overtime.
    """
    rows = []
    for emp in employees:
        paying_total = emp["total_hours"] - emp["pto"]   # exclude PTO from work hours
        regular = min(paying_total, OT_THRESHOLD)
        ot = max(0.0, paying_total - OT_THRESHOLD)
        rows.append({
            "Employee":    emp["name"],
            "Regular":     regular,
            "Overtime":    ot,
            "PTO":         emp["pto"],
            "DT":          emp["dt"],
            "Total Hours": emp["total_hours"],
            # Validation flag
            "_ot_flag":    ot > 0,
            "_pto_flag":   emp["pto"] > 0,
        })
    return rows


# ─────────────────────────────────────────────────────────────────────────────
# EXCEL BUILDER
# ─────────────────────────────────────────────────────────────────────────────

def _c(ws, r, c, val="", fill=None, font=None, align="left", fmt=None,
       bold=False, wrap=False, border=True):
    cell = ws.cell(row=r, column=c, value=val)
    if fill:  cell.fill = fill
    if font:  cell.font = font
    elif bold: cell.font = BOLD_FONT
    cell.alignment = Alignment(horizontal=align, vertical="center", wrap_text=wrap)
    if fmt:   cell.number_format = fmt
    if border: cell.border = BOX
    return cell


def build_excel(recap_rows: list[dict], payroll_rows: list[dict],
                employees: list[dict]) -> bytes:
    wb = Workbook()
    wb.remove(wb.active)
    _sheet_weekly_recap(wb, recap_rows, employees)
    _sheet_payroll_hours(wb, payroll_rows)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ── Tab 1: Weekly Recap ──────────────────────────────────────────────────────

RECAP_COLS = [
    ("A", "MWO",           16),
    ("B", "Tech",          20),
    ("C", "Address",       28),
    ("D", "Unit",          10),
    ("E", "Check-In",      10),
    ("F", "Actual",        10),
    ("G", "Paying",        10),
    ("H", "Billable",      10),
    ("I", "",               6),   # intentionally blank per SOP
    ("J", "Notes",         30),
    ("K", "Trade",         18),
    ("L", "Work Performed",38),
]


def _sheet_weekly_recap(wb: Workbook, rows: list[dict], employees: list[dict]):
    ws = wb.create_sheet("Weekly Recap")
    ws.sheet_view.showGridLines = False

    # ── Title row ─────────────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    title = ws.cell(row=1, column=1, value="Weekly Recap – Planet Synergy PM Report")
    title.font = Font(bold=True, size=13, color="FFFFFF")
    title.fill = DARK_BLUE
    title.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Period label (derive from first employee)
    period = ""
    if employees:
        period = f"Pay Period: {employees[0]['period_start']}  –  {employees[0]['period_end']}"
    ws.merge_cells("A2:L2")
    pc = ws.cell(row=2, column=1, value=period)
    pc.font = Font(bold=True, size=10, color="FFFFFF")
    pc.fill = MED_BLUE
    pc.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 16

    # ── Column headers (row 3) ────────────────────────────────────────────
    for idx, (_, hdr, width) in enumerate(RECAP_COLS, 1):
        ltr = get_column_letter(idx)
        ws.column_dimensions[ltr].width = width
        _c(ws, 3, idx, hdr, fill=DARK_BLUE, font=WHITE_FONT, align="center", wrap=True)
    ws.row_dimensions[3].height = 28
    ws.freeze_panes = "A4"

    # ── Data rows, grouped by technician ─────────────────────────────────
    # Sort: Tech A–Z, then date, then check-in time
    sorted_rows = sorted(rows, key=lambda r: (r["Tech"], r["_date_iso"], r["Check-In"]))

    data_row = 4
    emp_groups: dict[str, list[dict]] = defaultdict(list)
    for r in sorted_rows:
        emp_groups[r["Tech"]].append(r)

    # ── SOP validation summary dict (for Streamlit display)
    flags: list[dict] = []

    for emp_name, emp_rows in emp_groups.items():
        fill_idx = 0
        f_start = data_row   # for subtotal formula

        for r in emp_rows:
            row_fill = ALT_FILL if fill_idx % 2 == 0 else WHITE_FILL
            fill_idx += 1

            # SOP flag overrides (priority order: red > orange > yellow > purple)
            paying_gt_billable  = r["Paying"] > r["Billable"] and r["Billable"] > 0
            meld_missing        = r["MWO"] and not r["_meld_found"]
            nb_no_note          = r.get("_nb_missing_note", False)
            flat_bill           = r.get("_flat_bill", False)

            if meld_missing:
                row_fill = RED_FILL
                flags.append({"type": "Missing Meld", "employee": emp_name,
                              "meld": r["MWO"], "date": r["_date"],
                              "hours": r["Actual"]})
            elif nb_no_note:
                row_fill = ORANGE_FILL
                flags.append({"type": "⚠ Non-Billable – No Note (Actionable Error §5.5)",
                              "employee": emp_name, "meld": r["MWO"],
                              "date": r["_date"], "hours": r["Actual"]})
            elif paying_gt_billable:
                row_fill = YELLOW_FILL
                flags.append({"type": "Paying > Billable", "employee": emp_name,
                              "meld": r["MWO"], "date": r["_date"],
                              "actual": r["Actual"], "billable": r["Billable"]})
            elif flat_bill:
                row_fill = PURPLE_FILL

            vals = [
                r["MWO"], r["Tech"], r["Address"], r["Unit"],
                r["Check-In"],
                r["Actual"],    # F
                r["Paying"],    # G
                r["Billable"],  # H
                "",             # I  blank
                r["Notes"],     # J
                r["Trade"],     # K
                r["Work Performed"],  # L
            ]
            for c_idx, val in enumerate(vals, 1):
                is_num = c_idx in (6, 7, 8)
                _c(ws, data_row, c_idx, val,
                   fill=row_fill,
                   align="center" if is_num else "left",
                   fmt="0.00" if is_num else None,
                   wrap=(c_idx in (3, 10, 12)))
            data_row += 1

        # ── Subtotal row per tech (§5.7) ──────────────────────────────────
        f_end = data_row - 1
        col_f = get_column_letter(6)
        col_g = get_column_letter(7)
        col_h = get_column_letter(8)

        _c(ws, data_row, 1, f"Subtotal – {emp_name}", fill=GREEN_FILL,
           bold=True, align="left")
        ws.merge_cells(f"A{data_row}:E{data_row}")
        ws.cell(row=data_row, column=1).fill = GREEN_FILL

        for ci, col_ltr in [(6, col_f), (7, col_g), (8, col_h)]:
            cell = ws.cell(row=data_row, column=ci,
                           value=f"=SUM({col_ltr}{f_start}:{col_ltr}{f_end})")
            cell.fill = GREEN_FILL
            cell.font = BOLD_FONT
            cell.number_format = "0.00"
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.border = BOX

        for ci in [9, 10, 11, 12]:
            ws.cell(row=data_row, column=ci).fill = GREEN_FILL
            ws.cell(row=data_row, column=ci).border = BOX

        ws.row_dimensions[data_row].height = 16
        data_row += 1   # blank separator
        _c(ws, data_row, 1, "", fill=WHITE_FILL, border=False)
        data_row += 1

    # ── Grand totals ──────────────────────────────────────────────────────
    gt_row = data_row
    ws.merge_cells(f"A{gt_row}:E{gt_row}")
    _c(ws, gt_row, 1, "GRAND TOTAL", fill=DARK_BLUE, font=WHITE_FONT, align="left")
    for ci in range(2, 6):
        ws.cell(row=gt_row, column=ci).fill = DARK_BLUE
        ws.cell(row=gt_row, column=ci).border = BOX

    for ci, ltr in [(6, "F"), (7, "G"), (8, "H")]:
        cell = ws.cell(row=gt_row, column=ci,
                       value=f"=SUMIF(B4:B{data_row-1},\"<>\"&\"Subtotal*\","
                             f"{ltr}4:{ltr}{data_row-1})")
        cell.fill = DARK_BLUE
        cell.font = WHITE_FONT
        cell.number_format = "0.00"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX

    for ci in [9, 10, 11, 12]:
        ws.cell(row=gt_row, column=ci).fill = DARK_BLUE
        ws.cell(row=gt_row, column=ci).border = BOX

    ws.row_dimensions[gt_row].height = 18

    # ── Legend ────────────────────────────────────────────────────────────
    leg_row = gt_row + 2
    ws.cell(row=leg_row, column=1, value="Legend:").font = BOLD_FONT
    items = [
        (RED_FILL,    "Meld ID not found in CSV – verify entry (§6)"),
        (ORANGE_FILL, "Non-billable entry missing explanation – actionable error (§5.5)"),
        (YELLOW_FILL, "Paying > Billable – manager review required (§6)"),
        (PURPLE_FILL, "Flat-bill entry – 'Not billable- to be flat bill' (§5.5)"),
        (GREEN_FILL,  "Technician subtotal row (§5.7)"),
    ]
    for i, (fill, label) in enumerate(items):
        ws.cell(row=leg_row + i + 1, column=1).fill = fill
        ws.cell(row=leg_row + i + 1, column=1).border = BOX
        ws.cell(row=leg_row + i + 1, column=2, value=label)

    # Store flag list on the worksheet for later use in Streamlit
    ws._sop_flags = flags


# ── Tab 2: Payroll Hours ─────────────────────────────────────────────────────

def _sheet_payroll_hours(wb: Workbook, rows: list[dict]):
    ws = wb.create_sheet("Payroll Hours")
    ws.sheet_view.showGridLines = False

    # Title
    ws.merge_cells("A1:F1")
    t = ws.cell(row=1, column=1, value="Payroll Hours Spreadsheet  (§5.8)")
    t.font = Font(bold=True, size=13, color="FFFFFF")
    t.fill = DARK_BLUE
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 22

    # Note
    ws.merge_cells("A2:F2")
    n = ws.cell(row=2, column=1,
                value="Regular ≤ 40 hrs  |  OT = hours > 40  |  PTO excluded from OT calc")
    n.font = Font(italic=True, size=9, color="FFFFFF")
    n.fill = MED_BLUE
    n.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[2].height = 14

    # Headers
    hdrs = ["Employee", "Regular", "Overtime", "PTO", "DT", "Total Hours"]
    widths = [24, 12, 12, 10, 10, 14]
    for c, (h, w) in enumerate(zip(hdrs, widths), 1):
        ws.column_dimensions[get_column_letter(c)].width = w
        _c(ws, 3, c, h, fill=DARK_BLUE, font=WHITE_FONT, align="center")
    ws.row_dimensions[3].height = 26
    ws.freeze_panes = "A4"

    # Data
    for r_idx, row in enumerate(rows, 4):
        fill = ALT_FILL if r_idx % 2 == 0 else WHITE_FILL
        _c(ws, r_idx, 1, row["Employee"], fill=fill, bold=True)
        for c_idx, key in enumerate(["Regular","Overtime","PTO","DT","Total Hours"], 2):
            cell = _c(ws, r_idx, c_idx, row[key], fill=fill,
                      align="center", fmt="0.00")
        # Highlight OT
        if row["_ot_flag"]:
            ws.cell(row=r_idx, column=3).fill = YELLOW_FILL
            ws.cell(row=r_idx, column=3).font = BOLD_FONT
        if row["_pto_flag"]:
            ws.cell(row=r_idx, column=4).fill = LIGHT_BLUE

    # Totals row
    tot = len(rows) + 4
    ws.merge_cells(f"A{tot}:A{tot}")
    _c(ws, tot, 1, "TOTAL", fill=DARK_BLUE, font=WHITE_FONT, bold=True, align="center")
    for c in range(2, 7):
        ltr = get_column_letter(c)
        cell = ws.cell(row=tot, column=c,
                       value=f"=SUM({ltr}4:{ltr}{tot-1})")
        cell.fill = DARK_BLUE
        cell.font = WHITE_FONT
        cell.number_format = "0.00"
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = BOX


# ─────────────────────────────────────────────────────────────────────────────
# STREAMLIT UI
# ─────────────────────────────────────────────────────────────────────────────

def main():
    st.set_page_config(
        page_title="Payroll Reconciliation – Stonelink",
        page_icon="📊",
        layout="wide",
    )

    st.title("📊 Weekly Payroll Reconciliation")
    st.caption("Stonelink Property Management  ·  TSheets × Property Meld  ·  SOP-Compliant")

    # ── File uploads ──────────────────────────────────────────────────────
    c1, c2 = st.columns(2)
    with c1:
        pdf_file = st.file_uploader(
            "TSheets / QB Time Payroll PDF",
            type=["pdf"],
            help="Weekly payroll report exported from QuickBooks Time — system of record for actual hours (SOP §4)",
        )
    with c2:
        csv_file = st.file_uploader(
            "Property Meld Report CSV",
            type=["csv"],
            help="Either the melds_report CSV or the Work Log Summary CSV from Property Meld",
        )

    if not pdf_file or not csv_file:
        st.info("Upload both files above to generate the payroll report.")
        with st.expander("SOP Summary (Updated 3-26-2026)"):
            st.markdown("""
**Three-tier hours system (§5.5):**
| Column | Name | Source | Rule |
|--------|------|--------|------|
| F | **Actual** | TSheets | Exact hours worked — must match TSheets total |
| G | **Paying** | Coordinator | Approved payable hours — may differ from Actual if unreasonable |
| H | **Billable** | Computed | Min 1 hr, rounded **UP** to nearest 0.25 hr; 0 for non-billable |

**Key SOP rules (Updated §5.5):**
- Paying ≤ Billable (flag if violated — manager review required)
- PTO does **not** count toward overtime
- OT = Paying hours over 40 per week
- Non-billable → Billable = 0 + **detailed explanation REQUIRED** in Notes col J (failure if missing)
- Flat billing requires **exact phrase**: `Not billable- to be flat bill` (no variation)
- All adjustments (Paying ≠ Actual) require a note + manager approval

**Flag colors in Excel output:**
- 🔴 Red = Meld ID not found in CSV
- 🟠 Orange = Non-billable entry missing explanation note *(actionable error)*
- 🟡 Yellow = Paying > Billable *(manager review)*
- 🟣 Purple = Flat-bill entry
- 🟢 Green = Technician subtotal row
            """)
        return

    # ── Parse & process ───────────────────────────────────────────────────
    with st.spinner("Parsing TSheets PDF…"):
        try:
            employees = parse_qb_pdf(pdf_file)
        except Exception as e:
            st.error(f"PDF parse error: {e}")
            return

    # ── DEBUG: show raw PDF text + parsed employee summaries ──────────────
    with st.expander("🔍 Debug – Raw PDF Text & Parsed Summaries (close when done)", expanded=False):
        pdf_file.seek(0)
        import pdfplumber as _plumber
        with _plumber.open(pdf_file) as _pdf:
            _raw = "\n".join(p.extract_text() or "" for p in _pdf.pages)
        st.text_area("Raw PDF text (first 4000 chars)", _raw[:4000], height=300)
        st.markdown("**Parsed employee summaries:**")
        for e in employees:
            st.json({k: v for k, v in e.items() if k != "shifts"})
        pdf_file.seek(0)

    with st.spinner("Loading Meld CSV…"):
        try:
            melds_df = load_melds(csv_file)
        except Exception as e:
            st.error(f"CSV load error: {e}")
            return

    with st.spinner("Applying SOP rules…"):
        recap_rows   = build_recap_rows(employees, melds_df)
        payroll_rows = build_payroll_rows(employees)

    # ── Metrics bar (SOP §8 Summary Reporting) ────────────────────────────
    st.divider()
    m1, m2, m3, m4, m5, m6, m7, m8 = st.columns(8)
    m1.metric("Employees", len(employees))
    m2.metric("Time Entries", len(recap_rows))

    total_actual   = sum(r["Actual"]   for r in recap_rows)
    total_billable = sum(r["Billable"] for r in recap_rows)
    flat_bill_hrs  = sum(r["Actual"] for r in recap_rows if r.get("_flat_bill"))
    missing_melds  = sum(1 for r in recap_rows if r["MWO"] and not r["_meld_found"])
    total_ot_hrs   = sum(r["Overtime"] for r in payroll_rows)

    m3.metric("Total Actual Hrs",    f"{total_actual:.2f}")
    m4.metric("Total Billable Hrs",  f"{total_billable:.2f}")
    m5.metric("Flat Billed Hrs",     f"{flat_bill_hrs:.2f}")
    m6.metric("⚠ Unmatched Melds",  missing_melds)
    m7.metric("Total OT Hrs",        f"{total_ot_hrs:.2f}")

    nb_no_note_count = sum(1 for r in recap_rows if r.get("_nb_missing_note"))
    m8.metric("🚨 Missing NB Notes", nb_no_note_count)

    # ── SOP validation checks (§6 Conditional Logic + §5.5 Controls) ────
    violations = []
    actionable_errors = []
    for r in recap_rows:
        if r["Paying"] > r["Billable"] and r["Billable"] > 0:
            violations.append(f"**{r['Tech']}** – {r['_date']} – {r['MWO']}: "
                               f"Paying {r['Paying']:.2f} > Billable {r['Billable']:.2f} – manager review required")
        if r["Billable"] == 0 and not r["_non_bill"] and r["MWO"]:
            violations.append(f"**{r['Tech']}** – {r['_date']} – {r['MWO']}: "
                               f"Billable = 0 but not marked non-billable")
        if r.get("_nb_missing_note"):
            actionable_errors.append(f"**{r['Tech']}** – {r['_date']} – {r['MWO'] or '(no meld)'}: "
                                     f"Non-billable entry has NO explanation note (§5.5 HARD REQUIREMENT)")

    if actionable_errors:
        with st.expander(f"🚨 {len(actionable_errors)} ACTIONABLE ERRORS – Non-Billable Notes Missing (§5.5)", expanded=True):
            st.error("SOP §5.5 requires a detailed note for EVERY non-billable entry. These entries will FAIL audit.")
            for e in actionable_errors:
                st.error(e)

    if violations:
        with st.expander(f"⚠ {len(violations)} SOP Violations (§6 Conditional Logic)", expanded=True):
            for v in violations:
                st.warning(v)

    # ── Preview tabs ──────────────────────────────────────────────────────
    t1, t2, t3 = st.tabs(["Weekly Recap (preview)", "Payroll Hours", "Flags & Exceptions"])

    with t1:
        disp_cols = ["Tech","_date","MWO","Address","Unit","Check-In",
                     "Actual","Paying","Billable","Notes","Trade"]
        df_disp = pd.DataFrame(recap_rows)[disp_cols].rename(
            columns={"_date": "Date"})

        def _row_color(row):
            if row["MWO"] and not pd.isna(row["MWO"]):
                # can't easily check _meld_found here, just show data
                pass
            return [""] * len(row)

        st.dataframe(df_disp, use_container_width=True, hide_index=True,
                     column_config={
                         "Actual":   st.column_config.NumberColumn(format="%.2f"),
                         "Paying":   st.column_config.NumberColumn(format="%.2f"),
                         "Billable": st.column_config.NumberColumn(format="%.2f"),
                     })
        st.caption("Column G (Paying) starts equal to Actual. Adjust in the downloaded Excel where time seems unreasonable — manager approval required per SOP §5.5.")

    with t2:
        df_pay = pd.DataFrame(payroll_rows).drop(columns=["_ot_flag","_pto_flag"])
        st.dataframe(df_pay, use_container_width=True, hide_index=True,
                     column_config={
                         "Regular":     st.column_config.NumberColumn(format="%.2f"),
                         "Overtime":    st.column_config.NumberColumn(format="%.2f"),
                         "PTO":         st.column_config.NumberColumn(format="%.2f"),
                         "DT":          st.column_config.NumberColumn(format="%.2f"),
                         "Total Hours": st.column_config.NumberColumn(format="%.2f"),
                     })
        st.caption("OT = hours over 40 from Paying column. PTO excluded from OT calc per SOP §5.8.")

    with t3:
        # ── Actionable errors: non-billable with no note ───────────────────
        nb_no_note_rows = [r for r in recap_rows if r.get("_nb_missing_note")]
        if nb_no_note_rows:
            st.error(f"🚨 {len(nb_no_note_rows)} non-billable entries are missing an explanation note (§5.5 HARD REQUIREMENT)")
            st.dataframe(pd.DataFrame(nb_no_note_rows)[["Tech","_date","MWO","Actual","_raw_notes"]]
                         .rename(columns={"_date":"Date","_raw_notes":"Technician Notes"}),
                         use_container_width=True, hide_index=True)
        else:
            st.success("✅ All non-billable entries have explanation notes.")

        # ── Flat-bill entries ──────────────────────────────────────────────
        flat_rows = [r for r in recap_rows if r.get("_flat_bill")]
        if flat_rows:
            st.subheader(f"Flat-Bill Entries ({len(flat_rows)})  –  'Not billable- to be flat bill'")
            flat_hrs = sum(r["Actual"] for r in flat_rows)
            st.caption(f"Total flat-billed hours: **{flat_hrs:.2f}**")
            st.dataframe(pd.DataFrame(flat_rows)[["Tech","_date","MWO","Actual","Notes"]]
                         .rename(columns={"_date":"Date"}),
                         use_container_width=True, hide_index=True)

        # ── Unmatched melds ────────────────────────────────────────────────
        unmatched = [r for r in recap_rows if r["MWO"] and not r["_meld_found"]]
        if unmatched:
            st.error(f"{len(unmatched)} shift entries reference a Meld ID not found in the CSV:")
            st.dataframe(pd.DataFrame(unmatched)[["Tech","_date","MWO","Actual","_raw_notes"]]
                         .rename(columns={"_date":"Date","_raw_notes":"Notes"}),
                         use_container_width=True, hide_index=True)
        else:
            st.success("✅ All Meld IDs matched to the CSV.")

        # ── Non-billable summary ───────────────────────────────────────────
        nb_rows = [r for r in recap_rows if r["_non_bill"]]
        if nb_rows:
            st.subheader(f"All Non-Billable Entries ({len(nb_rows)})")
            st.dataframe(pd.DataFrame(nb_rows)[["Tech","_date","MWO","Actual","Notes"]]
                         .rename(columns={"_date":"Date"}),
                         use_container_width=True, hide_index=True)

        # ── OT employees ───────────────────────────────────────────────────
        ot_list = [r for r in payroll_rows if r["_ot_flag"]]
        if ot_list:
            st.subheader("Employees with Overtime")
            st.dataframe(pd.DataFrame(ot_list)[["Employee","Regular","Overtime","PTO","Total Hours"]],
                         use_container_width=True, hide_index=True)

    # ── Download ──────────────────────────────────────────────────────────
    st.divider()
    with st.spinner("Building Excel workbook…"):
        excel_bytes = build_excel(recap_rows, payroll_rows, employees)

    week = employees[0]["period_start"].replace("/", "-") if employees else ""
    st.download_button(
        label="⬇ Download Payroll Report (.xlsx)",
        data=excel_bytes,
        file_name=f"payroll_report_{week}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        type="primary",
    )

    st.caption(
        "**Post-download:** Review Column G (Paying) for any unreasonable durations. "
        "Yellow rows = Paying > Billable → manager review required. "
        "Red rows = Meld ID not in CSV → verify. "
        "Add notes in Column J for all adjustments per SOP §5.5–5.6."
    )


if __name__ == "__main__":
    main()
