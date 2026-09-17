
import streamlit as st
import pandas as pd
import re
import pdfplumber
import fitz
from datetime import datetime, timedelta, time
from collections import defaultdict, Counter
from io import BytesIO
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="MACL vs RAMIS Reconciliation", layout="wide")

ARR_START = 4*60 + 45
ARR_END   = 15*60 + 45
DEP_START = 9*60
DEP_END   = 23*60 + 59

DAY_SHEETS = {
    "MON":"MONDAY","TUE":"TUESDAY","WED":"WEDNESDAY","THU":"THURSDAY",
    "FRI":"FRIDAY","SAT":"SATURDAY","SUN":"SUNDAY"
}

# Known / special pairing rules can be extended here.
FORCED_DEP = {
    "6E1127":"6E1128",
    "6E1129":"6E1130",
    "6E1131":"6E1134",
    "6E1133":"6E1132",
    "UL101":"UL102",
    "UL103":"UL104",
    "QR670":"QR671",
    "QR672":"QR673",
    "QR674":"QR675",
    "QR676":"QR677",
    "EK652":"EK652D",
    "EK656":"EK657",
    "EK658":"EK659",
    "EK660":"EK661",
    "FZ1207":"FZ1208",
    "FZ1353":"FZ1354",
    "FZ1570":"FZ1570D",
    "FZ1025":"FZ1025D",
}

def excel_date(v):
    # Safely ignore blank Excel cells, pandas NaN and NaT values.
    if v is None:
        return None
    try:
        if pd.isna(v):
            return None
    except Exception:
        pass
    if str(v).strip() == "":
        return None
    if isinstance(v, datetime):
        return v.replace(hour=0, minute=0, second=0, microsecond=0)
    if isinstance(v, (int, float)):
        try:
            return datetime(1899, 12, 30) + timedelta(days=float(v))
        except (ValueError, TypeError, OverflowError):
            return None
    s = str(v).strip()
    for fmt in ("%d.%m.%y", "%d.%m.%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(s, fmt)
        except Exception:
            pass
    return None

def parse_eff(v):
    if v is None:
        return None, None
    s = str(v).strip().replace("–", "-")
    if not s or s == "-":
        return None, None
    parts = [p.strip() for p in s.split("-") if p.strip()]
    if len(parts) == 1:
        d = excel_date(parts[0])
        return d, d
    return excel_date(parts[0]), excel_date(parts[-1])

def fmt_date(d):
    return d.strftime("%d.%m.%y") if d else ""

def fmt_days_ops(v):
    """Always return DAYS OF OPS as exactly 7 text characters."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    s = str(v).strip()
    if s.endswith(".0"):
        s = s[:-2]
    digits = re.sub(r"[^0-9]", "", s)
    return digits.zfill(7)[-7:] if digits else ""

def fmt_eff(s, e):
    if not s:
        return ""
    if e and e.date() != s.date():
        return f"{fmt_date(s)} - {fmt_date(e)}"
    return fmt_date(s)

def macl_time(v):
    if v in (None, "", "-"):
        return None
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, (int, float)):
        n = int(v)
        h, m = divmod(n, 100)
        if 0 <= h <= 23 and 0 <= m <= 59:
            return f"{h:02d}:{m:02d}"
    s = str(v).strip()
    if re.fullmatch(r"\d{1,4}", s):
        n = int(s)
        h, m = divmod(n, 100)
        if 0 <= h <= 23 and 0 <= m <= 59:
            return f"{h:02d}:{m:02d}"
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(s, fmt).strftime("%H:%M")
        except Exception:
            pass
    return None

def ramis_time(v):
    if v in (None, ""):
        return None
    if isinstance(v, (int, float)):
        sec = round(float(v) * 86400)
        return f"{(sec//3600)%24:02d}:{(sec%3600)//60:02d}"
    return macl_time(v)

def in_window(t, start_min, end_min):
    if not t:
        return False
    mins = int(t[:2]) * 60 + int(t[3:5])
    return start_min <= mins <= end_min

def expand_pair(flt):
    """
    Explicit MACL pair always wins.
    Forced mappings are used only when MACL does not explicitly provide a departure.
    """
    s = str(flt).strip().upper()

    if "-" not in s:
        return s, FORCED_DEP.get(s)

    left, right = s.split("-", 1)
    left = left.strip()
    right = right.strip()

    if right == "D":
        return left, left + "D"

    if any(ch.isalpha() for ch in right):
        return left, right

    m = re.match(r"^([A-Z0-9]*?[A-Z])?(\d+)([A-Z]*)$", left)
    if m and right.isdigit():
        prefix = m.group(1) or ""
        base = m.group(2)
        dep_digits = right if len(right) >= len(base) else base[:-len(right)] + right
        return left, prefix + dep_digits

    airline = re.match(r"^([A-Z0-9]{2})", left)
    if airline and right.isdigit():
        return left, airline.group(1) + right

    return left, FORCED_DEP.get(left)

def overlaps(rec, ms, me):
    return rec["start"] <= me and rec["end"] >= ms

def latest_applicable(records, ms, me):
    candidates = [r for r in records if overlaps(r, ms, me)]
    if not candidates:
        return None
    return max(candidates, key=lambda r: (r["start"], r["end"]))

WEEKDAY_NO = {
    "MONDAY":0, "TUESDAY":1, "WEDNESDAY":2, "THURSDAY":3,
    "FRIDAY":4, "SATURDAY":5, "SUNDAY":6
}

def operating_dates(day, ms, me):
    target = WEEKDAY_NO.get(str(day).upper())
    if target is None or ms is None or me is None:
        return []
    d = ms
    while d <= me and d.weekday() != target:
        d += timedelta(days=1)
    dates = []
    while d <= me:
        dates.append(d)
        d += timedelta(days=7)
    return dates

def record_for_date(records, op_date):
    """
    When multiple RAMIS records cover the same operating date,
    use the one with the latest Start Date.
    """
    candidates = [r for r in records if r["start"] <= op_date <= r["end"]]
    if not candidates:
        return None
    return max(candidates, key=lambda r: (r["start"], r["end"]))

def consolidated_active_dates_for_row(macl_row, all_macl_rows):
    """
    FINAL SOURCE-LEVEL MACL CONSOLIDATION — ALL AIRLINES

    Resolve overlapping day-wise MACL rows BEFORE deciding that RAMIS should be
    checked/removed.

    For the same airline + flight + weekday + STA/STD:
    - if another non-cancelled/non-cargo MACL row overlaps the questioned date,
      that date is still authorized by MACL;
    - do NOT depend on the other row's own ops_dates result here, because that
      creates a circular false-negative;
    - explicit red/cancelled MACL rows are never used to keep a movement active.

    This fixes cases such as CONDOR DE2320-1:
      22.09.26 standalone row
      31.03.26-20.10.26 continuing row
    The continuing row covers 22.09.26, therefore RAMIS must not be flagged for
    removal merely because the standalone row was not active in Days of OPS.
    """
    airline = str(macl_row.get("airline", "")).strip().upper()
    flt = str(macl_row.get("flt", "")).strip().upper()
    day = str(macl_row.get("day", "")).strip().upper()
    sta = macl_row.get("sta") or ""
    std = macl_row.get("std") or ""

    wanted_dates = operating_dates(day, macl_row["start"], macl_row["end"])
    if not wanted_dates:
        return []

    active = set()

    for other in all_macl_rows:
        if other is macl_row:
            continue
        if str(other.get("airline", "")).strip().upper() != airline:
            continue
        if str(other.get("flt", "")).strip().upper() != flt:
            continue
        if str(other.get("day", "")).strip().upper() != day:
            continue
        if (other.get("sta") or "") != sta or (other.get("std") or "") != std:
            continue
        if other.get("cancelled") or other.get("cargo"):
            continue

        # Source-level overlap: another valid day-wise MACL row itself is enough
        # to establish continuing MACL authorization for the overlapping date.
        other_dates = operating_dates(day, other["start"], other["end"])

        for d in wanted_dates:
            if d in other_dates:
                active.add(d)

    return sorted(active)



def missing_ramis_operating_dates(records, day, ms, me, op_dates=None):
    """
    Return required MACL operating dates for which RAMIS has NO applicable record.
    This is coverage-only; timing differences are handled separately.
    """
    ops = list(op_dates) if op_dates is not None else operating_dates(day, ms, me)
    return [d for d in ops if record_for_date(records, d) is None]


def fmt_date_span(dates):
    """Compact display for a set of missing operating dates."""
    if not dates:
        return ""
    dates = sorted(dates)
    if len(dates) == 1:
        return dates[0].strftime("%d.%m.%y")
    return f"{dates[0].strftime('%d.%m.%y')} - {dates[-1].strftime('%d.%m.%y')}"


def actionable_missing_dates(dates, as_of=None):
    """
    Missing RAMIS operating dates create actions only for today/future dates.
    Historical gaps remain non-actionable.
    """
    if as_of is None:
        as_of = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    return [d for d in dates if d >= as_of]


def latest_timing_regime(records, day, ms, me, op_dates=None):
    """
    Determine the latest/current RAMIS timing regime within the MACL period.

    Rules:
    - Work only on actual scheduled weekdays.
    - Start from the latest covered operating date and work backwards.
    - Split/adjacent records with the SAME timing are one regime.
    - Stop when an older different timing is reached.
    - Older timing before the latest regime is historical and ignored.
    - A later/earlier RAMIS boundary by itself is not a change.
    """
    ops = list(op_dates) if op_dates is not None else operating_dates(day, ms, me)
    if not ops:
        return None

    timeline = [(d, record_for_date(records, d)) for d in ops]

    latest_idx = None
    for i in range(len(timeline) - 1, -1, -1):
        if timeline[i][1] is not None:
            latest_idx = i
            break
    if latest_idx is None:
        return None

    latest_rec = timeline[latest_idx][1]
    regime_time = latest_rec["time"]

    used = []
    seen = set()
    first_idx = latest_idx

    i = latest_idx
    while i >= 0:
        d, rec = timeline[i]
        if rec is None:
            break
        if rec["time"] != regime_time:
            break

        key = (rec["fid"], rec["type"], rec["start"], rec["end"], rec["time"])
        if key not in seen:
            seen.add(key)
            used.append(rec)

        first_idx = i
        i -= 1

    display = max(used, key=lambda r: (r["start"], r["end"])) if used else latest_rec

    return {
        "time": regime_time,
        "display": display,
        "used": used,
        "first_operating_date": timeline[first_idx][0],
        "last_operating_date": timeline[latest_idx][0],
    }

def numeric_distance(a, b):
    ma = re.search(r"(\d+)", str(a))
    mb = re.search(r"(\d+)", str(b))
    if not ma or not mb:
        return 999999
    return abs(int(ma.group(1)) - int(mb.group(1)))

def infer_departure_id(airline, day, arr_id, std, ms, me, by_airline_day_type):
    candidates = [
        r for r in by_airline_day_type.get((airline, day, "DEPARTURE"), [])
        if overlaps(r, ms, me)
    ]
    ids = {r["fid"] for r in candidates}

    if arr_id in FORCED_DEP and FORCED_DEP[arr_id] in ids:
        return FORCED_DEP[arr_id]

    if arr_id in ids:
        return arr_id
    if arr_id + "D" in ids:
        return arr_id + "D"

    m = re.match(r"^(.+?)(\d+)$", arr_id)
    if m:
        nxt = f"{m.group(1)}{int(m.group(2))+1:0{len(m.group(2))}d}"
        if nxt in ids:
            return nxt

    exact = [r for r in candidates if std and r["time"] == std]
    if exact:
        exact.sort(key=lambda r: (numeric_distance(arr_id, r["fid"]), -r["start"].toordinal()))
        return exact[0]["fid"]

    return None

def parse_ramis(file_bytes):
    df = pd.read_excel(BytesIO(file_bytes), dtype=object)
    df.columns = [str(c).strip() for c in df.columns]

    required = ["Flight ID","Type","Start Date","End Date","AirLine ID","Scheduled Day","Scheduled Time"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("RAMIS file missing required columns: " + ", ".join(missing))

    records = []
    for _, row in df.iterrows():
        fid = row.get("Flight ID")
        typ = row.get("Type")
        day = row.get("Scheduled Day")
        if pd.isna(fid) or pd.isna(typ) or pd.isna(day):
            continue

        st = excel_date(row.get("Start Date"))
        en = excel_date(row.get("End Date"))
        # Ignore incomplete RAMIS records rather than stopping the whole reconciliation.
        if st is None or en is None:
            continue

        type_raw = str(typ).strip().upper()
        if type_raw in ("DEPART", "DEP", "DEPARTURE"):
            type_norm = "DEPARTURE"
        elif type_raw in ("ARR", "ARRIVAL"):
            type_norm = "ARRIVAL"
        else:
            type_norm = type_raw

        records.append({
            "fid": str(fid).strip().upper(),
            "type": type_norm,
            "start": st,
            "end": en,
            "airline": str(row.get("AirLine ID")).strip().upper() if not pd.isna(row.get("AirLine ID")) else "",
            "day": str(day).strip().upper(),
            "time": ramis_time(row.get("Scheduled Time")),
        })

    lookup = defaultdict(list)
    by_airline_day_type = defaultdict(list)
    for r in records:
        lookup[(r["day"], r["type"], r["fid"])].append(r)
        by_airline_day_type[(r["airline"], r["day"], r["type"])].append(r)

    for k in lookup:
        lookup[k].sort(key=lambda r: (r["start"], r["end"]))

    return lookup, by_airline_day_type

def is_red_fill(cell):
    fill = cell.fill
    if not fill or fill.fill_type != "solid":
        return False

    color = fill.fgColor
    if color is None:
        return False

    if color.type == "rgb" and color.rgb:
        rgb = str(color.rgb).upper()[-6:]
        try:
            r = int(rgb[0:2],16)
            g = int(rgb[2:4],16)
            b = int(rgb[4:6],16)
            return r >= 200 and g <= 90 and b <= 90
        except Exception:
            return False

    if color.type == "indexed" and color.indexed in (10,):
        return True
    return False

def macl_row_cancelled(ws, row_num):
    red_count = 0
    for col in range(2, 11):
        if is_red_fill(ws.cell(row=row_num, column=col)):
            red_count += 1
    return red_count >= 2

def days_ops_time_value(v):
    """Convert Excel/PDF time values to the app's HH:MM text format."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ""
    if isinstance(v, datetime):
        return v.strftime("%H:%M")
    if isinstance(v, time):
        return v.strftime("%H:%M")
    if isinstance(v, (int, float)):
        # Excel fractional-day time.
        if 0 <= float(v) < 1:
            total_minutes = int(round(float(v) * 24 * 60)) % (24 * 60)
            return f"{total_minutes // 60:02d}:{total_minutes % 60:02d}"
        # Numeric HHMM such as 1625.
        s = str(int(v)).zfill(4)
        if len(s) == 4 and s.isdigit():
            hh, mm = int(s[:2]), int(s[2:])
            if 0 <= hh <= 23 and 0 <= mm <= 59:
                return f"{hh:02d}:{mm:02d}"
    s = str(v).strip()
    if not s or s.upper() in ("NAN", "NONE", "-"):
        return ""
    if re.fullmatch(r"\d{1,2}:\d{2}", s):
        hh, mm = s.split(":")
        return f"{int(hh):02d}:{int(mm):02d}"
    if re.fullmatch(r"\d{3,4}", s):
        s = s.zfill(4)
        hh, mm = int(s[:2]), int(s[2:])
        if 0 <= hh <= 23 and 0 <= mm <= 59:
            return f"{hh:02d}:{mm:02d}"
    return s


def parse_days_ops_calendar(wb):
    """Read DAYS OF OPS as the operational-calendar validation layer."""
    target_name = None
    for name in wb.sheetnames:
        norm = re.sub(r"\s+", " ", str(name).strip().upper())
        if norm == "DAYS OF OPS":
            target_name = name
            break
    if target_name is None:
        return []

    ws = wb[target_name]
    start_col = 2
    for c in range(1, min(ws.max_column, 12) + 1):
        if str(ws.cell(row=1, column=c).value).strip().upper() == "AIRLINE":
            start_col = c
            break

    records = []
    for row_num in range(1, ws.max_row + 1):
        airline = ws.cell(row=row_num, column=start_col).value
        days_v = ws.cell(row=row_num, column=start_col + 1).value
        flt = ws.cell(row=row_num, column=start_col + 4).value
        sta_v = ws.cell(row=row_num, column=start_col + 5).value
        std_v = ws.cell(row=row_num, column=start_col + 6).value
        eff_v = ws.cell(row=row_num, column=start_col + 7).value
        days = fmt_days_ops(days_v)
        if not airline or not flt or len(days) != 7:
            continue
        ms, me = parse_eff(eff_v)
        if not ms or not me:
            continue
        red_count = 0
        for c in range(start_col, min(start_col + 9, ws.max_column + 1)):
            if is_red_fill(ws.cell(row=row_num, column=c)):
                red_count += 1
        records.append({
            "airline": str(airline).strip().upper(),
            "flt": str(flt).strip().upper(),
            "days_ops": days,
            "sta": days_ops_time_value(sta_v) or "",
            "std": days_ops_time_value(std_v) or "",
            "start": ms,
            "end": me,
            "cancelled": red_count >= 2,
            "source_row": row_num,
        })
    return records


def days_ops_includes_day(days_ops, day):
    pos = WEEKDAY_NO.get(str(day).upper())
    if pos is None:
        return False
    s = fmt_days_ops(days_ops)
    return len(s) == 7 and s[pos] != "0"



def days_ops_source_issues(r):
    """Return definite source-quality issues for one DAYS OF OPS record."""
    issues = []
    start = r.get("start")
    end = r.get("end")
    if start and end and start > end:
        issues.append("INVALID EFFECTIVE RANGE: START DATE AFTER END DATE")

    # For a one-day record, the actual date must agree with the 7-digit OPS code.
    if start and end and start == end:
        enabled = [i for i, ch in enumerate(fmt_days_ops(r.get("days_ops"))) if ch != "0"]
        if enabled and start.weekday() not in enabled:
            issues.append("DATE DOES NOT MATCH DAYS OF OPS")

    return issues


def relevant_days_ops_source_issues(macl_row, days_ops_calendar):
    """
    Scope source warnings to the current MACL period only.

    Multi-day effective start/end boundaries do not themselves have to fall on
    the operating weekday. The 7-digit DAYS OF OPS code defines the actual
    operating dates inside the effective window.
    """
    airline = str(macl_row.get("airline", "")).strip().upper()
    flt = str(macl_row.get("flt", "")).strip().upper()
    day = str(macl_row.get("day", "")).strip().upper()
    sta = macl_row.get("sta") or ""
    std = macl_row.get("std") or ""
    ms = macl_row.get("start")
    me = macl_row.get("end")

    issues = []

    for r in days_ops_calendar:
        if str(r.get("airline", "")).strip().upper() != airline:
            continue
        if str(r.get("flt", "")).strip().upper() != flt:
            continue
        if sta and (r.get("sta") or "") != sta:
            continue
        if std and (r.get("std") or "") != std:
            continue
        if not days_ops_includes_day(r.get("days_ops"), day):
            continue

        rec_issues = days_ops_source_issues(r)
        if not rec_issues:
            continue

        rs, re_ = r.get("start"), r.get("end")

        # Normal source ranges must overlap the current MACL period.
        if rs and re_ and rs <= re_:
            if ms and me and not (rs <= me and re_ >= ms):
                continue

        # Reversed source ranges are relevant only when this MACL row shares
        # one of their malformed boundaries. This prevents an unrelated bad
        # historical record from contaminating a valid current schedule.
        if rs and re_ and rs > re_:
            if not (ms and me and (ms in (rs, re_) or me in (rs, re_))):
                continue

        issues.extend(rec_issues)

    return list(dict.fromkeys(issues))


def validated_operating_dates(macl_row, days_ops_calendar):
    """
    FINAL DAYS-OF-OPS STATUS RULE — ALL AIRLINES

    1. RED is the only cancellation colour.
    2. For the same airline + flight + weekday + relevant timing regime, overlapping
       Days-of-OPS records are resolved per actual operating date by:
         a) latest start date;
         b) if tied, shortest / most-specific effective range;
         c) if still tied, later source row.
    3. The winning record controls the date:
         - winning RED row => cancelled;
         - winning non-RED row => active.
    4. Only the timing side(s) inside the TMA reference window are used to identify
       the relevant regime. An out-of-window arrival/departure must not distort the
       Days-of-OPS match.

    This allows a later white active exception to override a broad red cancellation,
    and also allows a later red cancellation to override a broad white active period.
    """
    flt = str(macl_row["flt"]).strip().upper()
    airline = str(macl_row["airline"]).strip().upper()
    day = str(macl_row["day"]).strip().upper()
    macl_sta = macl_row.get("sta") or ""
    macl_std = macl_row.get("std") or ""

    sta_required = in_window(macl_sta, ARR_START, ARR_END)
    std_required = in_window(macl_std, DEP_START, DEP_END)

    flight_matches = [r for r in days_ops_calendar if r["flt"] == flt]
    if not flight_matches:
        return None

    airline_matches = [r for r in flight_matches if r["airline"] == airline]
    records = airline_matches if airline_matches else flight_matches

    def same_regime(r):
        # Compare only the movement side(s) that TMA actually requires in RAMIS.
        if sta_required and (r.get("sta") or "") != macl_sta:
            return False
        if std_required and (r.get("std") or "") != macl_std:
            return False

        # If neither side is inside the TMA window, fall back to exact timetable
        # matching so unrelated regimes are not mixed.
        if not sta_required and not std_required:
            if macl_sta and (r.get("sta") or "") != macl_sta:
                return False
            if macl_std and (r.get("std") or "") != macl_std:
                return False

        return True

    wanted_dates = operating_dates(day, macl_row["start"], macl_row["end"])
    active_dates = []

    for d in wanted_dates:
        candidates = [
            r for r in records
            if r["start"] <= d <= r["end"]
            and days_ops_includes_day(r["days_ops"], day)
            and same_regime(r)
            and not days_ops_source_issues(r)
        ]

        # No valid cancellation/override record for this date:
        # keep the day-wise MACL movement active.
        if not candidates:
            active_dates.append(d)
            continue

        def precedence(r):
            span_days = (r["end"] - r["start"]).days
            return (
                r["start"],
                -span_days,
                r.get("source_row", 0),
            )

        winning = max(candidates, key=precedence)

        # Only RED means cancelled. Any other winning row is active.
        if not winning.get("cancelled", False):
            active_dates.append(d)

    return active_dates



PDF_WEEKDAY_SHEETS = {
    "MONDAY": "MON", "TUESDAY": "TUE", "WEDNESDAY": "WED",
    "THURSDAY": "THU", "FRIDAY": "FRI", "SATURDAY": "SAT", "SUNDAY": "SUN",
}
PDF_MACL_HEADERS = ["AIRLINE", "DAYS OF OPS", "A/C TYPE", "ROUTE", "FLT NO", "STA", "STD", "EFFECTIVE", "SEATS"]


def _pdf_clean_row(row):
    vals = [("" if v is None else str(v).strip()) for v in row]
    while len(vals) > 9 and vals and vals[0] == "":
        vals = vals[1:]
    return vals


def _pdf_is_schedule_row(vals):
    return len(vals) == 9 and bool(re.fullmatch(r"\d{7}", vals[1] or ""))


def _pdf_row_fill_color(fitz_page, bbox):
    """Return an Excel RGB fill that approximates the MACL PDF source row."""
    y0, y1 = bbox[1], bbox[3]
    row_h = max(1.0, y1 - y0)
    best_fill = None
    best_overlap = 0.0

    for drawing in fitz_page.get_drawings():
        fill = drawing.get("fill")
        rect = drawing.get("rect")
        if not fill or rect is None:
            continue

        fr, fg, fb = [float(x) for x in fill]

        # IMPORTANT: MACL PDFs contain large black structural/table drawings.
        # They can overlap every row and were incorrectly winning over the real
        # RED/GREEN/BLUE/YELLOW fill. Ignore near-black structural fills entirely.
        if fr <= 0.15 and fg <= 0.15 and fb <= 0.15:
            continue

        # Source row fills span most of the table width; ignore small cells/marks.
        if (rect.x1 - rect.x0) < 380:
            continue

        overlap = max(0.0, min(y1, rect.y1) - max(y0, rect.y0))
        if overlap >= row_h * 0.35 and overlap > best_overlap:
            best_fill = fill
            best_overlap = overlap

    if not best_fill:
        return None

    r, g, b = [float(x) for x in best_fill]

    # Exact MACL source palette. Near-white MUST remain unfilled.
    if r >= 0.94 and g >= 0.94 and b >= 0.94:
        return None

    def dist(rgb):
        return ((r-rgb[0])**2 + (g-rgb[1])**2 + (b-rgb[2])**2) ** 0.5

    palette = [
        ("FF0000", (1.000, 0.000, 0.000)),  # Cancelled
        ("C6E0B4", (0.776, 0.878, 0.706)),  # Ops Completed
        ("DDEBF7", (0.867, 0.922, 0.969)),  # Changes
        ("FFFF00", (1.000, 1.000, 0.000)),  # New Application
    ]
    best_name, best_d = min(((name, dist(rgb)) for name, rgb in palette), key=lambda x: x[1])
    return best_name if best_d <= 0.18 else None


def _extract_pdf_schedule_rows(pdf_page, fitz_page):
    rows = []
    seen = set()
    for table in pdf_page.find_tables():
        extracted = table.extract()
        if not extracted:
            continue
        for row_obj, raw in zip(table.rows, extracted):
            vals = _pdf_clean_row(raw)
            if not _pdf_is_schedule_row(vals):
                continue
            key = (tuple(vals), round(float(row_obj.bbox[1]), 1))
            if key in seen:
                continue
            seen.add(key)
            rows.append((vals, _pdf_row_fill_color(fitz_page, row_obj.bbox)))
    return rows



def _apply_pdf_row_style(ws, row_num, fill, start_col=2, end_col=10):
    """Apply the MACL PDF source highlight across the complete converted row."""
    if not fill:
        return
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row_num, col)
        cell.fill = PatternFill("solid", fgColor=fill)
        cell.font = Font(color="000000", bold=(fill == "FF0000"))


def _style_pdf_sheet(ws, header_row, start_col=2):
    header_fill = PatternFill("solid", fgColor="253B6E")
    for idx, header in enumerate(PDF_MACL_HEADERS, start_col):
        c = ws.cell(header_row, idx, header)
        c.fill = header_fill
        c.font = Font(color="FFFFFF", bold=True)
        c.alignment = Alignment(horizontal="center", vertical="center")
    widths = {2:25,3:13,4:12,5:24,6:14,7:9,8:9,9:22,10:10}
    for col_idx, width in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def convert_macl_pdf_to_xlsx(pdf_bytes):
    """
    Convert the standard MACL International Summer/Winter PDF into the exact
    workbook layout consumed by this reconciliation app.

    Output sheets:
      MON, TUE, WED, THU, FRI, SAT, SUN, DAYS OF OPS

    Source row colours are preserved where detectable; in particular RED rows
    remain RED so the locked cancellation logic can read them from Excel.
    """
    if not pdf_bytes:
        raise ValueError("The MACL PDF is empty.")

    pdf_stream = BytesIO(pdf_bytes)
    fitz_doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    if len(fitz_doc) < 9:
        raise ValueError("MACL PDF layout not recognised: expected weekday and Days of OPS pages.")

    wb = Workbook()
    wb.remove(wb.active)
    summary = {"weekday_rows": 0, "days_ops_rows": 0, "red_rows": 0, "pages": len(fitz_doc)}
    weekday_rows = {short: [] for short in PDF_WEEKDAY_SHEETS.values()}
    days_ops_rows = []

    with pdfplumber.open(pdf_stream) as pdf:
        # Locate weekday pages by their printed titles rather than fixed page number.
        weekday_page_indexes = {}
        for page_index, page in enumerate(pdf.pages):
            txt = (page.extract_text() or "").upper()
            for full_day, short_day in PDF_WEEKDAY_SHEETS.items():
                if re.search(rf"\b{full_day}\s*\([1-7]\)", txt):
                    weekday_page_indexes[short_day] = page_index

        missing = [d for d in PDF_WEEKDAY_SHEETS.values() if d not in weekday_page_indexes]
        if missing:
            raise ValueError("MACL PDF conversion stopped: weekday pages not found for " + ", ".join(missing))

        # IMPORTANT: each weekday can span more than one PDF page.  The printed
        # weekday title appears only on the first page; continuation pages still
        # belong to that weekday until the next weekday title begins.  Earlier
        # versions read only the title page, which silently dropped the lower
        # half/continuation pages.  Preserve EVERY schedule row in the section.
        # V25: assign EVERY physical PDF page to a weekday section.  Continuation
        # pages have no MONDAY/TUESDAY/etc title, so they inherit the most recent
        # weekday heading until the next heading (or DAYS OF OPS) appears.
        # This is deliberately page-driven rather than assuming one page per day.
        page_owner = {}
        current_day = None
        for page_index, page in enumerate(pdf.pages):
            txt = (page.extract_text() or "").upper()
            if "DAYS OF OPS" in txt and ("SUMMER" in txt or "WINTER" in txt):
                current_day = None
                break
            found_day = None
            for full_day, short_day in PDF_WEEKDAY_SHEETS.items():
                if re.search(rf"\b{full_day}\s*\([1-7]\)", txt):
                    found_day = short_day
                    break
            if found_day:
                current_day = found_day
            if current_day:
                page_owner[page_index] = current_day

        # Extract each owned page once, including title pages AND continuation pages.
        page_counts = {short: [] for short in PDF_WEEKDAY_SHEETS.values()}
        for page_index in sorted(page_owner):
            short_day = page_owner[page_index]
            extracted = _extract_pdf_schedule_rows(pdf.pages[page_index], fitz_doc[page_index])
            weekday_rows[short_day].extend(extracted)
            page_counts[short_day].append((page_index + 1, len(extracted)))
            summary["weekday_rows"] += len(extracted)
            summary["red_rows"] += sum(1 for _, fill in extracted if fill == "FF0000")

        # V26 continuation-page safety:
        # A physical page without its own weekday heading inherits the most
        # recently detected weekday. It must never be silently skipped.
        for page_index, short_day in sorted(page_owner.items()):
            txt = (pdf.pages[page_index].extract_text() or "").upper()
            has_weekday_heading = any(
                re.search(rf"\b{full_day}\s*\([1-7]\)", txt)
                for full_day in PDF_WEEKDAY_SHEETS
            )
            if not has_weekday_heading:
                inherited = _extract_pdf_schedule_rows(
                    pdf.pages[page_index], fitz_doc[page_index]
                )
                if not inherited:
                    raise ValueError(
                        f"MACL PDF conversion stopped: PDF page {page_index + 1} "
                        f"was inherited as {short_day} continuation, but no schedule "
                        f"rows were extracted. The converter will not create an "
                        f"incomplete Excel file."
                    )

        # V27 INDEPENDENT SOURCE-LINE VALIDATION.
        # Do NOT validate the converter by calling the same table extractor again.
        # Count schedule-looking source lines independently from PDF text: every
        # genuine MACL weekday movement contains a 7-digit DAYS OF OPS mask.
        # This catches a continuation page even if table extraction silently skips it.
        source_line_counts = {}
        for short_day in ["MON","TUE","WED","THU","FRI","SAT","SUN"]:
            owned = [i for i, d in page_owner.items() if d == short_day]
            if not owned:
                raise ValueError(f"MACL PDF conversion stopped: no PDF pages assigned to {short_day}.")
            expected = 0
            page_detail = []
            for page_index in owned:
                txt = pdf.pages[page_index].extract_text() or ""
                source_lines = [
                    ln for ln in txt.splitlines()
                    if re.search(r"\b\d{7}\b", ln)
                    and not ln.strip().upper().startswith("AIRLINE ")
                ]
                expected += len(source_lines)
                page_detail.append((page_index + 1, len(source_lines)))
            actual = len(weekday_rows[short_day])
            source_line_counts[short_day] = expected
            if actual != expected:
                raise ValueError(
                    f"MACL PDF conversion stopped: {short_day} is incomplete. "
                    f"Independent PDF source count = {expected} rows across "
                    f"{page_detail}; converted rows = {actual}. "
                    f"No incomplete workbook will be released."
                )

        summary["source_weekday_counts"] = source_line_counts
        summary["weekday_counts"] = {d: len(weekday_rows[d]) for d in weekday_rows}
        summary["weekday_pages"] = page_counts

        # Days of OPS begins on the page explicitly labelled DAYS OF OPS and ends
        # before the domestic schedule section (recognised by '# AIRLINE').
        days_start = None
        domestic_start = len(pdf.pages)
        for page_index, page in enumerate(pdf.pages):
            txt = (page.extract_text() or "").upper()
            if days_start is None and "DAYS OF OPS" in txt and ("SUMMER" in txt or "WINTER" in txt):
                days_start = page_index
            if days_start is not None and page_index > days_start:
                if re.search(r"#\s+AIRLINE\s+DAYS\s+OF\s+OPS", txt):
                    domestic_start = page_index
                    break

        if days_start is None:
            raise ValueError("MACL PDF conversion stopped: DAYS OF OPS section was not found.")

        for page_index in range(days_start, domestic_start):
            extracted = _extract_pdf_schedule_rows(pdf.pages[page_index], fitz_doc[page_index])
            days_ops_rows.extend(extracted)
            summary["days_ops_rows"] += len(extracted)
            summary["red_rows"] += sum(1 for _, fill in extracted if fill == "FF0000")

    if summary["days_ops_rows"] == 0:
        raise ValueError("MACL PDF conversion stopped: no Days of OPS rows could be extracted.")

    # Build day sheets exactly as parse_macl expects: data begins row 7, B:J.
    day_full_name = {v:k for k,v in PDF_WEEKDAY_SHEETS.items()}
    for short_day in ["MON","TUE","WED","THU","FRI","SAT","SUN"]:
        ws = wb.create_sheet(short_day)
        ws.merge_cells(start_row=1, start_column=2, end_row=1, end_column=10)
        ws.cell(1,2, f"{day_full_name[short_day]} - CONVERTED FROM MACL PDF")
        ws.cell(1,2).font = Font(bold=True, size=14)
        ws.cell(2,2, "Source: MACL PDF")
        ws.cell(3,2, "Conversion preserves PDF schedule values and detectable source row colours.")
        _style_pdf_sheet(ws, 6, 2)
        out_row = 7
        for vals, fill in weekday_rows[short_day]:
            for j, val in enumerate(vals, 2):
                ws.cell(out_row, j, val)
            _apply_pdf_row_style(ws, out_row, fill, 2, 10)
            out_row += 1
        ws.freeze_panes = "B7"
        # V26: filter available on every converted weekday sheet.
        if out_row > 7:
            ws.auto_filter.ref = f"B6:J{out_row - 1}"

    # DAYS OF OPS: preserve the PDF's airline-by-airline visual organisation.
    # Row 1 is kept as a hidden machine-readable header for reconciliation.
    ops = wb.create_sheet("DAYS OF OPS")
    _style_pdf_sheet(ops, 1, 2)
    ops.row_dimensions[1].hidden = True

    airline_groups = []
    group_index = {}
    for vals, fill in days_ops_rows:
        airline = str(vals[0]).strip()
        key = airline.upper()
        if key not in group_index:
            group_index[key] = len(airline_groups)
            airline_groups.append([airline, []])
        airline_groups[group_index[key]][1].append((vals, fill))

    out_row = 3
    navy = PatternFill("solid", fgColor="253B6E")
    thin = Side(style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    for airline, group_rows in airline_groups:
        ops.merge_cells(start_row=out_row, start_column=2, end_row=out_row, end_column=10)
        title = ops.cell(out_row, 2, airline)
        title.font = Font(bold=True, size=12, color="000000")
        title.alignment = Alignment(horizontal="center", vertical="center")
        out_row += 1

        for idx, header in enumerate(PDF_MACL_HEADERS, 2):
            c = ops.cell(out_row, idx, header)
            c.fill = navy
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center", vertical="center")
            c.border = border
        out_row += 1

        for vals, fill in group_rows:
            for j, val in enumerate(vals, 2):
                c = ops.cell(out_row, j, val)
                c.border = border
                c.alignment = Alignment(vertical="center")
            _apply_pdf_row_style(ops, out_row, fill, 2, 10)
            out_row += 1

        out_row += 2

    ops.freeze_panes = "B3"
    # V26: filter also available on DAYS OF OPS.
    if ops.max_row >= 1:
        ops.auto_filter.ref = f"B1:J{ops.max_row}"

    # Conversion audit sheet for staff confidence / troubleshooting.
    audit = wb.create_sheet("PDF CONVERSION CHECK", 0)
    audit["A1"] = "MACL PDF CONVERSION CHECK"
    audit["A1"].font = Font(bold=True, size=14, color="FFFFFF")
    audit["A1"].fill = PatternFill("solid", fgColor="1F4E78")
    audit_rows = [
        ("PDF pages", summary["pages"]),
        ("Weekday rows converted", summary["weekday_rows"]),
        ("MON rows / PDF pages", f"{summary['weekday_counts'].get('MON',0)} / {summary['weekday_pages'].get('MON',[])}"),
        ("TUE rows / PDF pages", f"{summary['weekday_counts'].get('TUE',0)} / {summary['weekday_pages'].get('TUE',[])}"),
        ("WED rows / PDF pages", f"{summary['weekday_counts'].get('WED',0)} / {summary['weekday_pages'].get('WED',[])}"),
        ("THU rows / PDF pages", f"{summary['weekday_counts'].get('THU',0)} / {summary['weekday_pages'].get('THU',[])}"),
        ("FRI rows / PDF pages", f"{summary['weekday_counts'].get('FRI',0)} / {summary['weekday_pages'].get('FRI',[])}"),
        ("SAT rows / PDF pages", f"{summary['weekday_counts'].get('SAT',0)} / {summary['weekday_pages'].get('SAT',[])}"),
        ("SUN rows / PDF pages", f"{summary['weekday_counts'].get('SUN',0)} / {summary['weekday_pages'].get('SUN',[])}"),
        ("Independent PDF weekday source rows", " | ".join(
            f"{d}: {summary.get('source_weekday_counts', {}).get(d, 0)}"
            for d in ["MON","TUE","WED","THU","FRI","SAT","SUN"]
        )),
        ("Converted weekday rows", " | ".join(
            f"{d}: {summary.get('weekday_counts', {}).get(d, 0)}"
            for d in ["MON","TUE","WED","THU","FRI","SAT","SUN"]
        )),
        ("Days of OPS rows converted", summary["days_ops_rows"]),
        ("Red/cancelled source rows detected", summary["red_rows"]),
        ("Colour validation", "WHITE / RED / GREEN / BLUE / YELLOW source fills preserved; black PDF structure ignored"),
        ("Days of OPS layout", "Airline sections reproduced with repeated headers"),
        ("Validation", "PASSED - weekday pages and Days of OPS found"),
    ]
    for r, (k, v) in enumerate(audit_rows, 3):
        audit.cell(r,1,k).font = Font(bold=True)
        audit.cell(r,2,v)
    audit.column_dimensions["A"].width = 34
    audit.column_dimensions["B"].width = 80
    # V26: filter on the conversion audit/check sheet as well.
    audit.auto_filter.ref = f"A2:B{audit.max_row}"

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out.getvalue(), summary


def parse_macl(file_bytes):
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    days_ops_calendar = parse_days_ops_calendar(wb)
    rows = []

    for sheet_name, dayname in DAY_SHEETS.items():
        if sheet_name not in wb.sheetnames:
            continue

        ws = wb[sheet_name]

        for row_num in range(7, ws.max_row + 1):
            airline = ws.cell(row=row_num, column=2).value
            days_ops = ws.cell(row=row_num, column=3).value
            flt = ws.cell(row=row_num, column=6).value
            sta_v = ws.cell(row=row_num, column=7).value
            std_v = ws.cell(row=row_num, column=8).value
            eff_v = ws.cell(row=row_num, column=9).value
            seat_v = ws.cell(row=row_num, column=10).value

            if airline is None or flt is None:
                continue

            sta = macl_time(sta_v)
            std = macl_time(std_v)
            ms, me = parse_eff(eff_v)

            if not ms or not me:
                continue

            sta_req = in_window(sta, ARR_START, ARR_END)
            std_req = in_window(std, DEP_START, DEP_END)
            if not sta_req and not std_req:
                continue

            macl_rec = {
                "airline": str(airline).strip(),
                "day": dayname,
                "days_ops": fmt_days_ops(days_ops),
                "flt": str(flt).strip().upper(),
                "sta": sta or "",
                "std": std or "",
                "start": ms,
                "end": me,
                "cancelled": macl_row_cancelled(ws, row_num),
                "cargo": str(seat_v).strip().upper() == "CARGO" if seat_v is not None else False,
            }
            ops_dates = validated_operating_dates(macl_rec, days_ops_calendar)
            macl_rec["ops_dates"] = ops_dates
            macl_rec["ops_calendar_used"] = ops_dates is not None

            source_issues = []
            if ms > me:
                source_issues.append("INVALID MACL EFFECTIVE RANGE: START DATE AFTER END DATE")
            source_issues.extend(relevant_days_ops_source_issues(macl_rec, days_ops_calendar))
            macl_rec["source_issues"] = list(dict.fromkeys(source_issues))
            rows.append(macl_rec)

    return rows

def read_days_ops_calendar(file_bytes):
    wb = load_workbook(BytesIO(file_bytes), data_only=True)
    return parse_days_ops_calendar(wb)


def _active_macl_covers_date(macl_rows, airline, flt, day, sta, std, d):
    airline_u = str(airline).strip().upper()
    flt_u = str(flt).strip().upper()
    day_u = str(day).strip().upper()

    for m in macl_rows:
        if str(m.get("airline", "")).strip().upper() != airline_u:
            continue
        if str(m.get("flt", "")).strip().upper() != flt_u:
            continue
        if str(m.get("day", "")).strip().upper() != day_u:
            continue
        if m.get("cargo") or m.get("cancelled"):
            continue
        if (m.get("sta") or "") != (sta or "") or (m.get("std") or "") != (std or ""):
            continue
        if not (m.get("start") <= d <= m.get("end")):
            continue
        if m.get("ops_calendar_used", False):
            if d in (m.get("ops_dates") or []):
                return True
        else:
            return True
    return False


def _days_ops_winning_record(records, day, d, sta, std):
    sta_required = in_window(sta, ARR_START, ARR_END)
    std_required = in_window(std, DEP_START, DEP_END)

    def same_regime(r):
        if sta_required and (r.get("sta") or "") != (sta or ""):
            return False
        if std_required and (r.get("std") or "") != (std or ""):
            return False
        if not sta_required and not std_required:
            if sta and (r.get("sta") or "") != sta:
                return False
            if std and (r.get("std") or "") != std:
                return False
        return True

    candidates = [
        r for r in records
        if r["start"] <= d <= r["end"]
        and days_ops_includes_day(r["days_ops"], day)
        and same_regime(r)
        and not days_ops_source_issues(r)
    ]
    if not candidates:
        return None

    def precedence(r):
        span_days = (r["end"] - r["start"]).days
        return (r["start"], -span_days, r.get("source_row", 0))

    return max(candidates, key=precedence)


def append_active_ramis_residual_warnings(out, macl_rows, lookup, days_ops_calendar):
    """
    SECOND RECONCILIATION PASS — ALL AIRLINES.

    Finds CURRENT/FUTURE RAMIS records that still exist where MACL DAYS OF OPS
    has the movement cancelled/no longer active. Fully historical RAMIS records
    ending before today are ignored.
    """
    today = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    existing = {
        (
            str(r.get("AIRLINE","")).strip().upper(),
            str(r.get("DAY","")).strip().upper(),
            str(r.get("FLT NO","")).strip().upper(),
            str(r.get("RAMIS FLT NO","")).strip().upper(),
        )
        for r in out if r.get("STATUS") == "RAMIS CHECK REQUIRED"
    }

    groups = defaultdict(list)
    for r in days_ops_calendar:
        groups[(r["airline"], r["flt"])].append(r)

    weekday_names = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"]

    for (airline, macl_flt), all_records in groups.items():
        if not any(r.get("cancelled") and not days_ops_source_issues(r) for r in all_records):
            continue

        regimes = defaultdict(list)
        for r in all_records:
            regimes[(r.get("sta") or "", r.get("std") or "")].append(r)

        for (sta, std), records in regimes.items():
            arr_id, dep_id = expand_pair(macl_flt)
            sta_req = in_window(sta, ARR_START, ARR_END)
            std_req = in_window(std, DEP_START, DEP_END)

            # Same-MACL-flight-number D segregation rule.
            if "-" not in str(macl_flt):
                if sta_req and std_req:
                    dep_id = arr_id + "D"
                elif std_req and not sta_req:
                    dep_id = arr_id
                else:
                    dep_id = None

            for day in weekday_names:
                candidate_dates = set()
                for r in records:
                    if days_ops_source_issues(r) or r["end"] < today:
                        continue
                    if not days_ops_includes_day(r["days_ops"], day):
                        continue
                    start = max(r["start"], today)
                    candidate_dates.update(operating_dates(day, start, r["end"]))

                if not candidate_dates:
                    continue

                cancelled_dates = []
                for d in sorted(candidate_dates):
                    winning = _days_ops_winning_record(records, day, d, sta, std)
                    if not winning or not winning.get("cancelled", False):
                        continue
                    if _active_macl_covers_date(macl_rows, airline, macl_flt, day, sta, std, d):
                        continue
                    cancelled_dates.append(d)

                if not cancelled_dates:
                    continue

                arr_matches, dep_matches = [], []

                if sta_req and arr_id:
                    for rr in lookup.get((day, "ARRIVAL", arr_id), []):
                        # A cancelled MACL timing may only flag the SAME RAMIS timing.
                        # Example: cancelled 09:45 must not flag active replacement 09:40.
                        if (rr.get("time") or "") != (sta or ""):
                            continue
                        if rr["end"] >= today and any(rr["start"] <= d <= rr["end"] for d in cancelled_dates):
                            arr_matches.append(rr)

                if std_req and dep_id:
                    for rr in lookup.get((day, "DEPARTURE", dep_id), []):
                        # A cancelled MACL timing may only flag the SAME RAMIS timing.
                        if (rr.get("time") or "") != (std or ""):
                            continue
                        if rr["end"] >= today and any(rr["start"] <= d <= rr["end"] for d in cancelled_dates):
                            dep_matches.append(rr)

                # Departure-only fallback to base flight ID.
                if std_req and not sta_req and not dep_matches and arr_id:
                    for rr in lookup.get((day, "DEPARTURE", arr_id), []):
                        if (rr.get("time") or "") != (std or ""):
                            continue
                        if rr["end"] >= today and any(rr["start"] <= d <= rr["end"] for d in cancelled_dates):
                            dep_matches.append(rr)
                    if dep_matches:
                        dep_id = arr_id

                if not arr_matches and not dep_matches:
                    continue

                arr_rr = max(arr_matches, key=lambda r: (r["start"], r["end"])) if arr_matches else None
                dep_rr = max(dep_matches, key=lambda r: (r["start"], r["end"])) if dep_matches else None

                display_dates = [
                    d for d in cancelled_dates
                    if (arr_rr and arr_rr["start"] <= d <= arr_rr["end"])
                    or (dep_rr and dep_rr["start"] <= d <= dep_rr["end"])
                ]
                if not display_dates:
                    continue

                ramis_flt = macl_flt if arr_rr and dep_rr else arr_id if arr_rr else dep_id if dep_rr else ""
                signature = (airline, day, macl_flt, ramis_flt)
                if signature in existing:
                    continue
                existing.add(signature)

                current_records = [r for r in (arr_rr, dep_rr) if r]
                ramis_start = min(r["start"] for r in current_records)
                ramis_end = max(r["end"] for r in current_records)

                out.append({
                    "AIRLINE": airline,
                    "DAY": day,
                    "DAYS OF OPS": fmt_days_ops(next(
                        (r["days_ops"] for r in records if r.get("cancelled") and days_ops_includes_day(r["days_ops"], day)),
                        ""
                    )),
                    "FLT NO": macl_flt,
                    "STA": sta,
                    "STD": std,
                    "EFFECTIVE": fmt_eff(min(display_dates), max(display_dates)),
                    "RAMIS AIRLINE": airline,
                    "RAMIS DAY": day,
                    "RAMIS FLT NO": ramis_flt,
                    "RAMIS STA": arr_rr["time"] if arr_rr else "",
                    "RAMIS STD": dep_rr["time"] if dep_rr else "",
                    "RAMIS EFFECTIVE": fmt_eff(ramis_start, ramis_end),
                    "NEW FLT NO": "",
                    "NEW STA": "",
                    "NEW STD": "",
                    "NEW EFFECTIVE": "",
                    "STATUS": "RAMIS CHECK REQUIRED",
                    "COMMENTS": "CANCELLED / NO ACTIVE MACL MOVEMENT - ACTIVE RAMIS RECORD STILL EXISTS",
                })

    return out



def ramis_effective_from_selected(arr_sel, dep_sel):
    """
    RAMIS EFFECTIVE DISPLAY — SOURCE ONLY.

    Never copy or manufacture a MACL effective date into the RAMIS section.

    - If ARR and DEP selected RAMIS records have the same source range, show it once.
    - If only one side exists, show that actual RAMIS source range.
    - If ARR and DEP source ranges differ, show their common applicable RAMIS period
      as one clean range.
    """
    if arr_sel and dep_sel:
        common_start = max(arr_sel["start"], dep_sel["start"])
        common_end = min(arr_sel["end"], dep_sel["end"])
        if common_start <= common_end:
            return fmt_eff(common_start, common_end)

        # Defensive fallback only: selected records should normally overlap on the
        # operating date. If they do not, keep the display source-based.
        arr_eff = fmt_eff(arr_sel["start"], arr_sel["end"])
        dep_eff = fmt_eff(dep_sel["start"], dep_sel["end"])
        return f"{arr_eff} / {dep_eff}"
    if arr_sel:
        return fmt_eff(arr_sel["start"], arr_sel["end"])
    if dep_sel:
        return fmt_eff(dep_sel["start"], dep_sel["end"])
    return ""



def is_special_date_override(ms, me, arr_sel, dep_sel, sta_changed, std_changed):
    """
    Detect a MACL special/shorter period inside a broader RAMIS regime.

    Example:
      RAMIS 19.07.26-27.09.26 = 11:25 / 13:15
      MACL  13.09.26 only      = 10:10 / 12:10

    This is not merely a generic time change: staff must update only the
    MACL exception period, not overwrite the complete broader RAMIS regime.
    """
    if not (sta_changed or std_changed) or not ms or not me:
        return False

    selected = [r for r in (arr_sel, dep_sel) if r]
    if not selected:
        return False

    # Every selected RAMIS side used for the timing comparison must actually
    # cover the complete MACL exception period.
    if not all(r["start"] <= ms and r["end"] >= me for r in selected):
        return False

    # At least one real RAMIS source range must be broader than the MACL period.
    return any(r["start"] < ms or r["end"] > me for r in selected)


def reconcile(macl_rows, lookup, by_airline_day_type, days_ops_calendar=None):
    out = []

    for m in macl_rows:
        airline_name = m["airline"]
        day = m["day"]
        flt = m["flt"]
        sta = m["sta"]
        std = m["std"]
        ms = m["start"]
        me = m["end"]

        sta_req = in_window(sta, ARR_START, ARR_END)
        std_req = in_window(std, DEP_START, DEP_END)

        arr_id, dep_id = expand_pair(flt)

        # LOCKED SAME-FLIGHT-NUMBER RULE - APPLIES TO ALL AIRLINES.
        # When MACL uses one identical flight number for both ARR and DEP:
        #   * both ARR and DEP inside TMA windows -> RAMIS ARR keeps base ID,
        #     RAMIS DEP uses base ID + "D" only to segregate the two records.
        #   * only ARR is inside TMA window -> use base ID only; no D required.
        #   * only DEP is inside TMA window -> use base ID only; no D required.
        # Explicit MACL pairs such as G9093-4 / FZ1207-8 are NOT changed by this rule.
        macl_text = str(flt).strip().upper()
        same_macl_number = "-" not in macl_text

        # Also support an explicitly repeated pair such as FZ1026-1026.
        if "-" in macl_text:
            _left, _right = macl_text.split("-", 1)
            _left = _left.strip()
            _right = _right.strip()
            if _right == _left:
                same_macl_number = True
            else:
                _tmp_arr, _tmp_dep = expand_pair(macl_text)
                if _tmp_dep == _tmp_arr:
                    same_macl_number = True

        if same_macl_number:
            if sta_req and std_req:
                dep_id = arr_id + "D"
            elif std_req and not sta_req:
                dep_id = arr_id
            else:
                dep_id = None

        airline_match = re.match(r"^([A-Z0-9]{2})", arr_id)
        airline_code = airline_match.group(1) if airline_match else ""

        if std_req and not dep_id:
            dep_id = infer_departure_id(
                airline_code, day, arr_id, std, ms, me, by_airline_day_type
            )

        arr_records = lookup.get((day, "ARRIVAL", arr_id), []) if sta_req else []
        dep_records = lookup.get((day, "DEPARTURE", dep_id), []) if std_req and dep_id else []

        validated_dates = m.get("ops_dates") if m.get("ops_calendar_used", False) else None

        arr_regime = latest_timing_regime(
            arr_records, day, ms, me, validated_dates
        ) if sta_req else None
        dep_regime = latest_timing_regime(
            dep_records, day, ms, me, validated_dates
        ) if std_req else None

        # DEPARTURE-ONLY FALLBACK:
        # When MACL arrival is outside the TMA arrival window, RAMIS may keep the
        # departure under the MACL base flight number rather than the paired DEP ID.
        if std_req and not sta_req and dep_regime is None and arr_id:
            base_dep_records = lookup.get((day, "DEPARTURE", arr_id), [])
            base_dep_regime = latest_timing_regime(
                base_dep_records, day, ms, me, validated_dates
            )
            if base_dep_regime is not None:
                dep_id = arr_id
                dep_records = base_dep_records
                dep_regime = base_dep_regime

        arr_sel = arr_regime["display"] if arr_regime else None
        dep_sel = dep_regime["display"] if dep_regime else None
        selected = [r for r in (arr_sel, dep_sel) if r]

        # REFERENCE-ONLY RAMIS DISPLAY:
        # Even when a MACL side is outside the TMA comparison window, show the
        # actual RAMIS source time if it exists. This must NOT create a change
        # action for that out-of-window side.
        arr_ref_regime = None
        dep_ref_regime = None

        if not sta_req and arr_id:
            arr_ref_records = lookup.get((day, "ARRIVAL", arr_id), [])
            arr_ref_regime = latest_timing_regime(
                arr_ref_records, day, ms, me, validated_dates
            )

        if not std_req and dep_id:
            dep_ref_records = lookup.get((day, "DEPARTURE", dep_id), [])
            dep_ref_regime = latest_timing_regime(
                dep_ref_records, day, ms, me, validated_dates
            )

        display_arr_sel = arr_sel or (arr_ref_regime["display"] if arr_ref_regime else None)
        display_dep_sel = dep_sel or (dep_ref_regime["display"] if dep_ref_regime else None)

        # COVERAGE CHECK:
        # A later RAMIS record must not hide earlier required MACL operating dates
        # where no RAMIS record exists at all.
        arr_missing_dates = missing_ramis_operating_dates(
            arr_records, day, ms, me, validated_dates
        ) if sta_req else []
        dep_missing_dates = missing_ramis_operating_dates(
            dep_records, day, ms, me, validated_dates
        ) if std_req else []

        # BACK-DATE RULE:
        # Historical missing dates do not create Action Required items.
        # Only today/future coverage gaps remain actionable.
        today_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
        arr_missing_dates = actionable_missing_dates(arr_missing_dates, today_start)
        dep_missing_dates = actionable_missing_dates(dep_missing_dates, today_start)

        current_flt = (
            flt if display_arr_sel and display_dep_sel
            else arr_id if display_arr_sel
            else dep_id if display_dep_sel
            else ""
        )
        current_sta = (
            arr_regime["time"] if arr_regime
            else arr_ref_regime["time"] if arr_ref_regime
            else ""
        )
        current_std = (
            dep_regime["time"] if dep_regime
            else dep_ref_regime["time"] if dep_ref_regime
            else ""
        )

        # RAMIS EFFECTIVE reflects the actual source records being displayed.
        current_eff = ramis_effective_from_selected(display_arr_sel, display_dep_sel)

        comments = []

        # Informational only: show that RAMIS has a source time on a side which is
        # outside the TMA comparison window. Do not create STA/STD change from it.
        if arr_ref_regime is not None:
            comments.append(
                f"RAMIS ARRIVAL {arr_ref_regime['time']} SHOWN FOR REFERENCE - "
                f"MACL STA {sta} OUTSIDE TMA ARRIVAL WINDOW"
            )
        if dep_ref_regime is not None:
            comments.append(
                f"RAMIS DEPARTURE {dep_ref_regime['time']} SHOWN FOR REFERENCE - "
                f"MACL STD {std} OUTSIDE TMA DEPARTURE WINDOW"
            )
        new_flt = new_sta = new_std = new_eff = ""

        cutoff = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0) - timedelta(days=1)
        is_expired = me <= cutoff
        is_cancelled = bool(m.get("cancelled", False))
        is_cargo = bool(m.get("cargo", False)) or "CARGO" in airline_name.upper()

        if m.get("source_issues"):
            # Do not silently repair a bad MACL / DAYS OF OPS source row.
            # Show the RAMIS record already selected by the normal lookup, if any,
            # and send the row to staff review instead of generating an automatic action.
            status = "MACL SOURCE CHECK REQUIRED"
            comments = ["MACL DATE / DAYS OF OPS IS INCORRECT - GO BACK AND CHECK DAYS OF OPS"]
            comments.extend(m.get("source_issues", []))
            new_flt = new_sta = new_std = new_eff = ""

        elif is_cargo:
            status = "IGNORED"
            comments = ["CARGO FLIGHT - IGNORED / NO ACTION REQUIRED"]

        elif is_expired:
            status = "EXPIRED"
            comments = ["PAST SCHEDULE - NO ACTION REQUIRED"]

        elif is_cancelled:
            if selected:
                status = "CANCELLED"
                new_flt = "CANCEL / REMOVE"
                comments = ["CANCELLED IN MACL - REMOVE FROM RAMIS"]
            else:
                status = "CANCELLED - NO ACTION"
                comments = ["CANCELLED IN MACL - NO APPLICABLE RAMIS RECORD"]

        elif m.get("ops_calendar_used", False) and not m.get("ops_dates"):
            # First consolidate the complete day-wise MACL schedule. A special row
            # must not cancel RAMIS if another active MACL row with the same
            # airline/flight/day/timing still covers the date.
            consolidated_dates = consolidated_active_dates_for_row(m, macl_rows)

            if consolidated_dates:
                check_arr_regime = latest_timing_regime(
                    arr_records, day, ms, me, consolidated_dates
                ) if sta_req else None
                check_dep_regime = latest_timing_regime(
                    dep_records, day, ms, me, consolidated_dates
                ) if std_req else None

                if check_arr_regime or check_dep_regime:
                    arr_regime = check_arr_regime
                    dep_regime = check_dep_regime
                    arr_sel = arr_regime["display"] if arr_regime else None
                    dep_sel = dep_regime["display"] if dep_regime else None
                    selected = [r for r in (arr_sel, dep_sel) if r]
                    current_flt = (
                        flt if arr_sel and dep_sel
                        else arr_id if arr_sel
                        else dep_id if dep_sel
                        else ""
                    )
                    current_sta = arr_regime["time"] if arr_regime else ""
                    current_std = dep_regime["time"] if dep_regime else ""
                    current_eff = ramis_effective_from_selected(arr_sel, dep_sel)
                    status = "NO CHANGE"
                    comments = ["-"]
                    # This row is already resolved by another active MACL row.
                    out.append({
                        "AIRLINE": airline_name, "DAY": day,
                        "DAYS OF OPS": fmt_days_ops(m["days_ops"]), "FLT NO": flt,
                        "STA": sta, "STD": std, "EFFECTIVE": fmt_eff(ms, me),
                        "RAMIS AIRLINE": airline_name if selected else "",
                        "RAMIS DAY": day if selected else "",
                        "RAMIS FLT NO": current_flt,
                        "RAMIS STA": current_sta, "RAMIS STD": current_std,
                        "RAMIS EFFECTIVE": current_eff,
                        "NEW FLT NO": "", "NEW STA": "", "NEW STD": "",
                        "NEW EFFECTIVE": "", "STATUS": status,
                        "COMMENTS": ", ".join(comments),
                    })
                    continue

            # No other active MACL row covers it. Now check RAMIS independently.
            # A cancelled/non-active MACL movement must never hide a RAMIS record.
            check_arr_regime = latest_timing_regime(arr_records, day, ms, me) if sta_req else None
            check_dep_regime = latest_timing_regime(dep_records, day, ms, me) if std_req else None

            # Preserve the existing departure-only/base-flight fallback here too.
            if std_req and not sta_req and check_dep_regime is None and dep_id != arr_id:
                base_dep_records = lookup.get((airline_name, day, arr_id, "DEPARTURE"), [])
                check_base_dep = latest_timing_regime(base_dep_records, day, ms, me)
                if check_base_dep is not None:
                    dep_id = arr_id
                    check_dep_regime = check_base_dep

            check_arr_sel = check_arr_regime["display"] if check_arr_regime else None
            check_dep_sel = check_dep_regime["display"] if check_dep_regime else None
            check_selected = [r for r in (check_arr_sel, check_dep_sel) if r]

            if check_selected:
                # Show the actual RAMIS record and flag it prominently for review.
                arr_regime = check_arr_regime
                dep_regime = check_dep_regime
                arr_sel = check_arr_sel
                dep_sel = check_dep_sel
                selected = check_selected

                current_flt = (
                    flt if arr_sel and dep_sel
                    else arr_id if arr_sel
                    else dep_id if dep_sel
                    else ""
                )
                current_sta = arr_regime["time"] if arr_regime else ""
                current_std = dep_regime["time"] if dep_regime else ""
                current_eff = ramis_effective_from_selected(arr_sel, dep_sel)

                status = "RAMIS CHECK REQUIRED"
                comments = ["NO ACTIVE MACL MOVEMENT - RAMIS RECORD STILL EXISTS"]
            else:
                # No RAMIS record: leave RAMIS columns blank and no action is required.
                current_flt = current_sta = current_std = current_eff = ""
                status = "CANCELLED - NO ACTION"
                comments = ["NO ACTIVE MACL MOVEMENT / NO RAMIS RECORD"]

        else:
            arr_absent = sta_req and arr_regime is None
            dep_absent = std_req and dep_regime is None

            if arr_absent and dep_absent:
                status = "NEW FLIGHT"
                new_flt = flt
                new_sta = sta
                new_std = std
                new_eff = fmt_eff(ms, me)
                comments = ["NEW FLIGHT"]
            else:
                sta_changed = False
                std_changed = False

                if arr_absent:
                    comments.append("MISSING RAMIS ARRIVAL")
                    new_flt = arr_id
                    new_sta = sta
                else:
                    if arr_missing_dates:
                        comments.append(
                            "MISSING RAMIS ARRIVAL FOR " + fmt_date_span(arr_missing_dates)
                        )
                        new_flt = new_flt or arr_id
                        new_sta = sta
                        new_eff = fmt_eff(ms, me)
                    if sta_req and arr_regime["time"] != sta:
                        comments.append("STA CHANGE")
                        new_sta = sta
                        sta_changed = True

                if dep_absent:
                    comments.append("MISSING RAMIS DEPARTURE")
                    if not new_flt:
                        new_flt = dep_id or flt
                    new_std = std
                else:
                    if dep_missing_dates:
                        comments.append(
                            "MISSING RAMIS DEPARTURE FOR " + fmt_date_span(dep_missing_dates)
                        )
                        if not new_flt:
                            new_flt = dep_id or flt
                        new_std = std
                        new_eff = fmt_eff(ms, me)
                    if std_req and dep_regime["time"] != std:
                        comments.append("STD CHANGE")
                        new_std = std
                        std_changed = True

                # SPECIAL DATE / SHORT-PERIOD OVERRIDE:
                # MACL introduces a different timing inside a broader real RAMIS
                # regime. Keep the actual RAMIS source range in the RAMIS section,
                # but tell staff that the revised timing applies ONLY to the MACL
                # exception period.
                if is_special_date_override(
                    ms, me, arr_sel, dep_sel, sta_changed, std_changed
                ):
                    detail = ", ".join(comments)
                    comments = [
                        "SPECIAL DATE OVERRIDE - MACL HAS A SEPARATE MOVEMENT "
                        f"WITHIN AN EXISTING RAMIS EFFECTIVE RANGE; {detail}; "
                        "VERIFY AND UPDATE RAMIS FOR THIS MACL DATE/PERIOD ONLY"
                    ]
                    new_eff = fmt_eff(ms, me)

                # No EFFECTIVE CHANGE merely because normal source boundaries differ.
                actionable_comments = [
                    c for c in comments
                    if not c.startswith("RAMIS ARRIVAL ")
                    and not c.startswith("RAMIS DEPARTURE ")
                ]
                status = "CHANGE" if actionable_comments else "NO CHANGE"
                if not comments:
                    comments = ["-"]

        out.append({
            "AIRLINE": airline_name,
            "DAY": day,
            "DAYS OF OPS": m["days_ops"],
            "FLT NO": flt,
            "STA": sta,
            "STD": std,
            "EFFECTIVE": fmt_eff(ms, me),
            "RAMIS AIRLINE": airline_name if selected else "",
            "RAMIS DAY": day if selected else "",
            "RAMIS FLT NO": current_flt,
            "RAMIS STA": current_sta,
            "RAMIS STD": current_std,
            "RAMIS EFFECTIVE": current_eff,
            "NEW FLT NO": new_flt,
            "NEW STA": new_sta,
            "NEW STD": new_std,
            "NEW EFFECTIVE": new_eff,
            "STATUS": status,
            "COMMENTS": ", ".join(comments),
        })

    if days_ops_calendar:
        out = append_active_ramis_residual_warnings(
            out, macl_rows, lookup, days_ops_calendar
        )

    return out

def build_excel(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "MACL vs RAMIS"
    act = wb.create_sheet("ACTION REQUIRED")
    summ = wb.create_sheet("SUMMARY")
    logic = wb.create_sheet("LOCKED LOGIC")

    blue = "5B9BD5"
    blue2 = "D9EAF7"
    purple = "D9CCE3"
    green = "C6E0B4"
    pink = "F4CCCC"
    yellow = "FFD966"
    palegreen = "E2F0D9"
    orange = "F8CBAD"
    grey = "E7E6E6"

    headers = [
        "AIRLINE","DAY","DAYS OF OPS","FLT NO","STA","STD","EFFECTIVE","",
        "AIRLINE","DAY","FLT NO","STA","STD","EFFECTIVE",
        "NEW FLT NO","NEW STA","NEW STD","NEW EFFECTIVE","STATUS","COMMENTS"
    ]

    def make_sheet(sh, data):
        sh.merge_cells("A1:G1")
        sh.merge_cells("I1:N1")
        sh.merge_cells("O1:R1")
        sh.merge_cells("S1:T1")
        sh["A1"]="MACL WINTER SCHEDULE"
        sh["I1"]="RAMIS SCHEDULE TIME"
        sh["O1"]="RAMIS REVISED TIME"
        sh["S1"]="CHANGES"

        for cell in ("A1","I1","O1","S1"):
            sh[cell].font = Font(bold=True)
            sh[cell].alignment = Alignment(horizontal="center")

        sh["A1"].fill = PatternFill("solid", fgColor=blue)
        sh["I1"].fill = PatternFill("solid", fgColor=purple)
        sh["O1"].fill = PatternFill("solid", fgColor=green)
        sh["S1"].fill = PatternFill("solid", fgColor=pink)

        for c,h in enumerate(headers,1):
            sh.cell(2,c,h)
            sh.cell(2,c).font = Font(bold=True)
            sh.cell(2,c).alignment = Alignment(horizontal="center")
            if c <= 7:
                sh.cell(2,c).fill = PatternFill("solid", fgColor=blue2)
            elif 9 <= c <= 14:
                sh.cell(2,c).fill = PatternFill("solid", fgColor=purple)
            elif 15 <= c <= 18:
                sh.cell(2,c).fill = PatternFill("solid", fgColor=green)
            elif c >= 19:
                sh.cell(2,c).fill = PatternFill("solid", fgColor=pink)

        for r_idx, rec in enumerate(data, 3):
            vals = [
                rec["AIRLINE"],rec["DAY"],rec["DAYS OF OPS"],rec["FLT NO"],
                rec["STA"],rec["STD"],rec["EFFECTIVE"],"",
                rec["RAMIS AIRLINE"],rec["RAMIS DAY"],rec["RAMIS FLT NO"],
                rec["RAMIS STA"],rec["RAMIS STD"],rec["RAMIS EFFECTIVE"],
                rec["NEW FLT NO"],rec["NEW STA"],rec["NEW STD"],rec["NEW EFFECTIVE"],
                rec["STATUS"],rec["COMMENTS"]
            ]
            for c,v in enumerate(vals,1):
                sh.cell(r_idx,c,v)

            status = rec["STATUS"]
            if status == "NO CHANGE":
                sh.cell(r_idx,19).fill = PatternFill("solid", fgColor=palegreen)
            elif status == "CHANGE":
                sh.cell(r_idx,19).fill = PatternFill("solid", fgColor=yellow)
            elif status == "NEW FLIGHT":
                sh.cell(r_idx,19).fill = PatternFill("solid", fgColor=orange)
            elif status == "IGNORED":
                for c in range(1, 21):
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="E7E6E6")
                    sh.cell(r_idx,c).font = Font(color="7F7F7F")
                sh.cell(r_idx,19).font = Font(color="7F7F7F", bold=True)
            elif status == "EXPIRED":
                for c in range(1, 21):
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="D9D9D9")
                    sh.cell(r_idx,c).font = Font(color="7F7F7F")
            elif status == "MACL SOURCE CHECK REQUIRED":
                for c in range(1, 21):
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="FFD966")
                    sh.cell(r_idx,c).font = Font(color="9C5700", bold=True)
                for c in [19,20]:
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="C65911")
                    sh.cell(r_idx,c).font = Font(color="FFFFFF", bold=True)

            elif status == "RAMIS CHECK REQUIRED":
                # Strong warning: MACL has no active movement but RAMIS still has data.
                for c in range(1, 21):
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="F4B183")
                    sh.cell(r_idx,c).font = Font(color="9C0006", bold=True)
                for c in [9,10,11,12,13,14,19,20]:
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="FF0000")
                    sh.cell(r_idx,c).font = Font(color="FFFFFF", bold=True)

            elif status in ("CANCELLED", "CANCELLED - NO ACTION"):
                for c in range(1, 21):
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="F4CCCC")
                for c in list(range(1, 8)) + [19, 20]:
                    sh.cell(r_idx,c).fill = PatternFill("solid", fgColor="FF0000")
                    sh.cell(r_idx,c).font = Font(color="FFFFFF", bold=True)
                if status == "CANCELLED":
                    sh.cell(r_idx,15).fill = PatternFill("solid", fgColor="C00000")
                    sh.cell(r_idx,15).font = Font(color="FFFFFF", bold=True)

            if rec["NEW STA"]:
                sh.cell(r_idx,16).fill = PatternFill("solid", fgColor=yellow)
            if rec["NEW STD"]:
                sh.cell(r_idx,17).fill = PatternFill("solid", fgColor=yellow)
            if rec["NEW EFFECTIVE"]:
                sh.cell(r_idx,18).fill = PatternFill("solid", fgColor=pink)

        sh.freeze_panes = "A3"
        sh.auto_filter.ref = f"A2:T{max(2, sh.max_row)}"

        widths = {
            "A":23,"B":12,"C":13,"D":16,"E":9,"F":9,"G":23,"H":2,
            "I":23,"J":12,"K":16,"L":10,"M":10,"N":23,
            "O":16,"P":10,"Q":10,"R":23,"S":13,"T":34
        }
        for col,w in widths.items():
            sh.column_dimensions[col].width = w

    make_sheet(ws, rows)
    action_rows = [r for r in rows if r["STATUS"] not in ("NO CHANGE","IGNORED","EXPIRED","CANCELLED - NO ACTION")]
    make_sheet(act, action_rows)

    cnt = Counter(r["STATUS"] for r in rows)
    summ_data = [
        ["MACL vs RAMIS FINAL RECONCILIATION",""],
        ["Reference / Master","MACL Winter Schedule"],
        ["Checked System","RAMIS Connecting Flight Plans"],
        ["Latest RAMIS rule","Use the latest overlapping RAMIS record by Start Date."],
        ["Timing rule","MACL STA/STD are compared against that latest applicable RAMIS record."],
        ["Effective rule","Do not flag a change because RAMIS starts earlier/later or ends earlier than MACL. Flag EFFECTIVE CHANGE only when the latest applicable RAMIS record continues beyond the MACL end date."],
        ["",""],
        ["RESULT","COUNT"],
        ["Total checked",len(rows)],
        ["NO CHANGE",cnt.get("NO CHANGE",0)],
        ["CHANGE",cnt.get("CHANGE",0)],
        ["NEW FLIGHT",cnt.get("NEW FLIGHT",0)],
        ["IGNORED",cnt.get("IGNORED",0)],
        ["EXPIRED / PAST",cnt.get("EXPIRED",0)],
    ]
    for r,row in enumerate(summ_data,1):
        summ.cell(r,1,row[0]); summ.cell(r,2,row[1])
    summ.column_dimensions["A"].width = 35
    summ.column_dimensions["B"].width = 105
    summ["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    summ["A1"].fill = PatternFill("solid", fgColor=blue)

    logic_rows = [
        ["FINAL LOCKED LOGIC – MACL vs RAMIS",""],
        ["1. Master","MACL is the official master/reference."],
        ["2. Latest record","For each matching flight/day/type, select the latest RAMIS record that overlaps the MACL period. Latest = greatest RAMIS Start Date."],
        ["3. Historical records","Do not combine old RAMIS effective ranges into one cell."],
        ["4. STA/STD","Compare MACL timing only with the latest applicable RAMIS record."],
        ["5. Start-date difference","RAMIS may start before or after MACL. This alone is not a change."],
        ["6. Earlier RAMIS end","RAMIS may end before MACL. This alone is not a change when the selected latest record overlaps MACL and the timing is correct."],
        ["7. Effective conflict","Flag EFFECTIVE CHANGE only when the selected latest RAMIS record continues beyond MACL's end date."],
        ["8. Missing record","If no applicable RAMIS ARR/DEP record exists, flag it as missing."],
        ["9. New flight","If neither applicable side exists, status = NEW FLIGHT."],
        ["10. Expired schedules","If MACL Effective End Date is on or before D-1, status = EXPIRED. The row remains in the full report, is dimmed grey, and is excluded from ACTION REQUIRED."],
        ["11. Days of Ops","DAYS OF OPS is always displayed as exactly 7 digits/text characters, preserving leading zeros."],
        ["12. Cargo","Cargo is ignored."],
        ["13. Days of OPS","Validate actual operating dates across MON-SUN; red Days-of-OPS rows are cancelled and do not bridge continuity."],
        ["14. Same flight D rule","Same MACL flight number: both ARR+DEP inside TMA window => DEP uses D; only one side inside => base number only, no D."],
        ["15. MACL source validation","If MACL / DAYS OF OPS contains a definite source error such as START DATE AFTER END DATE, flag MACL SOURCE CHECK REQUIRED and tell staff to go back and check DAYS OF OPS. An invalid red row cannot cancel a movement."],
        ["16. No active MACL / RAMIS check","If a VALID red DAYS OF OPS movement is cancelled but RAMIS still has an applicable record, show RAMIS data and flag RAMIS CHECK REQUIRED. If RAMIS has no record, leave RAMIS columns blank and use CANCELLED - NO ACTION."],
        ["17. RAMIS residual scan","After normal reconciliation, scan current/future RAMIS records for flights cancelled/no longer active in MACL. A cancelled regime only matches the SAME RAMIS timing; replacement timings are not falsely flagged. Fully historical RAMIS records are ignored."],
        ["18. RAMIS source fidelity","RAMIS STA/STD/EFFECTIVE values come only from actual selected RAMIS source records. When ARR and DEP source ranges differ, the EFFECTIVE cell shows their common applicable RAMIS period as one range. MACL dates are never copied into the RAMIS section."],
        ["19. Special date override","If MACL has a single-date/shorter period with different timing inside a broader actual RAMIS range, keep the true RAMIS source range, set revised timing/effective to the MACL exception period, and comment SPECIAL DATE OVERRIDE - update this date/period only."],
        ["20. Partial RAMIS coverage","If RAMIS covers only part of the required MACL operating dates, flag the uncovered dates as MISSING RAMIS ARRIVAL/DEPARTURE even when a later RAMIS record exists. Timing differences on covered dates are still reported separately."],
        ["21. Out-of-window RAMIS visibility","If RAMIS has an actual ARR/DEP record for a side outside the TMA comparison window, show that RAMIS time/effective range for reference and state in COMMENTS that the MACL side is outside the TMA window. It is informational only and must not create a timing-change action."],
        ["22. Historical coverage gaps","Past missing RAMIS operating dates do not create Action Required items. Partial-coverage warnings are generated only for today/future required operating dates; historical RAMIS data may still be shown for reference."],
        ["23. Direction","Revised values show what RAMIS should become based on MACL."],
    ]
    for r,row in enumerate(logic_rows,1):
        logic.cell(r,1,row[0]); logic.cell(r,2,row[1])
    logic.column_dimensions["A"].width = 28
    logic.column_dimensions["B"].width = 110
    logic["A1"].font = Font(bold=True, color="FFFFFF", size=14)
    logic["A1"].fill = PatternFill("solid", fgColor="1F4E78")

    out = BytesIO()
    wb.save(out)
    out.seek(0)
    return out

st.title("TMA – MACL vs RAMIS Reconciliation")
st.caption("MACL is the master schedule. Upload MACL as Excel or PDF; PDF is converted and validated before reconciliation.")

col1, col2 = st.columns(2)
with col1:
    macl_file = st.file_uploader(
        "1. Upload MACL Schedule (Excel or PDF)",
        type=["xlsx", "pdf"],
        key="macl",
        help="Upload the MACL Excel directly, or upload the standard MACL PDF and the app will convert it automatically."
    )
with col2:
    ramis_file = st.file_uploader("2. Upload RAMIS Connecting Flight Plans", type=["xlsx"], key="ramis")

macl_workbook_bytes = None
macl_input_type = None
pdf_conversion_summary = None

if macl_file is not None:
    macl_name = (macl_file.name or "").lower()
    if macl_name.endswith(".pdf"):
        macl_input_type = "PDF"
        try:
            with st.spinner("Converting MACL PDF to reconciliation-ready Excel..."):
                macl_workbook_bytes, pdf_conversion_summary = convert_macl_pdf_to_xlsx(macl_file.getvalue())
            st.success(
                f"MACL PDF converted successfully — "
                f"{pdf_conversion_summary['weekday_rows']} weekday rows and "
                f"{pdf_conversion_summary['days_ops_rows']} Days of OPS rows loaded."
            )
            st.download_button(
                "Download Converted MACL Excel",
                data=macl_workbook_bytes,
                file_name="MACL_CONVERTED_FROM_PDF.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
                key="download_converted_macl"
            )
            st.caption(
                f"PDF validation: {pdf_conversion_summary['pages']} pages read; "
                f"{pdf_conversion_summary['red_rows']} red/cancelled source rows detected."
            )
        except Exception as conversion_error:
            st.error("MACL PDF conversion could not be validated. Reconciliation has been stopped for this file.")
            st.exception(conversion_error)
            macl_workbook_bytes = None
    else:
        macl_input_type = "Excel"
        macl_workbook_bytes = macl_file.getvalue()

with st.expander("Current locked reconciliation logic"):
    st.markdown("""
- MACL is the official reference. MACL input may be **XLSX or the standard MACL PDF**; PDF is converted internally before reconciliation.
- Select the **latest RAMIS record that overlaps the MACL effective period**.
- Do not merge historical RAMIS effective periods.
- Compare STA/STD against the latest applicable RAMIS record.
- A RAMIS record starting earlier or later than MACL is acceptable.
- A RAMIS record ending earlier than MACL is acceptable if it overlaps and the timing is correct.
- **EFFECTIVE CHANGE is raised only when the latest applicable RAMIS record continues beyond the MACL end date.**
- Missing RAMIS arrival/departure records are flagged separately.
- MACL schedules ending on or before **D-1** are marked **EXPIRED**, dimmed grey, and excluded from Action Required.
- **DAYS OF OPS always displays 7 digits**, including leading zeros.
- **DAYS OF OPS validates the actual active operating dates across all seven weekdays.**
- RED DAYS OF OPS movements are cancelled and do not bridge operating continuity.
- RAMIS is reconciled only against the genuine active MACL operating dates for that weekday/range.
- The locked same-flight-number `D` rule remains unchanged.
- MACL rows with **SEAT = CARGO** are shown as `IGNORED`, dimmed grey, and excluded from Action Required.
""")

if macl_workbook_bytes and ramis_file:
    if st.button("Run Reconciliation", type="primary", use_container_width=True):
        try:
            macl_bytes = macl_workbook_bytes
            ramis_bytes = ramis_file.getvalue()

            with st.spinner("Reading schedules and reconciling..."):
                lookup, by_airline_day_type = parse_ramis(ramis_bytes)
                macl_rows = parse_macl(macl_bytes)
                days_ops_calendar = read_days_ops_calendar(macl_bytes)
                result = reconcile(
                    macl_rows, lookup, by_airline_day_type, days_ops_calendar
                )

            cnt = Counter(r["STATUS"] for r in result)

            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Total Checked", len(result))
            c2.metric("No Change", cnt.get("NO CHANGE",0))
            c3.metric("Changes", cnt.get("CHANGE",0))
            c4.metric("New Flights", cnt.get("NEW FLIGHT",0))

            df = pd.DataFrame(result)
            tabs = st.tabs(["Action Required","Full Reconciliation","No Change"])
            with tabs[0]:
                action_df = df[~df["STATUS"].isin(["NO CHANGE","IGNORED","EXPIRED","CANCELLED - NO ACTION"])]
                st.dataframe(action_df, use_container_width=True, height=520)
            with tabs[1]:
                st.dataframe(df, use_container_width=True, height=520)
            with tabs[2]:
                st.dataframe(df[df["STATUS"]=="NO CHANGE"], use_container_width=True, height=520)

            excel_out = build_excel(result)
            st.download_button(
                "Download Final Reconciliation Excel",
                data=excel_out.getvalue(),
                file_name="MACL_vs_RAMIS_FINAL_RECONCILIATION.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            st.success("Reconciliation completed.")
        except Exception as e:
            st.exception(e)
