#!/usr/bin/env python3
"""
Rebuilds the Medline Capacity Planning workbook from a raw Raydar's export.

The export can cover any span - a single week, a month, a quarter, or a
full year of daily records - the script does not assume a fixed cadence.

Inputs:
  - the CURRENT live Google Sheet, downloaded as .xlsx (read-only source of
    truth for the Name list, Activity list, and the Daily/Weekly/Monthly
    logs)
  - a raw Raydar's export .xlsx covering whatever date range the user has

Output:
  - a brand new .xlsx with Data / Summary / Daily / Weekly / Monthly tabs,
    ready to import into the same Google Sheet via
    File > Import > Replace spreadsheet.
    - Data / Summary reflect the FULL span of the uploaded file (sorted by
      Name then Activity; Summary's header names the actual date range
      instead of assuming "last week").
    - Daily / Weekly / Monthly are flat logs, each row bucketed by that
      row's own Start Date, so a single upload spanning a year populates
      all three at the right grain in one run. Re-running for a period
      already logged replaces that period's rows instead of duplicating.

Also prints a JSON report to stdout (date range, new people, new/unmapped
activities, totals) so the calling agent can summarize the run.
"""
import argparse
import json
import re
import sys
from datetime import datetime, timedelta
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.worksheet.formula import ArrayFormula
from openpyxl.styles import Font

SCRIPT_DIR = Path(__file__).resolve().parent
CONFIG_DIR = SCRIPT_DIR.parent / "config"

HEADER_FONT = Font(bold=True)


def load_json(path):
    with open(path) as f:
        return json.load(f)


def normalize_key(s):
    return re.sub(r"\s+", " ", str(s or "").strip()).lower()


def title_case_fallback(raw):
    """Best-effort canonicalization for a Task string with no alias entry."""
    words = re.sub(r"\s+", " ", str(raw or "").strip()).split(" ")
    out = []
    ACRONYMS = {"qa", "pii", "fmea", "asg", "spt", "pmo", "ip"}
    for w in words:
        lw = w.lower()
        if lw in ACRONYMS:
            out.append(lw.upper())
        elif "-" in w:
            out.append("-".join(p[:1].upper() + p[1:].lower() if p else p for p in w.split("-")))
        else:
            out.append(w[:1].upper() + w[1:].lower() if w else w)
    return " ".join(out)


def parse_time_tracking_to_hours(value):
    """Fallback parser for strings like '1h 5m 37s' -> decimal hours."""
    if value is None:
        return None
    s = str(value)
    h = re.search(r"(\d+)\s*h", s)
    m = re.search(r"(\d+)\s*m", s)
    sec = re.search(r"(\d+)\s*s", s)
    total = 0.0
    if h:
        total += int(h.group(1))
    if m:
        total += int(m.group(1)) / 60.0
    if sec:
        total += int(sec.group(1)) / 3600.0
    return total if (h or m or sec) else None


def find_header_row(ws, expected_any):
    """Find the row index whose cell values (lower/stripped) intersect expected_any."""
    for r in range(1, min(ws.max_row, 10) + 1):
        vals = {normalize_key(c.value) for c in ws[r] if c.value is not None}
        if vals & expected_any:
            return r
    return 1


def read_weekly_export(path):
    wb = load_workbook(path, data_only=True)
    ws = wb.worksheets[0]
    header_row = find_header_row(ws, {"task", "task assigned to", "overall time"})
    headers = {}
    for cell in ws[header_row]:
        if cell.value is not None:
            headers[normalize_key(cell.value)] = cell.column

    def col(*names):
        for n in names:
            if n in headers:
                return headers[n]
        return None

    c_name = col("task assigned to")
    c_role = col("analyst designation")
    c_task = col("task")
    c_hours = col("overall time")
    c_time_tracking = col("time tracking")
    c_date = col("start date")

    missing = [label for label, c in [
        ("Task Assigned to", c_name), ("Task", c_task)
    ] if c is None]
    if missing:
        raise ValueError(
            f"Export is missing required column(s): {', '.join(missing)}. "
            f"Found headers: {sorted(headers.keys())}"
        )

    rows = []
    for r in range(header_row + 1, ws.max_row + 1):
        name = ws.cell(r, c_name).value
        if name is None or str(name).strip() == "":
            continue
        task = ws.cell(r, c_task).value
        role_raw = ws.cell(r, c_role).value if c_role else None
        hours = ws.cell(r, c_hours).value if c_hours else None
        if hours is None and c_time_tracking:
            hours = parse_time_tracking_to_hours(ws.cell(r, c_time_tracking).value)
        if hours is None:
            continue
        date_val = ws.cell(r, c_date).value if c_date else None
        rows.append({
            "name": str(name).strip(),
            "role_raw": str(role_raw).strip() if role_raw else "",
            "task_raw": str(task).strip() if task else "",
            "hours": float(hours),
            "date": date_val if isinstance(date_val, datetime) else None,
        })
    return rows, ws, header_row


def normalize_rows(raw_rows, activity_aliases, role_aliases, report):
    out = []
    for row in raw_rows:
        key = normalize_key(row["task_raw"])
        if key in activity_aliases:
            activity = activity_aliases[key]
        else:
            activity = title_case_fallback(row["task_raw"])
            report["unmapped_activities"].add(row["task_raw"])

        rkey = normalize_key(row["role_raw"])
        role = role_aliases.get(rkey, title_case_fallback(row["role_raw"]) if row["role_raw"] else "")

        out.append({
            "name": row["name"],
            "role": role,
            "activity": activity,
            "hours": row["hours"],
            "date": row["date"],
        })
    return out


# ---- Period bucketing: every row is bucketed by ITS OWN date, so one run
# can span a week, a quarter, or a year and still land in the right
# day/week/month buckets. ----

def period_day(d):
    return d.strftime("%Y-%m-%d")


def period_week(d):
    week_start = d - timedelta(days=d.weekday())  # Monday
    week_end = week_start + timedelta(days=6)
    return f"{week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}"


def period_month(d):
    return d.strftime("%Y-%m")


def date_range_label(normalized_rows, override):
    if override:
        return override
    dates = [r["date"] for r in normalized_rows if isinstance(r["date"], datetime)]
    if not dates:
        return "This Period"
    lo, hi = min(dates), max(dates)
    if lo.date() == hi.date():
        return lo.strftime("%b %d, %Y")
    return f"{lo.strftime('%b %d, %Y')} to {hi.strftime('%b %d, %Y')}"


def aggregate_by_period(normalized_rows, period_fn):
    agg = {}
    skipped = 0
    for row in normalized_rows:
        if not isinstance(row["date"], datetime):
            skipped += 1
            continue
        period = period_fn(row["date"])
        key = (period, row["name"], row["role"], row["activity"])
        agg[key] = agg.get(key, 0.0) + row["hours"]
    return agg, skipped


def merge_period_rows(existing_rows, agg):
    """Replace any existing (period, name, role, activity) rows whose period
    is covered by this run's data; keep every other existing row untouched;
    append the new aggregates. Re-running the same period is idempotent.

    Only safe for an ATOMIC period (Daily: a single date is either fully
    covered by an upload or not at all). It is NOT safe for Weekly/Monthly,
    where two different uploads can each cover *part* of the same week or
    month - replacing would silently drop whichever upload's hours aren't
    in the current run. Use `reaggregate_from_daily` for those instead."""
    new_periods = {k[0] for k in agg.keys()}
    kept = [row for row in existing_rows if row[0] not in new_periods]
    added = [
        (period, name, role, activity, round(hours, 3))
        for (period, name, role, activity), hours in agg.items()
    ]
    return sorted(kept + added, key=lambda r: (r[0], r[1], r[3]))


def reaggregate_from_daily(daily_rows, period_fn, legacy_rows):
    """Derive Weekly/Monthly totals by re-summing the authoritative Daily
    log (already merged across every run to date), so a period split across
    multiple uploads accumulates correctly instead of the later upload
    overwriting the earlier one's hours.

    `legacy_rows` are old Weekly/Monthly rows from before Daily tracking
    existed - kept as-is for any period with no Daily rows behind it, so
    that pre-existing history isn't silently dropped."""
    agg = {}
    for date_str, name, role, activity, hours in daily_rows:
        try:
            d = datetime.strptime(str(date_str), "%Y-%m-%d")
        except ValueError:
            continue
        period = period_fn(d)
        key = (period, name, role, activity)
        agg[key] = agg.get(key, 0.0) + hours
    covered_periods = {k[0] for k in agg.keys()}
    legacy_kept = [row for row in legacy_rows if row[0] not in covered_periods]
    recomputed = [
        (period, name, role, activity, round(hours, 3))
        for (period, name, role, activity), hours in agg.items()
    ]
    return sorted(legacy_kept + recomputed, key=lambda r: (r[0], r[1], r[3]))


def read_current_sheet(path):
    """Pull the bits of the live sheet we need to preserve: name list,
    activity list, and the Daily/Weekly/Monthly logs (if present)."""
    result = {
        "names": [],
        "activities": [],
        "daily_rows": [],
        "weekly_rows": [],
        "monthly_rows": [],
    }
    if not path:
        return result
    wb = load_workbook(path, data_only=False)
    if "Summary" in wb.sheetnames:
        ws = wb["Summary"]
        r = 3
        while ws.cell(r, 2).value:  # column B
            result["names"].append(str(ws.cell(r, 2).value).strip())
            r += 1
        r = 3
        while ws.cell(r, 9).value:  # column I
            result["activities"].append(str(ws.cell(r, 9).value).strip())
            r += 1

    def read_log(sheet_name):
        rows = []
        if sheet_name in wb.sheetnames:
            ws = wb[sheet_name]
            for r in range(2, ws.max_row + 1):
                vals = [ws.cell(r, c).value for c in range(1, 6)]
                if vals[0] is None:
                    continue
                rows.append(tuple(vals))
        return rows

    result["daily_rows"] = read_log("Daily")
    result["weekly_rows"] = read_log("Weekly") or read_log("History")  # "History" = old tab name
    result["monthly_rows"] = read_log("Monthly")
    return result


def build_workbook(normalized_rows, weekly_raw_ws, weekly_header_row, current, period_label, report):
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Data tab: the FULL span of this upload, sorted by Name then
    # Activity (doubles as the per-resource-per-activity view). ----
    ws_data = wb.create_sheet("Data")
    ws_data.append(["Name", "Role", "Activity", "Hours"])
    for c in ws_data[1]:
        c.font = HEADER_FONT
    for row in sorted(normalized_rows, key=lambda r: (r["name"], r["activity"])):
        ws_data.append([row["name"], row["role"], row["activity"], round(row["hours"], 3)])
    ws_data.column_dimensions["A"].width = 24
    ws_data.column_dimensions["B"].width = 18
    ws_data.column_dimensions["C"].width = 40
    ws_data.column_dimensions["D"].width = 10
    n_data = len(normalized_rows)

    # ---- Raydars export tab (raw audit copy of the uploaded file) ----
    ws_raw = wb.create_sheet("Raydars export")
    for row in weekly_raw_ws.iter_rows(min_row=weekly_header_row, values_only=True):
        ws_raw.append(list(row))
    for c in ws_raw[1]:
        c.font = HEADER_FONT

    # ---- Merge name / activity lists (preserve existing order, append new) ----
    names_seen = list(current["names"])
    for row in normalized_rows:
        if row["name"] not in names_seen:
            names_seen.append(row["name"])
            report["new_people"].add(row["name"])

    activities_seen = list(current["activities"])
    data_activities = {row["activity"] for row in normalized_rows}
    for activity in sorted(data_activities):
        if activity not in activities_seen:
            activities_seen.append(activity)
            report["new_activities"].add(activity)

    # ---- Summary tab (formula-driven; header names the actual date range
    # covered by this upload rather than assuming "last week") ----
    ws_sum = wb.create_sheet("Summary")
    ws_sum["B2"] = "Name"
    ws_sum["C2"] = "Role"
    ws_sum["D2"] = f"Total Hours ({period_label})"
    ws_sum["E2"] = "Major Activity"
    ws_sum["F2"] = "Hours on Major Activity"
    ws_sum["G2"] = "% of Total"
    ws_sum["I2"] = "Activity"
    ws_sum["J2"] = f"Total Hours ({period_label})"
    ws_sum["K2"] = "% of Team Time"
    for coord in ["B2", "C2", "D2", "E2", "F2", "G2", "I2", "J2", "K2"]:
        ws_sum[coord].font = HEADER_FONT

    for i, name in enumerate(names_seen):
        r = 3 + i
        ws_sum.cell(r, 2, name)  # B
        ws_sum.cell(r, 3, ArrayFormula(
            f"C{r}", f'=IFERROR(INDEX(Data!$B:$B,MATCH(B{r},Data!$A:$A,0)),"")'
        ))
        ws_sum.cell(r, 4, f"=SUMIF(Data!$A:$A,B{r},Data!$D:$D)").number_format = "0.00"
        ws_sum.cell(r, 5, ArrayFormula(
            f"E{r}",
            f'=IFERROR(INDEX(Data!$C:$C,MATCH(1,(Data!$A:$A=B{r})*(Data!$D:$D=MAXIFS(Data!$D:$D,Data!$A:$A,B{r})),0)),"")'
        ))
        ws_sum.cell(r, 6, f"=IFERROR(MAXIFS(Data!$D:$D,Data!$A:$A,B{r}),0)").number_format = "0.00"
        ws_sum.cell(r, 7, f"=IFERROR(F{r}/D{r},0)").number_format = "0%"

    for i, activity in enumerate(activities_seen):
        r = 3 + i
        ws_sum.cell(r, 9, activity)  # I
        ws_sum.cell(r, 10, f"=SUMIF(Data!$C:$C,I{r},Data!$D:$D)").number_format = "0.00"
        ws_sum.cell(r, 11, f"=IFERROR(J{r}/SUM(Data!$D:$D),0)").number_format = "0.0%"

    ws_sum.column_dimensions["B"].width = 22
    ws_sum.column_dimensions["C"].width = 16
    ws_sum.column_dimensions["E"].width = 30
    ws_sum.column_dimensions["I"].width = 45

    # ---- Daily / Weekly / Monthly logs: each row bucketed by its OWN date,
    # so one upload spanning a year fills in the right buckets across all
    # three grains in one run.
    #
    # Daily is atomic (a date is either fully covered by an upload or not),
    # so it merges by simple replace-on-match. Weekly/Monthly are then
    # RECOMPUTED from the merged Daily log rather than merged directly -
    # otherwise a week/month split across two uploads (e.g. days 1-3 from
    # one file, days 4-5 from another) would have the second upload's
    # partial total silently overwrite the first's instead of adding to it.
    daily_agg, skipped_no_date = aggregate_by_period(normalized_rows, period_day)
    merged_daily_rows = merge_period_rows(current["daily_rows"], daily_agg)
    weekly_rows = reaggregate_from_daily(merged_daily_rows, period_week, current["weekly_rows"])
    monthly_rows = reaggregate_from_daily(merged_daily_rows, period_month, current["monthly_rows"])

    def write_log(sheet_name, header_label, rows):
        ws = wb.create_sheet(sheet_name)
        ws.append([header_label, "Name", "Role", "Activity", "Hours"])
        for c in ws[1]:
            c.font = HEADER_FONT
        for row in rows:
            ws.append(list(row))
        ws.column_dimensions["A"].width = 24
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["D"].width = 40
        return ws.max_row - 1

    n_daily = write_log("Daily", "Date", merged_daily_rows)
    n_weekly = write_log("Weekly", "Week", weekly_rows)
    n_monthly = write_log("Monthly", "Month", monthly_rows)

    return wb, {
        "team_total_hours": round(sum(r["hours"] for r in normalized_rows), 2),
        "people_count": len(names_seen),
        "activity_count": len(activities_seen),
        "records_in_upload": n_data,
        "rows_missing_date": skipped_no_date,
        "daily_log_rows": n_daily,
        "weekly_log_rows": n_weekly,
        "monthly_log_rows": n_monthly,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--current", help="Current live sheet, downloaded as .xlsx")
    ap.add_argument("--weekly", required=True, help="Raw Raydar's export .xlsx (any date span)")
    ap.add_argument("--output", required=True, help="Path to write the rebuilt workbook")
    ap.add_argument(
        "--period-label",
        help="Override the auto-detected date-range label shown in the Summary "
             "header. Display only - Daily/Weekly/Monthly bucketing always uses "
             "each row's own date regardless of this.",
    )
    args = ap.parse_args()

    activity_aliases = load_json(CONFIG_DIR / "activity_aliases.json")
    activity_aliases = {k: v for k, v in activity_aliases.items() if not k.startswith("_")}
    role_aliases = load_json(CONFIG_DIR / "role_aliases.json")
    role_aliases = {k: v for k, v in role_aliases.items() if not k.startswith("_")}

    raw_rows, weekly_ws, header_row = read_weekly_export(args.weekly)
    if not raw_rows:
        print(json.dumps({"error": "No usable rows found in the export."}))
        sys.exit(1)

    report = {"unmapped_activities": set(), "new_people": set(), "new_activities": set()}
    normalized_rows = normalize_rows(raw_rows, activity_aliases, role_aliases, report)
    period_label = date_range_label(normalized_rows, args.period_label)

    current = read_current_sheet(args.current)
    wb, stats = build_workbook(normalized_rows, weekly_ws, header_row, current, period_label, report)

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    out_report = {
        "output_file": str(args.output),
        "period_label": period_label,
        **stats,
        "new_people": sorted(report["new_people"]),
        "new_activities": sorted(report["new_activities"]),
        "unmapped_activities": sorted(report["unmapped_activities"]),
    }
    print(json.dumps(out_report, indent=2))


if __name__ == "__main__":
    main()
