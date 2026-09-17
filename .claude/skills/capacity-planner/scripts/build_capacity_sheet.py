#!/usr/bin/env python3
"""
Rebuilds the Medline Capacity Planning workbook for one week.

Inputs:
  - the CURRENT live Google Sheet, downloaded as .xlsx (read-only source of
    truth for the Name list, Activity list, "Potentially to be reduced"
    flags, and History log)
  - THIS WEEK'S raw Raydar's export .xlsx

Output:
  - a brand new .xlsx with Data / Per Resource - Per Activity / Summary /
    History / Anas Format tabs, ready to import into the same Google Sheet
    via File > Import > Replace spreadsheet.

Also prints a JSON report to stdout (new people, new/unmapped activities,
totals) so the calling agent can summarize the run for the user.
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
from openpyxl.utils import get_column_letter

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
            f"Weekly export is missing required column(s): {', '.join(missing)}. "
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
            "date": date_val,
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


def week_ending_label(rows, override):
    if override:
        return override
    dates = [r["date"] for r in rows if isinstance(r["date"], datetime)]
    if not dates:
        return "Unknown Week"
    max_date = max(dates)
    # Week starting Monday
    week_start = max_date - timedelta(days=max_date.weekday())
    week_end = week_start + timedelta(days=6)
    return f"{week_start.strftime('%Y-%m-%d')} to {week_end.strftime('%Y-%m-%d')}"


def read_current_sheet(path):
    """Pull the bits of the live sheet we need to preserve: name list, activity
    list + reduce-flags, and History log (if present)."""
    result = {
        "names": [],
        "activities": [],  # list of (activity, flag)
        "history_rows": [],  # (week_ending, name, role, activity, hours)
        "anas_format_rows": None,
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
            activity = str(ws.cell(r, 9).value).strip()
            flag = ws.cell(r, 12).value  # column L
            result["activities"].append((activity, flag))
            r += 1
    if "History" in wb.sheetnames:
        ws = wb["History"]
        for r in range(2, ws.max_row + 1):
            vals = [ws.cell(r, c).value for c in range(1, 6)]
            if vals[0] is None:
                continue
            result["history_rows"].append(tuple(vals))
    if "Anas Format" in wb.sheetnames:
        ws = wb["Anas Format"]
        rows = []
        for r in range(1, ws.max_row + 1):
            rows.append([ws.cell(r, c).value for c in range(1, ws.max_column + 1)])
        result["anas_format_rows"] = rows
    return result


def build_workbook(normalized_rows, weekly_raw_ws, weekly_header_row, current, week_label, report):
    wb = Workbook()
    wb.remove(wb.active)

    # ---- Data tab (this week only, matches "Total Hours (Last Week)") ----
    ws_data = wb.create_sheet("Data")
    ws_data.append(["Name", "Role", "Activity", "Hours"])
    for c in ws_data[1]:
        c.font = HEADER_FONT
    for row in normalized_rows:
        ws_data.append([row["name"], row["role"], row["activity"], round(row["hours"], 3)])
    ws_data.column_dimensions["A"].width = 24
    ws_data.column_dimensions["B"].width = 18
    ws_data.column_dimensions["C"].width = 40
    ws_data.column_dimensions["D"].width = 10
    n_data = len(normalized_rows)

    # ---- Raydars export tab (raw audit copy of this week's upload) ----
    ws_raw = wb.create_sheet("Raydars export")
    for row in weekly_raw_ws.iter_rows(min_row=weekly_header_row, values_only=True):
        ws_raw.append(list(row))
    for c in ws_raw[1]:
        c.font = HEADER_FONT

    # ---- Per Resource - Per Activity tab ----
    ws_pr = wb.create_sheet("Per Resource - Per Activity")
    ws_pr.append([None, "Name", "Role", "Activity", "Hours per Week"])
    for c in ws_pr[1]:
        c.font = HEADER_FONT
    for row in sorted(normalized_rows, key=lambda r: (r["name"], r["activity"])):
        ws_pr.append([None, row["name"], row["role"], row["activity"], round(row["hours"], 3)])
    ws_pr.column_dimensions["B"].width = 24
    ws_pr.column_dimensions["C"].width = 18
    ws_pr.column_dimensions["D"].width = 40

    # ---- Merge name / activity lists (preserve existing order, append new) ----
    names_seen = list(current["names"])
    for row in normalized_rows:
        if row["name"] not in names_seen:
            names_seen.append(row["name"])
            report["new_people"].add(row["name"])

    existing_activity_names = [a for a, _ in current["activities"]]
    flag_by_activity = {a: f for a, f in current["activities"]}
    activities_seen = list(existing_activity_names)
    data_activities = {row["activity"] for row in normalized_rows}
    for activity in sorted(data_activities):
        if activity not in activities_seen:
            activities_seen.append(activity)
            report["new_activities"].add(activity)

    # ---- Summary tab (formula-driven, same pattern as the original sheet) ----
    ws_sum = wb.create_sheet("Summary")
    ws_sum["B2"] = "Name"
    ws_sum["C2"] = "Role"
    ws_sum["D2"] = "Total Hours (Last Week)"
    ws_sum["E2"] = "Major Activity"
    ws_sum["F2"] = "Hours on Major Activity"
    ws_sum["G2"] = "% of Total"
    ws_sum["I2"] = "Activity"
    ws_sum["J2"] = "Total Hours (Week)"
    ws_sum["K2"] = "% of Team Time"
    ws_sum["L2"] = "Potentially to be reduced"
    for coord in ["B2", "C2", "D2", "E2", "F2", "G2", "I2", "J2", "K2", "L2"]:
        ws_sum[coord].font = HEADER_FONT

    last_data_row = n_data + 1
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
        flag = flag_by_activity.get(activity)
        if flag:
            ws_sum.cell(r, 12, flag)  # L
    last_activity_row = 2 + len(activities_seen)
    if activities_seen:
        ws_sum.cell(3, 13, f'=SUMIF(L3:L{last_activity_row}, "Y", J3:J{last_activity_row})')

    ws_sum.column_dimensions["B"].width = 22
    ws_sum.column_dimensions["C"].width = 16
    ws_sum.column_dimensions["E"].width = 30
    ws_sum.column_dimensions["I"].width = 45

    # ---- History tab (Week Ending, Name, Role, Activity, Hours) ----
    ws_hist = wb.create_sheet("History")
    ws_hist.append(["Week", "Name", "Role", "Activity", "Hours"])
    for c in ws_hist[1]:
        c.font = HEADER_FONT
    agg = {}
    for row in normalized_rows:
        key = (row["name"], row["role"], row["activity"])
        agg[key] = agg.get(key, 0.0) + row["hours"]
    existing_history = [h for h in current["history_rows"] if h[0] != week_label]
    for h in existing_history:
        ws_hist.append(list(h))
    for (name, role, activity), hours in sorted(agg.items()):
        ws_hist.append([week_label, name, role, activity, round(hours, 3)])
    ws_hist.column_dimensions["A"].width = 24
    ws_hist.column_dimensions["B"].width = 24
    ws_hist.column_dimensions["D"].width = 40

    # ---- Anas Format tab (static manual-entry template, passthrough) ----
    ws_anas = wb.create_sheet("Anas Format")
    template = current["anas_format_rows"] or [
        ["Name", "Rol", "Activities", "Hours per week"],
    ]
    for row in template:
        ws_anas.append(row)
    for c in ws_anas[1]:
        c.font = HEADER_FONT

    return wb, {
        "team_total_hours": round(sum(r["hours"] for r in normalized_rows), 2),
        "people_count": len(names_seen),
        "activity_count": len(activities_seen),
        "records_this_week": n_data,
    }


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--current", help="Current live sheet, downloaded as .xlsx")
    ap.add_argument("--weekly", required=True, help="This week's raw Raydar's export .xlsx")
    ap.add_argument("--output", required=True, help="Path to write the rebuilt workbook")
    ap.add_argument("--week-label", help="Override the auto-detected week label")
    args = ap.parse_args()

    activity_aliases = load_json(CONFIG_DIR / "activity_aliases.json")
    activity_aliases = {k: v for k, v in activity_aliases.items() if not k.startswith("_")}
    role_aliases = load_json(CONFIG_DIR / "role_aliases.json")
    role_aliases = {k: v for k, v in role_aliases.items() if not k.startswith("_")}

    raw_rows, weekly_ws, header_row = read_weekly_export(args.weekly)
    if not raw_rows:
        print(json.dumps({"error": "No usable rows found in the weekly export."}))
        sys.exit(1)

    report = {"unmapped_activities": set(), "new_people": set(), "new_activities": set()}
    normalized_rows = normalize_rows(raw_rows, activity_aliases, role_aliases, report)
    week_label = week_ending_label(raw_rows, args.week_label)

    current = read_current_sheet(args.current)
    wb, stats = build_workbook(normalized_rows, weekly_ws, header_row, current, week_label, report)

    Path(args.output).parent.mkdir(parents=True, exist_ok=True)
    wb.save(args.output)

    out_report = {
        "output_file": str(args.output),
        "week_label": week_label,
        **stats,
        "new_people": sorted(report["new_people"]),
        "new_activities": sorted(report["new_activities"]),
        "unmapped_activities": sorted(report["unmapped_activities"]),
    }
    print(json.dumps(out_report, indent=2))


if __name__ == "__main__":
    main()
