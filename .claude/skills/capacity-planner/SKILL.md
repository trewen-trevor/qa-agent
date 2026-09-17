---
name: capacity-planner
description: Refresh the Medline team Capacity Planning Google Sheet from a weekly raw Raydar's time-tracking export. Cleans and normalizes the export (name/role/activity naming drift), rebuilds the Data / Per Resource - Per Activity / Summary tabs, and appends a dated row to a History tab for trend tracking over time. Use when the user says "run capacity planner", "update the Medline capacity sheet", "refresh capacity planning", or hands over this week's Raydar's export alongside a link to the Capacity Planning Google Sheet.
---

# Capacity Planner (Medline)

Turns a weekly raw Raydar's export into a rebuilt Capacity Planning workbook
that gets imported back into the same live Google Sheet.

## Why this exists / how it works

The Capacity Planning sheet is formula-driven: the `Summary` and
`Per Resource - Per Activity` tabs are all `SUMIF`/`INDEX`/`MATCH`/`MAXIFS`
formulas that read off a flat `Data` tab (Name, Role, Activity, Hours). The
raw Raydar's export uses inconsistent naming ("Product qa" vs "Product QA",
"Workspace Conig" typo, "call" vs "Calls", "Associate Data Analyst" folded
into "Data Analyst", etc.) — cleaning that up by hand every week is the
actual manual work today.

**Known tool limitation — read this before promising anything to the user:**
there is no connected Google Sheets values-write API in this environment
(only Google Drive file-level operations: read/download/create/share/trash,
metadata-only update). That means this skill **cannot** edit cells of the
live sheet in place. The agreed workflow is a full-workbook rebuild that the
user imports over the existing sheet via Google Sheets' own
**File > Import > Replace spreadsheet**, which preserves the same URL. Do
not attempt to "just update the Google Sheet directly" — it isn't possible
with the connectors available here. If the user wants zero-manual-step
automation later, that requires a Google Sheets API service account or an
Apps Script bound to the sheet (out of scope for this skill as built).

## Inputs needed each run

1. The **Google Sheet link** to the live Capacity Planning sheet (the user
   shares this every time — do not assume it's the same as a prior run).
2. **This week's raw Raydar's export** (.xlsx), attached/uploaded by the
   user in the conversation.

If either is missing, ask for it before doing anything else.

## Procedure

1. **Resolve the file ID** from the Google Sheet URL (the segment between
   `/d/` and the next `/`).

2. **Download the current live sheet** so the rebuild can preserve its
   Name list, Activity list, "Potentially to be reduced" flags, and History
   log:
   - Call `mcp__Google_Drive__download_file_content` with that `fileId` and
     `exportMimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"`.
   - Decode the base64 result and save it to the scratchpad, e.g.
     `current_live_sheet.xlsx`.
   - If the download fails (wrong link, no access), stop and tell the user
     — don't silently build from scratch, since that would blow away their
     Name/Activity ordering and reducible-hours flags.

3. **Save the user's uploaded weekly Raydar's export** to the scratchpad,
   e.g. `weekly_export.xlsx`.

4. **Run the build script**:
   ```
   python3 .claude/skills/capacity-planner/scripts/build_capacity_sheet.py \
     --current <scratchpad>/current_live_sheet.xlsx \
     --weekly <scratchpad>/weekly_export.xlsx \
     --output <scratchpad>/Capacity_Planning_Medline_<week-label>.xlsx
   ```
   It prints a JSON report: week label, team total hours, headcount,
   activity count, and — importantly — `new_people`, `new_activities`, and
   `unmapped_activities`.

   Optionally pass `--week-label "YYYY-MM-DD to YYYY-MM-DD"` to override the
   auto-detected week (auto-detection uses the Monday-Sunday week containing
   the latest `Start Date` in the export).

5. **Review the report before handing off:**
   - `unmapped_activities`: raw Task strings with no entry in
     `config/activity_aliases.json`. The script falls back to a
     best-effort title-case, but flag these to the user — if it's a
     genuinely new activity, that's fine (it gets added), but if it's just
     another spelling of an existing one, add a permanent mapping to
     `config/activity_aliases.json` (raw string, lowercased, → canonical
     name) and re-run so it doesn't create a duplicate bucket in Summary.
   - `new_people` / `new_activities`: call these out explicitly — a new
     hire or a brand-new activity bucket is exactly the kind of change a
     capacity planner should surface, not bury in a diff.
   - Sanity-check `team_total_hours` and `people_count` against what the
     user expects for a normal week (e.g. flag if someone's weekly total
     looks implausibly low/high — see "Suggested checks" below).

6. **Hand off the file.** Use `SendUserFile` to send the generated
   workbook. In the same message, give the user the one manual step:
   > Open the Capacity Planning Google Sheet → File → Import → Upload →
   > select this file → choose **Replace spreadsheet** → Import data.
   > This keeps the same sheet link; Summary and the other tabs recalculate
   > automatically since they're formulas reading off the Data tab.

7. **Summarize the week** in chat: team total hours, any new people/activities,
   top activity by hours, and anything flagged as unmapped that needs a
   permanent alias.

## Suggested checks before flagging a week as "done"

- Does `records_this_week` roughly match expectations (no obviously
  truncated export)?
- Is anyone's weekly total near 0 (didn't log time) or implausibly high
  (>50h/week)?
- Do `new_activities` look like real new work streams, or typos that should
  go into `config/activity_aliases.json` instead?

## Maintaining the alias configs

- `config/activity_aliases.json` — raw Task string (lowercased, trimmed) →
  canonical Activity name shown on the sheet.
- `config/role_aliases.json` — raw Analyst Designation → canonical Role
  (e.g. "Associate Data Analyst" is deliberately folded into "Data Analyst"
  to match the account's existing reporting convention; "Senior Analyst"
  stays distinct).

Edit these directly (they're plain JSON) whenever the export introduces a
new spelling of an existing name. This is expected periodic maintenance,
not a bug — call it out to the user the first time it happens so they know
the file exists and can ask for edits.

## Known pre-existing data quirk

The original sheet's Activity list already contains a legacy combined
bucket, `App QA / Product QA`, alongside separate `App QA` and `Product QA`
entries — historical rows are inconsistently split across the three. This
skill's alias table deliberately routes new raw "App qa"/"Product qa" rows
into the clean, separate `App QA` / `Product QA` buckets going forward and
does not add new hours to the legacy combined one. Mention this once to the
user; it's a data-quality artifact from before this skill existed, not
something the script got wrong.
