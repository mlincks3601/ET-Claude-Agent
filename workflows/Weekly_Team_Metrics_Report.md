# Workflow: Weekly/Recurring Team Metrics Report

**When to use this:** Maddie provides (or points to) raw team metric data —
ticket volumes, CRM activity, workflow completion rates, headcount/utilization
— and wants it turned into her standard reporting format.

## Inputs needed from Maddie

- The source data (uploaded file, or a description of where to pull it from).
- Which metrics matter this cycle and any target/threshold values to flag
  against.
- The reporting period (week, month, quarter) and whether this replaces or
  appends to a prior report.

## Steps

1. Clarify which metrics and time window before building anything.
2. Load and validate the source data — flag missing or suspicious values back
   to Maddie rather than silently dropping them.
3. Build one Excel workbook with a tab per metric group and a Summary tab
   first, listing headline numbers and any values that missed target.
4. Keep formatting plain: bold header row only, no heavy color coding unless
   Maddie asks for it, consistent column widths.
5. Save to output/<Reporting_Period_Name>/ (e.g. output/Team_Metrics_2026-W33/).
6. If this is the first time running this exact report, ask Maddie whether it
   should become a scheduled task for future weeks.
