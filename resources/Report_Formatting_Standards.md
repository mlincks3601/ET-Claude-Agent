# Report Formatting Standards

Reference doc for the standing formatting rules in CLAUDE.MD, spelled out so
any workflow can point here instead of repeating them.

## Excel reports and summaries

- One workbook, one tab per logical group (metric group, category, etc.), plus
  a Summary tab as the first tab whenever there's more than one data tab.
- Bold header row, frozen top row, column widths sized to content — no color
  coding or conditional formatting unless specifically requested.
- Formulas over hardcoded values where a number is derived from other cells
  in the same workbook.

## Word document summaries

- Concise, prose-based, straight to the point — no bolded words, no bullet-
  heavy formatting, no embedded raw data tables (raw data belongs in the
  Excel companion file, not the summary).
- Always end with a Sources section: title plus URL for every external fact
  or statistic used.

## Code and scripts

- Python by default for scripts and automation unless another language is
  specified (HTML, Java, etc.).
- Save any generator script alongside the report/workbook it produces, in the
  same output/<Topic>/ folder, so the output can be regenerated later.

## File organization

- Finished deliverables go in output/<Topic_Name>/, named for the
  conversation topic (matches existing Remote_Work_Trends and
  Copilot_Forms_Flow_Approvals folders).
- Reusable recipes (plain-English instructions the agent follows) go in
  workflows/.
- Reference docs, templates, and reusable tools/scripts go in resources/.
- Any time a file or folder is created or moved, say so before finishing.
