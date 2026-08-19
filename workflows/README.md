# Workflows Index

This folder holds ready-to-run instruction files. Tell the agent which one to
use (by name, or just describe the task and it will match one), and it will
follow that recipe rather than improvising from scratch. Every workflow below
inherits the standing rules in the root CLAUDE.MD (clarify before starting,
show the plan first, Excel-with-tabs for reports, Word-doc-no-report-data for
summaries, Python by default, save to output/<Topic>/, cite sources) — the
files below only add the steps specific to that job.

## Available workflows

**Research_Brief_Report.md** — Turn a research topic into a cited, executive-
ready Word summary. This is the recipe behind the existing
Remote_Work_Trends and Copilot_Forms_Flow_Approvals reports in output/.

**Weekly_Team_Metrics_Report.md** — Compile recurring team metrics into the
standard tabbed Excel workbook (Summary tab + one tab per metric group).

**HubSpot_Duplicate_Email_Matching.md** — Reconcile a prospect/marketing email
list against a HubSpot contacts export using the model in
resources/HubSpot_Matching_Model/.

## Adding a new workflow

Copy the shape of an existing file: a one-line "when to use this," the exact
inputs needed from Maddie, the steps in order, and the expected output format.
Keep the standing CLAUDE.MD rules out of individual workflow files — they
already apply everywhere.
