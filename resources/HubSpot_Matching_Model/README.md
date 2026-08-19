# HubSpot Duplicate & Email-Matching Model

A reusable model that reconciles any marketing/prospect email list against a
HubSpot contacts export. Run it whenever a new list comes in.

## What it produces

One Excel workbook with five tabs:
- **Summary** — headline counts (list size, in HubSpot, net-new, clean names, need cleanup)
- **In HubSpot** — every list email that already exists, with its HubSpot contact info + a record count for duplicates
- **Net-new** — emails not in HubSpot, each flagged as a *known account (new contact)* or a *new account*
- **Clean Names** — net-new `first.last@` addresses parsed into First / Last / Email / Company, ready to import
- **Cleanup** — the rest, with empty name fields, a confidence rating, and a per-row hint for manual cleanup

## The rules it follows

- Matches by **email only** (exact, case-insensitive) — never by company name
- **Names derived only for clean `first.last` addresses**; everything else is left **blank** (a wrong mail-merge name is worse than a blank) and bucketed by confidence
- **Role / shared mailboxes are flagged** (parking@, permits@, admin, etc.)
- **Duplicates are kept intact**
- Auto-fixes common typos: a comma where a dot belongs in a domain, and two addresses fused together

## How to run it

Needs Python with `pandas` and `openpyxl`:

```
pip install pandas openpyxl
python hubspot_match_model.py <hubspot_export.xlsx> <target_list.xlsx> [output.xlsx]
```

The target list can be either a clean two-column sheet (company, emails) **or** a
grouped sheet (Category > Region > Org rows with multiple emails packed into one cell).

If your list uses different section headers, edit `CATEGORY_HEADERS` /
`REGION_HEADERS` / `ROLE_TERMS` near the top of the script.

You can also just upload both files to Claude and say "run the matching model" —
no local setup needed, or point the agent at workflows/HubSpot_Duplicate_Email_Matching.md.

---

## Setting this up as a standalone Claude Project (optional)

If you ever want this model to live in its own dedicated Claude Project
(separate from the general agent workspace) rather than as a workflow inside
this one:

1. In the Claude sidebar, click **New Project** and name it e.g. *HubSpot List Matching*.
2. Add `hubspot_match_model.py` and this README to the project knowledge.
3. Paste the text below into the project's **custom instructions** so every chat
   in the project runs the same way.

### Suggested project custom instructions

> This project reconciles marketing/prospect email lists against HubSpot contact
> exports. When I upload two spreadsheets, treat the HubSpot export as the source
> of truth and the other as the target list to check.
>
> Match by email address only (exact, case-insensitive), never by company name.
> Split results into emails already in HubSpot (pull back their contact info, keep
> duplicate records intact) and net-new emails (flag whether the domain is a known
> account). For net-new emails, derive First/Last names ONLY for clean first.last
> addresses; leave names blank rather than guess, and bucket the rest by confidence
> into a cleanup tab. Flag role/shared mailboxes. Deliver a formatted .xlsx with
> Summary, In HubSpot, Net-new, Clean Names, and Cleanup tabs. The
> hubspot_match_model.py file in this project encodes the exact logic — use it.
