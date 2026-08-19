# Workflow: HubSpot Duplicate & Email-Matching Model

**When to use this:** Maddie has a new marketing/prospect email list and
wants to know which addresses are already HubSpot contacts and which are
net-new, cleaned up for import.

## Inputs needed from Maddie

- The HubSpot contacts export (.xlsx) — treated as the source of truth.
- The target list to check (.xlsx) — either a clean two-column sheet
  (company, emails) or a grouped sheet (Category > Region > Org with emails
  packed into cells).

## Steps

1. Confirm both files are available (uploaded, or reachable via the device
   bridge).
2. Run resources/HubSpot_Matching_Model/hubspot_match_model.py against them:
   `python hubspot_match_model.py <hubspot_export.xlsx> <target_list.xlsx> [output.xlsx]`
   If the list uses non-standard section headers, adjust CATEGORY_HEADERS /
   REGION_HEADERS / ROLE_TERMS at the top of that script first.
3. The script matches by email only (exact, case-insensitive) — never by
   company name — and keeps duplicate records intact.
4. Confirm the output workbook has all five tabs: Summary, In HubSpot,
   Net-new, Clean Names, Cleanup. Spot-check a handful of rows in each before
   delivering.
5. Save to output/HubSpot_Matching_<ListName>_<Date>/ and deliver the
   workbook to Maddie.

See resources/HubSpot_Matching_Model/README.md for the full spec (matching
rules, name-parsing rules, role-mailbox flagging, typo auto-fixes).
