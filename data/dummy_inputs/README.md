# Dummy inputs: CleoUSA, month end July 2026

All files in this folder are made-up test data. They copy the column layout and value
formats of the real Citrus exports and the LabConnect inventory report.

Regenerate with:

```
.venv/bin/python scripts/generate_dummy_data.py
```

The seed is fixed, so every run gives the same files.

## Snapshot

- Export time: 2026-08-03 13:56 UTC, the first working day after the July month end.
- Status fields show the state on that day. Dated events cover August 2024 to 3 August 2026.

## How the files link

Enrollments per site and month match the site performance workbook. One simulated journey per participant drives every file:

pre-screen → consent → enrolment and blood draw → surgery → pathology report upload → study completion

| File | Rule |
|---|---|
| PARTICIPANTS_OVERVIEW | One row per participant. Status is the furthest milestone reached by the export time. |
| DATA_COMPLETENESS__CRF | 11 forms per participant. "Unavailable" means the form is not open yet for that participant. Forms with an Open query show "Open Query". |
| DATA_COMPLETENESS__EPRO | 2 ePRO forms per consented participant. Each row has exactly one outcome count set to 1. |
| DATA_QUERY | Queries start after consent. Age runs to "Updated at" for Closed queries and to the export time for Open and Reviewing queries. |
| Inclusion_Criteria / Exclusion_Criteria | A row exists only when someone started the form. |
| PostSurgical_Findings | A row exists once the surgery date is entered. COMPLETE means the pathology report is uploaded, and "Completed At" is the upload date. |
| Inventory report | One requisition per enrolled participant. DX1 and DX2 share the enrolment date as the collection date. Subject ID holds the participant ID. |

## Scenarios built in for the report to catch

- DX1 registered later than 14 days after collection.
- DX2 registered later than 28 days after collection.
- DX1 and DX2 registered on the same day.
- Pairs with one sample missing.
- Two enrolled participants with no sample registered.
- Recent collections where DX2 is not registered yet. These are pending, not missing.
- Participants with a surgery date but no pathology report.
- Stale CONSENT participants who never enrolled.
- Long-open queries from 2025.

## Reference data

- `data/reference/site_lookup.csv` comes from the CLEO-001 site performance workbook: site code, name, type,
  activation date, monthly target, tracker flag and termination.
- `data/reference/site_monthly_enrollment.csv` holds the workbook's enrollments per site per month.
  The generator matches these counts exactly up to the export date.
- `data/reference/site_academic_status.csv` holds academic status, from web research or an educated guess.
