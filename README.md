# CleoUSA monthly trial summary

Builds the Cleo Diagnostics (CSP-01-001) monthly trial summary from Citrus exports and the
LabConnect inventory report. All numbers are computed in code. A person reviews the output before
anything goes to the sponsor.

## Run

```
python3 -m venv .venv
.venv/bin/pip install openpyxl python-docx matplotlib
.venv/bin/python scripts/import_site_list.py         # site list workbook -> data/reference/
.venv/bin/python scripts/generate_dummy_data.py      # dummy inputs in data/dummy_inputs/
.venv/bin/python scripts/build_report.py --month 2026-07
```

Use `--inputs <folder>` to point at real exports. Site details come from the CLEO-001 site performance
workbook via `scripts/import_site_list.py`. Academic status lives in `data/reference/site_academic_status.csv`,
with the basis for each call.

## Outputs (outputs/<YYYY-MM>/)

| File | Use |
|---|---|
| `CleoUSA_Monthly_Summary_<Mon><YYYY>_Sponsor.docx` and `.xlsx` | Sponsor version: no data completeness, no participant-level data |
| `CleoUSA_Monthly_Summary_<Mon><YYYY>_Internal.docx` and `.xlsx` | Internal version: every section, plus participant-level action lists in Excel |
| `CleoUSA_Monthly_Totals_<Mon><YYYY>.csv` | Totals only, no participant rows. This is the file that may go into Claude under SOP-91 |

Use `--audience sponsor` or `--audience internal` to build one version only.

## Trial Summary Builder (browser tool for non-technical users)

Published as a private claude.ai artifact: https://claude.ai/artifact/8CEatoK8RNhrVHtxwZq4WD

It reads the exports inside the browser, runs the same calculations as the Python report, and saves the Word,
Excel and totals files through the artifact's downloads capability. No data is uploaded and no AI reads it.

| Path | Role |
|---|---|
| `app/src/engine.js` | JavaScript port of `metrics.py` and `narrative.py`, plus file recognition and input checks |
| `app/src/model.js` | The content toggles and the sponsor and internal defaults |
| `app/src/writers.js` | Word (docx 8.5.0), Excel (ExcelJS 4.4.0), totals CSV and charts |
| `app/src/ui.js`, `app/src/page.html` | Page behaviour and layout |
| `app/build.py` | Inlines the scripts and site list into `app/dist/trial-summary-builder.html` and stages example files |
| `app/test/parity.js` | Fails if any JavaScript total differs from the Python totals file |
| `app/test/writers.js` | Builds sponsor and internal files in Node and scans sponsor files for participant data |
| `app/test/sponsor_text.js` | Checks that the sponsor highlights and action items match the Python report word for word |
| `app/test/charts.js` | Checks that the six story charts (data, colours, takeaway titles) match `trialsummary/charts.py` |
| `app/test/digest.js` | Checks that the Word report's "Inside this report" rows and section key figures match `narrative.digest()` |

```
.venv/bin/python scripts/build_report.py --month 2026-07   # Python totals, the parity reference
node app/test/parity.js 2026-07                            # JavaScript engine must match
python3 app/build.py                                       # assemble the artifact page
```

Opening the page with `#selftest` at the end of the address builds every output from the example data and
writes the files into the page as base64, for automated checks.

## Code

| Path | Role |
|---|---|
| `trialsummary/load.py` | Finds and parses the input files |
| `trialsummary/metrics.py` | Every metric and its denominator |
| `trialsummary/narrative.py` | Highlights, action items and definitions, built from the numbers |
| `trialsummary/charts.py` | Six story charts with takeaway titles, drawn in pixel coordinates; `app/src/writers.js` mirrors it |
| `trialsummary/excel.py`, `word.py` | Output writers |
