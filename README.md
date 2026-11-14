# Workday HR Analytics Dashboard

A Streamlit people-analytics dashboard that pulls Workday HR data (demo dataset or file upload) into SQL-driven KPIs, attrition-driver views and a Python risk model, with a formula-driven Excel report export.

## Features
- **Overview:** headcount, active, terminations, attrition rate, average age/tenure/pay; attrition by organization, gender, age band and job profile
- **Attrition drivers:** overtime, travel, satisfaction, work-life balance, tenure, stock options, org x level heatmap
- **Compensation:** pay by level and attrition status, attrition by pay band, tenure vs pay
- **Attrition risk (ML):** logistic regression with ROC AUC, precision/recall, top drivers and a top-25 at-risk list with plain-language flags
- **SQL & export:** read-only SQL explorer over the `workers` table; Excel download (live-formula KPIs and breakdowns)
- Sidebar filters: supervisory org, job profile, gender, age band, education

## Quick start
```bash
pip install -r requirements.txt
streamlit run app.py
```
The app opens on the bundled demo data. No Workday access is needed to try it.

## Data sources (sidebar)
| Source | What it does |
|---|---|
| Demo | Bundled HR-analytics dataset (1,470 workers) mapped to Workday-style fields |
| Upload | CSV/XLSX export from Workday; for workbooks, the `Workers` sheet is used if present |

**Required fields:** `worker_id`, `supervisory_org`, `job_profile`, `worker_status`, `monthly_comp`, `age`.

## Synthetic dataset
`python generate_synthetic_workday.py` creates `data/synthetic_workday_hr_dataset.xlsx` (2,500 fictional workers, seed 42) with hire/termination dates, exit reasons, locations and built-in attrition patterns. It can be downloaded from the app's sidebar, and its `Workers` sheet loads back through Upload. All data is random; findings from it are illustrative only.

## Project structure
| File | Role |
|---|---|
| `app.py` | Streamlit UI |
| `workday_connector.py` | Demo and upload loaders; field mapping; derived fields |
| `data_layer.py` | SQLite `workers` table and parameterised SQL |
| `risk_model.py` | Attrition model and risk flags |
| `excel_report.py` | Excel results workbook |
| `generate_synthetic_workday.py` | Synthetic data generator |

## Notes and limitations
- Attrition rate = leavers / headcount in the segment. This differs from a segment's share of all leavers.
- The demo dataset is a public HR-analytics sample (see its original LICENSE), not Postman data.
- The model uses a random train/test split and includes demographic fields (gender, age, marital status). Remove them before any use on real employees, and validate on a time-based split once hire/termination dates are available.
- Precision for leavers is modest. Use risk scores to prioritise conversations, not to make decisions about individuals.

## Privacy and security
Individual risk scores are sensitive. Restrict access, host behind your SSO, and limit access to the exported files.

## Roadmap
Voluntary vs involuntary modelling, time-based validation, minimum-group-size suppression.
