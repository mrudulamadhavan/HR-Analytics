"""Workday connector.

DEMO mode  : maps the bundled HR-analytics dataset onto Workday-style field names.
UPLOAD mode: loads a CSV/XLSX export from Workday (or the synthetic dataset).
Both return the same canonical schema (see CANONICAL).
"""
from __future__ import annotations
import io, os
import numpy as np
import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))

# Demo-source header -> canonical (Workday-style) field
DEMO_MAP = {
    "emp no": "worker_id", "Employee Number": "employee_number", "Gender": "gender", "Age": "age",
    "Department": "supervisory_org", "Job Role": "job_profile", "Job Level": "management_level",
    "Marital Status": "marital_status", "Education": "education", "Education Field": "field_of_study",
    "Business Travel": "business_travel", "Over Time": "overtime", "Monthly Income": "monthly_comp",
    "Percent Salary Hike": "pct_salary_hike", "Years At Company": "years_at_company",
    "Years In Current Role": "years_in_role", "Years Since Last Promotion": "years_since_promotion",
    "Years With Curr Manager": "years_with_manager", "Total Working Years": "total_working_years",
    "Distance From Home": "distance_from_home", "Job Satisfaction": "job_satisfaction",
    "Environment Satisfaction": "environment_satisfaction", "Work Life Balance": "work_life_balance",
    "Performance Rating": "performance_rating", "Job Involvement": "job_involvement",
    "Relationship Satisfaction": "relationship_satisfaction", "Stock Option Level": "stock_option_level",
    "Training Times Last Year": "training_times_last_year", "Num Companies Worked": "num_companies_worked",
    "Attrition": "worker_status",
}
# Common Workday report column headers -> canonical
WORKDAY_MAP = {
    "Employee ID": "worker_id", "Worker ID": "worker_id", "Supervisory Organization": "supervisory_org",
    "Job Profile": "job_profile", "Management Level": "management_level", "Gender": "gender",
    "Worker Status": "worker_status", "Active Status": "worker_status", "Age": "age",
    "Total Base Pay - Monthly": "monthly_comp", "Years of Service": "years_at_company",
}
REQUIRED = ["worker_id", "supervisory_org", "job_profile", "worker_status", "monthly_comp", "age"]
CANONICAL = list(DEMO_MAP.values())

def _derive(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    s = df["worker_status"].astype(str).str.strip().str.lower()
    df["terminated"] = s.isin({"yes", "terminated", "inactive", "false", "0"}).astype(int)
    df["worker_status"] = np.where(df["terminated"] == 1, "Terminated", "Active")
    df["age_band"] = pd.cut(df["age"], [0, 24, 34, 44, 54, 200],
                            labels=["Under 25", "25-34", "35-44", "45-54", "55+"]).astype(str)
    if "years_at_company" in df:
        df["tenure_band"] = pd.cut(df["years_at_company"], [-1, 1, 4, 9, 200],
                                   labels=["0-1 yrs", "2-4 yrs", "5-9 yrs", "10+ yrs"]).astype(str)
    df["comp_band"] = pd.cut(df["monthly_comp"], [0, 2999, 5999, 8999, 11999, 1e9],
                             labels=["<3k", "3k-6k", "6k-9k", "9k-12k", "12k+"]).astype(str)
    return df

def _finish(df: pd.DataFrame) -> pd.DataFrame:
    missing = [c for c in REQUIRED if c not in df.columns]
    if missing:
        raise ValueError(f"The file is missing required fields {missing}. "
                         "Rename the columns to match (e.g. worker_id, supervisory_org, job_profile, worker_status, monthly_comp, age).")
    return _derive(df)

def load_demo() -> pd.DataFrame:
    raw = pd.read_csv(os.path.join(HERE, "data", "hr_analytics_raw.csv"))
    return _finish(raw.rename(columns=DEMO_MAP))

def load_upload(file) -> pd.DataFrame:
    """CSV/XLSX export from Workday (or any file already using the canonical or demo headers)."""
    if file.name.lower().endswith(".csv"):
        df = pd.read_csv(file)
    else:
        xl = pd.ExcelFile(file)
        sheet = "Workers" if "Workers" in xl.sheet_names else xl.sheet_names[0]
        df = xl.parse(sheet)
    df = df.rename(columns={**DEMO_MAP, **WORKDAY_MAP})
    return _finish(df)
