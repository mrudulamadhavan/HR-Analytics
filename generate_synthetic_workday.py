"""Synthetic Workday-style worker dataset (fully fictional; no real people or company data).
Usage: python generate_synthetic_workday.py  ->  data/synthetic_workday_hr_dataset.xlsx (+ _workers.csv)
Sheet 1 'Workers' uses the dashboard's field names, so it can be loaded via the app's Upload option."""
import numpy as np, pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SEED, N = 42, 2500
AS_OF = pd.Timestamp("2026-11-14")
WINDOW_START = pd.Timestamp("2025-01-01")       # leavers included from this date
TARGET_ATTR = 0.15
rng = np.random.default_rng(SEED)

ORGS = {  # org: (weight, overtime_p, [(job_profile, level, base monthly USD, job_family)])
 "Engineering": (.38, .28, [("Software Engineer",2,7000,"Engineering"),("Senior Software Engineer",3,9500,"Engineering"),
     ("Staff Engineer",5,13500,"Engineering"),("Engineering Manager",4,12000,"Engineering"),("QA Engineer",2,5500,"Engineering"),("Site Reliability Engineer",3,9000,"Engineering")]),
 "Product & Design": (.10, .22, [("Product Manager",3,9500,"Product"),("Senior Product Manager",4,12000,"Product"),
     ("Product Designer",2,7500,"Design"),("Group Product Manager",5,14500,"Product")]),
 "Sales": (.16, .22, [("Sales Development Rep",1,4500,"Sales"),("Account Executive",3,8500,"Sales"),
     ("Sales Manager",4,11500,"Sales"),("Solutions Engineer",3,9000,"Sales")]),
 "Customer Success": (.09, .35, [("Customer Success Manager",2,6500,"Customer Success"),("Support Engineer",2,5200,"Support"),("CS Manager",4,10000,"Customer Success")]),
 "Marketing": (.07, .20, [("Marketing Manager",3,8000,"Marketing"),("Content Strategist",2,6000,"Marketing"),
     ("Developer Advocate",3,8500,"Marketing"),("Demand Gen Specialist",2,6200,"Marketing")]),
 "Solutions & Support": (.10, .33, [("Technical Support Specialist",1,4200,"Support"),("Implementation Consultant",3,8000,"Services"),("Support Team Lead",3,7500,"Support")]),
 "G&A": (.10, .15, [("HR Business Partner",3,8000,"People"),("Recruiter",2,6000,"People"),("Financial Analyst",2,6500,"Finance"),
     ("Legal Counsel",4,12500,"Legal"),("IT Support Specialist",2,5000,"IT")]),
}
CC = {"Engineering":"CC1000","Product & Design":"CC2000","Sales":"CC3000","Customer Success":"CC4000","Marketing":"CC5000","Solutions & Support":"CC6000","G&A":"CC7000"}
LOCS = {"Bengaluru":(.52,.50,"India"),"San Francisco":(.16,1.45,"United States"),"New York":(.08,1.35,"United States"),
        "London":(.08,1.15,"United Kingdom"),"Remote - India":(.10,.48,"India"),"Remote - US":(.06,1.25,"United States")}
def pick(d, n, p): return rng.choice(list(d), n, p=np.array(p)/np.sum(p))

# ---- structure
org = pick(ORGS, N, [v[0] for v in ORGS.values()])
rows = []
for o in org:
    prof, lvl, base, fam = ORGS[o][2][rng.integers(len(ORGS[o][2]))]
    rows.append((prof, lvl, base, fam))
df = pd.DataFrame(rows, columns=["job_profile","management_level","base","job_family"]); df["supervisory_org"] = org
loc = pick(LOCS, N, [v[0] for v in LOCS.values()]); df["location"] = loc
df["country"] = [LOCS[l][2] for l in loc]
loc_mult = np.array([LOCS[l][1] for l in loc])
df["monthly_comp"] = np.round(df.base * loc_mult * rng.lognormal(0, .12, N), -1).astype(int)
df["comp_ratio"] = df.monthly_comp / (df.base * loc_mult)

# ---- demographics & career
df["age"] = np.clip(np.round(22 + df.management_level * 2.6 + rng.normal(4, 3.5, N)), 21, 62).astype(int)
df["gender"] = rng.choice(["Male","Female","Non-binary"], N, p=[.55,.43,.02])
df["marital_status"] = rng.choice(["Single","Married","Divorced"], N, p=[.40,.52,.08])
df["education"] = rng.choice(["High School","Associates Degree","Bachelor's Degree","Master's Degree","Doctoral Degree"], N, p=[.03,.07,.55,.32,.03])
df["field_of_study"] = [rng.choice(["Computer Science","Engineering","Business","Design","Communications","Other"],
                        p=[.55,.15,.12,.05,.05,.08] if o in ("Engineering","Product & Design") else [.12,.08,.45,.08,.17,.10]) for o in org]
df["total_working_years"] = np.clip(df.age - 22 + rng.integers(-2, 2, N), 0, df.age - 18).astype(int)
ten_asof = np.minimum(rng.gamma(1.6, 2.0, N), np.minimum(df.total_working_years + .5, 11))
hire = AS_OF - pd.to_timedelta(ten_asof * 365.25, unit="D")
df["hire_date"] = hire.dt.normalize()
df["num_companies_worked"] = np.clip(np.round(df.total_working_years / 4 + rng.normal(0, 1, N)), 0, 9).astype(int)
df["distance_from_home"] = np.clip(np.round(rng.gamma(2, 4.5, N)), 1, 50).astype(int)

# ---- engagement / work pattern
ot_p = np.array([ORGS[o][1] for o in org]); df["overtime"] = np.where(rng.random(N) < ot_p, "Yes", "No")
trav_p = {"Sales":[.08,.52,.40]}
df["business_travel"] = [rng.choice(["Non-Travel","Travel_Rarely","Travel_Frequently"], p=trav_p.get(o,[.16,.72,.12])) for o in org]
def lik(n, p=(.10,.22,.40,.28)): return rng.choice([1,2,3,4], n, p=p)
df["job_satisfaction"], df["environment_satisfaction"], df["relationship_satisfaction"], df["job_involvement"], df["work_life_balance"] = (lik(N), lik(N), lik(N), lik(N,(.04,.14,.58,.24)), lik(N,(.06,.22,.55,.17)))
df["performance_rating"] = rng.choice([1,2,3,4], N, p=[.03,.17,.60,.20])
df["stock_option_level"] = rng.choice([0,1,2,3], N, p=[.30,.40,.20,.10])
df["training_times_last_year"] = rng.choice(range(7), N, p=[.04,.06,.20,.32,.20,.12,.06])
df["pct_salary_hike"] = np.clip(np.round(rng.normal(12 + (df.performance_rating - 3) * 3, 3)), 5, 30).astype(int)

# ---- attrition (driven by realistic factors; intercept tuned to TARGET_ATTR)
t = ten_asof
promo_gap = np.minimum(ten_asof * rng.uniform(.3, 1, N), ten_asof)
logit = (0.95*(df.overtime=="Yes") + 0.55*(df.job_satisfaction<=2) + 0.40*(df.work_life_balance<=2)
         + 0.55*(df.business_travel=="Travel_Frequently") - 0.17*np.minimum(t,10) + 0.75*(t<1.5)
         + 0.35*(promo_gap>=3) - 1.3*(df.comp_ratio-1) - 0.25*(df.stock_option_level>=2) + 0.012*df.distance_from_home
         + 0.30*(df.performance_rating<=2) + rng.normal(0, .35, N)).values
lo, hi = -8, 2
for _ in range(60):
    mid = (lo+hi)/2
    (lo, hi) = (mid, hi) if (1/(1+np.exp(-(logit+mid)))).mean() < TARGET_ATTR else (lo, mid)
p = 1/(1+np.exp(-(logit+mid)))
term = rng.random(N) < p
earliest = (df.hire_date + pd.Timedelta(days=60)).clip(lower=WINDOW_START)
latest = AS_OF - pd.Timedelta(days=7)
term &= (earliest < latest).values
span = (latest - earliest).dt.days.clip(lower=1).values
term_date = earliest + pd.to_timedelta((rng.random(N) * span).astype(int), unit="D")
df["termination_date"] = pd.Series(np.where(term, term_date, pd.NaT), dtype="datetime64[ns]")
df["worker_status"] = np.where(term, "Terminated", "Active")
end = df.termination_date.fillna(AS_OF)
df["years_at_company"] = np.round(((end - df.hire_date).dt.days / 365.25), 1)
tn = np.maximum(df.years_at_company.round().astype(int), 0)
df["years_in_role"] = np.minimum(np.round(tn * rng.beta(2, 2, N)), tn).astype(int)
df["years_since_promotion"] = np.minimum(np.round(df.years_in_role * rng.uniform(.3, 1, N)), tn).astype(int)
df["years_with_manager"] = np.minimum(np.round(tn * rng.beta(2, 3, N)), tn).astype(int)
df["years_at_company"] = tn  # whole years, as in the dashboard dataset

# ---- termination reasons
vol = rng.random(N) > (0.08 + 0.35*(df.performance_rating<=2))
reasons_v = ["Better Opportunity","Compensation","Career Change","Relocation","Personal Reasons","Work-Life Balance","Lack of Growth"]
reasons_i = ["Performance","Restructuring / Layoff","Conduct","End of Contract"]
df["termination_category"] = np.where(term, np.where(vol, "Voluntary", "Involuntary"), "")
df["termination_reason"] = [("" if not tm else (rng.choice(reasons_v, p=[.28,.18,.10,.08,.10,.14,.12]) if v else rng.choice(reasons_i, p=[.50,.30,.12,.08])))
                            for tm, v in zip(term, vol)]

# ---- ids, managers, misc
order = df.sort_values("hire_date").index
ids = pd.Series(range(100001, 100001 + N), index=order)
df["worker_id"] = ["W-%d" % ids[i] for i in df.index]
df["employee_number"] = ids.values[np.argsort(order)] if False else [int(ids[i]) for i in df.index]
mgr_pool = df[df.management_level >= 4].groupby("supervisory_org").worker_id.apply(list).to_dict()
df["manager_worker_id"] = [rng.choice(mgr_pool.get(o, df.worker_id.tolist())) for o in df.supervisory_org]
df["pay_grade"] = ["G%d" % l for l in df.management_level]
df["cost_center"] = df.supervisory_org.map(CC)
df["fte"] = rng.choice([1.0, .8, .6], N, p=[.95,.04,.01])
df["time_type"] = np.where(df.fte == 1.0, "Full-time", "Part-time")
df["hire_source"] = rng.choice(["Employee Referral","Job Board","Agency","Campus","Direct Sourcing"], N, p=[.28,.30,.12,.10,.20])
df["currency"] = "USD"
df["as_of_date"] = AS_OF

CORE = ["worker_id","employee_number","gender","age","supervisory_org","job_profile","management_level","marital_status","education",
        "field_of_study","business_travel","overtime","monthly_comp","pct_salary_hike","years_at_company","years_in_role",
        "years_since_promotion","years_with_manager","total_working_years","distance_from_home","job_satisfaction",
        "environment_satisfaction","work_life_balance","performance_rating","job_involvement","relationship_satisfaction",
        "stock_option_level","training_times_last_year","num_companies_worked","worker_status"]
EXT = ["worker_id","hire_date","termination_date","termination_category","termination_reason","job_family","pay_grade","cost_center",
       "location","country","manager_worker_id","fte","time_type","hire_source","currency","as_of_date"]
df = df.sort_values("worker_id").reset_index(drop=True)
core, ext = df[CORE], df[EXT]

DICT = [("Workers","worker_id","Workday Worker ID (synthetic, W-1xxxxx)"),("Workers","supervisory_org","Supervisory organization (department)"),
 ("Workers","job_profile","Workday job profile"),("Workers","management_level","Job level 1-5 (higher = more senior)"),
 ("Workers","monthly_comp","Base monthly compensation in USD (location-adjusted)"),("Workers","worker_status","Active or Terminated as of the as-of date"),
 ("Workers","years_at_company","Whole years of service to as-of date, or to termination date for leavers"),
 ("Workers","job_satisfaction / environment_satisfaction / relationship_satisfaction / job_involvement / work_life_balance","Survey scores 1 (low) - 4 (high)"),
 ("Workers","performance_rating","Latest rating 1 (low) - 4 (high)"),
 ("Workers","other columns","Same meaning as in the dashboard's schema (age, gender, overtime, business_travel, etc.)"),
 ("Workday_Extended","hire_date / termination_date","Hire date; termination date (blank if active)"),
 ("Workday_Extended","termination_category / termination_reason","Voluntary vs involuntary and reason (blank if active)"),
 ("Workday_Extended","location / country","Work location and country; compensation scales with location"),
 ("Workday_Extended","manager_worker_id","Synthetic manager (level 4+ worker in the same org)"),
 ("Workday_Extended","job_family, pay_grade, cost_center, fte, time_type, hire_source, currency, as_of_date","Additional Workday-style attributes"),
 ("About","Synthetic data","All records are randomly generated (seed 42). Relationships (overtime, satisfaction, tenure, pay -> attrition) are built in, so findings are illustrative, not real."),
 ("About","Window","Includes workers active on 2026-11-14 plus leavers from 2025-01-01 to 2026-11-07.")]
OUT = "data/synthetic_workday_hr_dataset.xlsx"
with pd.ExcelWriter(OUT, engine="openpyxl", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as xw:
    core.to_excel(xw, sheet_name="Workers", index=False); ext.to_excel(xw, sheet_name="Workday_Extended", index=False)
    pd.DataFrame(DICT, columns=["Sheet","Field","Description"]).to_excel(xw, sheet_name="Data_Dictionary", index=False)
core.to_csv("data/synthetic_workday_workers.csv", index=False)

wb = load_workbook(OUT)
for ws in wb:
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="212121"); c.alignment = Alignment(wrap_text=True, vertical="center")
    for row in ws.iter_rows(min_row=2):
        for c in row: c.font = Font(name="Arial", size=10)
    for j in range(1, ws.max_column + 1):
        L = get_column_letter(j); w = max(len(str(c.value or "")) for c in ws[L][:200])
        ws.column_dimensions[L].width = min(70, max(12, w + 3))
    ws.freeze_panes = "A2"
    if ws.title != "Data_Dictionary": ws.auto_filter.ref = ws.dimensions
# Summary with live formulas
s = wb.create_sheet("Summary", 0); s.sheet_view.showGridLines = False
n = len(core) + 1
s["B2"] = "Synthetic Workday HR dataset – summary"; s["B2"].font = Font(name="Arial", size=16, bold=True, color="FF6C37")
s["B3"] = "Fictional data for demos/testing. Formulas below read the Workers / Workday_Extended sheets."; s["B3"].font = Font(name="Arial", size=9, italic=True, color="666666")
items = [("Workers (rows)", f"=COUNTA(Workers!A2:A{n})", "#,##0"),
         ("Active", f'=COUNTIFS(Workers!AD2:AD{n},"Active")', "#,##0"),
         ("Terminated", f'=COUNTIFS(Workers!AD2:AD{n},"Terminated")', "#,##0"),
         ("Attrition rate (leavers / rows)", "=IF(C5=0,0,C7/C5)", "0.0%"),
         ("Voluntary leavers", f'=COUNTIFS(Workday_Extended!D2:D{n},"Voluntary")', "#,##0"),
         ("Involuntary leavers", f'=COUNTIFS(Workday_Extended!D2:D{n},"Involuntary")', "#,##0"),
         ("Voluntary share of leavers", "=IF(C7=0,0,C9/C7)", "0.0%"),
         ("Avg monthly comp (USD)", f"=AVERAGE(Workers!M2:M{n})", "#,##0")]
for i, (a, f, fmt) in enumerate(items, 5):
    s.cell(i, 2, a).font = Font(name="Arial", bold=True); c = s.cell(i, 3, f); c.number_format = fmt; c.font = Font(name="Arial")
s.column_dimensions["A"].width = 3; s.column_dimensions["B"].width = 34; s.column_dimensions["C"].width = 16
wb.save(OUT)
print(core.shape, ext.shape, "attrition %.1f%%" % (100*(df.worker_status=="Terminated").mean()))
