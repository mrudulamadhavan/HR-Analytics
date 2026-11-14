"""SQL layer: workers table in SQLite, parameterised queries, whitelisted dimensions."""
import sqlite3
import pandas as pd

DIMS = ["supervisory_org", "job_profile", "gender", "age_band", "marital_status", "education",
        "field_of_study", "business_travel", "overtime", "management_level", "tenure_band",
        "comp_band", "job_satisfaction", "work_life_balance", "environment_satisfaction",
        "stock_option_level", "performance_rating"]
FILTERS = ["supervisory_org", "gender", "job_profile", "age_band", "education"]

class HRDB:
    def __init__(self, df: pd.DataFrame):
        self.conn = sqlite3.connect(":memory:", check_same_thread=False)
        df.to_sql("workers", self.conn, index=False, if_exists="replace")

    def q(self, sql: str, params=()):
        return pd.read_sql_query(sql, self.conn, params=params)

    @staticmethod
    def where(filters: dict):
        parts, params = [], []
        for col, vals in filters.items():
            if col in FILTERS and vals:
                parts.append(f"{col} IN ({','.join('?' * len(vals))})")
                params += list(vals)
        return (" WHERE " + " AND ".join(parts)) if parts else "", params

    def kpis(self, filters):
        w, p = self.where(filters)
        return self.q(f"""SELECT COUNT(*) AS headcount,
            COALESCE(SUM(1 - terminated),0) AS active,
            COALESCE(SUM(terminated),0) AS terminations,
            ROUND(100.0 * AVG(terminated), 2) AS attrition_rate,
            ROUND(AVG(age), 1) AS avg_age,
            ROUND(AVG(years_at_company), 1) AS avg_tenure,
            ROUND(AVG(monthly_comp), 0) AS avg_comp,
            ROUND(AVG(job_satisfaction), 2) AS avg_job_sat
            FROM workers{w}""", p).iloc[0]

    def by(self, dim, filters, order="attrition_pct DESC"):
        assert dim in DIMS, "dimension not allowed"
        assert order in ("attrition_pct DESC", "segment", "headcount DESC")
        w, p = self.where(filters)
        return self.q(f"""SELECT {dim} AS segment, COUNT(*) AS headcount,
            SUM(terminated) AS terminations,
            ROUND(100.0 * AVG(terminated), 1) AS attrition_pct,
            ROUND(AVG(monthly_comp), 0) AS avg_comp,
            ROUND(AVG(job_satisfaction), 2) AS avg_job_sat
            FROM workers{w} GROUP BY {dim} ORDER BY {order}""", p)

    def heat(self, filters):
        w, p = self.where(filters)
        return self.q(f"""SELECT supervisory_org, management_level,
            ROUND(100.0*AVG(terminated),1) AS attrition_pct, COUNT(*) AS headcount
            FROM workers{w} GROUP BY supervisory_org, management_level""", p)

    def frame(self, filters):
        w, p = self.where(filters)
        return self.q(f"SELECT * FROM workers{w}", p)

    def adhoc(self, sql: str):
        s = sql.strip().rstrip(";")
        if not s.lower().startswith(("select", "with")) or ";" in s:
            raise ValueError("Only a single SELECT / WITH query is allowed.")
        return self.q(s)
