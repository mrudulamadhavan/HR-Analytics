import os, warnings; warnings.filterwarnings("ignore")
import pandas as pd
import plotly.express as px
import streamlit as st

import workday_connector as wd
from data_layer import HRDB, FILTERS
from risk_model import fit_and_score, risk_flags
from excel_report import build_workbook

ORANGE, GREY = "#FF6C37", "#9E9E9E"
SEQ = ["#FFE0D2", "#FFA980", ORANGE, "#C94A17"]
st.set_page_config(page_title="HR Analytics Dashboard", page_icon="📊", layout="wide")
st.markdown("""<style>
.block-container{padding-top:1.4rem} h1,h2,h3{letter-spacing:-.01em}
[data-testid="stMetricValue"]{font-size:1.6rem}
</style>""", unsafe_allow_html=True)

# ------------------------------------------------------------------ data source
st.sidebar.title("Workday HR")
mode = st.sidebar.radio("Data source", ["Demo (mapped HR dataset)", "Upload Workday export"])

_syn = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "synthetic_workday_hr_dataset.xlsx")
if os.path.exists(_syn):
    with open(_syn, "rb") as _f:
        st.sidebar.download_button("⬇ Download synthetic dataset (Excel)", _f.read(), "synthetic_workday_hr_dataset.xlsx",
                                   "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   help="2,500 fictional workers. Upload it back via 'Upload Workday export' to try the dashboard.")

@st.cache_data(show_spinner="Loading data…", ttl=900)
def get_demo(): return wd.load_demo()

df, label = None, ""
try:
    if mode.startswith("Demo"):
        df, label = get_demo(), "Demo dataset (HR-analytics project, mapped to Workday field names)"
    else:
        up = st.sidebar.file_uploader("CSV / XLSX", type=["csv", "xlsx"])
        if up is None:
            st.info("Upload a Workday worker export to begin."); st.stop()
        df, label = wd.load_upload(up), f"Upload: {up.name}"
except Exception as e:
    st.error(f"Could not load data: {e}"); st.stop()

db = HRDB(df)

# ------------------------------------------------------------------ filters
st.sidebar.subheader("Filters")
filters = {}
for col, nice in [("supervisory_org", "Supervisory organization"), ("job_profile", "Job profile"),
                  ("gender", "Gender"), ("age_band", "Age band"), ("education", "Education")]:
    if col in df:
        filters[col] = st.sidebar.multiselect(nice, sorted(df[col].dropna().unique()))
st.sidebar.caption(f"Source: {label}")

k = db.kpis(filters)
if k.headcount == 0:
    st.warning("No workers match the filters."); st.stop()

st.title("People Analytics – Workforce & Attrition")
st.caption("Reporting view for HR leadership · SQL (SQLite) for aggregation · Python for modelling · Plotly visuals")

tabs = st.tabs(["Overview", "Attrition drivers", "Compensation", "Attrition risk (ML)", "SQL & export"])

def bar(d, x, y, title, horizontal=False, fmt=None, color=ORANGE):
    fig = px.bar(d, x=y if horizontal else x, y=x if horizontal else y, orientation="h" if horizontal else "v",
                 title=title, text=y, color_discrete_sequence=[color])
    fig.update_traces(texttemplate=fmt or "%{text}", textposition="outside", cliponaxis=False)
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=50, b=10), yaxis_title=None, xaxis_title=None,
                      plot_bgcolor="rgba(0,0,0,0)")
    return fig

# ------------------------------------------------------------------ Overview
with tabs[0]:
    c = st.columns(4)
    c[0].metric("Headcount", f"{int(k.headcount):,}", border=True)
    c[1].metric("Active", f"{int(k.active):,}", border=True)
    c[2].metric("Terminations", f"{int(k.terminations):,}", border=True)
    c[3].metric("Attrition rate", f"{k.attrition_rate:.1f}%", border=True)
    c = st.columns(4)
    c[0].metric("Avg age", f"{k.avg_age:.1f}", border=True)
    c[1].metric("Avg tenure (yrs)", f"{k.avg_tenure:.1f}", border=True)
    c[2].metric("Avg monthly comp", f"{k.avg_comp:,.0f}", border=True)
    c[3].metric("Avg job satisfaction (1–4)", f"{k.avg_job_sat:.2f}", border=True)

    a, b = st.columns(2)
    d = db.by("supervisory_org", filters)
    a.plotly_chart(bar(d, "segment", "attrition_pct", "Attrition rate by supervisory org (%)", fmt="%{text:.1f}%"), width="stretch")
    g = db.by("gender", filters)
    fig = px.pie(g, names="segment", values="terminations", hole=.55, title="Terminations by gender",
                 color_discrete_sequence=[ORANGE, "#212121", GREY])
    fig.update_layout(height=340, margin=dict(l=10, r=10, t=50, b=10)); b.plotly_chart(fig, width="stretch")

    a, b = st.columns(2)
    ag = db.by("age_band", filters, "segment")
    a.plotly_chart(bar(ag, "segment", "attrition_pct", "Attrition rate by age band (%)", fmt="%{text:.1f}%"), width="stretch")
    fig = px.bar(ag, x="segment", y="headcount", title="Headcount by age band", text="headcount", color_discrete_sequence=[GREY])
    fig.update_traces(textposition="outside"); fig.update_layout(height=340, margin=dict(l=10, r=10, t=50, b=10), xaxis_title=None, yaxis_title=None, plot_bgcolor="rgba(0,0,0,0)")
    b.plotly_chart(fig, width="stretch")

    jr = db.by("job_profile", filters)
    st.plotly_chart(bar(jr.sort_values("attrition_pct"), "segment", "attrition_pct", "Attrition rate by job profile (%)", horizontal=True, fmt="%{text:.1f}%"), width="stretch")
    st.caption("Note: attrition **rate** = leavers ÷ headcount in that segment. Shares of total leavers (e.g. R&D ≈ 56% of all leavers) are a different measure.")

# ------------------------------------------------------------------ Drivers
with tabs[1]:
    st.subheader("What separates leavers from stayers?")
    pairs = [("overtime", "Overtime"), ("business_travel", "Business travel"), ("job_satisfaction", "Job satisfaction (1–4)"),
             ("work_life_balance", "Work-life balance (1–4)"), ("tenure_band", "Tenure band"), ("stock_option_level", "Stock option level")]
    cols = st.columns(2)
    for i, (dim, nice) in enumerate(pairs):
        d = db.by(dim, filters, "segment")
        cols[i % 2].plotly_chart(bar(d, "segment", "attrition_pct", f"Attrition % by {nice.lower()}", fmt="%{text:.1f}%"), width="stretch")
    h = db.heat(filters)
    if not h.empty:
        piv = h.pivot(index="supervisory_org", columns="management_level", values="attrition_pct")
        fig = px.imshow(piv, text_auto=".1f", aspect="auto", color_continuous_scale=SEQ, title="Attrition % – supervisory org × management level")
        fig.update_layout(height=320, margin=dict(l=10, r=10, t=50, b=10)); st.plotly_chart(fig, width="stretch")
        st.caption("Small cells (few workers) swing widely – check headcount before acting on a single cell.")

# ------------------------------------------------------------------ Compensation
with tabs[2]:
    fdf = db.frame(filters)
    a, b = st.columns(2)
    fig = px.box(fdf, x="management_level", y="monthly_comp", color="worker_status", title="Monthly comp by management level",
                 color_discrete_map={"Active": GREY, "Terminated": ORANGE}, category_orders={"management_level": sorted(fdf.management_level.unique())})
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10)); a.plotly_chart(fig, width="stretch")
    cb = db.by("comp_band", filters, "segment")
    order = ["<3k", "3k-6k", "6k-9k", "9k-12k", "12k+"]
    cb["segment"] = pd.Categorical(cb["segment"], order, ordered=True); cb = cb.sort_values("segment")
    b.plotly_chart(bar(cb, "segment", "attrition_pct", "Attrition % by monthly comp band", fmt="%{text:.1f}%"), width="stretch")
    fig = px.scatter(fdf, x="years_at_company", y="monthly_comp", color="worker_status", opacity=.6,
                     color_discrete_map={"Active": GREY, "Terminated": ORANGE}, title="Tenure vs monthly comp")
    fig.update_layout(height=380, margin=dict(l=10, r=10, t=50, b=10)); st.plotly_chart(fig, width="stretch")

# ------------------------------------------------------------------ ML
with tabs[3]:
    @st.cache_data(show_spinner="Training model…")
    def run_model(frame): return fit_and_score(frame)
    metrics, coef, scored = run_model(df)
    st.caption("Logistic regression, class-balanced, 75/25 stratified split. Trained on all workers; scores shown for current filters.")
    m = st.columns(4)
    for col, (name, v) in zip(m, metrics.items()): col.metric(name, f"{v:.2f}", border=True)
    st.caption("Precision is modest because leavers are ~16% of workers – use scores to prioritise conversations, not as a verdict on any individual.")
    a, b = st.columns([1, 1.4])
    cf = coef.sort_values("coefficient")
    fig = px.bar(cf, x="coefficient", y="feature", orientation="h", title="Top model drivers (+ raises risk)",
                 color=(cf.coefficient > 0).map({True: "Raises risk", False: "Lowers risk"}),
                 color_discrete_map={"Raises risk": ORANGE, "Lowers risk": GREY})
    fig.update_layout(height=520, margin=dict(l=10, r=10, t=50, b=10), yaxis_title=None, legend_title=None); a.plotly_chart(fig, width="stretch")
    ids = db.frame(filters)["worker_id"]
    act = scored[(scored.terminated == 0) & scored.worker_id.isin(ids)].sort_values("risk_score", ascending=False).head(25).copy()
    act["risk_flags"] = act.apply(risk_flags, axis=1)
    show = act[["worker_id", "supervisory_org", "job_profile", "management_level", "monthly_comp", "years_at_company", "risk_score", "risk_flags"]]
    b.markdown("**Top 25 active workers by modelled attrition risk**")
    b.dataframe(show, hide_index=True, width="stretch", height=520,
                column_config={"risk_score": st.column_config.ProgressColumn("Risk %", min_value=0, max_value=100, format="%.0f")})
    st.session_state["risk_export"] = (show, metrics, coef)

# ------------------------------------------------------------------ SQL & export
with tabs[4]:
    st.subheader("SQL explorer")
    st.caption("Table `workers` · single read-only SELECT/WITH queries.")
    sql = st.text_area("Query", """SELECT supervisory_org, job_profile, COUNT(*) AS headcount,
       SUM(terminated) AS leavers, ROUND(100.0*AVG(terminated),1) AS attrition_pct
FROM workers
GROUP BY supervisory_org, job_profile
ORDER BY attrition_pct DESC""", height=140)
    if st.button("Run query"):
        try: st.dataframe(db.adhoc(sql), hide_index=True, width="stretch")
        except Exception as e: st.error(str(e))
    st.divider()
    st.subheader("Excel results workbook")
    fdf = db.frame(filters)
    risk_part = st.session_state.get("risk_export")
    xl = build_workbook(fdf, *(risk_part if risk_part else (None, None, None)), source_label=label)
    st.download_button("Download Excel (current filters)", xl, "workday_hr_results.xlsx",
                       "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", type="primary")
    st.caption("KPI and breakdown cells are live formulas over the Data sheet; risk scores are model outputs.")
