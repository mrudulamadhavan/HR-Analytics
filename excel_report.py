"""Builds the Excel results workbook. Summary tables are live formulas over the Data sheet."""
import io
import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

ORANGE, DARK, LIGHT = "FF6C37", "212121", "FFF1EA"
F = "Arial"
SEGMENTS = [("Supervisory Organization", "supervisory_org"), ("Job Profile", "job_profile"),
            ("Gender", "gender"), ("Age Band", "age_band"), ("Overtime", "overtime"),
            ("Business Travel", "business_travel"), ("Education", "education"),
            ("Marital Status", "marital_status"), ("Tenure Band", "tenure_band")]

def _hdr(ws, row, col, text):
    c = ws.cell(row=row, column=col, value=text)
    c.font = Font(name=F, bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor=DARK)
    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

def build_workbook(df: pd.DataFrame, risk: pd.DataFrame | None = None, model_metrics: dict | None = None,
                   coef: pd.DataFrame | None = None, source_label: str = "Demo dataset") -> bytes:
    wb = Workbook()
    ws_sum, ws_seg, ws_data = wb.active, wb.create_sheet("Breakdowns"), wb.create_sheet("Data")
    ws_sum.title = "Summary"
    thin = Side(style="thin", color="DDDDDD"); box = Border(left=thin, right=thin, top=thin, bottom=thin)

    # ---- Data sheet (raw rows) ----
    cols = list(df.columns)
    for j, c in enumerate(cols, 1):
        _hdr(ws_data, 1, j, c)
    for i, row in enumerate(df.itertuples(index=False), 2):
        for j, v in enumerate(row, 1):
            ws_data.cell(row=i, column=j, value=v.item() if hasattr(v, "item") else v).font = Font(name=F, size=10)
    n = len(df) + 1
    ws_data.freeze_panes = "A2"; ws_data.auto_filter.ref = f"A1:{get_column_letter(len(cols))}{n}"
    for j, c in enumerate(cols, 1):
        ws_data.column_dimensions[get_column_letter(j)].width = max(12, min(26, len(c) + 3))

    def rng(name):
        L = get_column_letter(cols.index(name) + 1)
        return f"Data!${L}$2:${L}${n}"
    status, comp, age, ten, sat = rng("worker_status"), rng("monthly_comp"), rng("age"), rng("years_at_company"), rng("job_satisfaction")

    # ---- Summary ----
    ws_sum.sheet_view.showGridLines = False
    ws_sum["B2"] = "Workday HR Analytics – Results"; ws_sum["B2"].font = Font(name=F, size=18, bold=True, color=ORANGE)
    ws_sum["B3"] = f"Source: {source_label}  |  Rows in Data sheet: {len(df):,}  |  All KPIs are live formulas over the Data sheet"
    ws_sum["B3"].font = Font(name=F, size=9, italic=True, color="666666")
    kpis = [("Headcount", f"=COUNTA({status})", "#,##0"),
            ("Active Workers", f'=COUNTIFS({status},"Active")', "#,##0"),
            ("Terminations", f'=COUNTIFS({status},"Terminated")', "#,##0"),
            ("Attrition Rate", "=IF(C6=0,0,C8/C6)", "0.0%"),
            ("Average Age", f"=AVERAGE({age})", "0.0"),
            ("Average Tenure (yrs)", f"=AVERAGE({ten})", "0.0"),
            ("Avg Monthly Comp", f"=AVERAGE({comp})", "#,##0"),
            ("Avg Job Satisfaction (1-4)", f"=AVERAGE({sat})", "0.00")]
    _hdr(ws_sum, 5, 2, "KPI"); _hdr(ws_sum, 5, 3, "Value")
    for i, (lab, fml, fmt) in enumerate(kpis, 6):
        a = ws_sum.cell(row=i, column=2, value=lab); b = ws_sum.cell(row=i, column=3, value=fml)
        a.font = Font(name=F, bold=True); b.font = Font(name=F, color="000000", bold=True)
        b.number_format = fmt; a.border = b.border = box; a.fill = PatternFill("solid", fgColor=LIGHT)
    ws_sum.column_dimensions["A"].width = 3; ws_sum.column_dimensions["B"].width = 30; ws_sum.column_dimensions["C"].width = 18

    r = 16
    ws_sum.cell(row=r, column=2, value="Notes").font = Font(name=F, bold=True)
    notes = ["Attrition Rate = Terminations / Headcount. Segment tables are on the 'Breakdowns' sheet.",
             "Monthly comp is in the source currency units (the demo dataset is USD-denominated).",
             "Demo data = HR-analytics project dataset mapped to Workday-style fields; it is not a live Workday tenant extract."]
    for k, t in enumerate(notes, r + 1):
        ws_sum.cell(row=k, column=2, value=t).font = Font(name=F, size=9, color="555555")

    # ---- Breakdowns (formulas) ----
    ws_seg.sheet_view.showGridLines = False
    row = 2
    for title, col in SEGMENTS:
        if col not in df: continue
        seg = rng(col)
        ws_seg.cell(row=row, column=1, value=title).font = Font(name=F, bold=True, size=12, color=ORANGE)
        for j, h in enumerate([title, "Headcount", "Terminations", "Attrition %", "Avg Monthly Comp", "Avg Job Sat"], 1):
            _hdr(ws_seg, row + 1, j, h)
        order = (df.groupby(col)["terminated"].mean().sort_values(ascending=False).index.tolist())
        first = row + 2
        for k, v in enumerate(order):
            rr = first + k
            ws_seg.cell(row=rr, column=1, value=str(v))
            ws_seg.cell(row=rr, column=2, value=f"=COUNTIFS({seg},$A{rr})")
            ws_seg.cell(row=rr, column=3, value=f'=COUNTIFS({seg},$A{rr},{status},"Terminated")')
            ws_seg.cell(row=rr, column=4, value=f"=IF(B{rr}=0,0,C{rr}/B{rr})").number_format = "0.0%"
            ws_seg.cell(row=rr, column=5, value=f"=IFERROR(AVERAGEIFS({comp},{seg},$A{rr}),0)").number_format = "#,##0"
            ws_seg.cell(row=rr, column=6, value=f"=IFERROR(AVERAGEIFS({sat},{seg},$A{rr}),0)").number_format = "0.00"
            for j in range(1, 7):
                c = ws_seg.cell(row=rr, column=j); c.font = Font(name=F, size=10); c.border = box
        last = first + len(order) - 1
        if col in ("supervisory_org", "age_band", "job_profile", "overtime"):
            ch = BarChart(); ch.type = "bar" if col == "job_profile" else "col"
            ch.title = f"Attrition % by {title}"; ch.legend = None; ch.roundedCorners = False
            ch.add_data(Reference(ws_seg, min_col=4, min_row=row + 1, max_row=last), titles_from_data=True)
            ch.set_categories(Reference(ws_seg, min_col=1, min_row=first, max_row=last))
            ch.x_axis.delete = False; ch.y_axis.delete = False; ch.y_axis.numFmt = "0%"
            ch.series[0].graphicalProperties.solidFill = ORANGE
            ch.height, ch.width = 7.5, 14
            ws_seg.add_chart(ch, f"H{row}")
            row = max(last + 3, row + 17)
        else:
            row = last + 3
    for j, w in enumerate([30, 12, 14, 12, 18, 12], 1):
        ws_seg.column_dimensions[get_column_letter(j)].width = w

    # ---- Risk sheet (model output = values by nature) ----
    if risk is not None:
        ws_r = wb.create_sheet("At-Risk Workers")
        heads = ["Worker ID", "Supervisory Org", "Job Profile", "Mgmt Level", "Monthly Comp", "Tenure (yrs)", "Risk Score %", "Risk Flags"]
        for j, h in enumerate(heads, 1): _hdr(ws_r, 1, j, h)
        for i, t in enumerate(risk.itertuples(index=False), 2):
            for j, v in enumerate(t, 1):
                c = ws_r.cell(row=i, column=j, value=v.item() if hasattr(v, "item") else v); c.font = Font(name=F, size=10)
        for j, w in enumerate([12, 18, 26, 12, 14, 12, 12, 60], 1): ws_r.column_dimensions[get_column_letter(j)].width = w
        ws_r.freeze_panes = "A2"
        k = len(risk) + 4
        ws_r.cell(row=k, column=1, value="Model: logistic regression (class-balanced), scores are model outputs computed in Python, not formulas.").font = Font(name=F, size=9, italic=True)
        if model_metrics:
            for m, (name, val) in enumerate(model_metrics.items(), k + 1):
                ws_r.cell(row=m, column=1, value=name).font = Font(name=F, size=10, bold=True)
                c = ws_r.cell(row=m, column=3, value=round(float(val), 3)); c.font = Font(name=F, size=10); c.number_format = "0.000"
    if coef is not None:
        ws_c = wb.create_sheet("Model Drivers")
        _hdr(ws_c, 1, 1, "Feature"); _hdr(ws_c, 1, 2, "Coefficient (standardised)")
        for i, t in enumerate(coef.itertuples(index=False), 2):
            ws_c.cell(row=i, column=1, value=t[0]).font = Font(name=F, size=10)
            c = ws_c.cell(row=i, column=2, value=round(float(t[1]), 3)); c.font = Font(name=F, size=10)
        ws_c.column_dimensions["A"].width = 38; ws_c.column_dimensions["B"].width = 26
        ws_c["D1"] = "Positive = raises attrition risk; negative = lowers it."; ws_c["D1"].font = Font(name=F, size=9, italic=True)

    buf = io.BytesIO(); wb.save(buf); return buf.getvalue()
