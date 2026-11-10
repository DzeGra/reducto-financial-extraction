"""
Builds the LP earnings workbook from a Reducto Extract result.

Input : amazon_raw, the Reducto extraction payload (each field = {"value": ..., "citations": [...]})
Output: an Excel workbook with the tabs the LP brief asks for
        (Financials, Segments, Cash Flow, Guidance, Notes)
        plus KPIs, Disclosures, Validation and Extraction Audit.

Every number in the workbook comes from the extraction. Ratios and growth rates are
Excel formulas, so an analyst can click any cell and see how it was calculated.
"""

import requests
from datetime import date
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 1. Reading the Reducto result
# ---------------------------------------------------------------------------

def load_payload(result):
    """Turn a Reducto extract response or job into the field dictionary.

    Reducto returns small results inline and large results as a URL.
    This walks the response until it finds the extracted fields.
    """
    data = result.model_dump() if hasattr(result, "model_dump") else result
    for _ in range(5):
        if isinstance(data, list) and len(data) == 1:
            data = data[0]
        if isinstance(data, dict) and "q4_2024_total_net_sales" in data:
            return data
        if isinstance(data, dict) and isinstance(data.get("url"), str):
            response = requests.get(data["url"], timeout=120)
            response.raise_for_status()
            data = response.json()
        elif isinstance(data, dict) and "result" in data:
            data = data["result"]
    raise ValueError("Could not find extracted fields in the Reducto response")


def unwrap(item):
    """Return the extracted value without its citation wrapper."""
    return item["value"] if isinstance(item, dict) and "value" in item else item


def unwrap_record(rec, key_field):
    """List items have their own 'value' property, so only unwrap a wrapper around the record."""
    if isinstance(rec, dict) and key_field not in rec and isinstance(rec.get("value"), dict):
        return rec["value"]
    return rec


def values_from(raw):
    """Field ID -> plain value, ready for calculations."""
    return {key: unwrap(item) for key, item in raw.items()}


def first_citation(item):
    cites = item.get("citations") if isinstance(item, dict) else None
    return cites[0] if cites else {}


def source_page(item):
    return first_citation(item).get("bbox", {}).get("page")


def confidence(item):
    cite = first_citation(item)
    score = (cite.get("granular_confidence") or {}).get("extract_confidence")
    level = item.get("confidence") if isinstance(item, dict) else None
    return score, level


def kpi_records(raw):
    """Flatten the KPI list into simple dictionaries."""
    records = []
    for rec in unwrap(raw.get("operating_metrics_and_kpis")) or []:
        rec = unwrap_record(rec, "metric")
        if not isinstance(rec, dict):
            continue
        records.append({
            "metric": unwrap(rec.get("metric")),
            "details": unwrap(rec.get("details")),
            "value": unwrap(rec.get("value")),
            "unit": unwrap(rec.get("unit")),
            "period": unwrap(rec.get("period")),
            "page": source_page(rec.get("value")) or source_page(rec.get("metric")),
        })
    return records


QUARTERS = ["Q3 2023", "Q4 2023", "Q1 2024", "Q2 2024", "Q3 2024", "Q4 2024"]


def kpi_series(records, metric_start):
    """Six-quarter series for one supplemental metric (TTM or quarterly)."""
    series = {}
    for r in records:
        if r["metric"] and r["metric"].startswith(metric_start):
            quarter = (r["period"] or "").replace("TTM as of ", "")
            if quarter in QUARTERS:
                series[quarter] = r["value"]
    return [series.get(q) for q in QUARTERS]


# ---------------------------------------------------------------------------
# 2. Readable names (column A) alongside field IDs (kept for the audit trail)
# ---------------------------------------------------------------------------

PERIOD_LABELS = {"q4_2023": "Q4 2023", "q4_2024": "Q4 2024", "fy2023": "FY2023", "fy2024": "FY2024",
                 "dec2023": "Dec 31, 2023", "dec2024": "Dec 31, 2024", "q1_2025": "Q1 2025", "q1_2024": "Q1 2024",
                 "fy2025": "FY2025"}

NAME_LABELS = {
    "total_net_sales": "Net sales", "operating_income": "Operating income", "net_income": "Net income",
    "basic_eps": "Basic EPS", "diluted_eps": "Diluted EPS", "operating_cash_flow": "Operating cash flow",
    "stock_based_compensation": "Stock-based compensation",
    "north_america_net_sales": "North America net sales", "international_net_sales": "International net sales",
    "aws_net_sales": "AWS net sales", "north_america_operating_income": "North America operating income",
    "international_operating_income": "International operating income (loss)",
    "aws_operating_income": "AWS operating income", "free_cash_flow_ttm": "Free cash flow (TTM)",
    "purchases_property_equipment": "Capital expenditures (purchases of property & equipment)",
    "net_purchases_property_equipment": "Capex net of proceeds & incentives (TTM)",
    "finance_lease_equipment_acquired": "Equipment acquired under finance leases (TTM)",
    "finance_lease_principal_repayments": "Principal repayments of finance leases",
    "long_term_debt_repayments": "Repayments of long-term debt",
    "long_term_debt_proceeds": "Proceeds from long-term debt",
    "financing_obligations_principal_repayments": "Principal repayments of financing obligations",
    "purchases_marketable_securities": "Purchases of marketable securities",
    "sales_maturities_marketable_securities": "Sales & maturities of marketable securities",
    "cash_cash_equivalents": "Cash and cash equivalents", "marketable_securities": "Marketable securities",
    "long_term_debt": "Long-term debt", "current_portion_long_term_debt": "Current portion of long-term debt",
    "guidance_net_sales_low": "Net sales guidance (low)", "guidance_net_sales_high": "Net sales guidance (high)",
    "guidance_sales_yoy_growth_low_pct": "Net sales growth guidance (low)",
    "guidance_sales_yoy_growth_high_pct": "Net sales growth guidance (high)",
    "guidance_operating_income_low": "Operating income guidance (low)",
    "guidance_operating_income_high": "Operating income guidance (high)",
    "guidance_fx_impact_usd": "FX impact on net sales guidance",
    "guidance_fx_impact_basis_points": "FX impact on net sales guidance (bps)",
    "operating_income_guidance_comparison": "Operating income (comparison in guidance)",
    "leap_year_net_sales_impact": "Leap Year net sales benefit",
    "guidance_assumptions": "Guidance assumptions", "guidance_margin_commentary": "Margin commentary",
    "expected_capital_investment_commentary": "Capital investment commentary",
    "strategic_partnerships_and_investments": "Strategic partnerships & investments",
    "other_material_business_developments": "Other business developments",
    "operating_metrics_and_kpis": "Operating metrics & KPIs",
}


def label(field_id):
    for prefix, period in PERIOD_LABELS.items():
        if field_id.startswith(prefix + "_"):
            name = field_id[len(prefix) + 1:]
            return f"{NAME_LABELS.get(name, name.replace('_', ' ').capitalize())} — {period}"
    return NAME_LABELS.get(field_id, field_id.replace("_", " ").capitalize())


# ---------------------------------------------------------------------------
# 3. Styling helpers
# ---------------------------------------------------------------------------

NAVY, LIGHT, GREY = "17365D", "DCE6F1", "6B7280"
FMT_M, FMT_EPS, FMT_PCT, FMT_PP = '#,##0;(#,##0)', '$0.00', '0.0%;(0.0%)', '+0.0" pp";-0.0" pp"'


class Sheet:
    """Small wrapper that writes rows top to bottom and remembers where things are."""

    def __init__(self, wb, title, heading, subtitle, widths):
        self.ws = wb.create_sheet(title)
        self.ws.sheet_view.showGridLines = False
        self.ws["A1"] = heading
        self.ws["A1"].font = Font(size=15, bold=True, color=NAVY)
        self.ws["A2"] = subtitle
        self.ws["A2"].font = Font(italic=True, color=GREY)
        for i, w in enumerate(widths, 1):
            self.ws.column_dimensions[get_column_letter(i)].width = w
        self.row = 4

    def section(self, text):
        cell = self.ws.cell(self.row, 1, text)
        cell.font = Font(bold=True, size=12, color=NAVY)
        self.row += 1

    def header(self, labels):
        for i, text in enumerate(labels, 1):
            c = self.ws.cell(self.row, i, text)
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor=NAVY)
            c.alignment = Alignment(horizontal="center" if i > 1 else "left", wrap_text=True, vertical="center")
        self.row += 1
        return self.row - 1

    def add(self, values, formats=None, bold=False):
        for i, v in enumerate(values, 1):
            c = self.ws.cell(self.row, i, v)
            if formats and i - 1 < len(formats) and formats[i - 1]:
                c.number_format = formats[i - 1]
            if bold:
                c.font = Font(bold=True)
                c.fill = PatternFill("solid", fgColor=LIGHT)
        self.row += 1
        return self.row - 1

    def gap(self, n=1):
        self.row += n


def show_axes(chart):
    """openpyxl 3.1 hides chart axes in Excel unless told otherwise."""
    chart.x_axis.delete = False
    chart.y_axis.delete = False
    return chart


def growth(new, old):
    return f'=IFERROR({new}/{old}-1,"")'


def omit_row(*vals):
    """Brief: data missing from the source document is left out."""
    return all(v is None for v in vals)


# ---------------------------------------------------------------------------
# 4. Tabs
# ---------------------------------------------------------------------------

def financials_tab(wb, v, raw):
    s = Sheet(wb, "Financials", "Amazon.com, Inc. | Financial Overview",
              "USD millions except per-share data · Source: Q4 2024 Earnings Release (Feb 6, 2025)",
              [44, 13, 13, 12, 13, 13, 12, 9, 38])
    cols = ["Metric", "Q4 2023", "Q4 2024", "Q4 YoY", "FY2023", "FY2024", "FY YoY", "Page", "Field ID"]
    s.section("Income statement")
    s.header(cols)
    rows = {}
    for key, fmt in [("total_net_sales", FMT_M), ("operating_income", FMT_M), (None, None),
                     ("net_income", FMT_M), ("basic_eps", FMT_EPS), ("diluted_eps", FMT_EPS)]:
        if key is None:  # operating margin, calculated
            ns, oi = rows["total_net_sales"], rows["operating_income"]
            s.add(["Operating margin (calculated)", f"=B{oi}/B{ns}", f"=C{oi}/C{ns}", f"=(C{s.row}-B{s.row})*100",
                   f"=E{oi}/E{ns}", f"=F{oi}/F{ns}", f"=(F{s.row}-E{s.row})*100", None, "calculated"],
                  [None, FMT_PCT, FMT_PCT, FMT_PP, FMT_PCT, FMT_PCT, FMT_PP])
            continue
        r = s.row
        rows[key] = r
        s.add([NAME_LABELS[key], v.get(f"q4_2023_{key}"), v.get(f"q4_2024_{key}"), growth(f"C{r}", f"B{r}"),
               v.get(f"fy2023_{key}"), v.get(f"fy2024_{key}"), growth(f"F{r}", f"E{r}"),
               source_page(raw.get(f"fy2024_{key}")), f"*_{key}"],
              [None, fmt, fmt, FMT_PCT, fmt, fmt, FMT_PCT])

    s.gap()
    s.section("Cash flow")
    s.header(cols)
    for key in ["operating_cash_flow", "purchases_property_equipment", "net_purchases_property_equipment",
                "free_cash_flow_ttm", "stock_based_compensation"]:
        vals = [v.get(f"q4_2023_{key}"), v.get(f"q4_2024_{key}"), v.get(f"fy2023_{key}"), v.get(f"fy2024_{key}")]
        if omit_row(*vals):
            continue
        r = s.row
        q_growth = growth(f"C{r}", f"B{r}") if vals[0] and vals[1] else None
        s.add([NAME_LABELS[key], vals[0], vals[1], q_growth, vals[2], vals[3], growth(f"F{r}", f"E{r}"),
               source_page(raw.get(f"fy2024_{key}")), f"*_{key}"],
              [None, FMT_M, FMT_M, FMT_PCT, FMT_M, FMT_M, FMT_PCT])
    s.ws.freeze_panes = "B4"
    return rows


def segments_tab(wb, v, raw):
    s = Sheet(wb, "Segments", "Amazon.com, Inc. | Segment Performance",
              "USD millions · Source: Segment Information, Q4 2024 Earnings Release p.8",
              [26, 13, 13, 12, 13, 13, 12, 13, 13, 13, 12])
    segs = [("North America", "north_america"), ("International", "international"), ("AWS", "aws")]

    s.section("Net sales by segment")
    s.header(["Segment", "Q4 2023", "Q4 2024", "Q4 YoY", "FY2023", "FY2024", "FY YoY"])
    first = s.row
    for name, key in segs:
        r = s.row
        s.add([name, v[f"q4_2023_{key}_net_sales"], v[f"q4_2024_{key}_net_sales"], growth(f"C{r}", f"B{r}"),
               v[f"fy2023_{key}_net_sales"], v[f"fy2024_{key}_net_sales"], growth(f"F{r}", f"E{r}")],
              [None, FMT_M, FMT_M, FMT_PCT, FMT_M, FMT_M, FMT_PCT])
    last = s.row - 1
    tot = s.row
    s.add(["Total"] + [f"=SUM({c}{first}:{c}{last})" if c in "BCEF" else growth(f"{'C' if c == 'D' else 'F'}{tot}", f"{'B' if c == 'D' else 'E'}{tot}")
                       for c in "BCDEFG"],
          [None, FMT_M, FMT_M, FMT_PCT, FMT_M, FMT_M, FMT_PCT], bold=True)
    sales_first, sales_last, sales_total = first, last, tot

    s.gap()
    s.section("Operating income and margin by segment")
    s.header(["Segment", "Q4 2023 OI", "Q4 2024 OI", "Q4 2023 margin", "Q4 2024 margin",
              "FY2023 OI", "FY2024 OI", "FY2023 margin", "FY2024 margin", "FY margin change"])
    first = s.row
    for i, (name, key) in enumerate(segs):
        r, sr = s.row, sales_first + i
        s.add([name, v[f"q4_2023_{key}_operating_income"], v[f"q4_2024_{key}_operating_income"],
               f"=B{r}/B{sr}", f"=C{r}/C{sr}", v[f"fy2023_{key}_operating_income"], v[f"fy2024_{key}_operating_income"],
               f"=F{r}/E{sr}", f"=G{r}/F{sr}", f"=(I{r}-H{r})*100"],
              [None, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_PP])
    margin_first, margin_last = first, s.row - 1
    r = s.row
    s.add(["Total", f"=SUM(B{first}:B{r-1})", f"=SUM(C{first}:C{r-1})", f"=B{r}/B{sales_total}", f"=C{r}/C{sales_total}",
           f"=SUM(F{first}:F{r-1})", f"=SUM(G{first}:G{r-1})", f"=F{r}/E{sales_total}", f"=G{r}/F{sales_total}", f"=(I{r}-H{r})*100"],
          [None, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_PP], bold=True)

    s.gap()
    s.section("Revenue mix (share of total net sales)")
    s.header(["Segment", "Q4 2023", "Q4 2024", "FY2023", "FY2024"])
    for i, (name, _) in enumerate(segs):
        sr = sales_first + i
        s.add([name, f"=B{sr}/B${sales_total}", f"=C{sr}/C${sales_total}", f"=E{sr}/E${sales_total}", f"=F{sr}/F${sales_total}"],
              [None, FMT_PCT, FMT_PCT, FMT_PCT, FMT_PCT])

    # Chart 1: FY net sales by segment
    ws = s.ws
    c1 = BarChart()
    c1.title, c1.y_axis.title, c1.height, c1.width = "Net sales by segment (USD M)", "USD millions", 7.5, 15
    c1.add_data(Reference(ws, min_col=5, max_col=6, min_row=sales_first - 1, max_row=sales_last), titles_from_data=True)
    c1.set_categories(Reference(ws, min_col=1, min_row=sales_first, max_row=sales_last))
    ws.add_chart(show_axes(c1), "M4")

    # Chart 2: FY operating margin by segment
    c2 = BarChart()
    c2.title, c2.height, c2.width = "Operating margin by segment", 7.5, 15
    c2.y_axis.number_format = "0%"
    c2.add_data(Reference(ws, min_col=8, max_col=9, min_row=margin_first - 1, max_row=margin_last), titles_from_data=True)
    c2.set_categories(Reference(ws, min_col=1, min_row=margin_first, max_row=margin_last))
    ws.add_chart(show_axes(c2), "M20")
    ws.freeze_panes = "B4"


def cash_flow_tab(wb, v, raw, kpis):
    s = Sheet(wb, "Cash Flow", "Amazon.com, Inc. | Cash Flow, Debt Service and Liquidity",
              "USD millions · Cash outflows shown as positive amounts · Source: Q4 2024 Earnings Release pp.5, 9, 10",
              [52, 13, 13, 12, 13, 13, 12, 9])

    def block(title, keys, prefixes=("fy2023", "fy2024"), heads=("FY2023", "FY2024")):
        s.section(title)
        s.header(["Metric", *heads, "YoY", "Page"])
        start = s.row
        for key in keys:
            a, b = v.get(f"{prefixes[0]}_{key}"), v.get(f"{prefixes[1]}_{key}")
            if omit_row(a, b):
                continue
            r = s.row
            s.add([NAME_LABELS[key], a, b, growth(f"C{r}", f"B{r}") if a and b else None,
                   source_page(raw.get(f"{prefixes[1]}_{key}"))], [None, FMT_M, FMT_M, FMT_PCT])
        return start, s.row - 1

    block("Cash flow highlights (twelve months ended Dec 31)",
          ["operating_cash_flow", "purchases_property_equipment", "net_purchases_property_equipment",
           "free_cash_flow_ttm", "stock_based_compensation", "purchases_marketable_securities",
           "sales_maturities_marketable_securities"])
    s.gap()
    a, b = block("Debt repayments and lease obligations (twelve months ended Dec 31)",
                 ["long_term_debt_repayments", "finance_lease_principal_repayments",
                  "financing_obligations_principal_repayments", "finance_lease_equipment_acquired",
                  "long_term_debt_proceeds"])
    r = s.row
    s.add(["Total debt and lease repayments (calculated)", f"=SUM(B{a}:B{a+2})", f"=SUM(C{a}:C{a+2})",
           growth(f"C{r}", f"B{r}")], [None, FMT_M, FMT_M, FMT_PCT], bold=True)
    s.gap()
    a, b = block("Available cash and marketable securities (balance sheet)",
                 ["cash_cash_equivalents", "marketable_securities", "long_term_debt", "current_portion_long_term_debt"],
                 prefixes=("dec2023", "dec2024"), heads=("Dec 31, 2023", "Dec 31, 2024"))
    r = s.row
    s.add(["Total cash and marketable securities (calculated)", f"=B{a}+B{a+1}", f"=C{a}+C{a+1}",
           growth(f"C{r}", f"B{r}")], [None, FMT_M, FMT_M, FMT_PCT], bold=True)

    s.gap()
    s.section("Trailing-twelve-month trend (Supplemental Financial Information p.10)")
    head = s.header(["Metric (TTM)", *QUARTERS])
    for name, metric in [("Operating cash flow", "Operating cash flow -- trailing"),
                         ("Capex net of proceeds & incentives", "Purchases of property and equipment, net"),
                         ("Free cash flow", "Free cash flow -- TTM")]:
        s.add([name, *kpi_series(kpis, metric)], [None] + [FMT_M] * 6)
    ws = s.ws
    chart = LineChart()
    chart.title, chart.y_axis.title, chart.height, chart.width = "Cash flow vs. capex, TTM (USD M)", "USD millions", 8, 18
    chart.add_data(Reference(ws, min_col=1, max_col=7, min_row=head + 1, max_row=head + 3), from_rows=True, titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=2, max_col=7, min_row=head))
    for series in chart.series:
        series.smooth = False
    chart.width = 22
    ws.add_chart(show_axes(chart), f"A{s.row + 1}")
    ws.freeze_panes = "B4"


def guidance_tab(wb, v, raw):
    s = Sheet(wb, "Guidance", "Amazon.com, Inc. | Forward Guidance",
              "First Quarter 2025 guidance as issued Feb 6, 2025 · Source: Q4 2024 Earnings Release p.3",
              [44, 13, 13, 14, 40, 40, 9])
    s.section("Quantitative guidance — Q1 2025")
    s.header(["Metric", "Low", "High", "Unit", "Field ID (low)", "Field ID (high)", "Page"])
    for name, key, unit, scale, fmt in [
        ("Net sales", "net_sales", "USD millions", 1, FMT_M),
        ("Net sales growth vs. Q1 2024", "sales_yoy_growth", "%", 0.01, FMT_PCT),
        ("Operating income", "operating_income", "USD millions", 1, FMT_M),
    ]:
        lo_id = f"q1_2025_guidance_{key}_low" + ("_pct" if key == "sales_yoy_growth" else "")
        hi_id = f"q1_2025_guidance_{key}_high" + ("_pct" if key == "sales_yoy_growth" else "")
        lo, hi = v.get(lo_id), v.get(hi_id)
        if omit_row(lo, hi):
            continue
        s.add([name, lo * scale if lo is not None else None, hi * scale if hi is not None else None, unit,
               lo_id, hi_id, source_page(raw.get(lo_id))], [None, fmt, fmt])

    s.gap()
    s.section("FX, comparisons and assumptions")
    s.header(["Metric", "Value", "", "Unit", "Field ID", "", "Page"])
    for name, key, unit, fmt in [
        ("FX impact on Q1 2025 net sales", "q1_2025_guidance_fx_impact_usd", "USD millions", FMT_M),
        ("FX impact on Q1 2025 net sales", "q1_2025_guidance_fx_impact_basis_points", "basis points", "#,##0;(#,##0)"),
        ("Q1 2024 operating income (comparison base)", "q1_2024_operating_income_guidance_comparison", "USD millions", FMT_M),
        ("Leap Year net sales benefit in Q1 2024", "q1_2024_leap_year_net_sales_impact", "USD millions", FMT_M),
        ("Margin commentary", "q1_2025_guidance_margin_commentary", "text", None),
        ("Capital investment commentary (2025)", "fy2025_expected_capital_investment_commentary", "text", None),
    ]:
        if v.get(key) is None:
            continue
        s.add([name, v[key], None, unit, key, None, source_page(raw.get(key))], [None, fmt])
    if v.get("q1_2025_guidance_assumptions"):
        r = s.add(["Guidance assumes", v["q1_2025_guidance_assumptions"], None, "text",
                   "q1_2025_guidance_assumptions", None, source_page(raw.get("q1_2025_guidance_assumptions"))])
        s.ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=3)
        s.ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical="top")
        s.ws.row_dimensions[r].height = 45
    s.ws.freeze_panes = "B4"


SECTION_RULES = [
    ("North America", lambda m: m.startswith("North America Segment")),
    ("International", lambda m: m.startswith("International Segment")),
    ("AWS segment", lambda m: m.startswith("AWS Segment")),
    ("Net sales by type", lambda m: m.split(" --")[0].split(" (")[0] in
     {"Online stores", "Physical stores", "Third-party seller services", "Advertising services",
      "Subscription services", "AWS", "Other"}),
    ("Operating expenses", lambda m: m.startswith(("Cost of sales", "Fulfillment", "Technology and infrastructure",
                                                   "Sales and marketing", "General and administrative",
                                                   "Total stock-based compensation"))),
    ("Operational KPIs", lambda m: m.startswith(("WW shipping", "WW paid units", "WW seller unit", "Employees"))),
    ("Cash flows and shares", lambda m: m.startswith(("Operating cash flow", "Purchases of property", "Principal",
                                                      "Equipment acquired", "Free cash flow", "Common shares",
                                                      "Stock-based awards"))),
    ("Results of operations", lambda m: True),
]


def kpis_tab(wb, kpis):
    s = Sheet(wb, "KPIs", "Amazon.com, Inc. | Key Business Metrics",
              "Source: Q4 2024 Earnings Release highlights (p.2) and Supplemental Financial Information (pp.10–13)",
              [22, 60, 14] + [12] * 6 + [8])
    highlights = [r for r in kpis if r["page"] == 2]  # text highlights on p.2; tables start on p.10
    s.section("Highlights since the last earnings announcement")
    s.header(["Period", "Metric", "Value", "Unit"])
    for r in highlights:
        s.add([r["period"], r["metric"], r["value"], r["unit"]], [None, None, "#,##0.0##"])

    s.gap()
    s.section("Supplemental metrics by quarter")
    head = s.header(["Section", "Metric", "Unit", *QUARTERS, "Page"])
    table, order = {}, []
    for r in kpis:
        if r["page"] == 2 or not r["metric"]:
            continue
        quarter = (r["period"] or "").replace("TTM as of ", "")
        if quarter not in QUARTERS:
            continue
        if r["metric"] not in table:
            table[r["metric"]] = {"unit": r["unit"], "page": r["page"]}
            order.append(r["metric"])
        table[r["metric"]][quarter] = r["value"]
    sbc_row = None
    for metric in sorted(order, key=lambda m: [i for i, (_, rule) in enumerate(SECTION_RULES) if rule(m)][0]):
        section = next(name for name, rule in SECTION_RULES if rule(metric))
        row = table[metric]
        unit = row["unit"] or ""
        fmt = '0.0"%";(0.0"%")' if "percent" in unit.lower() or unit == "%" else ("$0.00" if "per share" in metric.lower() else "#,##0;(#,##0)")
        shown = metric.split(": — ")[-1]
        r = s.add([section, shown, "%" if fmt.startswith("0.0") else unit, *[row.get(q) for q in QUARTERS], row["page"]],
                  [None, None, None] + [fmt] * 6)
        if metric.startswith("Total stock-based compensation"):
            sbc_row = r

    if sbc_row:  # Chart 4: stock-based compensation trend
        ws = s.ws
        chart = BarChart()
        chart.title, chart.y_axis.title, chart.height, chart.width = "Stock-based compensation by quarter (USD M)", "USD millions", 7.5, 15
        chart.legend = None
        chart.add_data(Reference(ws, min_col=4, max_col=9, min_row=sbc_row), from_rows=True)
        chart.set_categories(Reference(ws, min_col=4, max_col=9, min_row=head))
        ws.add_chart(show_axes(chart), "L4")
    s.ws.freeze_panes = "C4"


def disclosures_tab(wb, raw):
    s = Sheet(wb, "Disclosures", "Amazon.com, Inc. | Partnerships, Contract Wins and Business Developments",
              "Qualitative items as stated in the release · Source: Q4 2024 Earnings Release pp.2–3",
              [30, 30, 80, 12, 16, 8, 12])
    s.header(["Category", "Topic", "Detail", "Value", "Unit", "Page", "Confidence"])
    for field in ["strategic_partnerships_and_investments", "other_material_business_developments"]:
        for rec in unwrap(raw.get(field)) or []:
            rec = unwrap_record(rec, "topic")
            score, level = confidence(rec.get("details") or {})
            r = s.add([NAME_LABELS[field], unwrap(rec.get("topic")), unwrap(rec.get("details")),
                       unwrap(rec.get("value")), unwrap(rec.get("unit")), source_page(rec.get("topic")),
                       score if score is not None else level], [None, None, None, "#,##0", None, None, "0%"])
            s.ws.cell(r, 3).alignment = Alignment(wrap_text=True, vertical="top")
    s.ws.freeze_panes = "A5"


def validation_checks(v, kpis):
    """Reconcile extracted figures against each other. Returns (check, calculated, reported)."""
    def seg_sum(period, metric):
        return sum(v[f"{period}_{s}_{metric}"] for s in ("north_america", "international", "aws"))

    q4 = lambda metric: kpi_series(kpis, metric)[-1]
    return [
        ("Q4 2024 segment net sales = total net sales", seg_sum("q4_2024", "net_sales"), v["q4_2024_total_net_sales"]),
        ("Q4 2023 segment net sales = total net sales", seg_sum("q4_2023", "net_sales"), v["q4_2023_total_net_sales"]),
        ("FY2024 segment net sales = total net sales", seg_sum("fy2024", "net_sales"), v["fy2024_total_net_sales"]),
        ("FY2023 segment net sales = total net sales", seg_sum("fy2023", "net_sales"), v["fy2023_total_net_sales"]),
        ("Q4 2024 segment operating income = total", seg_sum("q4_2024", "operating_income"), v["q4_2024_operating_income"]),
        ("FY2024 segment operating income = total", seg_sum("fy2024", "operating_income"), v["fy2024_operating_income"]),
        ("FY2024 FCF = operating cash flow − net capex", v["fy2024_operating_cash_flow"] - v["fy2024_net_purchases_property_equipment"], v["fy2024_free_cash_flow_ttm"]),
        ("FY2023 FCF = operating cash flow − net capex", v["fy2023_operating_cash_flow"] - v["fy2023_net_purchases_property_equipment"], v["fy2023_free_cash_flow_ttm"]),
        ("Q4 2024 net sales: supplemental table (p.10) = income statement (p.6)", q4("Worldwide (WW) net sales"), v["q4_2024_total_net_sales"]),
        ("FY2024 net income: TTM table (p.10) = income statement (p.6)", q4("Net income -- TTM"), v["fy2024_net_income"]),
        ("FY2024 operating cash flow: TTM table (p.10) = cash flow statement (p.5)", q4("Operating cash flow -- trailing"), v["fy2024_operating_cash_flow"]),
    ]


def validation_tab(wb, checks):
    s = Sheet(wb, "Validation", "Extraction Validation",
              "Each check reconciles extracted numbers against other numbers in the same document",
              [70, 15, 15, 12, 10])
    s.header(["Check", "Calculated", "Reported", "Difference", "Result"])
    for name, calc, reported in checks:
        r = s.row
        s.add([name, calc, reported, f"=B{r}-C{r}", f'=IF(ABS(D{r})<0.5,"PASS","FAIL")'], [None, FMT_M, FMT_M, FMT_M])
    r = s.row
    s.add([f"Checks passed", f'=COUNTIF(E5:E{r-1},"PASS")&" of {len(checks)}"'], bold=True)
    s.ws.freeze_panes = "A5"


def audit_tab(wb, raw, v):
    s = Sheet(wb, "Extraction Audit", "Reducto Extraction Audit",
              "Every extracted field with its source page, confidence and cited text (bounding boxes viewable in Reducto Studio)",
              [46, 52, 14, 12, 8, 12, 70])
    s.header(["Field ID", "Readable name", "Value", "Status", "Page", "Confidence", "Cited source text"])
    for key, item in raw.items():
        value = v.get(key)
        shown = f"{len(value)} records" if isinstance(value, list) else value
        score, level = confidence(item)
        s.add([key, label(key), shown, "Not in source" if value is None else "Extracted", source_page(item),
               score if score is not None else level, (first_citation(item).get("content") or "")[:300]],
              [None, None, "#,##0.##", None, None, "0%"])
    s.ws.freeze_panes = "B5"


def notes_tab(wb, v, raw, meta, n_checks_passed, n_checks):
    s = Sheet(wb, "Notes", "Notes, Assumptions and Adjustments", "Read with the Financials, Segments, Cash Flow and Guidance tabs",
              [34, 110])
    wrap = Alignment(wrap_text=True, vertical="top")

    def note(topic, text):
        r = s.add([topic, text])
        s.ws.cell(r, 2).alignment = wrap
        s.ws.cell(r, 1).alignment = wrap

    s.section("Source and method")
    s.header(["Item", "Detail"])
    note("Company", meta.get("company", "Amazon.com, Inc. (NASDAQ: AMZN)"))
    note("Source document", meta.get("source", "Q4 2024 Earnings Release, dated Feb 6, 2025 (14 pages)"))
    note("Extraction", f"Reducto Extract API with a {meta.get('n_fields', len(raw))}-field JSON schema and citations enabled. Job ID: {meta.get('job_id', 'n/a')}.")
    note("Workbook generated", meta.get("run_date", date.today().isoformat()))
    note("Validation", f"{n_checks_passed} of {n_checks} reconciliation checks passed (see Validation tab).")

    s.gap()
    s.section("Units and conventions")
    s.header(["Item", "Detail"])
    note("Units", "USD millions, except per-share amounts (USD) and items labeled otherwise.")
    note("Fiscal year", "Amazon's fiscal year is the calendar year. FY2024 = twelve months ended Dec 31, 2024.")
    note("Cash outflows", "Shown as positive amounts (the source shows them in parentheses).")
    note("Operating losses", "Shown as negative amounts (International segment, 2023).")
    note("Guidance", "Guidance amounts reported in billions in the source are converted to USD millions.")

    s.gap()
    s.section("Calculated fields (Excel formulas)")
    s.header(["Item", "Detail"])
    note("YoY growth", "(Current period ÷ prior-year period) − 1, calculated from extracted values.")
    note("Operating margin", "Operating income ÷ net sales. Margin changes are in percentage points (pp).")
    note("Revenue mix", "Segment net sales ÷ total net sales. Matches the company-reported mix on p.8.")
    note("Free cash flow", "As defined by Amazon: operating cash flow less purchases of property and equipment net of proceeds from sales and incentives, trailing twelve months. At Q4, TTM equals the fiscal year.")
    note("Total debt and lease repayments", "Sum of repayments of long-term debt, principal repayments of finance leases and principal repayments of financing obligations.")
    note("Total cash and marketable securities", "Cash and cash equivalents plus marketable securities.")

    s.gap()
    s.section("Omitted items (not disclosed in the source document)")
    s.header(["Field", "Detail"])
    for key, item in raw.items():
        if v.get(key) is None:
            note(label(key), f"Field '{key}' returned no value; the release does not report it.")
    note("Q4 capex and quarterly FCF", "Capex was extracted for the twelve-month periods; free cash flow is reported only on a TTM basis.")

    s.gap()
    s.section("Metrics disclosed as growth rates only")
    s.header(["Metric", "Detail"])
    note("U.S. Prime same-day/overnight items", "+65% vs. Q4 2023. Base volume undisclosed, so the absolute change is not derivable.")
    note("Kindle devices sold", "+30% YoY in Q4 2024. Unit volume undisclosed, so the absolute change is not derivable.")
    note("Thursday Night Football viewers", "13.2M average, +11% vs. 2023. Implied 2023 average ≈ 11.9M (13.2 ÷ 1.11).")
    note("WW paid units", "Reported as Y/Y growth only (KPIs tab). Unit volumes undisclosed.")


# ---------------------------------------------------------------------------
# 5. Main entry point
# ---------------------------------------------------------------------------

def build_workbook(amazon_raw, path="Amazon_Q4_2024_LP_Analysis.xlsx", meta=None):
    meta = meta or {}
    v = values_from(amazon_raw)
    kpis = kpi_records(amazon_raw)
    checks = validation_checks(v, kpis)
    passed = sum(abs(c - r) < 0.5 for _, c, r in checks)

    wb = Workbook()
    wb.remove(wb.active)
    financials_tab(wb, v, amazon_raw)
    segments_tab(wb, v, amazon_raw)
    cash_flow_tab(wb, v, amazon_raw, kpis)
    guidance_tab(wb, v, amazon_raw)
    notes_tab(wb, v, amazon_raw, meta, passed, len(checks))
    kpis_tab(wb, kpis)
    disclosures_tab(wb, amazon_raw)
    validation_tab(wb, checks)
    audit_tab(wb, amazon_raw, v)
    wb.save(path)
    return path, checks
