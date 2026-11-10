"""
Builds the LP earnings workbook for Apple's Q3 FY2024 Form 10-Q from a Reducto Extract result.

Input : fields, the Reducto extraction payload (each field = {"value": ..., "citations": [...]})
Output: an Excel workbook with the tabs the LP brief asks for
        (Financials, Segments, Cash Flow, Guidance, Notes) plus Validation.

Shared helpers (styling, the Sheet row writer, growth formulas, page lookups and the
Validation tab) come from build_workbook.py, the Amazon builder. Only the tabs are Apple-specific.
Every number comes from the extraction; growth rates and margins are Excel formulas.
"""

import re
from datetime import date

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment

from build_workbook import (FMT_EPS, FMT_M, FMT_PCT, FMT_PP, Sheet, growth, show_axes,
                            source_page, unwrap, unwrap_record, validation_tab, values_from)

COMPANY = "Apple Inc."
SOURCE = "Form 10-Q, quarter ended June 29, 2024"
PERIODS = ["q3_fy2023", "q3_fy2024", "ytd_fy2023", "ytd_fy2024"]
LABEL = {"q3_fy2023": "Q3 FY23", "q3_fy2024": "Q3 FY24", "ytd_fy2023": "9M FY23", "ytd_fy2024": "9M FY24"}
HEAD = ["Metric", "Q3 FY23", "Q3 FY24", "Q3 YoY", "9M FY23", "9M FY24", "9M YoY", "Page", "Field ID"]
FMT_SH = "#,##0"
FMTS = [None, FMT_M, FMT_M, FMT_PCT, FMT_M, FMT_M, FMT_PCT, None, None]
WRAP = Alignment(wrap_text=True, vertical="top")

REGIONS = [("Americas", "americas"), ("Europe", "europe"), ("Greater China", "greater_china"),
           ("Japan", "japan"), ("Rest of Asia Pacific", "rest_of_asia_pacific")]
CATEGORIES = [("iPhone", "iphone"), ("Mac", "mac"), ("iPad", "ipad"),
              ("Wearables, Home and Accessories", "wearables_home_accessories"), ("Services", "services_category")]


# ---------------------------------------------------------------------------
# Row helpers
# ---------------------------------------------------------------------------

def period_row(s, v, raw, name, key, fmt, bold=False):
    """One metric across the four periods, with YoY growth formulas."""
    a, b, c, d = [v.get(f"{p}_{key}") for p in PERIODS]
    r = s.row
    page = source_page(raw.get(f"q3_fy2024_{key}") or raw.get(f"ytd_fy2024_{key}"))
    s.add([name, a, b, growth(f"C{r}", f"B{r}") if a and b else None,
           c, d, growth(f"F{r}", f"E{r}") if c and d else None, page, f"*_{key}"],
          [None, fmt, fmt, FMT_PCT, fmt, fmt, FMT_PCT], bold=bold)
    return r


def ratio_row(s, name, top, bottom):
    """A calculated margin: row 'top' divided by row 'bottom', change in percentage points."""
    r = s.row
    s.add([name, f"=B{top}/B{bottom}", f"=C{top}/C{bottom}", f"=(C{r}-B{r})*100",
           f"=E{top}/E{bottom}", f"=F{top}/F{bottom}", f"=(F{r}-E{r})*100", None, "calculated"],
          [None, FMT_PCT, FMT_PCT, FMT_PP, FMT_PCT, FMT_PCT, FMT_PP])


def total_row(s, first, last):
    r = s.row
    s.add(["Total", f"=SUM(B{first}:B{last})", f"=SUM(C{first}:C{last})", growth(f"C{r}", f"B{r}"),
           f"=SUM(E{first}:E{last})", f"=SUM(F{first}:F{last})", growth(f"F{r}", f"E{r}")], FMTS, bold=True)
    return r


def text_row(s, topic, text, page):
    r = s.add([topic, text, None, page])
    s.ws.cell(r, 1).alignment = WRAP
    s.ws.cell(r, 2).alignment = WRAP
    s.ws.row_dimensions[r].height = max(18, 15 * (len(str(text)) // 95 + 1))


# ---------------------------------------------------------------------------
# Tabs
# ---------------------------------------------------------------------------

def financials_tab(wb, v, raw):
    s = Sheet(wb, "Financials", f"{COMPANY} | Financial Overview",
              f"USD millions except per-share data · Source: {SOURCE}",
              [46, 13, 13, 12, 13, 13, 12, 8, 34])
    s.section("Income statement")
    s.header(HEAD)
    rows = {}
    for name, key, fmt in [
        ("Total net sales", "total_net_sales", FMT_M),
        ("Products net sales", "products_net_sales", FMT_M),
        ("Services net sales", "services_net_sales", FMT_M),
        ("Gross margin", "gross_margin", FMT_M),
        ("Research and development", "research_and_development", FMT_M),
        ("Selling, general and administrative", "selling_general_administrative", FMT_M),
        ("Operating income", "operating_income", FMT_M),
        ("Provision for income taxes", "provision_for_income_taxes", FMT_M),
        ("Net income", "net_income", FMT_M),
        ("Basic EPS", "basic_eps", FMT_EPS),
        ("Diluted EPS", "diluted_eps", FMT_EPS),
        ("Diluted shares (thousands)", "diluted_shares", FMT_SH),
    ]:
        rows[key] = period_row(s, v, raw, name, key, fmt, bold=key in ("total_net_sales", "net_income"))
        if key == "gross_margin":
            ratio_row(s, "Gross margin % (calculated)", rows["gross_margin"], rows["total_net_sales"])
        if key == "operating_income":
            ratio_row(s, "Operating margin (calculated)", rows["operating_income"], rows["total_net_sales"])
        if key == "provision_for_income_taxes":
            pct = [x / 100 if x is not None else None for x in (v.get(f"{p}_effective_tax_rate_pct") for p in PERIODS)]
            r = s.row
            s.add(["Effective tax rate (as reported, MD&A)", pct[0], pct[1], f"=(C{r}-B{r})*100",
                   pct[2], pct[3], f"=(F{r}-E{r})*100",
                   source_page(raw.get("q3_fy2024_effective_tax_rate_pct")), "*_effective_tax_rate_pct"],
                  [None, FMT_PCT, FMT_PCT, FMT_PP, FMT_PCT, FMT_PCT, FMT_PP])

    s.gap()
    s.section("Cash flow (a 10-Q reports cash flows for the nine months only)")
    s.header(HEAD)
    ocf = period_row(s, v, raw, "Operating cash flow", "operating_cash_flow", FMT_M)
    capex = period_row(s, v, raw, "Capital expenditures (payments for PP&E)", "purchases_ppe", FMT_M)
    r = s.row
    s.add(["Free cash flow (calculated: operating cash flow − capex)", None, None, None,
           f"=E{ocf}-E{capex}", f"=F{ocf}-F{capex}", growth(f"F{r}", f"E{r}"), None, "calculated"],
          [None, FMT_M, FMT_M, FMT_PCT, FMT_M, FMT_M, FMT_PCT], bold=True)
    period_row(s, v, raw, "Share-based compensation", "share_based_compensation", FMT_M)
    s.ws.freeze_panes = "B4"


def segments_tab(wb, v, raw):
    s = Sheet(wb, "Segments", f"{COMPANY} | Segment Performance",
              f"USD millions · Source: {SOURCE} (Note 2 Revenue p.9, Note 11 Segments p.15, MD&A p.19)",
              [34, 13, 13, 12, 13, 13, 12, 8, 34, 13, 12])
    head = ["Segment"] + HEAD[1:]

    s.section("Net sales by reportable segment (geographic)")
    s.header(head)
    g_first = s.row
    for name, key in REGIONS:
        period_row(s, v, raw, name, f"{key}_net_sales", FMT_M)
    g_last = s.row - 1
    g_total = total_row(s, g_first, g_last)

    s.gap()
    s.section("Operating income and margin by segment")
    s.header(["Segment", "Q3 FY23 OI", "Q3 FY24 OI", "Q3 FY23 margin", "Q3 FY24 margin",
              "9M FY23 OI", "9M FY24 OI", "9M FY23 margin", "9M FY24 margin", "9M margin change"])
    margin_fmts = [None, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_M, FMT_M, FMT_PCT, FMT_PCT, FMT_PP]
    m_first = s.row
    for i, (name, key) in enumerate(REGIONS):
        r, sr = s.row, g_first + i
        s.add([name, v[f"q3_fy2023_{key}_operating_income"], v[f"q3_fy2024_{key}_operating_income"],
               f"=B{r}/B{sr}", f"=C{r}/C{sr}", v[f"ytd_fy2023_{key}_operating_income"], v[f"ytd_fy2024_{key}_operating_income"],
               f"=F{r}/E{sr}", f"=G{r}/F{sr}", f"=(I{r}-H{r})*100"], margin_fmts)
    m_last = s.row - 1
    r = s.row
    s.add(["Total segment operating income", f"=SUM(B{m_first}:B{m_last})", f"=SUM(C{m_first}:C{m_last})",
           f"=B{r}/B{g_total}", f"=C{r}/C{g_total}", f"=SUM(F{m_first}:F{m_last})", f"=SUM(G{m_first}:G{m_last})",
           f"=F{r}/E{g_total}", f"=G{r}/F{g_total}", f"=(I{r}-H{r})*100"], margin_fmts, bold=True)

    s.gap()
    s.section("Net sales by category (Note 2, Revenue)")
    s.header(["Category"] + HEAD[1:])
    c_first = s.row
    for name, key in CATEGORIES:
        period_row(s, v, raw, name, f"{key}_net_sales", FMT_M)
    c_last = s.row - 1
    c_total = total_row(s, c_first, c_last)

    s.gap()
    s.section("Gross margin by Products and Services (MD&A)")
    s.header(HEAD)
    for kind in ("products", "services"):
        gm = period_row(s, v, raw, f"{kind.capitalize()} gross margin", f"{kind}_gross_margin", FMT_M)
        ns = s.row
        s.add([f"{kind.capitalize()} net sales", *[v.get(f"{p}_{kind}_net_sales") for p in PERIODS[:2]], None,
               *[v.get(f"{p}_{kind}_net_sales") for p in PERIODS[2:]]], FMTS)
        ratio_row(s, f"{kind.capitalize()} gross margin % (calculated)", gm, ns)

    s.gap()
    s.section("Revenue mix (share of total net sales)")
    s.header(["Segment / category", "Q3 FY23", "Q3 FY24", "", "9M FY23", "9M FY24"])
    for first, last, tot in [(g_first, g_last, g_total), (c_first, c_last, c_total)]:
        for sr in range(first, last + 1):
            s.add([f"=A{sr}", f"=B{sr}/B${tot}", f"=C{sr}/C${tot}", None, f"=E{sr}/E${tot}", f"=F{sr}/F${tot}"],
                  [None, FMT_PCT, FMT_PCT, None, FMT_PCT, FMT_PCT])

    ws = s.ws
    for title, cols, first, last, anchor, pct in [
        ("Net sales by segment, nine months (USD M)", (5, 6), g_first, g_last, "L4", False),
        ("Operating margin by segment, nine months", (8, 9), m_first, m_last, "L20", True),
        ("Net sales by category, nine months (USD M)", (5, 6), c_first, c_last, "L36", False),
    ]:
        chart = BarChart()
        chart.title, chart.height, chart.width = title, 7.5, 15
        if pct:
            chart.y_axis.number_format = "0%"
        else:
            chart.y_axis.title = "USD millions"
        chart.add_data(Reference(ws, min_col=cols[0], max_col=cols[1], min_row=first - 1, max_row=last), titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=first, max_row=last))
        ws.add_chart(show_axes(chart), anchor)
    ws.freeze_panes = "B4"


def cash_flow_tab(wb, v, raw):
    s = Sheet(wb, "Cash Flow", f"{COMPANY} | Cash Flow, Debt Service and Liquidity",
              f"USD millions · Cash outflows shown as positive amounts · Source: {SOURCE} (pp.6, 8, 11, 20, 23)",
              [58, 14, 14, 12, 8, 34])

    def block(title, rows, prefixes=("ytd_fy2023", "ytd_fy2024"), heads=("9M FY23", "9M FY24")):
        """A two-period table; returns {field key: row number}."""
        s.section(title)
        s.header(["Metric", *heads, "Change", "Page", "Field ID"])
        where = {}
        for name, key in rows:
            a, b = v.get(f"{prefixes[0]}_{key}"), v.get(f"{prefixes[1]}_{key}")
            if a is None and b is None:
                continue  # brief: omit data the source does not report
            r = s.row
            where[key] = r
            s.add([name, a, b, growth(f"C{r}", f"B{r}") if a and b else None,
                   source_page(raw.get(f"{prefixes[1]}_{key}")) or source_page(raw.get(f"{prefixes[0]}_{key}")), f"*_{key}"],
                  [None, FMT_M, FMT_M, FMT_PCT])
        return where

    def calc(name, formula_b, formula_c):
        r = s.row
        s.add([name, formula_b, formula_c, growth(f"C{r}", f"B{r}"), None, "calculated"],
              [None, FMT_M, FMT_M, FMT_PCT], bold=True)
        return r

    w = block("Cash flow highlights (nine months)", [
        ("Cash generated by operating activities", "operating_cash_flow"),
        ("Depreciation and amortization", "depreciation_amortization"),
        ("Share-based compensation expense", "share_based_compensation"),
        ("Capital expenditures (payments for PP&E)", "purchases_ppe"),
        ("Purchases of marketable securities", "purchases_marketable_securities"),
        ("Proceeds from maturities of marketable securities", "maturities_marketable_securities"),
        ("Proceeds from sales of marketable securities", "sales_marketable_securities"),
        ("Cash generated by investing activities", "investing_cash_flow"),
        ("Cash paid for income taxes, net", "cash_paid_income_taxes"),
    ])
    o, c = w["operating_cash_flow"], w["purchases_ppe"]
    calc("Free cash flow (calculated: operating cash flow − capex)", f"=B{o}-B{c}", f"=C{o}-C{c}")

    s.gap()
    w = block("Debt service and capital return (nine months)", [
        ("Repayments of term debt", "term_debt_repayments"),
        ("Proceeds from issuance of term debt, net", "term_debt_proceeds"),
        ("Commercial paper, net (negative = net repayment)", "commercial_paper_net"),
        ("Payments for dividends and dividend equivalents", "dividends_paid"),
        ("Repurchases of common stock", "share_repurchases"),
        ("Taxes paid on net share settlement of equity awards", "net_share_settlement_taxes"),
        ("Cash used in financing activities (negative = outflow)", "financing_cash_flow"),
    ])
    t, cp = w["term_debt_repayments"], w["commercial_paper_net"]
    debt_row = calc("Total debt repayments (calculated: term debt + net commercial paper)",
                    f"=B{t}-MIN(B{cp},0)", f"=C{t}-MIN(C{cp},0)")
    d, rp = w["dividends_paid"], w["share_repurchases"]
    calc("Total capital returned (calculated: dividends + buybacks)", f"=B{d}+B{rp}", f"=C{d}+C{rp}")
    uses = {"capex": c, "debt": debt_row, "dividends": d, "buybacks": rp}

    s.gap()
    w = block("Available cash, marketable securities and debt (balance sheet)", [
        ("Cash and cash equivalents", "cash_cash_equivalents"),
        ("Marketable securities, current", "current_marketable_securities"),
        ("Marketable securities, non-current", "non_current_marketable_securities"),
        ("Commercial paper", "commercial_paper"),
        ("Term debt, current", "current_term_debt"),
        ("Term debt, non-current", "non_current_term_debt"),
        ("Of which restricted (State Aid escrow; included in marketable securities above)", "restricted_marketable_securities"),
    ], prefixes=("sep2023", "jun2024"), heads=("Sep 30, 2023", "Jun 29, 2024"))
    a, b2, c2 = w["cash_cash_equivalents"], w["current_marketable_securities"], w["non_current_marketable_securities"]
    cash_row = calc("Total cash and marketable securities (calculated)", f"=B{a}+B{b2}+B{c2}", f"=C{a}+C{b2}+C{c2}")
    p, q, r_ = w["commercial_paper"], w["current_term_debt"], w["non_current_term_debt"]
    debt_total = calc("Total debt (calculated: commercial paper + term debt)", f"=B{p}+B{q}+B{r_}", f"=C{p}+C{q}+C{r_}")
    calc("Net cash (calculated: cash and securities − debt)", f"=B{cash_row}-B{debt_total}", f"=C{cash_row}-C{debt_total}")

    s.gap()
    s.section("Capital return program and obligations (as stated for Q3 FY24 / June 29, 2024)")
    s.header(["Item", "Value", "Unit", "", "Page", "Field ID"])
    for name, key, unit, fmt in [
        ("Shares repurchased in Q3 FY24 (dollar amount, MD&A)", "q3_fy2024_share_repurchases_md_and_a", "USD millions", FMT_M),
        ("Shares repurchased in Q3 FY24 (count)", "q3_fy2024_shares_repurchased_thousands", "thousands of shares", FMT_SH),
        ("Quarterly dividend declared per share", "q3_fy2024_dividend_per_share", "USD", FMT_EPS),
        ("Dividends and dividend equivalents declared, Q3 FY24", "q3_fy2024_dividends_declared", "USD millions", FMT_M),
        ("New repurchase program authorized May 2, 2024", "may2024_new_repurchase_program", "USD millions", FMT_M),
        ("Remaining repurchase authorization, June 29, 2024", "jun2024_remaining_repurchase_authorization", "USD millions", FMT_M),
        ("Manufacturing purchase obligations, total", "jun2024_manufacturing_purchase_obligations", "USD millions", FMT_M),
        ("Manufacturing purchase obligations, due within 12 months", "jun2024_manufacturing_purchase_obligations_12m", "USD millions", FMT_M),
    ]:
        if v.get(key) is not None:
            s.add([name, v[key], unit, None, source_page(raw.get(key)), key], [None, fmt])

    ws = s.ws
    s.gap()
    s.section("Chart data: main uses of cash (nine months)")
    head = s.header(["Use of cash", "9M FY23", "9M FY24"])
    for name, key in [("Capital expenditures", "capex"), ("Debt repayments", "debt"),
                      ("Dividends", "dividends"), ("Share repurchases", "buybacks")]:
        s.add([name, f"=B{uses[key]}", f"=C{uses[key]}"], [None, FMT_M, FMT_M])
    chart = BarChart()
    chart.title, chart.y_axis.title, chart.height, chart.width = "Main uses of cash, nine months (USD M)", "USD millions", 8, 16
    chart.add_data(Reference(ws, min_col=2, max_col=3, min_row=head, max_row=head + 4), titles_from_data=True)
    chart.set_categories(Reference(ws, min_col=1, min_row=head + 1, max_row=head + 4))
    ws.add_chart(show_axes(chart), "H4")
    ws.freeze_panes = "B4"


def guidance_tab(wb, v, raw):
    s = Sheet(wb, "Guidance", f"{COMPANY} | Forward Guidance and Management Commentary",
              f"As stated in the filing, no editorial commentary · Source: {SOURCE}", [34, 100, 18, 8])

    s.section("Quantitative guidance")
    s.header(["Item", "Detail", "", "Page"])
    if v.get("quantitative_guidance") is None:
        text_row(s, "Revenue, margin, opex or capex guidance",
                 "None provided. Apple's Form 10-Q contains no numeric guidance for future periods "
                 "(field quantitative_guidance returned null).", None)
    else:
        text_row(s, "Quantitative guidance", v["quantitative_guidance"], source_page(raw.get("quantitative_guidance")))

    s.gap()
    s.section("Margin outlook, dividend policy and fiscal calendar (quoted from MD&A)")
    s.header(["Topic", "Statement", "", "Page"])
    for topic, key in [("Gross margin outlook", "gross_margin_outlook"), ("Dividend policy", "dividend_policy"),
                       ("Fiscal calendar", "fiscal_calendar_note")]:
        if v.get(key):
            text_row(s, topic, v[key], source_page(raw.get(key)))

    if v.get("fx_commentary"):
        s.gap()
        s.section("Foreign currency impact on segment net sales (Segment Operating Performance, MD&A)")
        s.header(["Segment", "Statement", "", "Page"])
        page = source_page(raw.get("fx_commentary"))
        # one row per region; handles plain sentences and "Region: ..." formatting
        for sentence in re.split(r"(?<=[.;])\s+(?=The |Americas|Europe|Greater China|Japan|Rest of Asia Pacific)",
                                 v["fx_commentary"]):
            region = next((n for n, _ in REGIONS if n in sentence), "All segments")
            text_row(s, region, sentence.strip(), page)

    for title, key in [("Legal and regulatory developments (Part II Item 1, Note 6)", "legal_and_regulatory_developments"),
                       ("Products and business announcements (MD&A)", "product_and_business_announcements")]:
        records = v.get(key) or []
        if not records:
            continue
        s.gap()
        s.section(title)
        s.header(["Topic", "Details", "Amount / period", "Page"])
        for rec in (unwrap_record(x, "topic") for x in records):
            amount, unit, period = (unwrap(rec.get(k)) for k in ("value", "unit", "period"))
            amount_text = f"{amount:,.0f}" if isinstance(amount, (int, float)) else amount
            amount_text = " ".join(str(x) for x in (amount_text, unit) if x)  # "10 percent of ..." reads as one phrase
            extra = " · ".join(x for x in (amount_text, str(period) if period else "") if x)
            details = unwrap(rec.get("details"))
            r = s.add([unwrap(rec.get("topic")), details, extra or None,
                       source_page(rec.get("topic")) or source_page(rec.get("details"))])
            for col in (1, 2, 3):
                s.ws.cell(r, col).alignment = WRAP
            s.ws.row_dimensions[r].height = max(18, 15 * (len(str(details)) // 95 + 1))
    s.ws.freeze_panes = "A4"


def notes_tab(wb, v, raw, meta):
    s = Sheet(wb, "Notes", "Notes, Assumptions and Adjustments",
              "Read with the Financials, Segments, Cash Flow and Guidance tabs", [36, 110])

    def note(topic, text):
        r = s.add([topic, text])
        s.ws.cell(r, 1).alignment = WRAP
        s.ws.cell(r, 2).alignment = WRAP
        s.ws.row_dimensions[r].height = max(18, 15 * (len(text) // 105 + 1))

    s.section("Source and method")
    s.header(["Item", "Detail"])
    note("Company", "Apple Inc. (NASDAQ: AAPL)")
    note("Source document", "Form 10-Q for the fiscal quarter ended June 29, 2024 (28 pages).")
    note("Extraction", f"Reducto Extract API with a {meta.get('n_fields', len(raw))}-field JSON schema and citations enabled. "
                       f"Job ID: {meta.get('job_id', 'n/a')}.")
    if meta.get("answer_key_total"):
        note("Verification", f"{meta['answer_key_matches']} of {meta['answer_key_total']} fields match an independently prepared answer key. "
                             f"{meta.get('on_page')} numbers were found printed on the exact page Reducto cited; "
                             f"{meta.get('nulls')} fields are correctly blank.")
    note("Validation", f"{meta.get('checks_passed')} of {meta.get('n_checks')} reconciliation checks passed (see Validation tab).")
    note("Workbook generated", meta.get("run_date", date.today().isoformat()))

    s.gap()
    s.section("Periods, units and conventions")
    s.header(["Item", "Detail"])
    note("Fiscal calendar", "Apple's fiscal year ends on the last Saturday of September. Q3 FY24 = three months ended June 29, 2024; "
                            "Q3 FY23 = three months ended July 1, 2023.")
    note("Year to date instead of full year", "A 10-Q reports no full-year results. The brief's full-year columns use the nine months "
                                              "ended June 29, 2024 and July 1, 2023 (labeled 9M).")
    note("Balance sheet dates", "Balance sheet items compare June 29, 2024 with September 30, 2023 (the last fiscal year end), as the filing does.")
    note("Units", "USD millions, except per-share amounts (USD) and share counts (thousands). Amounts the filing states in billions are converted to millions.")
    note("Signs", "Cash outflows are shown as positive amounts. Net commercial paper and cash used in financing activities keep the source's negative sign.")

    s.gap()
    s.section("Calculated fields (Excel formulas)")
    s.header(["Item", "Detail"])
    note("YoY growth", "(Current period ÷ same period a year earlier) − 1.")
    note("Margins", "Gross margin %, operating margin and segment margins = the line item ÷ the matching net sales. Changes are in percentage points (pp).")
    note("Free cash flow", "Cash generated by operating activities − payments for property, plant and equipment. Apple does not report free cash flow; this is a calculation, nine months only.")
    note("Debt repayments", "Repayments of term debt plus net repayments of commercial paper.")
    note("Capital returned", "Dividends paid plus repurchases of common stock.")
    note("Liquidity", "Total cash and marketable securities = cash and cash equivalents + current and non-current marketable securities. "
                      "Total debt = commercial paper + current and non-current term debt. Net cash = the difference.")
    note("Revenue mix", "Segment or category net sales ÷ total net sales.")

    s.gap()
    s.section("Source differences and adjustments")
    s.header(["Item", "Detail"])
    note("Share-based compensation", "Quarterly figures come from Note 9 (Share-Based Compensation); nine-month figures from the cash flow statement.")
    note("Effective tax rate", "As reported in MD&A (Provision for Income Taxes), not recalculated.")
    note("Quarterly cash flow", "The 10-Q presents cash flows for the nine months only, so quarterly operating cash flow, capex and free cash flow are left blank.")
    note("Qualitative fields", "Guidance-tab text is quoted from the filing with page citations. Wording can vary slightly between extraction runs, so these fields are reviewed by eye.")

    s.gap()
    s.section("Omitted items (not disclosed in the source document)")
    s.header(["Item", "Detail"])
    for key in raw:
        if v.get(key) is None and key != "quantitative_guidance":
            note(key, "Returned no value: the filing shows a dash or does not report it.")
    note("Lease obligations", "Requested by the brief. The 10-Q contains no lease disclosures, so they are omitted.")
    note("Unit sales and customer metrics", "Requested by the brief. Apple does not disclose unit sales or customer counts in this filing, so they are omitted.")
    note("Quantitative guidance", "Requested by the brief. The 10-Q provides none (see Guidance tab).")
    s.ws.freeze_panes = "A4"


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------

def apple_checks(v):
    """Reconcile extracted numbers against other numbers in the same filing: (check, calculated, reported)."""
    checks = []
    for p in PERIODS:
        L = LABEL[p]
        checks += [
            (f"{L}: 5 geographic segments add up to total net sales",
             sum(v[f"{p}_{k}_net_sales"] for _, k in REGIONS), v[f"{p}_total_net_sales"]),
            (f"{L}: 5 product categories add up to total net sales",
             sum(v[f"{p}_{k}_net_sales"] for _, k in CATEGORIES), v[f"{p}_total_net_sales"]),
            (f"{L}: Products + Services gross margin (MD&A) = total gross margin (income statement)",
             v[f"{p}_products_gross_margin"] + v[f"{p}_services_gross_margin"], v[f"{p}_gross_margin"]),
            (f"{L}: Gross margin − R&D − SG&A = operating income",
             v[f"{p}_gross_margin"] - v[f"{p}_research_and_development"] - v[f"{p}_selling_general_administrative"],
             v[f"{p}_operating_income"]),
            (f"{L}: Segment operating income − R&D − other corporate expenses = operating income",
             v[f"{p}_segment_operating_income_total"] - v[f"{p}_research_and_development"] - v[f"{p}_other_corporate_expenses"],
             v[f"{p}_operating_income"]),
            (f"{L}: 5 segments' operating income add up to total segment operating income",
             sum(v[f"{p}_{k}_operating_income"] for _, k in REGIONS), v[f"{p}_segment_operating_income_total"]),
        ]
    for p in ("ytd_fy2023", "ytd_fy2024"):
        checks.append((f"{LABEL[p]}: Operating + investing + financing cash flow = change in cash (ending − beginning)",
                       v[f"{p}_operating_cash_flow"] + v[f"{p}_investing_cash_flow"] + v[f"{p}_financing_cash_flow"],
                       v[f"{p}_ending_cash"] - v[f"{p}_beginning_cash"]))
    for p, d in (("sep2023", "Sep 30, 2023"), ("jun2024", "Jun 29, 2024")):
        checks.append((f"{d}: Total liabilities + shareholders' equity = total assets",
                       v[f"{p}_total_liabilities"] + v[f"{p}_shareholders_equity"], v[f"{p}_total_assets"]))
    return checks


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def build_apple_workbook(fields, path="Apple_Q3_FY2024_LP_Analysis.xlsx", meta=None):
    meta = dict(meta or {})
    v = values_from(fields)
    checks = apple_checks(v)
    meta["checks_passed"] = sum(abs(c - r) < 0.5 for _, c, r in checks)
    meta["n_checks"] = len(checks)

    wb = Workbook()
    wb.remove(wb.active)
    financials_tab(wb, v, fields)
    segments_tab(wb, v, fields)
    cash_flow_tab(wb, v, fields)
    guidance_tab(wb, v, fields)
    notes_tab(wb, v, fields, meta)
    validation_tab(wb, checks)
    wb.save(path)
    return path, checks
