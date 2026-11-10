"""Generates schema/apple_q3_fy2024_10q_schema.json (same contract style as the Amazon schema)."""
import json, sys

PERIODS = {
    "q3_fy2024": ("the three months ended June 29, 2024", "Three Months Ended June 29, 2024"),
    "q3_fy2023": ("the three months ended July 1, 2023", "Three Months Ended July 1, 2023"),
    "ytd_fy2024": ("the nine months ended June 29, 2024", "Nine Months Ended June 29, 2024"),
    "ytd_fy2023": ("the nine months ended July 1, 2023", "Nine Months Ended July 1, 2023"),
}
BALANCE = {"jun2024": "June 29, 2024", "sep2023": "September 30, 2023"}
NUM = {"type": ["number", "null"]}
STR = {"type": ["string", "null"]}
MILL = "Return the reported amount in USD millions, without currency symbols or commas."
NULL = "If unavailable, return null."

props = {}

def num(fid, desc):
    props[fid] = {**NUM, "description": desc}

def avoid(period):
    others = [f"'{PERIODS[q][1]}'" for q in PERIODS if q != period]
    return "Do not use the " + ", ".join(others[:-1]) + " or " + others[-1] + " columns."

# 1. Income statement (Condensed Consolidated Statements of Operations), four columns
IS = [
    ("total_net_sales", "Total net sales", MILL + " Do not use Products or Services net sales alone."),
    ("products_net_sales", "Net sales: Products", MILL + " Do not use total net sales or cost of sales."),
    ("services_net_sales", "Net sales: Services", MILL + " Do not use total net sales or cost of sales."),
    ("gross_margin", "total Gross margin (dollar amount)", MILL + " Do not return a percentage."),
    ("research_and_development", "Research and development operating expense", MILL),
    ("selling_general_administrative", "Selling, general and administrative operating expense", MILL),
    ("operating_income", "Operating income", MILL + " Do not use segment operating income."),
    ("provision_for_income_taxes", "Provision for income taxes", MILL),
    ("net_income", "Net income", MILL),
    ("basic_eps", "Earnings per share: Basic", "Return the reported USD per-share amount without currency symbols. Do not use diluted EPS."),
    ("diluted_eps", "Earnings per share: Diluted", "Return the reported USD per-share amount without currency symbols. Do not use basic EPS."),
    ("diluted_shares", "Shares used in computing earnings per share: Diluted", "Return the reported number in thousands of shares, without commas. Do not use basic shares."),
]
for p, (words, col) in PERIODS.items():
    for key, line, rule in IS:
        num(f"{p}_{key}", f"Extract Apple's {line} for {words} from the Condensed Consolidated Statements of Operations (column '{col}'). {rule} {avoid(p)} {NULL}")

# 2. Gross margin by Products / Services (MD&A Gross Margin table)
for p, (words, col) in PERIODS.items():
    for kind in ("Products", "Services"):
        num(f"{p}_{kind.lower()}_gross_margin",
            f"Extract the {kind} gross margin dollar amount for {words} from the Gross Margin table in Management's Discussion and Analysis. {MILL} Do not return the gross margin percentage. {avoid(p)} {NULL}")
    num(f"{p}_effective_tax_rate_pct",
        f"Extract Apple's effective tax rate for {words} from the Provision for Income Taxes table in Management's Discussion and Analysis. Return percentage points, e.g. 15.9 for 15.9%. Do not use the statutory federal rate. {avoid(p)} {NULL}")

# 3. Net sales by category (Note 2 Revenue)
CATS = [("iphone", "iPhone"), ("mac", "Mac"), ("ipad", "iPad"),
        ("wearables_home_accessories", "Wearables, Home and Accessories"), ("services_category", "Services")]
for p, (words, col) in PERIODS.items():
    for key, name in CATS:
        num(f"{p}_{key}_net_sales",
            f"Extract {name} net sales for {words} from the net sales by category table in Note 2, Revenue. {MILL} Do not use geographic segment sales or Products/Services totals. {avoid(p)} {NULL}")

# 4. Geographic reportable segments (Note: Segment Information and Geographic Data)
GEO = [("americas", "Americas"), ("europe", "Europe"), ("greater_china", "Greater China"),
       ("japan", "Japan"), ("rest_of_asia_pacific", "Rest of Asia Pacific")]
for p, (words, col) in PERIODS.items():
    for key, name in GEO:
        num(f"{p}_{key}_net_sales",
            f"Extract {name} segment net sales for {words} from the Segment Information and Geographic Data note. {MILL} Do not use product category sales or consolidated sales. {avoid(p)} {NULL}")
        num(f"{p}_{key}_operating_income",
            f"Extract {name} segment operating income for {words} from the Segment Information and Geographic Data note. {MILL} Do not use segment net sales or consolidated operating income. {avoid(p)} {NULL}")
    num(f"{p}_segment_operating_income_total",
        f"Extract total 'Segment operating income' for {words} from the reconciliation of segment operating income to the Condensed Consolidated Statements of Operations. {MILL} {avoid(p)} {NULL}")
    num(f"{p}_other_corporate_expenses",
        f"Extract 'Other corporate expenses, net' for {words} from the reconciliation of segment operating income to the Condensed Consolidated Statements of Operations. {MILL} Return the positive magnitude. {avoid(p)} {NULL}")

# 5. Cash flow statement (nine months only in a Q3 10-Q)
CF = [
    ("operating_cash_flow", "Cash generated by operating activities", ""),
    ("share_based_compensation", "Share-based compensation expense", ""),
    ("depreciation_amortization", "Depreciation and amortization", ""),
    ("purchases_ppe", "Payments for acquisition of property, plant and equipment", " Return the positive magnitude of the cash outflow."),
    ("purchases_marketable_securities", "Purchases of marketable securities", " Return the positive magnitude of the cash outflow."),
    ("maturities_marketable_securities", "Proceeds from maturities of marketable securities", ""),
    ("sales_marketable_securities", "Proceeds from sales of marketable securities", ""),
    ("investing_cash_flow", "Cash generated by investing activities", " Return a negative number only if the source shows parentheses."),
    ("net_share_settlement_taxes", "Payments for taxes related to net share settlement of equity awards", " Return the positive magnitude of the cash outflow."),
    ("dividends_paid", "Payments for dividends and dividend equivalents", " Return the positive magnitude of the cash outflow."),
    ("share_repurchases", "Repurchases of common stock", " Return the positive magnitude of the cash outflow."),
    ("term_debt_proceeds", "Proceeds from issuance of term debt, net", " Return null if the cell shows a dash."),
    ("term_debt_repayments", "Repayments of term debt", " Return the positive magnitude of the cash outflow."),
    ("commercial_paper_net", "Repayments of commercial paper, net", " Return a negative number for net repayments (the source shows parentheses)."),
    ("financing_cash_flow", "Cash used in financing activities", " Return a negative number (the source shows parentheses)."),
    ("beginning_cash", "Cash, cash equivalents and restricted cash, beginning balances", ""),
    ("ending_cash", "Cash, cash equivalents and restricted cash, ending balances", ""),
    ("cash_paid_income_taxes", "Supplemental cash flow disclosure: Cash paid for income taxes, net", ""),
]
for p in ("ytd_fy2024", "ytd_fy2023"):
    words, col = PERIODS[p]
    other = "July 1, 2023" if p == "ytd_fy2024" else "June 29, 2024"
    for key, line, rule in CF:
        num(f"{p}_{key}",
            f"Extract '{line}' for {words} from the Condensed Consolidated Statements of Cash Flows. {MILL}{rule} Do not use the nine months ended {other} column. {NULL}")

# 6. Balance sheet
BS = [
    ("cash_cash_equivalents", "Cash and cash equivalents", " Do not include marketable securities or restricted cash."),
    ("current_marketable_securities", "Marketable securities (current assets)", " Do not use non-current marketable securities."),
    ("non_current_marketable_securities", "Marketable securities (non-current assets)", " Do not use current marketable securities."),
    ("inventories", "Inventories", ""),
    ("total_assets", "Total assets", ""),
    ("deferred_revenue_current", "Deferred revenue (current liabilities)", ""),
    ("commercial_paper", "Commercial paper", ""),
    ("current_term_debt", "Term debt (current liabilities)", " Do not use non-current term debt."),
    ("non_current_term_debt", "Term debt (non-current liabilities)", " Do not use current term debt."),
    ("total_liabilities", "Total liabilities", ""),
    ("shareholders_equity", "Total shareholders' equity", ""),
]
for p, date in BALANCE.items():
    other = "September 30, 2023" if p == "jun2024" else "June 29, 2024"
    for key, line, rule in BS:
        num(f"{p}_{key}", f"Extract '{line}' as of {date} from the Condensed Consolidated Balance Sheets. {MILL}{rule} Do not use the {other} column. {NULL}")
    num(f"{p}_restricted_marketable_securities",
        f"Extract the amount of marketable securities restricted from general use related to the State Aid Decision and other agreements, as of {date}, from the Financial Instruments note. USD millions; convert billions to millions. {NULL}")

for p in ("q3_fy2024", "q3_fy2023"):
    words, col = PERIODS[p]
    num(f"{p}_share_based_compensation",
        f"Extract share-based compensation expense for {words} from the share-based compensation table in Note 9, Share-Based Compensation. {MILL} Do not use the income tax benefit or the nine-month figure. {avoid(p)} {NULL}")
for p, date in BALANCE.items():
    num(f"{p}_total_deferred_revenue",
        f"Extract total deferred revenue as of {date} as stated in Note 2, Revenue (current and non-current combined). USD millions; convert billions to millions. Do not use the current deferred revenue line on the balance sheet. {NULL}")

# 7. Capital return and obligations (MD&A and Part II)
num("q3_fy2024_share_repurchases_md_and_a", "Extract the dollar amount of common stock Apple repurchased during the third quarter of 2024 as stated in Liquidity and Capital Resources / Capital Return Program. USD millions; convert billions to millions. Do not use the nine-month cash flow figure. " + NULL)
num("q3_fy2024_shares_repurchased_thousands", "Extract the total number of shares purchased in the third quarter of 2024 from the Issuer Purchases of Equity Securities table in Part II, Item 2 (Total row). Return thousands of shares without commas. " + NULL)
num("q3_fy2024_dividend_per_share", "Extract the quarterly dividend per share declared, as stated in Capital Return Program or the Statements of Shareholders' Equity for the three months ended June 29, 2024. USD per share. " + NULL)
num("q3_fy2024_dividends_declared", "Extract dividends and dividend equivalents declared for the three months ended June 29, 2024 from the Condensed Consolidated Statements of Shareholders' Equity. " + MILL + " " + NULL)
num("jun2024_remaining_repurchase_authorization", "Extract the approximate dollar value of shares that may yet be purchased under the plans or programs as of the end of the quarter, from the Issuer Purchases of Equity Securities table (Part II, Item 2). " + MILL + " " + NULL)
num("may2024_new_repurchase_program", "Extract the size of the new share repurchase program authorized by the Board on May 2, 2024, from the footnote to the Issuer Purchases of Equity Securities table. USD millions; convert billions to millions. " + NULL)
num("jun2024_manufacturing_purchase_obligations", "Extract total manufacturing purchase obligations as of June 29, 2024 from Liquidity and Capital Resources. USD millions; convert billions to millions. " + NULL)
num("jun2024_manufacturing_purchase_obligations_12m", "Extract the portion of manufacturing purchase obligations payable within 12 months, as of June 29, 2024, from Liquidity and Capital Resources. USD millions; convert billions to millions. " + NULL)

# 8. Forward-looking and qualitative (Guidance tab)
props["gross_margin_outlook"] = {**STR, "description": "Quote management's forward-looking statement about future gross margins from the Gross Margin section of MD&A (for example expected volatility or downward pressure). Quote exactly; do not summarize risk factors. Return null if not disclosed."}
props["dividend_policy"] = {**STR, "description": "Quote the statement on Apple's intention to increase its dividend, from Capital Return Program in MD&A. Quote exactly. Return null if not disclosed."}
props["fx_commentary"] = {**STR, "description": "Transcribe the statements in the Segment Operating Performance section of MD&A about how foreign currency movements affected segment net sales in the third quarter and first nine months of 2024. No outside information. Return null if not disclosed."}
props["fiscal_calendar_note"] = {**STR, "description": "Transcribe the statement on the number of weeks in fiscal 2024 and fiscal 2023 and which quarter had the extra week, from the Fiscal Period section of MD&A or Note 1. Return null if not disclosed."}
props["quantitative_guidance"] = {**STR, "description": "Extract any explicit quantitative revenue, margin, operating expense or capital expenditure guidance for a future quarter or fiscal year stated in this Form 10-Q. Do not use historical results, purchase obligations or risk factors. Return null if the filing gives no quantitative guidance."}

ITEM = {"type": "object", "properties": {"topic": {"type": "string"}, "details": {"type": "string"},
        "value": NUM, "unit": STR, "period": STR}, "additionalProperties": False}
for fid, desc in [
    ("legal_and_regulatory_developments", "Extract each legal or regulatory matter described in Part II, Item 1 Legal Proceedings and the European Commission State Aid Decision in Note 6, Income Taxes (for example the Digital Markets Act investigations, the Department of Justice lawsuit, Epic Games, the European Commission State Aid Decision). Each item: topic, factual details with dates, any disclosed amount, period. No outside information. Return null if none."),
    ("product_and_business_announcements", "Extract the products, software and business announcements listed in the Business Seasonality and Product Introductions or Quarterly Highlights section of MD&A. Each item: topic (product name), details, period. Return null if none."),
]:
    props[fid] = {"type": ["array", "null"], "items": ITEM, "description": desc}

schema = {"schema": {"type": "object", "properties": props, "required": list(props)}}
out = sys.argv[1]
json.dump(schema, open(out, "w"), indent=2, ensure_ascii=False)
print(len(props), "fields ->", out)
