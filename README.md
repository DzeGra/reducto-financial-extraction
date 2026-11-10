# Reducto Financial Extraction

Turns a quarterly earnings release into the standardized Excel workbook an LP uses for portfolio monitoring and benchmarking, using Reducto's Extract API with citations.

Built for the Reducto Senior Solutions Engineer panel exercise.

## The problem

An LP tracks dozens of portfolio companies and public comparables. Each quarter, analysts copy figures from earnings releases into spreadsheets by hand. It's slow, error-prone, and inconsistent across companies and periods.

## The solution

```
Earnings PDF ──► Reducto upload ──► Reducto Extract ──► Validation ──► LP workbook (.xlsx)
                                    (JSON schema,       (12 reconciliation
                                     citations on)       checks)
```

1. **Extract:** one JSON schema defines every figure the LP asked for, where it appears in the document, its unit and its sign convention.
2. **Trace:** citations return the page, bounding box and confidence score for every value.
3. **Validate:** extracted numbers are reconciled against each other (segments sum to totals, FCF = operating cash flow − net capex, supplemental tables agree with the financial statements).
4. **Deliver:** an Excel workbook in the LP's requested format, with formulas, charts and a full audit trail.

## Output workbook

| Tab | Contents |
|---|---|
| Financials | Net sales, operating income and margin, net income, EPS, operating cash flow, capex, FCF: Q4 and full year, with YoY |
| Segments | North America, International, AWS: net sales, operating income, margin, revenue mix; two charts |
| Cash Flow | Cash flow highlights, debt and lease repayments, cash and marketable securities, TTM trend chart |
| Guidance | Q1 2025 guidance ranges, FX impact, comparison bases, assumptions |
| Notes | Source, method, units, calculated fields, omitted items, growth-only disclosures |
| KPIs | Release highlights plus every supplemental metric by quarter; stock-based compensation chart |
| Disclosures | Partnerships, contract wins and business developments |
| Validation | 12 reconciliation checks with PASS/FAIL |
| Extraction Audit | Every field with readable name, value, page, confidence and cited source text |

## Repository

```
reducto_financial_extraction.ipynb    Walkthrough notebook (run in Google Colab)
src/build_workbook.py                 Payload loading, validation and workbook builder
schema/amazon_q4_2024_schema.json     Reducto extraction schema (92 fields)
dev/development_notebook.ipynb        Development log: schema iterations and experiments
```

## Run it

1. Open `reducto_financial_extraction.ipynb` in Google Colab.
2. Add your Reducto API key to Colab Secrets as `REDUCTO_API_KEY`.
3. Run all cells and upload `AMZN-Q4-2024-Earnings-Release.pdf` when prompted.
4. The workbook downloads at the end. To reuse a completed extraction and save credits, set `REUSE_JOB_ID`.

## Design decisions

- **Schema descriptions name the exact table and line item.** Early runs mixed up three-month and twelve-month columns; naming the period and table resolved it.
- **Sign conventions are explicit.** Cash outflows return positive and operating losses return negative, so formulas work without cleanup.
- **Citations stay with the data.** `amazon_raw` keeps the full response; `amazon` holds plain values for math.
- **Formulas over hardcoded results.** Growth, margins and mix are live Excel formulas an analyst can audit.
- **Missing data is omitted and logged.** Fields the release leaves out appear in Notes, per the LP brief.
- **Extraction, validation and layout are separate,** so a new company needs a new schema and the rest carries over.

## Roadmap

- Apple Q3 FY2024 Form 10-Q through the same pipeline (SEC filing format, different segment structure)
- Batch processing with async jobs and webhooks for a full portfolio each quarter
- Review queue that routes low-confidence fields to an analyst
