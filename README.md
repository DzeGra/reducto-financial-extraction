# Reducto Financial Extraction

Turns quarterly financial documents into the standardized Excel workbook an LP uses for portfolio monitoring and benchmarking, using Reducto's Parse and Extract APIs with citations.

Two documents run through the same pipeline: **Amazon's Q4 2024 earnings release** and **Apple's Q3 FY2024 Form 10-Q**. Each company is reported on its own.

Built for the Reducto Senior Solutions Engineer panel exercise.

## The problem

An LP tracks dozens of portfolio companies and public comparables. Each quarter, analysts copy figures from earnings releases into spreadsheets by hand. It's slow, error-prone, and inconsistent across companies and periods.

## The solution

```
Earnings PDF ──► Reducto upload ──► Reducto Extract ──► Validation ──► LP workbook (.xlsx)
                                    (JSON schema,       (12 reconciliation
                                     citations on)       checks)
```

1. **Studio first:** parsed and extracted in Reducto Studio, took the generated code from Studio’s Code tab, then iterated on the schema with AI (50, then 80, then 90 fields, 92 in the final file).
2. **Extract:** one JSON schema defines every figure the LP asked for, where it appears in the document, its unit and its sign convention.
3. **Trace:** citations return the page, bounding box and confidence score for every value.
4. **Validate:** extracted numbers are reconciled against each other (segments sum to totals, FCF = operating cash flow − net capex, supplemental tables agree with the financial statements).
5. **Deliver:** an Excel workbook in the LP's requested format, with formulas, charts and a full audit trail.

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

## Apple Q3 FY2024 Form 10-Q

The same pipeline pointed at a different document type. Only the schema and the tabs are new; the helpers (styling, row writer, growth formulas, page lookups, Validation tab) are shared with the Amazon builder.

1. **Studio first:** parsed and extracted in Reducto Studio to tune settings and check citations visually, then took the generated code from Studio's Code tab into the notebook.
2. **Schema:** 207 fields across four periods (Q3 FY24, Q3 FY23 and both nine-month periods), written by `schema/generate_apple_schema.py`.
3. **Verify:** an independent answer key (`tests/apple_q3_fy2024_expected.json`) and a PDF page check (`src/verify_against_pdf.py`).
4. **Validate and deliver:** 28 reconciliation checks and the LP workbook.

| Result | |
|---|---|
| Answer key | 201 of 201 fields match, including 2 correct nulls |
| PDF page check | 199 of 199 numbers printed on the exact page Reducto cited |
| Reconciliation | 28 of 28 checks pass |

| Tab | Contents |
|---|---|
| Financials | Income statement for Q3 and nine months with YoY, gross and operating margin, effective tax rate; operating cash flow, capex, FCF, share-based compensation |
| Segments | Five geographic segments (net sales, operating income, margin), net sales by category, Products vs. Services gross margin, revenue mix; three charts |
| Cash Flow | Nine-month cash flow highlights, debt service and capital return, cash, marketable securities, debt and net cash, capital return program; uses-of-cash chart |
| Guidance | No quantitative guidance (stated), margin outlook, dividend policy, FX impact by segment, legal and regulatory items, product announcements |
| Notes | Source, method, verification results, periods and units, calculated fields, omitted items |
| Validation | 28 reconciliation checks with PASS/FAIL |

## Repository

```
reducto_financial_extraction.ipynb          Amazon walkthrough notebook (run in Google Colab)
apple_q3_fy2024_10q.ipynb                   Apple walkthrough notebook (run in Google Colab)
src/build_workbook.py                       Shared helpers plus the Amazon workbook builder
src/build_apple_workbook.py                 Apple workbook tabs and validation checks
src/verify_against_pdf.py                   Checks every extracted number against the source PDF
schema/amazon_q4_2024_schema.json           Amazon extraction schema (92 fields)
schema/apple_q3_fy2024_10q_schema.json      Apple extraction schema (207 fields)
schema/generate_apple_schema.py             Writes the Apple schema
tests/apple_q3_fy2024_expected.json         Apple answer key: every number with its PDF page
dev/development_notebook.ipynb              Amazon development log: schema iterations and experiments
dev/apple_build_notebook.ipynb              Apple step-by-step build, one cell per tab
output/Amazon_Q4_2024_LP_Analysis.xlsx      Amazon sample output from the final run (job df4c47ae)
output/Apple_Q3_FY2024_LP_Analysis.xlsx     Apple sample output (job 43142753)
```

## Sample output

`output/Amazon_Q4_2024_LP_Analysis.xlsx` (12 of 12 checks passed) and `output/Apple_Q3_FY2024_LP_Analysis.xlsx` (28 of 28 checks passed) are the workbooks from the final validated runs. Open them to see the results without running anything.

## Run it

1. Open `reducto_financial_extraction.ipynb` (Amazon) or `apple_q3_fy2024_10q.ipynb` (Apple) in Google Colab.
2. Add your Reducto API key to Colab Secrets as `REDUCTO_API_KEY`.
3. Run all cells and upload the PDF when prompted (`AMZN-Q4-2024-Earnings-Release.pdf` or `_10-Q-Q3-2024-As-Filed.pdf`).
4. The workbook downloads at the end. To reuse a completed extraction and save credits, set REUSE_JOB_ID. Both notebooks save jobs with persist_results, since Reducto deletes job results after 12 hours by default.

## Design decisions

- **Schema descriptions name the exact table and line item.** Early runs mixed up three-month and twelve-month columns; naming the period and table resolved it.
- **Sign conventions are explicit.** Cash outflows return positive and operating losses return negative, so formulas work without cleanup.
- **Citations stay with the data.** `amazon_raw` keeps the full response; `amazon` holds plain values for math.
- **Formulas over hardcoded results.** Growth, margins and mix are live Excel formulas an analyst can audit.
- **Missing data is omitted and logged.** Fields the release leaves out appear in Notes, per the LP brief.
- **Extraction, validation and layout are separate,** so a new company needs a new schema and the rest carries over.

## Roadmap

- Batch processing with async jobs and webhooks for a full portfolio each quarter
- Review queue that routes low-confidence fields to an analyst
- Completeness checks for array fields: KPI record counts varied across reruns (494, then 486) while all 12 checks passed
