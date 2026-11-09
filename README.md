
# Reducto Financial Document Extraction

## Project Overview

This project demonstrates how Reducto's document extraction API can support financial analysts by converting financial disclosures into structured data, validating key financial figures, and exporting results for further analysis.

Built as a technical demonstration for a Senior Solutions Engineer interview at Reducto.

## Business Use Case

A financial institution needs to reduce the manual effort involved in reviewing earnings releases and regulatory filings.

The solution demonstrates a workflow to:

1. Upload financial documents.
2. Extract structured financial information using Reducto.
3. Retrieve extracted values and supporting source citations.
4. Validate financial figures through reconciliation checks.
5. Export results into Excel for analyst review.

## Documents

- Amazon Q4 and FY2024 Earnings Release — implemented.
- Apple Q3 2024 Form 10-Q — planned.

## Technology

- Python
- Google Colab
- Reducto API and Studio
- Requests
- OpenPyXL
- GitHub

## Repository Contents

**reducto_financial_extraction.ipynb**

The development notebook includes extraction iterations, financial validation, workbook generation, and a dedicated technical walkthrough.

## Technical Workflow

### 1. Authentication

The notebook initializes the Reducto Python client using an API key stored securely in Google Colab Secrets.

### 2. Document Upload

The Amazon earnings release PDF is uploaded to Reducto using the API.

### 3. Structured Extraction

The walkthrough retrieves a previously completed Reducto extraction job containing 90 structured fields.

The original extraction response is preserved to retain supporting citations.

### 4. Financial Validation

Three reconciliation checks verify internal consistency:

| Validation | Calculated ($M) | Reported ($M) | Result |
|---|---:|---:|---|
| FY2024 segment revenue | 637,959 | 637,959 | PASS |
| FY2024 segment operating income | 68,593 | 68,593 | PASS |
| FY2024 free cash flow | 38,219 | 38,219 | PASS |

These checks confirm that selected extracted figures reconcile mathematically. They do not independently verify every extracted value against the source document.

### 5. Excel Export

The walkthrough generates `Amazon_Financial_Validation_Demo.xlsx`, containing the three financial checks.

A separate six-sheet Amazon analysis workbook was also developed for the LP-facing financial analysis.

## How to Run

1. Open `reducto_financial_extraction.ipynb` in Google Colab.
2. Install the Reducto Python SDK and other required packages if they are not already installed.
3. Add a Colab Secret named `REDUCTO_API_KEY` containing your own Reducto API key.
4. Enable notebook access to that secret.
5. Run the technical walkthrough section, uploading the Amazon earnings release when prompted.
6. Retrieve the completed extraction, execute the validation checks, and generate the Excel output.

**Important:** The demonstration retrieves an existing extraction using a hardcoded job ID. Access to that job may depend on the Reducto account and permissions. A new user would need to run an extraction in their own account and update the job ID.

## Design Decisions

- **Credential security:** API credentials are retrieved from Colab Secrets rather than hardcoded.
- **Traceability:** Original extraction responses retain citation metadata.
- **Validation:** Financial reconciliation checks provide a basic quality-control layer.
- **Cost efficiency:** The walkthrough reuses a completed extraction instead of repeatedly submitting extraction jobs.
- **Analyst usability:** Structured financial data is exported into Excel.

## Current Limitations

- The technical walkthrough demonstrates the Amazon workflow only.
- The notebook retains exploratory development cells and is not a fully automated production pipeline.
- Extraction accuracy requires further source-level verification.
- Apple 10-Q processing and the broader multi-document solution remain to be completed.

## Next Steps

- Implement Apple 10-Q extraction and validation.
- Extend the workflow to support multiple documents.
- Develop a consolidated analyst dashboard.
- Improve automated testing and error handling.
