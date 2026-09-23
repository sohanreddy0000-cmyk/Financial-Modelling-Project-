"""
Financial Extractor - ITC Ltd. (Profit & Loss Statement)
Reads multi-year PDF P&L statements via a SINGLE Gemini API call.

Requirements:
    pip install google-genai

Usage:
    python "ITC P&L Extractor.py"
"""

from google import genai
from google.genai import types
import json
import csv
import os
import time
import sys

# CONFIGURATION
API_KEY    = "YOUR_API_KEY_HERE"
MODEL_NAME = "gemini-3.6-flash"

BASE_DIR    = r"PATH OF THE FOLDER"
PDF_PATH    = os.path.join(BASE_DIR, "ITC P&L Statement Compiler.pdf")
OUTPUT_PATH = os.path.join(BASE_DIR, "ITC P&L Extractor.csv")

# How many times to retry if server returns 503 (overloaded), and how long to wait between retries
MAX_RETRIES      = 5
RETRY_WAIT_SEC   = 30   # seconds to wait before each retry

# MASTER PROMPT FOR P&L EXTRACTION
MASTER_PROMPT = """
You are a senior financial analyst and data extraction expert specializing in Indian corporate Profit & Loss (P&L) statements (both pre-Ind AS and post-Ind AS / Schedule VI formats).

You have been given the full PDF of ITC Ltd compiled Annual Report Statement of Profit & Loss / Income Statements spanning multiple years (2002 to 2026).

YOUR TASK:
Read EVERY SINGLE PAGE of this PDF carefully like a human reading a document.
Extract ALL Profit & Loss statement data for ALL years present in the PDF.
Consolidate everything into ONE perfectly structured JSON object.

OUTPUT FORMAT (STRICT - follow exactly):
Return ONLY a valid raw JSON object. No explanation text. No markdown. No code fences. Just the JSON.

Structure:
{
  "years": [2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
  "rows": [
    {
      "particular": "I. Revenue From Operations",
      "values": {"2002": "XXXX.XX", "2003": "XXXX.XX", "2026": "XXXX.XX"}
    },
    {
      "particular": "II. Other Income",
      "values": {"2002": "XXXX.XX", "2003": "XXXX.XX", "2026": "XXXX.XX"}
    }
  ]
}

PARTICULARS COLUMN - HOW TO BUILD IT:
1. Open the 2026 Statement of Profit and Loss (most recent year in the PDF).
2. Use its ENTIRE particulars list VERBATIM as the master row structure (from Revenue From Operations down to Earnings Per Share).
3. Preserve indentation using LEADING SPACES:
   Major Section Headers (e.g. EXPENSES, TAX EXPENSE) = 0 leading spaces
   Line items (e.g. Revenue From Operations, Cost of Materials Consumed) = 2 leading spaces
   Sub-items (e.g. Current Tax, Deferred Tax, Basic EPS, Diluted EPS) = 4 leading spaces
4. The 2026 particulars list is the ONLY row structure. Do NOT add extra rows from older years.

STARTING POINT RULE:
- Ignore ALL entries above Sales or Net Sales (e.g. gross turnover / excise duty reconciliations prior to net sales).
- START the extraction strictly from "Revenue From Operations" (also known as Sales / Net Sales).

MAPPING & ACCOUNTING INTELLIGENCE RULES:

RULE 1 - Revenue from Operations:
  In older formats: Net Sales / Income from Operations / Sales of Products & Services.
  Map directly to Revenue From Operations.

RULE 2 - Other Income:
  In older formats: Other Income / Interest & Dividend Income / Miscellaneous Income.
  Map directly to Other Income under Total Income.

RULE 3 - Expenditure Line Items:
  Map older Schedule VI expense items into their respective 2026 Ind AS line item slots:
    a) Cost of Materials Consumed (Raw Materials Consumed, Packing Materials)
    b) Purchases of Stock-in-Trade (Traded Goods Purchased)
    c) Changes in Inventories of Finished Goods, Work-in-Progress and Stock-in-Trade (Stock Differential / Inventory Changes)
    d) Employee Benefits Expense (Salaries, Wages, Bonus, Staff Welfare, Provident Fund)
    e) Finance Costs (Interest Expense, Borrowing Costs)
    f) Depreciation and Amortization Expense (Depreciation on assets)
    g) Other Expenses (Freight, Manufacturing Costs, Power & Fuel, Rent, Rates & Taxes, Repairs, Advertising, Selling expenses, Loss on Sale of Assets, etc.)
  UNMATCHED EXPENDITURE: If any expenditure line item in older years cannot be matched to a specific category above, ADD/COMBINE its value into "Other Expenses".

RULE 4 - Post-Expense Line Items (Total Expenses down to EPS):
  Use your accounting intelligence and domain knowledge to map all pre-Ind AS line items after Total Expenses to their corresponding 2026 Ind AS slots:
    a) Profit / (Loss) before exceptional items and tax = Total Income minus Total Expenses.
    b) Exceptional Items = Exceptional Items / Extra-ordinary items.
    c) Profit / (Loss) before tax (PBT) = Profit before tax.
    d) Tax Expense:
       - Current Tax (Provision for Income Tax / Tax for current year)
       - Deferred Tax (Deferred Tax Charge / Credit)
       - Prior Period / Earlier Years Tax Adjustments -> combine into Tax Expense / Current Tax as appropriate.
    e) Profit / (Loss) for the period (PAT / Net Profit) = Net Profit after tax.
    f) Other Comprehensive Income (OCI) = Items that will or will not be reclassified to profit or loss (only present in Ind AS years; empty string for pre-Ind AS years).
    g) Total Comprehensive Income = Net Profit + OCI (only present in Ind AS years; empty string for pre-Ind AS years).
    h) Earnings Per Share (EPS):
       - Basic EPS (Rs.)
       - Diluted EPS (Rs.)
       Note: Adjust for stock splits / bonus issues if already reflected in historical reported figures; otherwise copy exact reported EPS numbers.

GENERAL EXTRACTION RULES:
- Singular vs plural variations (e.g. Expense vs Expenses) treat as same.
- Copy all numeric values EXACTLY as written in the PDF (preserve reported unit, crores/lakhs, do NOT convert units).
- If a line item does not exist for a specific year, use an empty string "".
- Section header rows with no numeric value must use empty string "" for all years.
- Read EVERY page, ensure no page or year is missed.

FINAL REMINDER:
Return ONLY the raw JSON object. Nothing before it. Nothing after it.
"""


def upload_pdf(client, pdf_path):
    """Upload PDF and wait until Gemini has finished processing it."""
    print("=" * 62)
    print("  FINANCIAL EXTRACTOR - ITC Ltd. (P&L Statement)")
    print("=" * 62)
    print(f"\n[1/4] Uploading PDF : {os.path.basename(pdf_path)}")
    print(f"      File size     : {os.path.getsize(pdf_path) / (1024 * 1024):.2f} MB")

    uploaded = client.files.upload(file=pdf_path)
    print(f"      Uploaded as   : {uploaded.name}")

    print("[2/4] Waiting for Gemini to process the PDF", end="", flush=True)
    while uploaded.state.name == "PROCESSING":
        time.sleep(3)
        uploaded = client.files.get(name=uploaded.name)
        print(".", end="", flush=True)
    print()

    if uploaded.state.name == "FAILED":
        raise RuntimeError("Gemini failed to process the PDF file. Please try again.")

    print(f"      Status        : {uploaded.state.name}  OK")
    return uploaded


def call_gemini_once(client, uploaded_file, model_name):
    """
    Send the PDF + master prompt to Gemini in a single generate_content call.
    Automatically retries on 503 (server overload) up to MAX_RETRIES times.
    """
    print(f"\n[3/4] Sending the single API call to {model_name}")
    print("      Please wait - the model is reading every page of the PDF...")

    last_error = None

    for attempt in range(1, MAX_RETRIES + 1):
        if attempt > 1:
            print(f"      [Retry {attempt}/{MAX_RETRIES}] Server was busy. Waiting {RETRY_WAIT_SEC}s before retrying...")
            time.sleep(RETRY_WAIT_SEC)
            print(f"      [Retry {attempt}/{MAX_RETRIES}] Sending request again...")

        try:
            response = client.models.generate_content(
                model=model_name,
                contents=[uploaded_file, MASTER_PROMPT],
                config=types.GenerateContentConfig(
                    temperature=0,
                    response_mime_type="application/json",
                ),
            )
            raw = response.text.strip()
            print(f"      Response received  ({len(raw):,} characters)  OK")
            return raw

        except Exception as exc:
            error_str = str(exc)
            if "503" in error_str or "UNAVAILABLE" in error_str:
                print(f"      Server overloaded (503). Will retry...")
                last_error = exc
                continue
            else:
                raise

    raise RuntimeError(
        f"Gemini server returned 503 (overloaded) on all {MAX_RETRIES} attempts.\n"
        f"The model {model_name} is experiencing high demand.\n"
        f"Please try running the script again in a few minutes.\n"
        f"Last error: {last_error}"
    )


def parse_response(raw_text):
    """Parse the JSON from Gemini response, stripping markdown fences if present."""
    text = raw_text.strip()

    if text.startswith("```"):
        lines = text.split("\n")
        text = "\n".join(lines[1:-1]).strip()

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        debug_path = os.path.join(BASE_DIR, "debug_raw_pnl_response.txt")
        with open(debug_path, "w", encoding="utf-8") as fh:
            fh.write(raw_text)
        raise ValueError(
            f"Could not parse Gemini response as JSON.\n"
            f"Raw response saved to: {debug_path}\n"
            f"JSON error: {exc}"
        ) from exc

    if "years" not in data or "rows" not in data:
        raise ValueError("Gemini JSON is missing the required years or rows keys.")

    return data


def write_csv(data, output_path):
    """Write the consolidated multi-year P&L statement to a UTF-8 CSV file."""
    years = [str(y) for y in data["years"]]
    rows  = data["rows"]

    print(f"\n[4/4] Writing CSV output...")
    print(f"      Years  : {years[0]} to {years[-1]}  ({len(years)} columns)")
    print(f"      Rows   : {len(rows)}")

    with open(output_path, "w", newline="", encoding="utf-8-sig") as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(["Particulars"] + years)
        for row in rows:
            particular = row.get("particular", "")
            values     = row.get("values", {})
            row_data   = [particular] + [values.get(y, "") for y in years]
            writer.writerow(row_data)

    print(f"\n{'=' * 62}")
    print(f"  SUCCESS! File saved:")
    print(f"  {output_path}")
    print(f"{'=' * 62}\n")


def main():
    if not os.path.isfile(PDF_PATH):
        print(f"\nERROR: PDF not found:\n  {PDF_PATH}")
        sys.exit(1)

    client = genai.Client(api_key=API_KEY)

    try:
        uploaded_file = upload_pdf(client, PDF_PATH)
        raw_response  = call_gemini_once(client, uploaded_file, MODEL_NAME)
        data          = parse_response(raw_response)
        write_csv(data, OUTPUT_PATH)
    except Exception as exc:
        print(f"\nERROR: {exc}")
        sys.exit(1)
    finally:
        client.close()


if __name__ == "__main__":
    main()
