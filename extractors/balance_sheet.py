"""
Financial Extractor - ITC Ltd.
Reads multi-year PDF balance sheets via a SINGLE Gemini API call.

Requirements:
    pip install google-genai

Usage:
    python financial_extractor.py
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

BASE_DIR    = r"C:\Users\sohan\Desktop\Financial Modelling Project"
PDF_PATH    = os.path.join(BASE_DIR, "ITC Balancesheet Compiler.pdf")
OUTPUT_PATH = os.path.join(BASE_DIR, "Financial Extractor.csv")

# How many times to retry if server returns 503 (overloaded), and how long to wait between retries
MAX_RETRIES      = 5
RETRY_WAIT_SEC   = 30   # seconds to wait before each retry

# MASTER PROMPT
MASTER_PROMPT = """
You are a senior financial analyst and data extraction expert specializing in Indian corporate balance sheets (both pre-Ind AS and post-Ind AS formats).

You have been given the full PDF of ITC Ltd compiled Annual Report balance sheets spanning multiple years.

YOUR TASK:
Read EVERY SINGLE PAGE of this PDF carefully like a human reading a document.
Extract ALL balance sheet data for ALL years present in the PDF.
Consolidate everything into ONE perfectly structured JSON object.

OUTPUT FORMAT (STRICT - follow exactly):
Return ONLY a valid raw JSON object. No explanation text. No markdown. No code fences. Just the JSON.

Structure:
{
  "years": [2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
  "rows": [
    {
      "particular": "EQUITY AND LIABILITIES",
      "values": {"2002": "", "2003": "", "2026": ""}
    },
    {
      "particular": "  Shareholders Funds",
      "values": {"2002": "XXXX.XX", "2003": "XXXX.XX", "2026": "XXXX.XX"}
    },
    {
      "particular": "    Share Capital",
      "values": {"2002": "XXXX.XX", "2026": "XXXX.XX"}
    }
  ]
}

PARTICULARS COLUMN - HOW TO BUILD IT:
1. Open the 2026 balance sheet (most recent year in the PDF).
2. Copy its ENTIRE particulars list VERBATIM with exact spelling and exact casing.
3. Preserve indentation using LEADING SPACES:
   Section headers (e.g. NON-CURRENT ASSETS) = 0 leading spaces
   Sub-section headers (e.g. Financial Assets) = 2 leading spaces
   Line items (e.g. Investments) = 4 leading spaces
   Further sub-items = 6 leading spaces
4. The 2026 particulars list is the ONLY row structure.
   Do NOT add extra rows from other years UNLESS instructed by the mapping rules.

YEAR COLUMNS - HOW TO FILL THEM:
- For each particular row, find the matching value for each year.
- Copy the numeric value EXACTLY as written in the PDF (same unit, crores or lakhs, do NOT convert).
- If a particular does not exist in a given year, use empty string.
- Section header rows with no numeric value, always use empty string for all years.

MAPPING RULES - How older year formats map to the 2026 Ind AS structure:

RULE 1 - Share Capital:
  Old sheets may show Share Capital and Share Capital Suspense separately.
  ADD both values together. Place the combined total under Share Capital.
  Do NOT show them as two separate rows.

RULE 2 - Borrowings (Non-Current Liabilities):
  Old format: Secured Loans and Unsecured Loans are separate items.
  COMBINE both totals. Place under Borrowings within Non-Current Liabilities.

RULE 3 - Property Plant and Equipment (PPE):
  Old format shows: Gross Block / Less Accumulated Depreciation / Net Block.
  IGNORE Gross Block entirely. IGNORE Accumulated Depreciation entirely.
  Take ONLY Net Block and map it to PPE / Property Plant and Equipment.

RULE 4 - Provisions (Non-Current Liabilities):
  Provision for assets given on lease (old format)
  Map to Provisions under Non-Current Liabilities.

RULE 5 - Current Investments / Loans:
  Loans and advances in current investments = Loans in current investments.
  They are the SAME item. Use whichever value is present.

RULE 6 - Current Liabilities and Provisions (old format):
  Old format grouped them as one section: Current Liabilities and Provisions.
  Split correctly: Trade payables and Other current liabilities go to Current Liabilities section.
  Provisions for leave, gratuity etc go to Current Provisions section.

RULE 7 - Deferred Tax:
  Deferred Tax on the ASSET side of old sheet maps to Deferred Tax Assets (Net) under Non-Current Assets.
  Deferred Tax on the LIABILITY side of old sheet maps to Deferred Tax Liabilities (Net) under Non-Current Liabilities.

RULE 8 - Miscellaneous Expenditure:
  Miscellaneous Expenditure (old format) - ADD its value into Other Current Assets.

RULE 9 - Long Term Borrowings:
  Long Term Borrowings = Non-Current Borrowings. Treat identically.

RULE 10 - Terminology Equivalences:
  Long Term = Non-Current (always, without exception)
  Short Term = Current (always, without exception)
  Intangible Assets (old) maps to Other Intangible Assets.

RULE 11 - Goodwill:
  Goodwill on Consolidation (old) maps to Goodwill.

RULE 12 - Investments (Non-Current Financial Assets):
  Combine ALL three old items into ONE Investments row under Financial Assets in Non-Current Assets:
  Investment in Associates + Investment in Joint Ventures + Other Investments
  ADD all three together and place under Investments (Non-Current Financial Assets).

GENERAL RULES:
- Singular vs plural (e.g. Loan vs Loans) treat as same item.
- Minor naming variations that are clearly the same, treat as same.
- Everything NOT covered above, copy values verbatim as they appear.
- Do NOT skip any year. Do NOT summarize, round, or approximate any value.
- Read EVERY page, no page should be missed.

FINAL REMINDER:
Return ONLY the raw JSON object. Nothing before it. Nothing after it.
"""


def upload_pdf(client, pdf_path):
    """Upload PDF and wait until Gemini has finished processing it."""
    print("=" * 62)
    print("  FINANCIAL EXTRACTOR - ITC Ltd.")
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
    NOTE: This is still ONE logical AI call - we only retry if the server was
    too busy to even process the request (503 = server refused, not AI output).
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
            # Only retry on 503 (server overload) - everything else is a real error
            if "503" in error_str or "UNAVAILABLE" in error_str:
                print(f"      Server overloaded (503). Will retry...")
                last_error = exc
                continue
            else:
                # Real error - do not retry
                raise

    # All retries exhausted
    raise RuntimeError(
        f"Gemini server returned 503 (overloaded) on all {MAX_RETRIES} attempts.\n"
        f"The model gemini-3.8-flash is experiencing very high demand.\n"
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
        debug_path = os.path.join(BASE_DIR, "debug_raw_response.txt")
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
    """Write the consolidated multi-year balance sheet to a UTF-8 CSV file."""
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
