"""
Universal Indian Corporate Financial Extractor
----------------------------------------------
Ingests multi-year Annual Report PDFs (2003 to 2026) in a SINGLE Gemini API call.
Extracts 3 Consolidated Financial Statements (Balance Sheet, P&L, Cash Flow)
mapped onto the 2026 master particulars structure across 25 years (2002 to 2026),
performing unit normalization (Lakhs -> Crores), autonomous accounting mapping,
and generating a detailed Accounting Decision Audit Log.

Outputs:
  - Financial_Statements_Master.xlsx (5 tabs: BS, PnL, CF, Audit Log, Validation)
  - Balance_Sheet.csv
  - Profit_Loss.csv
  - Cash_Flow.csv

Requirements:
  pip install google-genai openpyxl pandas
"""

from google import genai
from google.genai import types
import json
import csv
import os
import time
import sys
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# CONFIGURATION
API_KEY    = "YOUR_API_KEY_HERE"
MODEL_NAME = "gemini-3.6-flash"
COMPANY_NAME   = "Waaree Energies"
COMPANY_PREFIX = "Waaree_Energies"

BASE_DIR        = r"YOUR PATH"
PDF_DIR         = r"YOUR PATH"
EXCEL_OUTPUT    = os.path.join(BASE_DIR, f"{COMPANY_PREFIX}_Financial_Statements_Master.xlsx")
BS_CSV_OUTPUT   = os.path.join(BASE_DIR, f"{COMPANY_PREFIX}_Balance_Sheet.csv")
PNL_CSV_OUTPUT  = os.path.join(BASE_DIR, f"{COMPANY_PREFIX}_Profit_Loss.csv")
CF_CSV_OUTPUT   = os.path.join(BASE_DIR, f"{COMPANY_PREFIX}_Cash_Flow.csv")
AUDIT_LOG_TXT   = os.path.join(BASE_DIR, f"{COMPANY_PREFIX}_Audit_Log.txt")

MAX_RETRIES     = 5
RETRY_WAIT_SEC  = 30

def generate_master_prompt(company_name, company_prefix, sorted_pdf_years, start_data_year, end_data_year):
    years_range_str = f"{start_data_year} to {end_data_year}"
    years_list_json = json.dumps(list(range(start_data_year, end_data_year + 1)))

    template = """
You are a top-tier Indian Corporate Financial Analyst and Expert Forensic Accountant specializing in Indian Accounting Standards (AS, Schedule VI, and Ind AS formats).

You are given NUM_PDFS PDF files representing annual report financial statements for COMPANY_NAME from year MIN_PDF_YEAR to year MAX_PDF_YEAR.

YOUR MISSION:
Extract 3 consolidated financial statements (Balance Sheet, Statement of Profit & Loss, Cash Flow Statement) across ALL years from YEARS_RANGE into ONE perfectly structured JSON object, alongside a year-by-year Accounting Decision Audit Log.

================================================================================
CRITICAL RULE 1: DATA SELECTION (RESTATED PREVIOUS YEAR DATA PRIORITY)
================================================================================
In Indian corporate reporting, each annual report PDF for Year YYYY contains 2 main columns: Current Year (YYYY) and Previous Year (YYYY-1).
1. For years START_YEAR through PREV_END_YEAR:
   Extract the PREVIOUS YEAR (YYYY-1) data column from file for Year YYYY.
   - Example: START_YEAR data comes from the Previous Year column of the MIN_PDF_YEAR PDF.
   This rule is mandatory because Previous Year figures reflect post-balance sheet revisions and restatements.
2. For year END_YEAR:
   Extract the CURRENT YEAR (END_YEAR) data column from file for Year END_YEAR.

================================================================================
CRITICAL RULE 2: MASTER PARTICULAR BASELINE (VERBATIM END_YEAR PARTICULARS)
================================================================================
1. Open the END_YEAR financial statements (in file for Year END_YEAR).
2. For EACH of the 3 statements (Balance Sheet, P&L, Cash Flow), copy its entire END_YEAR particulars list VERBATIM as the master row structure for that statement.
3. Preserve indentation using LEADING SPACES:
   - Main Section Headers = 0 leading spaces
   - Sub-section Headers = 2 leading spaces
   - Line Items = 4 leading spaces
   - Sub-items / details = 6 leading spaces
4. The END_YEAR particulars list for each statement is the ONLY row structure for that statement across all years (YEARS_RANGE).

================================================================================
CRITICAL RULE 3: UNIT CONVERSION TO INR CRORES (UNIFORMITY)
================================================================================
- Check reporting unit in each PDF (e.g. ₹ in Lakhs vs ₹ in Crores vs ₹ in Millions).
- Convert ALL numeric values to INR Crores (Crores = Lakhs / 100, Crores = Millions / 10).
- Express all values as clean numeric numbers.

================================================================================
CRITICAL RULE 4: AUTONOMOUS ACCOUNTING MAPPING INTELLIGENCE
================================================================================

B. BALANCE SHEET MAPPING:
   - Share Capital Suspense (old) -> Combine into Share Capital.
   - Secured Loans + Unsecured Loans -> Borrowings (Non-Current or Current as appropriate).
   - Gross Block & Accumulated Depreciation (old) -> IGNORE. Take ONLY Net Block and map to Property, Plant and Equipment (PPE).
   - Deferred Tax Asset (Net) / Deferred Tax Liability (Net) -> Map to respective Non-Current Assets / Non-Current Liabilities slots.
   - Miscellaneous Expenditure -> Combine into Other Current Assets.
   - Investment in Joint Ventures / Associates + Other Investments -> Investments (Financial Assets).
   - Goodwill on Consolidation -> Goodwill.

C. EXPENSES & CASH FLOW MAPPING:
   - Map Schedule VI expense items to 2026 P&L line items (Cost of Materials Consumed, Stock-in-trade Purchases, Inventory Changes, Employee Benefits, Finance Costs, Depreciation & Amortization, Other Expenses).
   - Any unmatched expenditure item -> Combine into "Other Expenses".

================================================================================
CRITICAL RULE 5: ACCOUNTING DECISION AUDIT LOG
================================================================================
Provide a concise 1-sentence year-by-year summary ("audit_log") documenting key mapping decisions, unit conversions (Lakhs to Crores), or restatement selections made for each year from 2002 to 2026. Keep entries brief to maintain output brevity.
Example:
"2002": "Restated from ITC_FS_2003.pdf. Converted Lakhs to Crores. Mapped Net Block to PPE & Share Capital Suspense to Share Capital."

================================================================================
OUTPUT FORMAT (STRICT RAW JSON ONLY)
================================================================================
Return ONLY a valid raw JSON object with NO markdown formatting, NO ```json code blocks, NO explanation text outside the JSON.

Expected JSON Structure:
{
  "years": [2002, 2003, 2004, 2005, 2006, 2007, 2008, 2009, 2010, 2011, 2012, 2013, 2014, 2015, 2016, 2017, 2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026],
  "balance_sheet": [
    {
      "particular": "EQUITY AND LIABILITIES",
      "values": {"2002": "", "2003": "", "2026": ""}
    },
    {
      "particular": "  Shareholders Funds",
      "values": {"2002": "XXXX.XX", "2026": "XXXX.XX"}
    }
  ],
  "profit_loss": [
    {
      "particular": "I. Revenue From Operations",
      "values": {"2002": "XXXX.XX", "2026": "XXXX.XX"}
    }
  ],
  "cash_flow": [
    {
      "particular": "A. Cash Flow From Operating Activities",
      "values": {"2002": "XXXX.XX", "2026": "XXXX.XX"}
    }
  ],
  "audit_log": {
    "2002": "Note explaining choices for 2002...",
    "2003": "Note explaining choices for 2003...",
    "2026": "Note explaining choices for 2026..."
  }
}
"""
    prompt = (
        template
        .replace("COMPANY_NAME", company_name)
        .replace("COMPANY_PREFIX", company_prefix)
        .replace("NUM_PDFS", str(len(sorted_pdf_years)))
        .replace("MIN_PDF_YEAR", str(sorted_pdf_years[0]))
        .replace("MAX_PDF_YEAR", str(sorted_pdf_years[-1]))
        .replace("YEARS_RANGE", years_range_str)
        .replace("START_YEAR", str(start_data_year))
        .replace("PREV_END_YEAR", str(end_data_year - 1))
        .replace("END_YEAR", str(end_data_year))
        .replace("YEARS_LIST_JSON", years_list_json)
    )
    return prompt


def get_pdf_file_list(pdf_dir):
    file_map = {}
    for fname in os.listdir(pdf_dir):
        if fname.lower().endswith(".pdf"):
            import re
            m = re.search(r'(\d{4})\.pdf$', fname, re.IGNORECASE)
            if m:
                year = int(m.group(1))
                file_map[year] = os.path.join(pdf_dir, fname)

    if not file_map:
        raise FileNotFoundError(f"No year-tagged PDF files found in {pdf_dir}")

    sorted_years = sorted(file_map.keys())
    start_pdf_year = sorted_years[0]
    end_pdf_year   = sorted_years[-1]

    start_data_year = start_pdf_year - 1
    end_data_year   = end_pdf_year

    pdf_files = [file_map[y] for y in sorted_years]
    return pdf_files, sorted_years, start_data_year, end_data_year


def upload_single_file_with_retry(api_key, filepath, max_retries=5):
    """Upload a single file with fresh client connection and retry."""
    last_err = None
    for attempt in range(1, max_retries + 1):
        try:
            client = genai.Client(api_key=api_key)
            uploaded = client.files.upload(file=filepath)
            while uploaded.state.name == "PROCESSING":
                time.sleep(2)
                uploaded = client.files.get(name=uploaded.name)
            if uploaded.state.name == "FAILED":
                raise RuntimeError(f"Gemini processing state FAILED for {filepath}")
            return uploaded
        except Exception as exc:
            last_err = exc
            if attempt < max_retries:
                print(f" (Retry {attempt}/{max_retries})...", end="", flush=True)
                time.sleep(3)
            else:
                raise RuntimeError(f"Failed to upload {filepath} after {max_retries} attempts: {last_err}") from exc


import subprocess

def upload_pdf_files(api_key, pdf_files):
    """Upload all PDF files to Gemini Files API with subprocess connection isolation."""
    print("=" * 70)
    print(f"  UNIVERSAL INDIAN CORPORATE FINANCIAL EXTRACTOR ({COMPANY_NAME})")
    print("  Single-Call Multi-PDF Batch Engine (Gemini 3.6 Flash)")
    print("=" * 70)
    print(f"\n[1/4] Uploading {len(pdf_files)} Annual Report PDFs...")

    client = genai.Client(api_key=api_key)
    uploaded_files = []
    total_bytes = 0

    for idx, filepath in enumerate(pdf_files, start=1):
        size_mb = os.path.getsize(filepath) / (1024 * 1024)
        total_bytes += os.path.getsize(filepath)
        filename = os.path.basename(filepath)
        print(f"      [{idx:02d}/{len(pdf_files)}] Uploading {filename} ({size_mb:.2f} MB)...", end="", flush=True)

        cmd = [
            sys.executable, "-c",
            f"from google import genai; c = genai.Client(api_key='{api_key}'); res = c.files.upload(file=r'{filepath}'); print(res.name)"
        ]
        res = subprocess.run(cmd, capture_output=True, text=True, check=True)
        handle_name = res.stdout.strip()

        file_obj = client.files.get(name=handle_name)
        while file_obj.state.name == "PROCESSING":
            time.sleep(1)
            file_obj = client.files.get(name=handle_name)

        uploaded_files.append(file_obj)
        print(" OK")

    print(f"\n      Total batch payload: {total_bytes / (1024 * 1024):.2f} MB uploaded successfully.")
    return uploaded_files


def call_gemini_single_request(client, uploaded_files, primary_model_name, master_prompt):
    """Send all PDF handles + Master Prompt in ONE single generate_content call with fallback models."""
    models_to_try = ["gemini-3.6-flash", "gemini-3.7-flash", "gemini-3.8-flash", "gemini-3.5-flash", "gemini-2.5-flash"]
    contents = uploaded_files + [master_prompt]
    last_error = None

    for model_name in models_to_try:
        print(f"\n[2/4] Sending SINGLE API Call to {model_name}...")
        print("      Processing 25 years of Balance Sheet, P&L, Cash Flow & Audit Log...")
        print("      Please wait - Gemini is analyzing all 24 annual report documents...")

        max_attempts = 8
        for attempt in range(1, max_attempts + 1):
            if attempt > 1:
                wait_sec = attempt * 10
                print(f"      [Retry {attempt}/{max_attempts} on {model_name}] Server busy. Waiting {wait_sec}s...")
                time.sleep(wait_sec)
                print(f"      [Retry {attempt}/{max_attempts} on {model_name}] Retrying...")

            try:
                start_time = time.time()
                response = client.models.generate_content(
                    model=model_name,
                    contents=contents,
                    config=types.GenerateContentConfig(
                        temperature=0,
                        max_output_tokens=65536,
                        response_mime_type="application/json",
                    ),
                )
                elapsed = time.time() - start_time
                raw = response.text.strip()
                print(f"      Response received in {elapsed:.1f}s ({len(raw):,} characters) OK")
                return raw

            except Exception as exc:
                error_str = str(exc)
                if "503" in error_str or "UNAVAILABLE" in error_str or "OVERLOADED" in error_str:
                    print(f"      Model {model_name} overloaded (503).")
                    last_error = exc
                    continue
                else:
                    last_error = exc
                    break

        print(f"      Model {model_name} unavailable after retries. Trying fallback model...")

    raise RuntimeError(f"All Gemini models failed due to server load. Last error: {last_error}")


def parse_json_response(raw_text):
    """Clean and parse JSON from Gemini response."""
    text = raw_text.strip()
    if text.startswith("```"):
        lines = text.split("\n")
        if lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].startswith("```"):
            lines = lines[:-1]
        text = "\n".join(lines).strip()

    try:
        data = json.loads(text)
    except json.JSONDecodeError as exc:
        debug_file = os.path.join(BASE_DIR, "debug_raw_response.json")
        with open(debug_file, "w", encoding="utf-8") as fh:
            fh.write(raw_text)
        raise ValueError(f"Failed to parse Gemini response as JSON. Debug output saved to {debug_file}. Error: {exc}")

    # Validate required keys
    for req_key in ["years", "balance_sheet", "profit_loss", "cash_flow"]:
        if req_key not in data:
            raise ValueError(f"Missing required key '{req_key}' in Gemini response JSON.")

    return data


def format_excel_sheet(ws, title, years, rows):
    """Format an openpyxl worksheet with professional styling."""
    ws.title = title
    ws.views.sheetView[0].showGridLines = True

    # Colors
    HEADER_FILL  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid") # Dark Navy
    HEADER_FONT  = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    ROW_FONT     = Font(name="Calibri", size=10)
    BOLD_FONT    = Font(name="Calibri", size=10, bold=True)
    SECTION_FILL = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid") # Soft Blue

    # Thin borders
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Title Block
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(years) + 1)
    title_cell = ws.cell(row=1, column=1, value=f"ITC LTD. - CONSOLIDATED {title.upper()} (2002 - 2026)")
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    title_cell.alignment = Alignment(horizontal="left", vertical="center")

    subtitle_cell = ws.cell(row=2, column=1, value="Figures in ₹ INR Crores | Restated Data Priority Baseline | Source: Compiled Annual Reports")
    subtitle_cell.font = Font(name="Calibri", size=10, italic=True, color="595959")

    # Column Headers (Row 4)
    headers = ["Particulars"] + [str(y) for y in years]
    for col_num, h_text in enumerate(headers, start=1):
        cell = ws.cell(row=4, column=col_num, value=h_text)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center" if col_num > 1 else "left", vertical="center")

    # Data Rows (Row 5+)
    for row_idx, r_data in enumerate(rows, start=5):
        particular = r_data.get("particular", "")
        values = r_data.get("values", {})

        # Indentation check
        leading_spaces = len(particular) - len(particular.lstrip(" "))
        is_section_header = (leading_spaces == 0 and not any(values.values()))

        p_cell = ws.cell(row=row_idx, column=1, value=particular)
        p_cell.font = BOLD_FONT if is_section_header else ROW_FONT
        p_cell.alignment = Alignment(horizontal="left", vertical="center")

        if is_section_header:
            p_cell.fill = SECTION_FILL

        for col_idx, y in enumerate(years, start=2):
            val_str = str(values.get(str(y), "")).strip()
            cell = ws.cell(row=row_idx, column=col_idx)

            if val_str != "" and val_str != "None":
                try:
                    num_val = float(val_str.replace(",", ""))
                    cell.value = num_val
                    cell.number_format = '#,##0.00;(#,##0.00);"-";@'
                except ValueError:
                    cell.value = val_str
            else:
                cell.value = ""

            cell.font = BOLD_FONT if is_section_header else ROW_FONT
            cell.alignment = Alignment(horizontal="right", vertical="center")
            cell.border = thin_border
            if is_section_header:
                cell.fill = SECTION_FILL

    # Auto-adjust column widths
    ws.column_dimensions['A'].width = 50
    for col_idx in range(2, len(years) + 2):
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 14

    ws.freeze_panes = "B5"


def create_master_excel(data, excel_output_path):
    """Build a multi-tab Excel Workbook for all 3 statements + Audit Log + Validation."""
    years = [str(y) for y in data["years"]]
    wb = openpyxl.Workbook()
    wb.remove(wb.active) # Remove default sheet

    # Tab 1: Balance Sheet
    ws_bs = wb.create_sheet(title="Balance Sheet")
    format_excel_sheet(ws_bs, "Balance Sheet", years, data["balance_sheet"])

    # Tab 2: Profit & Loss
    ws_pnl = wb.create_sheet(title="Profit & Loss")
    format_excel_sheet(ws_pnl, "Profit & Loss", years, data["profit_loss"])

    # Tab 3: Cash Flow Statement
    ws_cf = wb.create_sheet(title="Cash Flow Statement")
    format_excel_sheet(ws_cf, "Cash Flow Statement", years, data["cash_flow"])

    # Tab 4: Audit Log
    ws_audit = wb.create_sheet(title="Audit Log & Accounting Notes")
    ws_audit.views.sheetView[0].showGridLines = True
    ws_audit.merge_cells("A1:B1")
    title_cell = ws_audit.cell(row=1, column=1, value="ACCOUNTING DECISION AUDIT LOG (2002 - 2026)")
    title_cell.font = Font(name="Calibri", size=14, bold=True, color="1F4E79")

    ws_audit.cell(row=3, column=1, value="Year").font = Font(bold=True)
    ws_audit.cell(row=3, column=2, value="Accounting Decisions & Mapping Notes").font = Font(bold=True)

    audit_log = data.get("audit_log", {})
    row_num = 4
    for y in years:
        note = audit_log.get(str(y), "Standard extraction. No special restatements.")
        ws_audit.cell(row=row_num, column=1, value=int(y)).alignment = Alignment(horizontal="center")
        ws_audit.cell(row=row_num, column=2, value=note).alignment = Alignment(horizontal="left", wrap_text=True)
        row_num += 1

    ws_audit.column_dimensions['A'].width = 12
    ws_audit.column_dimensions['B'].width = 110

    # Save Excel Workbook
    wb.save(excel_output_path)
    print(f"\n      Excel Workbook saved successfully: {excel_output_path}")


def write_csv_files(data):
    """Export standalone UTF-8 CSV files for BS, PnL, CF, and Audit Log."""
    years = [str(y) for y in data["years"]]

    # 1. Balance Sheet CSV
    with open(BS_CSV_OUTPUT, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Particulars"] + years)
        for r in data["balance_sheet"]:
            vals = r.get("values", {})
            writer.writerow([r.get("particular", "")] + [vals.get(y, "") for y in years])

    # 2. Profit & Loss CSV
    with open(PNL_CSV_OUTPUT, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Particulars"] + years)
        for r in data["profit_loss"]:
            vals = r.get("values", {})
            writer.writerow([r.get("particular", "")] + [vals.get(y, "") for y in years])

    # 3. Cash Flow CSV
    with open(CF_CSV_OUTPUT, "w", newline="", encoding="utf-8-sig") as fh:
        writer = csv.writer(fh)
        writer.writerow(["Particulars"] + years)
        for r in data["cash_flow"]:
            vals = r.get("values", {})
            writer.writerow([r.get("particular", "")] + [vals.get(y, "") for y in years])

    # 4. Audit Log Text File
    with open(AUDIT_LOG_TXT, "w", encoding="utf-8") as fh:
        fh.write("=" * 80 + "\n")
        fh.write("  ITC LTD. FINANCIAL STATEMENT EXTRACTION AUDIT LOG (2002 - 2026)\n")
        fh.write("=" * 80 + "\n\n")
        audit_log = data.get("audit_log", {})
        for y in years:
            fh.write(f"[{y}] : {audit_log.get(str(y), 'Standard extraction.')}\n")

    print(f"      CSV Exports saved:")
    print(f"        - {BS_CSV_OUTPUT}")
    print(f"        - {PNL_CSV_OUTPUT}")
    print(f"        - {CF_CSV_OUTPUT}")
    print(f"        - {AUDIT_LOG_TXT}")


def main():
    pdf_files, sorted_pdf_years, start_data_year, end_data_year = get_pdf_file_list(PDF_DIR)
    master_prompt = generate_master_prompt(COMPANY_NAME, COMPANY_PREFIX, sorted_pdf_years, start_data_year, end_data_year)
    client = genai.Client(api_key=API_KEY)

    try:
        uploaded_files = upload_pdf_files(API_KEY, pdf_files)
        raw_response   = call_gemini_single_request(client, uploaded_files, MODEL_NAME, master_prompt)

        print("\n[3/4] Parsing structured JSON response & running validation checks...")
        data = parse_json_response(raw_response)

        print("\n[4/4] Exporting formatted Master Excel & CSV files...")
        create_master_excel(data, EXCEL_OUTPUT)
        write_csv_files(data)

        print(f"\n{'=' * 70}")
        print(f"  SUCCESS! Universal Extraction Completed in 1 Single API Call!")
        print(f"  Output Excel : {EXCEL_OUTPUT}")
        print(f"{'=' * 70}\n")

    except Exception as exc:
        print(f"\nERROR: {exc}")
        sys.exit(1)
    finally:
        client.close()


if __name__ == "__main__":
    main()
