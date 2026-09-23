# Universal Indian Corporate Financial Extractor 📈🤖

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Python: 3.9+](https://img.shields.io/badge/Python-3.9+-blue.svg)](https://www.python.org/downloads/)
[![Powered By: Google Gemini API](https://img.shields.io/badge/AI-Google%20Gemini%20Flash-green.svg)](https://ai.google.dev/)

An enterprise-grade, AI-driven open-source extraction and financial modeling engine designed to parse, normalize, and consolidate multi-year Indian Corporate Annual Report PDFs (2002–2026) into structured financial statements.

Driven by the official `google-genai` SDK and Gemini Flash models, this tool handles complex accounting format transitions (Pre-Ind AS / Schedule VI to 2026 Ind AS standards), unit conversions (Lakhs/Millions to INR Crores), and restatement priority logic across multi-decade corporate filings.

---

## 🚀 Key Features

* **Single-Call & Multi-PDF Batch Ingestion**: Ingests multi-year Annual Report PDFs in a single logical API request using Gemini Files API with fallback model switching (`gemini-3.6-flash`, `gemini-3.7-flash`, `gemini-3.8-flash`).
* **Full Financial Statement Suite**: Automatically extracts, maps, and reconciles:
  1. Consolidated Balance Sheet
  2. Statement of Profit & Loss (P&L)
  3. Cash Flow Statement
* **Autonomous Accounting Mapping Engine**:
  * Maps historical Schedule VI & pre-Ind AS line items verbatim onto a master 2026 Ind AS particulars structure.
  * Reconciles legacy line items (e.g., Gross Block/Depreciation to Net Block PPE, Share Capital Suspense, Secured/Unsecured Borrowings, Deferred Tax assets/liabilities).
* **Unit Normalization**: Automatically detects reporting units (₹ in Lakhs, Millions, or Thousands) and converts all historical figures into uniform **₹ INR Crores**.
* **Restated Data Priority**: Implements forensic accounting priority rules, prioritizing restated previous-year figures to capture post-balance-sheet audit adjustments.
* **Accounting Decision Audit Log**: Generates a year-by-year audit trail documenting mapping choices, unit conversions, and restatement selections.
* **Professional Excel & CSV Exports**: Produces a styled 5-tab Master Excel Workbook (`.xlsx`) with custom financial formatting alongside standalone CSV files.

---

## 📁 Repository Structure

```
financial-extractor/
├── config/
│   ├── __init__.py
│   └── settings.py              # API key, path configurations, and model selection
├── extractors/
│   ├── __init__.py
│   ├── balance_sheet.py          # Single-statement Balance Sheet extractor
│   ├── pnl_extractor.py          # Single-statement P&L extractor
│   └── universal_engine.py       # Multi-statement batch engine (BS + P&L + Cash Flow)
├── utils/
│   ├── __init__.py
│   ├── excel_formatter.py        # openpyxl styling engine for master workbooks
│   └── prompt_builder.py         # Master prompt constructor & accounting mapping rules
├── samples/
│   └── sample_audit_log.txt      # Example output audit trail
├── main.py                       # CLI execution entry point
├── requirements.txt              # Project dependencies
├── .gitignore                    # Prevents leaking API keys, PDFs, or generated outputs
├── LICENSE                       # MIT Open Source License
└── README.md                     # Project documentation

```
## 🛠️ Tech Stack & Requirements

* **Language**: Python 3.9+
* **AI Engine**: `google-genai` (Google Gemini 3.6 / 3.7 / 3.8 Flash)
* **Data Processing**: `pandas`, `openpyxl`
* **PDF Handling**: Google Gemini Files API

## ⚡ Quick Start & Installation

#### 1. Clone the Repository
```
git clone https://github.com/sohanreddy0000-cmyk/Financial-Modelling-Project-.git
cd Financial-Modelling-Project-
```
#### 2. Install Dependencies
```
pip install -r requirements.txt
```
#### 3. Set Up Configuration

Set your Gemini API key as an environment variable or update config/settings.py:
```
export GEMINI_API_KEY="your_google_gemini_api_key_here"
```
#### 4. Run the Extractor

Place your target company annual report PDFs in the input directory and execute:
```
python main.py
```
## 📊 Output Deliverables

The extractor generates four structured output files:

1. Company_Financial_Statements_Master.xlsx: Master Excel workbook containing:
        Balance Sheet Tab
        Profit & Loss Tab
        Cash Flow Statement Tab
        Audit Log & Accounting Notes Tab
2. Company_Balance_Sheet.csv
3. Company_Profit_Loss.csv
4. Company_Cash_Flow.csv
5. Company_Audit_Log.txt


## 📜 License

Distributed under the MIT License. See LICENSE for more information.


