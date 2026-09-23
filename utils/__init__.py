"""
Prompt Builder Module for Universal Financial Extractor.
Generates structured system and user prompts for extracting financial statements into standardized schemas.
"""

import json
from typing import Dict, Any, Optional
from config.settings import BASE_CURRENCY_UNIT, TARGET_ACCOUNTING_STANDARD


class FinancialPromptBuilder:
    """Constructs tailored prompts for extracting balance sheets, income statements, and cash flows."""

    def __init__(
        self,
        currency_unit: str = BASE_CURRENCY_UNIT,
        accounting_standard: str = TARGET_ACCOUNTING_STANDARD,
    ):
        self.currency_unit = currency_unit
        self.accounting_standard = accounting_standard

    def build_extraction_prompt(self, statement_type: str = "General Financial Statements") -> str:
        """
        Generates a comprehensive prompt instructing Gemini to parse tabular data from PDFs
        and return clean, strictly typed JSON output.
        """
        prompt = f"""
You are an expert financial analyst and data extraction system specializing in parsing complex corporate financial reports.

YOUR TASK:
Extract the financial data corresponding to '{statement_type}' from the attached PDF document(s).

EXTRACTION RULES & STANDARDS:
1. **Accounting Standard Compliance**: Normalize line items according to {self.accounting_standard}.
2. **Currency Standard**: Express all financial metrics in **{self.currency_unit}**. If original numbers are in thousands, lakhs, or millions, convert them accurately to {self.currency_unit}.
3. **Data Integrity**:
   - Extract exact values without guessing or halluncinating.
   - Retain precision up to 2 decimal places where applicable.
   - If a line item is explicitly missing or not reported, set its value to `null`.
   - Maintain historical comparability across provided periods (e.g., Current Year, Previous Year).
4. **Structural Normalization**:
   - Map variations of line item names to standard financial terms (e.g., "Revenue from Operations" vs "Turnover").

OUTPUT FORMAT:
Provide the final output STRICTLY as a valid JSON object matching the following structure:

```json
{{
  "metadata": {{
    "company_name": "String or null",
    "reporting_period": "String or null",
    "currency_unit": "{self.currency_unit}",
    "accounting_standard": "{self.accounting_standard}"
  }},
  "financial_data": {{
    "statement_type": "{statement_type}",
    "line_items": [
      {{
        "category": "String (e.g., Assets, Liabilities, Revenue, Expenses)",
        "standardized_name": "String",
        "original_label": "String",
        "current_period_value": number_or_null,
        "previous_period_value": number_or_null,
        "notes_reference": "String or null"
      }}
    ]
  }},
  "extraction_notes": ["Array of strings highlighting flags, anomalies, or currency conversions made"]
}}
