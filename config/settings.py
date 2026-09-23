"""
Global Configuration Settings for Universal Financial Extractor.
Handles environment variables, API keys, model parameters, and directory paths safely.
"""

import os
from pathlib import Path

# Base Directory Setup
BASE_DIR = Path(__file__).resolve().parent.parent
INPUT_DIR = BASE_DIR / "data" / "input_pdfs"
OUTPUT_DIR = BASE_DIR / "data" / "output"

# Ensure runtime input/output directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Google Gemini API Settings
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")

# Gemini Model Selection & Fallback Priority
DEFAULT_MODEL = "gemini-1.5-flash"
FALLBACK_MODELS = [
    "gemini-1.5-pro",
    "gemini-1.0-pro",
]

# API Execution & Server Retry Parameters
MAX_RETRIES = 5
RETRY_BACKOFF = 2.0  # Exponential backoff factor in seconds
TIMEOUT_SECONDS = 300  # Timeout for multi-PDF single-call uploads

# Financial Normalization Standards
BASE_CURRENCY_UNIT = "INR Crores"
TARGET_ACCOUNTING_STANDARD = "2026 Ind AS"
