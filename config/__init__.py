"""Configuration package for Universal Financial Extractor."""

from .settings import (
    BASE_DIR,
    INPUT_DIR,
    OUTPUT_DIR,
    GEMINI_API_KEY,
    DEFAULT_MODEL,
    FALLBACK_MODELS,
    MAX_RETRIES,
    RETRY_BACKOFF,
    BASE_CURRENCY_UNIT,
    TARGET_ACCOUNTING_STANDARD,
)

__all__ = [
    "BASE_DIR",
    "INPUT_DIR",
    "OUTPUT_DIR",
    "GEMINI_API_KEY",
    "DEFAULT_MODEL",
    "FALLBACK_MODELS",
    "MAX_RETRIES",
    "RETRY_BACKOFF",
    "BASE_CURRENCY_UNIT",
    "TARGET_ACCOUNTING_STANDARD",
]
