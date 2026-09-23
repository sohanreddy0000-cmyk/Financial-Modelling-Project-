"""
Main Execution Pipeline for Universal Financial Extractor.
Scans the input directory, invokes the Gemini extractor, and writes JSON outputs to disk.
"""

import json
import logging
from pathlib import Path
from config.settings import INPUT_DIR, OUTPUT_DIR
from extractors.gemini_extractor import GeminiFinancialExtractor

logging.basicConfig(level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")
logger = logging.getLogger(__name__)


def run_pipeline():
    """Runs the processing pipeline for all PDFs in data/input_pdfs/."""
    pdf_files = list(INPUT_DIR.glob("*.pdf"))
    
    if not pdf_files:
        logger.info(f"No PDF files found in '{INPUT_DIR}'. Please place PDF financial reports there.")
        return

    logger.info(f"Found {len(pdf_files)} PDF file(s) for processing.")
    extractor = GeminiFinancialExtractor()

    for pdf_file in pdf_files:
        logger.info(f"Starting pipeline processing for: {pdf_file.name}")
        output_file = OUTPUT_DIR / f"{pdf_file.stem}_extracted.json"

        try:
            results = extractor.extract_from_pdf(
                pdf_path=pdf_file,
                statement_type="Complete Financial Statements (Balance Sheet, P&L, Cash Flow)",
            )
            
            with open(output_file, "w", encoding="utf-8") as f:
                json.dump(results, f, indent=2, ensure_ascii=False)

            logger.info(f"Extraction successful! Output written to: {output_file}")

        except Exception as e:
            logger.error(f"Failed to process {pdf_file.name}: {e}")


if __name__ == "__main__":
    run_pipeline()
