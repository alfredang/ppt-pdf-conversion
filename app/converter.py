"""Conversion pipeline: PDF -> Extract -> Translate -> PPTX."""

import os
from app.pdf_extractor import extract_pdf
from app.translator import Translator
from app.pptx_generator import create_presentation


def convert_pdf_to_pptx(pdf_path, output_path, progress_callback=None):
    """Convert a PDF file to PPTX with Chinese-to-English translation.

    Args:
        pdf_path: Path to input PDF file
        output_path: Path for output PPTX file
        progress_callback: Optional callable(message: str) for progress updates
    """
    def _progress(msg):
        if progress_callback:
            progress_callback(msg)

    # Phase 1: Extract PDF content
    _progress("Starting PDF extraction...")
    pages_data = extract_pdf(pdf_path, progress_callback=_progress)
    _progress(f"Extracted {len(pages_data)} pages")

    # Phase 2: Translate Chinese to English
    _progress("Starting translation...")
    translator = Translator()
    pages_data = translator.translate_pages(pages_data, progress_callback=_progress)
    _progress("Translation complete")

    # Phase 3: Generate PPTX
    _progress("Generating PPTX...")
    prs = create_presentation(pages_data, progress_callback=_progress)

    # Ensure output directory exists
    os.makedirs(os.path.dirname(output_path) or '.', exist_ok=True)
    prs.save(output_path)
    _progress("Conversion complete!")

    return output_path
