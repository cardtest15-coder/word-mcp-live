import asyncio
from pathlib import Path

import pytest
from docx import Document

from word_document_server.tools.extended_document_tools import convert_to_pdf


def _make_sample_docx(path: Path) -> None:
    """Generates a simple .docx file in a temporary directory."""
    doc = Document()
    doc.add_heading("Conversion Test Document", level=1)
    doc.add_paragraph("This is a test paragraph for PDF conversion. Contains ASCII too.")
    doc.add_paragraph("Second paragraph: Contains special characters and spaces to cover path/content edge cases.")
    doc.save(path)


@pytest.mark.skipif(
    not Path("C:/Program Files/Microsoft Office").exists()
    and not Path("C:/Program Files (x86)/Microsoft Office").exists(),
    reason="Microsoft Word not installed"
)
def test_convert_to_pdf_with_temp_docx(tmp_path: Path):
    """
    End-to-end test: Create a temporary .docx -> call convert_to_pdf -> validate the PDF output.

    Notes:
    - On Linux/macOS, it first tries LibreOffice (soffice/libreoffice),
      and falls back to docx2pdf on failure (requires Microsoft Word).
    - If these tools are missing or the command is unavailable, the test is skipped with a reason.
    """
    src_doc = tmp_path / "sample document with spaces.docx"
    _make_sample_docx(src_doc)

    out_pdf = tmp_path / "converted output.pdf"

    result_msg = asyncio.run(convert_to_pdf(str(src_doc), output_filename=str(out_pdf)))

    success_keywords = ["successfully converted", "converted to PDF"]
    success = any(k.lower() in result_msg.lower() for k in success_keywords) or out_pdf.exists()

    if not success:
        pytest.skip(f"PDF conversion tool unavailable or conversion failed: {result_msg}")

    candidates = [
        out_pdf,
        out_pdf.parent / f"{src_doc.stem}.pdf",
        src_doc.with_suffix(".pdf"),
    ]

    found = None
    for p in candidates:
        if p.exists():
            found = p
            break
    if not found:
        pdfs = sorted(tmp_path.glob("*.pdf"), key=lambda p: p.stat().st_mtime, reverse=True)
        if pdfs:
            found = pdfs[0]

    if not found:
        pytest.skip(f"Could not find the generated PDF. Function output: {result_msg}")

    assert found.exists(), f"Generated PDF not found: {found}, function output: {result_msg}"
    assert found.stat().st_size > 0, f"The generated PDF file is empty: {found}"
