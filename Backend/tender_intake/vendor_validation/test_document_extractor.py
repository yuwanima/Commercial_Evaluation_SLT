import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT))

from docx import Document
from openpyxl import Workbook
from reportlab.pdfgen import canvas
from pypdf import PdfWriter

from Backend.tender_intake.vendor_validation.document_extractor import (
    extract_document,
)


def create_test_pdf(path):
    pdf = canvas.Canvas(str(path))
    pdf.drawString(72, 720, "Performance Security")
    pdf.drawString(72, 690, "The supplier shall submit security within 14 days.")
    pdf.showPage()
    pdf.save()


def create_empty_pdf(path):
    writer = PdfWriter()
    writer.add_blank_page(width=612, height=792)

    with path.open("wb") as file:
        writer.write(file)


def create_test_docx(path):
    document = Document()
    document.add_paragraph("Vendor Name: ABC Suppliers")
    table = document.add_table(rows=1, cols=2)
    table.cell(0, 0).text = "Clause"
    table.cell(0, 1).text = "Compliance"
    row = table.add_row()
    row.cells[0].text = "4.1"
    row.cells[1].text = "Complied"
    document.save(path)


def create_test_xlsx(path):
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Vendor Details"
    worksheet["A1"] = "Vendor Name"
    worksheet["B1"] = "ABC Suppliers"
    worksheet["A2"] = "Clause 4.1"
    worksheet["B2"] = "Complied"
    workbook.save(path)


def main():
    with tempfile.TemporaryDirectory() as temp_directory:
        folder = Path(temp_directory)

        pdf_path = folder / "sample_vendor.pdf"
        empty_pdf_path = folder / "empty_sample.pdf"
        docx_path = folder / "sample_vendor.docx"
        xlsx_path = folder / "sample_vendor.xlsx"

        create_test_pdf(pdf_path)
        create_empty_pdf(empty_pdf_path)
        create_test_docx(docx_path)
        create_test_xlsx(xlsx_path)

        # PDF test
        pdf_result = extract_document(pdf_path)
        pdf_text = " ".join(
            section["text"] for section in pdf_result["sections"]
        )

        assert "Performance Security" in pdf_text
        assert "14 days" in pdf_text
        assert pdf_result["sections"][0]["location"] == "Page 1"

        print("PASS: PDF text and page numbers extracted.")

        # Word test
        word_result = extract_document(docx_path)
        word_text = " ".join(
            section["text"] for section in word_result["sections"]
        )

        assert "ABC Suppliers" in word_text
        assert "Complied" in word_text
        assert any(
            section["location"].startswith("Table")
            for section in word_result["sections"]
        )

        print("PASS: Word paragraphs and tables extracted.")

        # Excel test
        excel_result = extract_document(xlsx_path)
        excel_text = " ".join(
            section["text"] for section in excel_result["sections"]
        )

        assert "ABC Suppliers" in excel_text
        assert "B2" in excel_text
        assert "Complied" in excel_text
        assert "Vendor Details" in str(excel_result["sections"])

        print("PASS: Excel worksheet names and cell references extracted.")

        # Empty/scanned PDF page test
        empty_result = extract_document(empty_pdf_path)

        assert len(empty_result["sections"]) == 1
        assert empty_result["sections"][0]["location"] == "Page 1"
        assert empty_result["sections"][0]["text"] == ""
        assert empty_result["sections"][0]["requires_ocr"] is True

        print("PASS: PDF page without readable text preserved and flagged for OCR.")

        print("\nAll four document extraction tests passed.")


if __name__ == "__main__":
    main()
