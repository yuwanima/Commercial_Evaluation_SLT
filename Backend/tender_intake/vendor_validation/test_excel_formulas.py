import sys
import tempfile
from pathlib import Path

PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT))

from openpyxl import Workbook

from Backend.tender_intake.vendor_validation.document_extractor import (
    extract_document,
)


with tempfile.TemporaryDirectory() as temp_directory:
    excel_path = Path(temp_directory) / "formula_test.xlsx"

    workbook = Workbook()
    worksheet = workbook.active
    worksheet.title = "Evaluation"

    worksheet["A1"] = "Bid Amount"
    worksheet["B1"] = 100000
    worksheet["A2"] = "Calculated Total"
    worksheet["B2"] = "=SUM(B1:B1)"

    workbook.save(excel_path)

    result = extract_document(excel_path)

    extracted_text = " ".join(
        section["text"] for section in result["sections"]
    )

    assert "SUM(B1:B1)" in extracted_text
    assert "formula" in extracted_text
    assert "cached result unavailable" in extracted_text

    print("PASS: Excel formula preserved.")
    print("PASS: Missing cached result identified correctly.")
    print("\nExcel formula test passed.")
