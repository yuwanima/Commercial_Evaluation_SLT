import tempfile
import unittest
from pathlib import Path

from docx import Document

from Backend.Price_Doc import pricesched


class PriceBidderIdentificationTests(unittest.TestCase):
    @staticmethod
    def create_price_document(path, frank_value=None, frank_label="Official Frank", description="Test item"):
        document = Document()
        if frank_value is not None:
            document.add_paragraph(f"{frank_label}: {frank_value}")
        table = document.add_table(rows=2, cols=8)
        headers = [
            "S/No", "Description", "Qty",
            "Unit Price (Rs.) [Excluding Taxes]", "VAT (%)", "NBT (%)",
            "Other Taxes (%)", "Unit Price (Rs.) [Including Taxes]",
        ]
        values = ["1", description, "1", "100", "0", "0", "0", "100"]
        for cell, value in zip(table.rows[0].cells, headers):
            cell.text = value
        for cell, value in zip(table.rows[1].cells, values):
            cell.text = value
        document.save(path)

    def test_official_frank_name_replaces_filename_name(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "anonymous.docx"
            self.create_price_document(path, "Acme Power (Pvt) Ltd")

            result = pricesched.run(
                [f"filename-derived={path}"],
                Path(directory) / "out",
            )

            self.assertEqual(result["rows"][0]["bidder"], "Acme Power (Pvt) Ltd")

    def test_blank_official_frank_keeps_filename_name(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "anonymous.docx"
            self.create_price_document(path)

            result = pricesched.run(
                [f"filename-derived={path}"],
                Path(directory) / "out",
            )

            self.assertEqual(result["rows"][0]["bidder"], "filename-derived")

    def test_official_frank_stamp_label_extracts_name_without_stamp_suffix(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "anonymous.docx"
            self.create_price_document(
                path,
                "Acme Power (Pvt) Ltd",
                frank_label="Official Frank / Stamp",
            )

            result = pricesched.run(
                [f"filename-derived={path}"],
                Path(directory) / "out",
            )

            self.assertEqual(result["rows"][0]["bidder"], "Acme Power (Pvt) Ltd")

    def test_gct_clause_text_is_not_a_foreign_price_header(self):
        self.assertFalse(
            pricesched._is_foreign_header(
                ["22", "CONFORMITY TO SPECIFICATIONS", "Complied", ""]
            )
        )
        self.assertTrue(
            pricesched._is_foreign_header(
                [
                    "S/No", "Description", "Qty", "Country of Origin",
                    "FOB / FCA Unit Price", "Total Freight",
                ]
            )
        )

    def test_price_output_uses_annexure_layout(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "anonymous.docx"
            self.create_price_document(path, "Acme Power (Pvt) Ltd")
            result = pricesched.run(
                [f"filename-derived={path}"],
                Path(directory) / "out",
            )

            from openpyxl import load_workbook
            workbook = load_workbook(result["file"], data_only=False)
            self.assertEqual(workbook.sheetnames, ["Acme Power (Pvt) Ltd"])
            sheet = workbook["Acme Power (Pvt) Ltd"]
            self.assertEqual(sheet["D9"].value, "Acme Power (Pvt) Ltd")
            self.assertEqual(sheet["E9"].value, "Test item")
            self.assertEqual(sheet["F9"].value, 1)
            self.assertEqual(sheet["O9"].value, 100)

    def test_price_output_has_separate_annexure_tab_per_company(self):
        with tempfile.TemporaryDirectory() as directory:
            first = Path(directory) / "first.docx"
            second = Path(directory) / "second.docx"
            self.create_price_document(first, "Acme Power (Pvt) Ltd", description="First item")
            self.create_price_document(second, "Beta Systems Ltd", description="Second item")
            result = pricesched.run(
                [f"first={first}", f"second={second}"],
                Path(directory) / "out",
            )

            from openpyxl import load_workbook
            workbook = load_workbook(result["file"], data_only=False)
            self.assertEqual(
                workbook.sheetnames,
                ["Acme Power (Pvt) Ltd", "Beta Systems Ltd"],
            )
            self.assertEqual(workbook["Acme Power (Pvt) Ltd"]["E9"].value, "First item")
            self.assertEqual(workbook["Beta Systems Ltd"]["E9"].value, "Second item")


if __name__ == "__main__":
    unittest.main()
