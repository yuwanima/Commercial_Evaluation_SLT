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
            # An LKR quotation belongs on the LKR side of the column, not under (USD).
            self.assertIsNone(sheet["O9"].value)
            self.assertEqual(sheet["P9"].value, 100)

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

    def test_price_output_header_and_column_widths(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "price.docx"
            doc = Document()
            doc.add_paragraph("TENDER FOR PURCHASING OF OPTICAL FIBER CABLES")
            doc.add_paragraph("Tender No: SLT/PROC/2026/099")
            doc.add_paragraph("Official Frank: Acme Power (Pvt) Ltd")
            table = doc.add_table(rows=2, cols=8)
            headers = [
                "S/No", "Description", "Qty",
                "Unit Price (Rs.) [Excluding Taxes]", "VAT (%)", "NBT (%)",
                "Other Taxes (%)", "Unit Price (Rs.) [Including Taxes]",
            ]
            values = ["1", "48-Core Aerial Armoured Optical Fiber Cable", "10", "1500000", "18", "0", "0", "1770000"]
            for cell, value in zip(table.rows[0].cells, headers):
                cell.text = value
            for cell, value in zip(table.rows[1].cells, values):
                cell.text = value
            doc.save(path)

            result = pricesched.run([str(path)], Path(directory) / "out")
            self.assertIn("SLT/PROC/2026/099", result["tenders"])
            self.assertIn("OPTICAL FIBER CABLES", result["purchasings"])

            from openpyxl import load_workbook
            workbook = load_workbook(result["file"], data_only=False)
            sheet = workbook["Acme Power (Pvt) Ltd"]

            # Header inspection
            self.assertIn("OPTICAL FIBER CABLES", sheet["B2"].value)
            self.assertIn("SLT/PROC/2026/099", sheet["B2"].value)

            # Column width inspection (e.g., Description and Unit Price columns are wide enough)
            self.assertGreaterEqual(sheet.column_dimensions["E"].width, 30)
            self.assertGreaterEqual(sheet.column_dimensions["O"].width, 20)
            self.assertGreaterEqual(sheet.column_dimensions["U"].width, 20)

            # Amount alignment inspection
            self.assertEqual(sheet["P9"].alignment.horizontal, "right")
            self.assertEqual(sheet["U9"].alignment.horizontal, "right")


    def test_sscl_and_vat_columns_hold_amounts_not_percentages(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "local.docx"
            doc = Document()
            doc.add_paragraph("Official Frank: Acme Power (Pvt) Ltd")
            table = doc.add_table(rows=2, cols=8)
            headers = [
                "S/No", "Description", "Qty",
                "Unit Price (Rs.) [Excluding Taxes]", "VAT (%)", "NBT (%)",
                "Other Taxes (%)", "Unit Price (Rs.) [Including Taxes]",
            ]
            values = ["1", "Test item", "1", "100000", "18", "2", "0", "120000"]
            for cell, value in zip(table.rows[0].cells, headers):
                cell.text = value
            for cell, value in zip(table.rows[1].cells, values):
                cell.text = value
            doc.save(path)

            result = pricesched.run([str(path)], Path(directory) / "out")

            from openpyxl import load_workbook
            workbook = load_workbook(result["file"], data_only=False)
            sheet = workbook["Acme Power (Pvt) Ltd"]
            # Both levies are charged on the quoted price:
            # SSCL = 100,000 x 2% = 2,000;  VAT = 100,000 x 18% = 18,000
            self.assertIsNone(sheet["O9"].value)
            self.assertEqual(sheet["P9"].value, 100000)
            self.assertAlmostEqual(sheet["R9"].value, 2000)
            self.assertAlmostEqual(sheet["S9"].value, 18000)
            # T = with SSCL without VAT = 102,000;  U = with SSCL & VAT = 120,000
            self.assertAlmostEqual(sheet["T9"].value, 102000)
            self.assertAlmostEqual(sheet["U9"].value, 120000)


if __name__ == "__main__":
    unittest.main()
