import datetime as dt
import tempfile
import unittest
from io import BytesIO
from unittest.mock import patch
from pathlib import Path

from docx import Document
from openpyxl import Workbook

from Backend.Price_Doc import pricesched
from Backend.Price_Doc import cbsl_web_rate
from Backend.Technical import convert


class PriceBidderIdentificationTests(unittest.TestCase):
    def test_cbsl_parser_returns_latest_rate_for_weekend_or_holiday(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.append(["Date", "Date", "United States", "United States"])
        sheet.append([None, "2025-08-06", 297.1004, 309.7260])
        sheet.append([None, "2025-08-07", 298.0000, 310.0000])
        stream = BytesIO()
        workbook.save(stream)
        stream.seek(0)

        rates = cbsl_web_rate.parse_usd(stream)
        self.assertEqual(rates[dt.date(2025, 8, 6)], (297.1004, 309.726))
        with patch.object(cbsl_web_rate, "_rates", return_value=rates):
            result = cbsl_web_rate.web_rate("2025-08-08", "average")
        self.assertEqual(result["rate_date"], "2025-08-07")
        self.assertEqual(result["rate"], 304.0)

    def test_exchange_rate_block_uses_cbs_web_rate_for_closing_and_po_dates(self):
        workbook = Workbook()
        sheet = workbook.active

        def fake_rate(day, kind):
            values = {
                ("2026-05-12", "buy"): 318.5751,
                ("2026-05-12", "sell"): 326.1274,
                ("2025-08-06", "buy"): 297.1004,
                ("2025-08-06", "sell"): 309.726,
            }
            return {"rate": values[(day.isoformat(), kind)]}

        with patch.object(pricesched, "cbsl_web_rate", side_effect=fake_rate):
            pricesched._populate_exchange_rates(
                sheet,
                closing_date="12-05-2026",
                po_dates=["06-08-2025"],
            )

        self.assertEqual(sheet["R4"].value, "1 USD (Tender Closing Date) = 12-05-26")
        self.assertEqual(sheet["S5"].value, 318.5751)
        self.assertEqual(sheet["T5"].value, 326.1274)
        self.assertEqual(sheet["U5"].value, 322.3513)
        self.assertEqual(sheet["V4"].value, "PO Date : 06-08-2025")
        self.assertEqual(sheet["V5"].value, 303.4132)

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
            self.assertEqual(sheet.max_row, 9)

    def test_price_output_has_only_one_row_per_described_item(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "anonymous.docx"
            document = Document()
            document.add_paragraph("Official Frank: Acme Power (Pvt) Ltd")
            table = document.add_table(rows=3, cols=8)
            headers = [
                "S/No", "Description", "Qty",
                "Unit Price (Rs.) [Excluding Taxes]", "VAT (%)", "NBT (%)",
                "Other Taxes (%)", "Unit Price (Rs.) [Including Taxes]",
            ]
            values = [
                ["1", "First item", "2", "100", "0", "0", "0", "100"],
                ["2", "Second item", "3", "200", "0", "0", "0", "200"],
            ]
            for cell, value in zip(table.rows[0].cells, headers):
                cell.text = value
            for row, row_values in zip(table.rows[1:], values):
                for cell, value in zip(row.cells, row_values):
                    cell.text = value
            document.save(path)

            result = pricesched.run([str(path)], Path(directory) / "out")

            from openpyxl import load_workbook
            sheet = load_workbook(result["file"], data_only=False).active
            self.assertEqual(sheet.max_row, 10)
            self.assertEqual([sheet.cell(row, 5).value for row in (9, 10)], ["First item", "Second item"])

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

    def test_evaluated_total_uses_required_qty_and_sscl_without_vat(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "local.docx"
            self.create_price_document(path, "Acme Power (Pvt) Ltd")
            document = Document(path)
            cells = document.tables[0].rows[1].cells
            cells[2].text = "10"
            cells[3].text = "100"
            cells[4].text = "18"
            cells[5].text = "2"
            cells[7].text = "120"
            document.save(path)

            result = pricesched.run([str(path)], Path(directory) / "out")

            self.assertEqual(result["rows"][0]["total"], 1200)
            self.assertEqual(result["rows"][0]["evaluated_total"], 1020)

    def test_previous_purchased_group_is_single_and_populated(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "foreign.docx"
            doc = Document()
            doc.add_paragraph("Official Frank: Acme Power (Pvt) Ltd")
            table = doc.add_table(rows=2, cols=10)
            headers = [
                "S/No", "Description", "Qty", "Country of Origin",
                "FOB/FCA Unit Price", "FOB/FCA Total Cost", "Total Insurance",
                "Total Freight", "Total CIF/CIP Cost", "Local Agent's Commission",
            ]
            values = ["1", "Optical module", "2", "China", "40", "80", "2", "3", "85", "0"]
            for cell, value in zip(table.rows[0].cells, headers):
                cell.text = value
            for cell, value in zip(table.rows[1].cells, values):
                cell.text = value
            doc.save(path)

            previous = [
                {"tender_no": "OLD/001", "cif_usd": 41.5, "po_date": "2025-01-10"},
                {"tender_no": "OLD/002", "cif_usd": 43.0, "po_date": "2024-06-15"},
            ]
            with patch.object(pricesched.T, "previous_purchased", return_value=previous):
                result = pricesched.run(
                    [f"current={path}"], Path(directory) / "out"
                )

            from openpyxl import load_workbook
            sheet = load_workbook(result["file"], data_only=False).active
            self.assertTrue(sheet.column_dimensions["Y"].hidden)
            self.assertTrue(sheet.column_dimensions["AA"].hidden)
            self.assertEqual(sheet["V9"].value, "OLD/001")
            self.assertEqual(sheet["W9"].value, 41.5)
            self.assertIsNone(sheet["Y9"].value)
            self.assertIsNone(sheet["Z9"].value)

    def test_pdf_price_table_continuation_pages_are_read(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "multi_page_price.docx"
            output = Path(directory) / "converted.docx"
            header = [
                "S/No", "Description", "Qty",
                "Unit Price (Rs.) [Excluding Taxes]", "VAT (%)", "NBT (%)",
                "Other Taxes (%)", "Unit Price (Rs.) [Including Taxes]",
            ]
            first_page = [header, ["1", "First page item", "1", "100", "0", "0", "0", "100"]]
            continuation_page = [["2", "Second page item", "2", "250", "0", "0", "0", "250"]]
            convert._pdf_to_docx(
                ["Tender No: SLT/PROC/2026/099", "Official Frank: Acme Power (Pvt) Ltd"],
                [first_page, continuation_page],
                output,
            )

            result = pricesched.parse_price(output)

            self.assertEqual(len(result["parts"]), 1)
            self.assertEqual(
                [item["desc"] for item in result["parts"][0]["items"]],
                ["First page item", "Second page item"],
            )


if __name__ == "__main__":
    unittest.main()
