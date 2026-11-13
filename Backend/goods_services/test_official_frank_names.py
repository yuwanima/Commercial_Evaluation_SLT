import unittest
import tempfile
from pathlib import Path

from docx import Document
from Backend.goods_services import consolidate


class OfficialFrankNameTests(unittest.TestCase):
    def test_official_stamp_label_extracts_name(self):
        text = "Official Stamp\nAcme Power (Pvt) Ltd"

        self.assertEqual(
            consolidate.extract_official_frank_name(text, "filename-derived"),
            "Acme Power (Pvt) Ltd",
        )

    def test_compound_official_label_is_detected_for_ocr(self):
        self.assertTrue(consolidate.has_official_frank_label("Official Frank / Stamp"))
        self.assertEqual(
            consolidate.extract_official_frank_ocr_name(
                "Official Frank / Stamp\nAcme Power (Pvt) Ltd",
                "filename-derived",
            ),
            "Acme Power (Pvt) Ltd",
        )

    def test_summary_exposes_name_and_source(self):
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "document.docx"
            Document().save(path)
            summary = consolidate.document_summary(
                path,
                "Acme Power (Pvt) Ltd",
                "GCT",
                name_source="Official frank / stamp / seal",
            )

            self.assertIn("Bidder name: Acme Power (Pvt) Ltd", summary["points"])
            self.assertIn(
                "Bidder name source: Official frank / stamp / seal",
                summary["points"],
            )


if __name__ == "__main__":
    unittest.main()
