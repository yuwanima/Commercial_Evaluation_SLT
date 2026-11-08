
import io
import json
import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

PROJECT_ROOT = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(PROJECT_ROOT))

from reportlab.pdfgen import canvas

from Backend.tender_intake.vendor_validation import submission_processor
from Backend.tender_intake.vendor_validation import submission_storage


class UploadedFile:
    def __init__(self, filename, content):
        self.filename = filename
        self.stream = io.BytesIO(content)


class SubmissionProcessorTests(unittest.TestCase):

    def test_pdf_upload_extraction_and_persistence(self):
        with tempfile.TemporaryDirectory() as temporary_dir:
            storage_path = Path(temporary_dir)
            pdf_path = storage_path / "sample.pdf"

            pdf = canvas.Canvas(str(pdf_path))
            pdf.drawString(
                100,
                750,
                "Vendor submission integration test",
            )
            pdf.save()

            uploaded_file = UploadedFile(
                "vendor_submission.pdf",
                pdf_path.read_bytes(),
            )

            with patch.object(
    submission_processor,
    "STORAGE_DIR",
    storage_path,
), patch.object(
    submission_storage,
    "STORAGE_DIR",
    storage_path,
):
                result = submission_processor.process_submission(
                    [uploaded_file],
                    category="Goods",
                )

            self.assertEqual(result["status"], "processed")
            self.assertEqual(
                result["summary"]["files_processed"],
                1,
            )
            self.assertGreater(
                result["summary"]["total_sections"],
                0,
            )

            submission_id = result["submission_id"]
            submission_dir = storage_path / submission_id

            self.assertTrue(
                (submission_dir / "vendor_submission.pdf").exists()
            )

            manifest_path = submission_dir / "manifest.json"
            extraction_path = submission_dir / "extraction.json"

            self.assertTrue(manifest_path.exists())
            self.assertTrue(extraction_path.exists())

            with manifest_path.open(encoding="utf-8") as file:
                manifest = json.load(file)

            with extraction_path.open(encoding="utf-8") as file:
                extraction = json.load(file)

            self.assertEqual(
                manifest["submission_id"],
                submission_id,
            )
            self.assertEqual(
                extraction["submission_id"],
                submission_id,
            )
            self.assertEqual(extraction["category"], "Goods")
            self.assertEqual(extraction["status"], "processed")

            print("PASS: PDF uploaded and extracted.")
            print("PASS: Submission manifest saved.")
            print("PASS: Extraction evidence saved.")
            print("PASS: Submission processor integration test.")


if __name__ == "__main__":
    unittest.main(verbosity=2)
