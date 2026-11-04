
import io
import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from Backend.tender_intake.vendor_validation import submission_storage


class UploadedFile:
    def __init__(self, filename, content):
        self.filename = filename
        self.stream = io.BytesIO(content)


class SubmissionStorageTests(unittest.TestCase):

    def setUp(self):
        self.temp_directory = tempfile.TemporaryDirectory()
        self.storage_path = Path(self.temp_directory.name)

        self.storage_patch = patch.object(
            submission_storage,
            "STORAGE_DIR",
            self.storage_path,
        )
        self.storage_patch.start()

    def tearDown(self):
        self.storage_patch.stop()
        self.temp_directory.cleanup()

    def make_file(self, filename, content):
        return UploadedFile(filename, content)

    def test_single_pdf_upload(self):
        files = [self.make_file("performance_security.pdf", b"%PDF-test")]
        result = submission_storage.save_submission(files)

        folder = self.storage_path / result["submission_id"]

        self.assertTrue((folder / "performance_security.pdf").is_file())
        self.assertTrue((folder / "manifest.json").is_file())

        print("PASS: Single PDF uploaded and saved.")

    def test_multiple_file_upload(self):
        files = [
            self.make_file("document.pdf", b"%PDF-test"),
            self.make_file("declaration.docx", b"word-test"),
            self.make_file("financial.xlsx", b"excel-test"),
        ]

        result = submission_storage.save_submission(files)
        folder = self.storage_path / result["submission_id"]

        self.assertEqual(len(result["files"]), 3)

        for item in result["files"]:
            self.assertTrue((folder / item["stored_name"]).is_file())

        print("PASS: Multiple documents saved in one submission.")

    def test_unique_submission_ids(self):
        first = submission_storage.save_submission([
            self.make_file("first.pdf", b"%PDF-one")
        ])

        second = submission_storage.save_submission([
            self.make_file("second.pdf", b"%PDF-two")
        ])

        self.assertNotEqual(first["submission_id"], second["submission_id"])

        print("PASS: Separate submissions receive unique IDs.")

    def test_unsupported_file_rejected(self):
        with self.assertRaises(ValueError):
            submission_storage.save_submission([
                self.make_file("malware.exe", b"test")
            ])

        print("PASS: Unsupported file type rejected.")

    def test_empty_file_rejected(self):
        with self.assertRaises(ValueError):
            submission_storage.save_submission([
                self.make_file("empty.pdf", b"")
            ])

        print("PASS: Empty file rejected.")

    def test_oversized_file_rejected(self):
        content = b"A" * (submission_storage.MAX_FILE_SIZE + 1)

        with self.assertRaises(ValueError):
            submission_storage.save_submission([
                self.make_file("large.pdf", content)
            ])

        print("PASS: Oversized file rejected.")

    def test_duplicate_filenames_rejected(self):
        files = [
            self.make_file("document.pdf", b"%PDF-one"),
            self.make_file("document.pdf", b"%PDF-two"),
        ]

        with self.assertRaises(ValueError):
            submission_storage.save_submission(files)

        print("PASS: Duplicate filenames rejected.")

    def test_manifest_contains_submission_metadata(self):
        result = submission_storage.save_submission(
            [self.make_file("offer.pdf", b"%PDF-test")],
            category="1. Goods & Services (Local or Foreign)",
        )

        manifest_path = (
            self.storage_path
            / result["submission_id"]
            / "manifest.json"
        )

        with manifest_path.open(encoding="utf-8") as file:
            manifest = json.load(file)

        self.assertEqual(
            manifest["category"],
            "1. Goods & Services (Local or Foreign)",
        )
        self.assertEqual(manifest["status"], "uploaded")
        self.assertEqual(len(manifest["files"]), 1)

        print("PASS: Submission manifest saved correctly.")


if __name__ == "__main__":
    unittest.main(verbosity=2)
