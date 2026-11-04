
import json
import os
import uuid
from datetime import datetime, timezone
from pathlib import Path

import re
import unicodedata

def secure_filename(filename):
    """Convert a filename into a safe filesystem filename."""

    filename = unicodedata.normalize("NFKD", filename)
    filename = filename.encode("ascii", "ignore").decode("ascii")
    filename = filename.replace("\\", "/").split("/")[-1]
    filename = filename.strip().replace(" ", "_")
    filename = re.sub(r"[^A-Za-z0-9_.-]", "", filename)
    filename = filename.lstrip(".")

    return filename

BASE_DIR = Path(__file__).resolve().parents[2]
STORAGE_DIR = BASE_DIR / "storage" / "vendor_submissions"

ALLOWED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".xlsm",
}

MAX_FILES = 20
MAX_FILE_SIZE = 10 * 1024 * 1024
MAX_TOTAL_SIZE = 50 * 1024 * 1024


def validate_files(files):
    """Validate uploaded files before saving anything."""

    if not files:
        raise ValueError("Please select at least one document.")

    if len(files) > MAX_FILES:
        raise ValueError(f"A maximum of {MAX_FILES} files is allowed.")

    total_size = 0
    validated = []
    filenames = set()

    for uploaded_file in files:
        original_name = uploaded_file.filename or ""

        if not original_name.strip():
            raise ValueError("One of the selected files has no filename.")

        safe_name = secure_filename(original_name)

        if not safe_name:
            raise ValueError("One of the filenames is invalid.")

        extension = Path(safe_name).suffix.lower()

        if extension not in ALLOWED_EXTENSIONS:
            raise ValueError(
                f"Unsupported file type: {safe_name}"
            )

        if safe_name.lower() in filenames:
            raise ValueError(
                f"Duplicate filename: {safe_name}"
            )

        filenames.add(safe_name.lower())

        content = uploaded_file.stream.read(MAX_FILE_SIZE + 1)

        if not content:
            raise ValueError(f"The file is empty: {safe_name}")

        if len(content) > MAX_FILE_SIZE:
            raise ValueError(
                f"{safe_name} exceeds the 10 MB file limit."
            )

        total_size += len(content)

        if total_size > MAX_TOTAL_SIZE:
            raise ValueError("The submission exceeds the 50 MB limit.")

        uploaded_file.stream.seek(0)

        validated.append({
            "original_name": original_name,
            "safe_name": safe_name,
            "extension": extension,
            "size": len(content),
        })

    return validated


def save_submission(files, category=None):
    """Save a validated group of vendor documents."""

    validated = validate_files(files)

    STORAGE_DIR.mkdir(parents=True, exist_ok=True)

    submission_id = uuid.uuid4().hex

    submission_dir = STORAGE_DIR / submission_id
    submission_dir.mkdir()

    saved_files = []

    try:
        for uploaded_file, metadata in zip(files, validated):
            destination = submission_dir / metadata["safe_name"]

            uploaded_file.stream.seek(0)

            with destination.open("xb") as output:
                while True:
                    block = uploaded_file.stream.read(1024 * 1024)

                    if not block:
                        break

                    output.write(block)

            saved_files.append({
                "original_name": metadata["original_name"],
                "stored_name": metadata["safe_name"],
                "extension": metadata["extension"],
                "size": metadata["size"],
            })

        manifest = {
            "submission_id": submission_id,
            "created_at": datetime.now(timezone.utc).isoformat(),
            "category": category,
            "status": "uploaded",
            "files": saved_files,
        }

        manifest_path = submission_dir / "manifest.json"

        with manifest_path.open("x", encoding="utf-8") as output:
            json.dump(manifest, output, indent=2, ensure_ascii=False)

        return manifest

    except Exception:
        import shutil
        shutil.rmtree(submission_dir, ignore_errors=True)
        raise
