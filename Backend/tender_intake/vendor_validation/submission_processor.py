
import json

from .document_extractor import extract_document
from .submission_storage import save_submission, STORAGE_DIR


def write_json_atomically(path, data):
    """Write JSON without leaving a partially written file."""

    temporary_path = path.with_suffix(path.suffix + ".tmp")

    with temporary_path.open("w", encoding="utf-8") as output:
        json.dump(data, output, indent=2, ensure_ascii=False)

    temporary_path.replace(path)


def process_submission(files, category=None):
    """Save vendor documents, extract evidence and persist results."""

    manifest = save_submission(files, category=category)
    submission_dir = STORAGE_DIR / manifest["submission_id"]

    results = []
    all_sections = []
    extraction_errors = []

    for file_metadata in manifest["files"]:
        file_path = submission_dir / file_metadata["stored_name"]

        try:
            extracted = extract_document(file_path)

            file_result = {
                "file_name": extracted["file_name"],
                "extension": extracted["extension"],
                "section_count": extracted["section_count"],
                "sections": extracted["sections"],
                "status": "extracted",
            }

            for section in extracted["sections"]:
                all_sections.append({
                    "file_name": extracted["file_name"],
                    "location": section["location"],
                    "text": section["text"],
                    "requires_ocr": section.get("requires_ocr", False),
                })

        except Exception:
            file_result = {
                "file_name": file_metadata["original_name"],
                "status": "extraction_failed",
                "error": (
                    "Unable to extract this document. "
                    "It may be corrupted, password-protected "
                    "or unsupported."
                ),
                "section_count": 0,
                "sections": [],
            }

            extraction_errors.append(file_result)

        results.append(file_result)

    ocr_pages = [
        {
            "file_name": item["file_name"],
            "location": item["location"],
        }
        for item in all_sections
        if item["requires_ocr"]
    ]

    if len(extraction_errors) == len(results):
        status = "processing_failed"
    elif extraction_errors or ocr_pages:
        status = "processed_needs_review"
    else:
        status = "processed"

    summary = {
        "files_processed": len(results) - len(extraction_errors),
        "files_failed": len(extraction_errors),
        "total_sections": len(all_sections),
        "pages_requiring_ocr": len(ocr_pages),
    }

    manifest["status"] = status
    manifest["extraction"] = summary

    extraction_data = {
        "submission_id": manifest["submission_id"],
        "category": category,
        "status": status,
        "files": results,
        "evidence": all_sections,
        "extraction_errors": extraction_errors,
        "pages_requiring_ocr": ocr_pages,
        "summary": summary,
    }

    write_json_atomically(
        submission_dir / "extraction.json",
        extraction_data,
    )

    write_json_atomically(
        submission_dir / "manifest.json",
        manifest,
    )

    return extraction_data
