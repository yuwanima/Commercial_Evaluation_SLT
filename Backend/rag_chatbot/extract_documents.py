
import json
from pathlib import Path
from datetime import datetime, timezone

from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook


# ============================================================
# PATH CONFIGURATION
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[2]

SOURCE_ROOT = (
    PROJECT_ROOT
    / "Backend"
    / "documents"
    / "rag_sources"
    / "GCC & SCC"
)

OUTPUT_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

OUTPUT_FILE = OUTPUT_DIR / "extracted_documents.jsonl"


# ============================================================
# DOCUMENT EXTRACTION FUNCTIONS
# ============================================================

def extract_pdf(file_path, category):
    """Extract text from each PDF page."""

    records = []
    reader = PdfReader(str(file_path))

    for page_index, page in enumerate(reader.pages):

        text = page.extract_text() or ""

        records.append({
            "category": category,
            "document_type": "PDF",
            "source_file": file_path.name,
            "page_number": page_index + 1,
            "worksheet": None,
            "content_type": "page_text",
            "text": text.strip()
        })

    return records


def extract_word(file_path, category):
    """Extract paragraphs and tables from a Word document."""

    records = []
    document = Document(str(file_path))

    # Extract paragraphs
    for index, paragraph in enumerate(document.paragraphs):

        text = paragraph.text.strip()

        if text:
            records.append({
                "category": category,
                "document_type": "DOCX",
                "source_file": file_path.name,
                "page_number": None,
                "worksheet": None,
                "content_type": "paragraph",
                "element_number": index + 1,
                "text": text
            })

    # Extract tables
    for table_index, table in enumerate(document.tables):

        for row_index, row in enumerate(table.rows):

            values = [
                cell.text.strip().replace("\n", " / ")
                for cell in row.cells
            ]

            if any(values):
                records.append({
                    "category": category,
                    "document_type": "DOCX",
                    "source_file": file_path.name,
                    "page_number": None,
                    "worksheet": None,
                    "content_type": "table_row",
                    "table_number": table_index + 1,
                    "row_number": row_index + 1,
                    "text": " | ".join(values)
                })

    return records


def extract_excel(file_path, category):
    """Extract populated rows from every Excel worksheet."""

    records = []

    workbook = load_workbook(
        filename=str(file_path),
        read_only=True,
        data_only=False
    )

    try:
        for worksheet in workbook.worksheets:

            for row_index, row in enumerate(
                worksheet.iter_rows(values_only=True),
                start=1
            ):

                values = [
                    str(value).strip()
                    if value is not None else ""
                    for value in row
                ]

                if any(values):

                    records.append({
                        "category": category,
                        "document_type": "XLSX",
                        "source_file": file_path.name,
                        "page_number": None,
                        "worksheet": worksheet.title,
                        "content_type": "worksheet_row",
                        "row_number": row_index,
                        "text": " | ".join(values)
                    })

    finally:
        workbook.close()

    return records


# ============================================================
# MAIN EXTRACTION PROCESS
# ============================================================

def main():

    if not SOURCE_ROOT.exists():
        raise FileNotFoundError(
            f"Source folder not found: {SOURCE_ROOT}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    all_records = []
    document_count = 0
    errors = []

    print("=" * 65)
    print("SLT PROCUREMENT DOCUMENT EXTRACTION")
    print("=" * 65)

    for category_folder in sorted(SOURCE_ROOT.iterdir()):

        if not category_folder.is_dir():
            continue

        category = category_folder.name

        print(f"\nCATEGORY: {category}")

        for file_path in sorted(category_folder.iterdir()):

            if not file_path.is_file():
                continue

            suffix = file_path.suffix.lower()

            if suffix not in [".pdf", ".docx", ".xlsx"]:
                continue

            try:

                if suffix == ".pdf":
                    records = extract_pdf(file_path, category)

                elif suffix == ".docx":
                    records = extract_word(file_path, category)

                elif suffix == ".xlsx":
                    records = extract_excel(file_path, category)

                else:
                    continue

                all_records.extend(records)
                document_count += 1

                print(
                    f"  OK: {file_path.name} "
                    f"({len(records)} records)"
                )

            except Exception as error:

                errors.append({
                    "file": str(file_path),
                    "error": str(error)
                })

                print(f"  ERROR: {file_path.name}")
                print(f"         {error}")

    # Save extracted records as JSON Lines
    with OUTPUT_FILE.open(
        "w",
        encoding="utf-8"
    ) as output:

        for record in all_records:

            record["extracted_at"] = datetime.now(
                timezone.utc
            ).isoformat()

            output.write(
                json.dumps(record, ensure_ascii=False) + "\n"
            )

    # Summary
    print("\n" + "=" * 65)
    print("EXTRACTION SUMMARY")
    print("=" * 65)

    print(f"Documents processed: {document_count}")
    print(f"Records extracted:   {len(all_records)}")
    print(f"Errors:              {len(errors)}")
    print(f"Output file:         {OUTPUT_FILE}")

    if errors:
        print("\nFILES WITH ERRORS:")

        for item in errors:
            print(item["file"])
            print(item["error"])

    else:
        print("\nAll supported documents extracted successfully.")


if __name__ == "__main__":
    main()
