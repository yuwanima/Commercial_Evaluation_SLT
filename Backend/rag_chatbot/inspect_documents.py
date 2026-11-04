
from pathlib import Path
from pypdf import PdfReader
from docx import Document
from openpyxl import load_workbook

PROJECT_ROOT = Path(__file__).resolve().parents[2]

SOURCE_ROOT = (
    PROJECT_ROOT
    / "Backend"
    / "documents"
    / "rag_sources"
    / "GCC & SCC"
)

REPORT_PATH = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
    / "document_inspection.txt"
)


def inspect_pdf(path):
    reader = PdfReader(str(path))

    print(f"Pages: {len(reader.pages)}")

    for page_number, page in enumerate(reader.pages[:2], start=1):
        text = page.extract_text() or ""
        print(f"\n--- Page {page_number} preview ---")
        print(text[:1200])


def inspect_word(path):
    document = Document(str(path))

    print(f"Paragraphs: {len(document.paragraphs)}")
    print(f"Tables: {len(document.tables)}")

    for paragraph in document.paragraphs[:15]:
        if paragraph.text.strip():
            print(paragraph.text[:500])

    for table_number, table in enumerate(document.tables, start=1):
        print(f"\nTable {table_number}:")
        for row in table.rows[:5]:
            print(" | ".join(cell.text[:150] for cell in row.cells))


def inspect_excel(path):
    workbook = load_workbook(
        filename=path,
        read_only=True,
        data_only=True
    )

    print(f"Worksheets: {len(workbook.sheetnames)}")

    for sheet in workbook.worksheets:
        print(f"\nWorksheet: {sheet.title}")
        print(f"Rows: {sheet.max_row}, Columns: {sheet.max_column}")

        for row in sheet.iter_rows(min_row=1, max_row=min(8, sheet.max_row),
                                   values_only=True):
            print(" | ".join(
                str(value)[:150] if value is not None else ""
                for value in row
            ))

    workbook.close()


def main():
    if not SOURCE_ROOT.exists():
        raise FileNotFoundError(
            f"Source folder not found: {SOURCE_ROOT}"
        )

    REPORT_PATH.parent.mkdir(parents=True, exist_ok=True)

    files = sorted(
        path for path in SOURCE_ROOT.rglob("*")
        if path.is_file()
        and path.suffix.lower() in {".pdf", ".docx", ".xlsx"}
    )

    with REPORT_PATH.open("w", encoding="utf-8") as report:
        import sys

        original_stdout = sys.stdout

        try:
            sys.stdout = report

            print("GCC & SCC DOCUMENT INSPECTION REPORT")
            print("=" * 70)
            print(f"Documents found: {len(files)}")

            for path in files:
                print("\n\n" + "=" * 70)
                print(f"CATEGORY: {path.parent.name}")
                print(f"FILE: {path.name}")
                print(f"TYPE: {path.suffix.upper()}")
                print("=" * 70)

                try:
                    if path.suffix.lower() == ".pdf":
                        inspect_pdf(path)

                    elif path.suffix.lower() == ".docx":
                        inspect_word(path)

                    elif path.suffix.lower() == ".xlsx":
                        inspect_excel(path)

                except Exception as error:
                    print(f"INSPECTION ERROR: {error}")

        finally:
            sys.stdout = original_stdout

    print(f"Inspection completed: {len(files)} files")
    print(f"Report saved to: {REPORT_PATH}")


if __name__ == "__main__":
    main()
