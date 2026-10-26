
from pathlib import Path
from pypdf import PdfReader

DOCUMENTS_DIR = Path("Backend") / "documents"

for pdf_path in sorted(DOCUMENTS_DIR.glob("*.pdf")):
    print("\n" + "=" * 70)
    print(f"File: {pdf_path.name}")

    try:
        reader = PdfReader(str(pdf_path))
        print(f"Total pages: {len(reader.pages)}")

        pages_with_text = 0
        total_characters = 0

        for page_number, page in enumerate(reader.pages, start=1):
            text = page.extract_text() or ""

            if text.strip():
                pages_with_text += 1
                total_characters += len(text)

            # Preview only the first two pages
            if page_number <= 2:
                print(f"\n--- Page {page_number} preview ---")
                print(text[:500] if text.strip() else "[No extractable text]")

        print("\nExtraction summary:")
        print(f"Pages containing text: {pages_with_text}/{len(reader.pages)}")
        print(f"Total extracted characters: {total_characters}")

    except Exception as error:
        print(f"ERROR reading PDF: {error}")
