
import json
import re
from pathlib import Path
from pypdf import PdfReader

DOCUMENTS_DIR = Path("Backend") / "documents"
OUTPUT_FILE = Path("Backend") / "storage" / "rag_chunks.jsonl"

CHUNK_SIZE = 1200
CHUNK_OVERLAP = 200


def split_text(text):
    """Split text into overlapping chunks without losing page metadata."""
    text = re.sub(r"\s+", " ", text).strip()
    chunks = []
    start = 0

    while start < len(text):
        end = min(start + CHUNK_SIZE, len(text))
        chunk = text[start:end].strip()

        if chunk:
            chunks.append(chunk)

        if end == len(text):
            break

        start = end - CHUNK_OVERLAP

    return chunks


def main():
    pdf_files = sorted(DOCUMENTS_DIR.glob("*.pdf"))

    if not pdf_files:
        print(f"No PDF files found in {DOCUMENTS_DIR}")
        return

    OUTPUT_FILE.parent.mkdir(parents=True, exist_ok=True)

    total_chunks = 0

    with OUTPUT_FILE.open("w", encoding="utf-8") as output:
        for pdf_path in pdf_files:
            print(f"\nProcessing: {pdf_path.name}")
            reader = PdfReader(str(pdf_path))
            document_chunks = 0

            for page_number, page in enumerate(reader.pages, start=1):
                text = page.extract_text() or ""

                if not text.strip():
                    print(f"Warning: No text on page {page_number}")
                    continue

                for chunk_number, chunk in enumerate(
                    split_text(text), start=1
                ):
                    record = {
                        "source": pdf_path.name,
                        "page": page_number,
                        "chunk": chunk_number,
                        "text": chunk,
                    }

                    output.write(
                        json.dumps(record, ensure_ascii=False) + "\n"
                    )

                    document_chunks += 1

            total_chunks += document_chunks
            print(f"Pages: {len(reader.pages)}")
            print(f"Chunks created: {document_chunks}")

    print(f"\nTotal chunks: {total_chunks}")
    print(f"Output saved to: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()
