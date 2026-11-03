import json
from pathlib import Path
from collections import Counter


INPUT_FILE = (
    Path(__file__).resolve().parents[1]
    / "storage"
    / "rag_chatbot"
    / "retrieval_units.jsonl"
)


REQUIRED_FIELDS = [
    "retrieval_id",
    "record_type",
    "category",
    "document_reference",
    "source_file",
    "clause_number",
    "start_page",
    "end_page",
    "text",
]


def prepare_chroma_record(record):

    for field in REQUIRED_FIELDS:
        if field not in record or record[field] is None:
            raise ValueError(
                f"Missing required field: {field}"
            )

    if not record["text"].strip():
        raise ValueError(
            f"Empty document text: {record['retrieval_id']}"
        )

    metadata = {
        "record_type": record["record_type"],
        "category": record["category"],
        "document_type": record["document_type"],
        "document_reference": record["document_reference"],
        "source_file": record["source_file"],
        "main_clause_number": record["main_clause_number"],
        "immediate_parent_number": record["immediate_parent_number"],
        "clause_number": record["clause_number"],
        "clause_title": record["clause_title"],
        "hierarchy_level": record["hierarchy_level"],
        "start_page": record["start_page"],
        "end_page": record["end_page"],
        "page_numbers": json.dumps(record["page_numbers"]),
    }

    warning = record.get("source_numbering_warning")

    if warning:
        metadata["source_numbering_warning"] = warning

    # Remove optional empty strings.
    metadata = {
        key: value
        for key, value in metadata.items()
        if value is not None and value != ""
    }

    # ChromaDB metadata values must be supported scalar values.
    for key, value in metadata.items():
        if not isinstance(value, (str, int, float, bool)):
            raise TypeError(
                f"Unsupported metadata value: {key} = {type(value)}"
            )

    return {
        "id": record["retrieval_id"],
        "document": record["text"],
        "metadata": metadata,
    }


def main():

    if not INPUT_FILE.exists():
        raise FileNotFoundError(INPUT_FILE)

    prepared_ids = set()
    record_types = Counter()
    categories = Counter()
    total = 0

    with INPUT_FILE.open("r", encoding="utf-8") as file:

        for line_number, line in enumerate(file, start=1):

            record = json.loads(line)
            prepared = prepare_chroma_record(record)

            if prepared["id"] in prepared_ids:
                raise ValueError(
                    f"Duplicate ID: {prepared['id']}"
                )

            prepared_ids.add(prepared["id"])

            record_types[record["record_type"]] += 1
            categories[record["category"]] += 1

            total += 1

            if total == 1:
                print("SAMPLE CHROMA RECORD")
                print("ID:", prepared["id"])
                print("DOCUMENT LENGTH:", len(prepared["document"]))
                print("METADATA:", prepared["metadata"])

    print("\n" + "=" * 55)
    print("CHROMA RECORD PREPARATION VALIDATION")
    print("=" * 55)

    print("Total prepared records:", total)
    print("Unique IDs:", len(prepared_ids))

    print("\nRECORD TYPES:")
    for name, count in sorted(record_types.items()):
        print(f"{name}: {count}")

    print("\nPROCUREMENT CATEGORIES:")
    for name, count in sorted(categories.items()):
        print(f"{name}: {count}")

    print("\nVALIDATION: PASSED")
    print("No database created or modified.")
    print("No Google API calls made.")


if __name__ == "__main__":
    main()
