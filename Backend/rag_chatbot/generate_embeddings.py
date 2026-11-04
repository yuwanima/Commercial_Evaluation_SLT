
import json
import time
from pathlib import Path

import rag_config
from embedding_client import (
    create_vertex_client,
    embed_text,
    prepare_document_input,
)


# --------------------------------------------------
# FILE PATHS
# --------------------------------------------------

RAG_DIR = Path(__file__).resolve().parent

PROJECT_ROOT = RAG_DIR.parents[1]

INPUT_FILE = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
    / "retrieval_units.jsonl"
)

OUTPUT_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

OUTPUT_FILE = OUTPUT_DIR / "document_embeddings.jsonl"


# --------------------------------------------------
# PROCESSING SETTINGS
# --------------------------------------------------

BATCH_SIZE = 10
MAX_RECORDS_PER_RUN = 662
MAX_RETRIES = 3
RETRY_DELAY_SECONDS = 3


# --------------------------------------------------
# FILE HELPERS
# --------------------------------------------------

def load_retrieval_units():
    if not INPUT_FILE.is_file():
        raise FileNotFoundError(
            f"Retrieval-unit file not found: {INPUT_FILE}"
        )

    records = []

    with INPUT_FILE.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            if not line.strip():
                continue

            try:
                record = json.loads(line)
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"Invalid JSON at line {line_number}"
                ) from exc

            records.append(record)

    if not records:
        raise ValueError("No retrieval units found.")

    ids = [record["retrieval_id"] for record in records]

    if len(ids) != len(set(ids)):
        raise ValueError("Duplicate retrieval IDs found.")

    return records


def load_completed_embeddings():
    completed = {}

    if not OUTPUT_FILE.exists():
        return completed

    with OUTPUT_FILE.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):
            if not line.strip():
                continue

            record = json.loads(line)

            retrieval_id = record["retrieval_id"]
            embedding = record["embedding"]

            if retrieval_id in completed:
                raise ValueError(
                    f"Duplicate embedding ID: {retrieval_id}"
                )

            if len(embedding) != rag_config.EMBEDDING_DIMENSIONS:
                raise ValueError(
                    f"Invalid embedding dimensions: {retrieval_id}"
                )

            completed[retrieval_id] = record

    return completed


def save_embedding(record, embedding):
    output_record = {
        "retrieval_id": record["retrieval_id"],
        "embedding_model": rag_config.EMBEDDING_MODEL,
        "embedding_dimensions": rag_config.EMBEDDING_DIMENSIONS,
        "embedding": embedding,
    }

    with OUTPUT_FILE.open("a", encoding="utf-8") as file:
        file.write(
            json.dumps(output_record, ensure_ascii=False) + "\n"
        )


# --------------------------------------------------
# EMBEDDING GENERATION
# --------------------------------------------------

def generate_one_embedding(client, record):
    title = (
        record.get("clause_title")
        or record.get("clause_number")
        or record.get("record_type")
        or "Procurement document"
    )

    text = prepare_document_input(
        text=record["text"],
        title=title,
    )

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            embedding = embed_text(client, text)

            if len(embedding) != rag_config.EMBEDDING_DIMENSIONS:
                raise ValueError(
                    "Unexpected embedding dimensions."
                )

            return embedding

        except Exception:
            if attempt == MAX_RETRIES:
                raise

            print(
                f"Temporary failure for "
                f"{record['retrieval_id']}. "
                f"Retry {attempt}/{MAX_RETRIES}..."
            )

            time.sleep(RETRY_DELAY_SECONDS * attempt)


def main():
    rag_config.validate_vertex_settings()
    rag_config.validate_embedding_settings()

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    records = load_retrieval_units()
    completed = load_completed_embeddings()

    valid_ids = {record["retrieval_id"] for record in records}

    unexpected_ids = set(completed) - valid_ids

    if unexpected_ids:
        raise ValueError(
            "Embedding output contains IDs not present in "
            "the retrieval-unit file."
        )

    pending = [
        record
        for record in records
        if record["retrieval_id"] not in completed
    ]

    pending = pending[:MAX_RECORDS_PER_RUN]

    print("=" * 60)
    print("PROCUREMENT DOCUMENT EMBEDDING GENERATION")
    print("=" * 60)
    print("Total retrieval units:", len(records))
    print("Already completed:", len(completed))
    print("Remaining:", len(pending))
    print("Embedding model:", rag_config.EMBEDDING_MODEL)
    print("Dimensions:", rag_config.EMBEDDING_DIMENSIONS)
    print("Batch size:", BATCH_SIZE)
    print("Output:", OUTPUT_FILE)
    print("=" * 60)

    if not pending:
        print("All retrieval units already have embeddings.")
        return

    confirmation = input(
        "Generate embeddings using Vertex AI? Type YES to continue: "
    )

    if confirmation != "YES":
        print("Cancelled. No embeddings generated.")
        return

    client = create_vertex_client()

    total = len(pending)

    for start in range(0, total, BATCH_SIZE):
        batch = pending[start:start + BATCH_SIZE]

        print(
            f"\nProcessing batch "
            f"{start // BATCH_SIZE + 1} "
            f"({len(batch)} records)"
        )

        for record in batch:
            embedding = generate_one_embedding(client, record)

            save_embedding(record, embedding)

            completed[record["retrieval_id"]] = {
                "retrieval_id": record["retrieval_id"],
                "embedding_model": rag_config.EMBEDDING_MODEL,
                "embedding_dimensions": rag_config.EMBEDDING_DIMENSIONS,
                "embedding": embedding,
            }

            print(
                f"Saved: {record['retrieval_id']} "
                f"({len(embedding)} dimensions)"
            )

        print(
            f"Progress: {len(completed)}/{len(records)}"
        )

    print("\nEMBEDDING GENERATION COMPLETED")
    print("Total embeddings:", len(completed))
    print("Output file:", OUTPUT_FILE)


if __name__ == "__main__":
    main()
