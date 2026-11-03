import json
import math
from pathlib import Path

import chromadb

import rag_config


# --------------------------------------------------
# PATHS AND SETTINGS
# --------------------------------------------------

RAG_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = RAG_DIR.parents[1]

STORAGE_DIR = PROJECT_ROOT / "Backend" / "storage" / "rag_chatbot"

RETRIEVAL_FILE = STORAGE_DIR / "retrieval_units.jsonl"
EMBEDDINGS_FILE = STORAGE_DIR / "document_embeddings.jsonl"

DATABASE_DIR = STORAGE_DIR / "chroma_db_734"

COLLECTION_NAME = "procurement_references"

BATCH_SIZE = 100


# --------------------------------------------------
# LOAD JSONL FILE
# --------------------------------------------------

def load_jsonl(file_path):
    if not file_path.is_file():
        raise FileNotFoundError(
            f"Required file not found: {file_path}"
        )

    records = []

    with file_path.open("r", encoding="utf-8") as file:
        for line_number, line in enumerate(file, start=1):

            if not line.strip():
                continue

            try:
                records.append(json.loads(line))

            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"Invalid JSON in {file_path.name}, "
                    f"line {line_number}"
                ) from exc

    return records


# --------------------------------------------------
# PREPARE CHROMADB METADATA
# --------------------------------------------------

def prepare_metadata(record):

    metadata_fields = [
        "record_type",
        "category",
        "document_type",
        "document_reference",
        "source_file",
        "main_clause_number",
        "immediate_parent_number",
        "clause_number",
        "clause_title",
        "hierarchy_level",
        "start_page",
        "end_page",
        "source_numbering_warning",
    ]

    metadata = {}

    for field in metadata_fields:

        value = record.get(field)

        if value is None:
            continue

        if isinstance(value, (str, int, float, bool)):
            metadata[field] = value

        else:
            metadata[field] = str(value)

    page_numbers = record.get("page_numbers")

    if page_numbers is not None:
        metadata["page_numbers"] = json.dumps(
            page_numbers,
            ensure_ascii=False
        )

    return metadata


# --------------------------------------------------
# VALIDATE INPUT RECORDS
# --------------------------------------------------

def validate_inputs(retrieval_units, embeddings):

    retrieval_map = {
        record["retrieval_id"]: record
        for record in retrieval_units
    }

    embedding_map = {
        record["retrieval_id"]: record
        for record in embeddings
    }

    retrieval_ids = set(retrieval_map)
    embedding_ids = set(embedding_map)

    if len(retrieval_map) != len(retrieval_units):
        raise ValueError("Duplicate retrieval IDs detected.")

    if len(embedding_map) != len(embeddings):
        raise ValueError("Duplicate embedding IDs detected.")

    if retrieval_ids != embedding_ids:
        raise ValueError(
            "Retrieval and embedding IDs do not match exactly."
        )

    expected_dimensions = rag_config.EMBEDDING_DIMENSIONS
    expected_model = rag_config.EMBEDDING_MODEL

    for retrieval_id, embedding_record in embedding_map.items():

        if embedding_record.get("embedding_model") != expected_model:
            raise ValueError(
                f"Incorrect embedding model: {retrieval_id}"
            )

        if embedding_record.get("embedding_dimensions") != expected_dimensions:
            raise ValueError(
                f"Incorrect embedding dimensions: {retrieval_id}"
            )

        vector = embedding_record.get("embedding")

        if not isinstance(vector, list):
            raise ValueError(
                f"Embedding is not a list: {retrieval_id}"
            )

        if len(vector) != expected_dimensions:
            raise ValueError(
                f"Incorrect vector length: {retrieval_id}"
            )

        if not all(
            isinstance(value, (int, float))
            and not isinstance(value, bool)
            and math.isfinite(value)
            for value in vector
        ):
            raise ValueError(
                f"Invalid vector values: {retrieval_id}"
            )

    for retrieval_id, record in retrieval_map.items():

        if not record.get("text", "").strip():
            raise ValueError(
                f"Empty retrieval text: {retrieval_id}"
            )

        if not record.get("category"):
            raise ValueError(
                f"Missing category: {retrieval_id}"
            )

    return retrieval_map, embedding_map


# --------------------------------------------------
# MAIN DATABASE BUILD
# --------------------------------------------------

def main():

    print("=" * 65)
    print("PROCUREMENT RAG - CHROMADB DATABASE BUILD")
    print("=" * 65)

    print("\nDATABASE PATH:")
    print(DATABASE_DIR)

    # Validate configuration without calling Google.
    rag_config.validate_embedding_settings()

    # Safety: refuse to overwrite an existing database.
    if DATABASE_DIR.exists():
        raise FileExistsError(
            "ChromaDB target directory already exists. "
            "No changes were made."
        )

    print("\nLoading retrieval units...")

    retrieval_units = load_jsonl(RETRIEVAL_FILE)

    print("Retrieval units:", len(retrieval_units))

    print("\nLoading document embeddings...")

    embeddings = load_jsonl(EMBEDDINGS_FILE)

    print("Generated embeddings:", len(embeddings))

    # Verify the complete dataset before database creation.
    retrieval_map, embedding_map = validate_inputs(
        retrieval_units,
        embeddings
    )

    print("\nINPUT VALIDATION: PASSED")

    print("Matched IDs:", len(retrieval_map))
    print("Embedding model:", rag_config.EMBEDDING_MODEL)
    print("Vector dimensions:", rag_config.EMBEDDING_DIMENSIONS)

    # Prepare all ChromaDB records before creating the database.
    ids = []
    documents = []
    metadatas = []
    vectors = []

    for retrieval_id, record in retrieval_map.items():

        embedding_record = embedding_map[retrieval_id]

        ids.append(retrieval_id)
        documents.append(record["text"])
        metadatas.append(prepare_metadata(record))
        vectors.append(embedding_record["embedding"])

    if len(ids) != 734:
        raise ValueError(
            f"Expected 662 records, found {len(ids)}."
        )

    print("\nPREPARATION VALIDATION: PASSED")
    print("Prepared IDs:", len(ids))
    print("Prepared documents:", len(documents))
    print("Prepared metadata records:", len(metadatas))
    print("Prepared vectors:", len(vectors))

    # Create the persistent database only after validation.
    print("\nCreating persistent ChromaDB database...")

    DATABASE_DIR.mkdir(parents=True, exist_ok=False)

    client = chromadb.PersistentClient(
        path=str(DATABASE_DIR)
    )

    collection = client.create_collection(
        name=COLLECTION_NAME,
        embedding_function=None,
        metadata={
            "hnsw:space": "cosine",
            "embedding_model": rag_config.EMBEDDING_MODEL,
            "embedding_dimensions": rag_config.EMBEDDING_DIMENSIONS,
            "record_count": len(ids),
        },
    )

    print("Collection created:", COLLECTION_NAME)

    # Insert pre-generated embeddings in manageable batches.
    total = len(ids)

    for start in range(0, total, BATCH_SIZE):

        end = min(start + BATCH_SIZE, total)

        collection.add(
            ids=ids[start:end],
            documents=documents[start:end],
            metadatas=metadatas[start:end],
            embeddings=vectors[start:end],
        )

        print(
            f"Inserted records: {end}/{total}"
        )

    # Verify stored database contents.
    stored_count = collection.count()

    if stored_count != total:
        raise RuntimeError(
            f"Database count mismatch: "
            f"expected {total}, found {stored_count}"
        )

    stored_sample = collection.get(
        ids=[ids[0]],
        include=["documents", "metadatas"]
    )

    if len(stored_sample["ids"]) != 1:
        raise RuntimeError(
            "Database sample verification failed."
        )

    if stored_sample["documents"][0] != documents[0]:
        raise RuntimeError(
            "Stored document does not match original text."
        )

    if stored_sample["metadatas"][0]["category"] != metadatas[0]["category"]:
        raise RuntimeError(
            "Stored metadata does not match original metadata."
        )

    print("\n" + "=" * 65)
    print("CHROMADB DATABASE BUILD COMPLETED")
    print("=" * 65)

    print("Database:", DATABASE_DIR)
    print("Collection:", COLLECTION_NAME)
    print("Stored records:", stored_count)
    print("Embedding model:", rag_config.EMBEDDING_MODEL)
    print("Vector dimensions:", rag_config.EMBEDDING_DIMENSIONS)
    print("Distance metric: cosine")

    print("\nDATABASE VALIDATION: PASSED")
    print("ORIGINAL RETRIEVAL UNITS: UNCHANGED")
    print("ORIGINAL EMBEDDINGS: UNCHANGED")
    print("NO NEW GOOGLE API CALLS MADE")


if __name__ == "__main__":
    main()