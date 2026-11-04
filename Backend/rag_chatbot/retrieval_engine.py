import json
from pathlib import Path

import chromadb

import rag_config
from embedding_client import (
    create_vertex_client,
    embed_text,
    prepare_query_input,
)


# --------------------------------------------------
# DATABASE SETTINGS
# --------------------------------------------------

RAG_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = RAG_DIR.parents[1]

STORAGE_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

DATABASE_DIR = STORAGE_DIR / "chroma_db_734"

COLLECTION_NAME = "procurement_references"

EXPECTED_RECORD_COUNT = 734

ALLOWED_CATEGORIES = [
    "1. Goods & Services (Local or Foreign)",
    "2. Goods (Local & Foreign)",
    "3. Goods & Services (Local)",
    "4. Goods (Local)",
    "5. Services (Local or Foreign)",
]


# --------------------------------------------------
# CONNECT TO CHROMADB
# --------------------------------------------------

def get_procurement_collection():

    if not DATABASE_DIR.is_dir():
        raise FileNotFoundError(
            f"ChromaDB database not found: {DATABASE_DIR}"
        )

    client = chromadb.PersistentClient(
        path=str(DATABASE_DIR)
    )

    collection = client.get_collection(
        name=COLLECTION_NAME,
        embedding_function=None,
    )

    if collection.count() != EXPECTED_RECORD_COUNT:
        raise RuntimeError(
            f"Expected {EXPECTED_RECORD_COUNT} records, "
            f"found {collection.count()}."
        )

    metadata = collection.metadata or {}

    if metadata.get("embedding_model") != rag_config.EMBEDDING_MODEL:
        raise RuntimeError("ChromaDB embedding model mismatch.")

    if metadata.get("embedding_dimensions") != rag_config.EMBEDDING_DIMENSIONS:
        raise RuntimeError("ChromaDB embedding dimensions mismatch.")

    return client, collection


# --------------------------------------------------
# VALIDATE USER QUESTION
# --------------------------------------------------

def validate_question(query):

    if not isinstance(query, str):
        raise ValueError("Question must be text.")

    query = query.strip()

    if not query:
        raise ValueError("Please enter a procurement question.")

    if len(query) > 8000:
        raise ValueError("Question is too long.")

    return query


# --------------------------------------------------
# RETRIEVE RELEVANT PROCUREMENT CLAUSES
# --------------------------------------------------

def retrieve_clauses(
    query,
    category=None,
    top_k=5,
    vertex_client=None,
):

    query = validate_question(query)

    if not isinstance(top_k, int) or isinstance(top_k, bool):
        raise ValueError("top_k must be an integer.")

    if not 1 <= top_k <= 20:
        raise ValueError("top_k must be between 1 and 20.")

    if category is not None and category not in ALLOWED_CATEGORIES:
        raise ValueError("Invalid procurement category.")

    rag_config.validate_embedding_settings()

    # Connect to the existing database.
    _, collection = get_procurement_collection()

    # Validate the selected category.
    where_filter = None

    if category is not None:
        where_filter = {
            "category": category
        }

    # Create the Vertex AI client only when needed.
    if vertex_client is None:
        vertex_client = create_vertex_client()

    # Convert the user's question into an embedding.
    formatted_query = prepare_query_input(query)

    query_vector = embed_text(
        vertex_client,
        formatted_query,
    )

    # Search using the existing document embeddings.
    search_arguments = {
        "query_embeddings": [query_vector],
        "n_results": top_k,
        "include": [
            "documents",
            "metadatas",
            "distances",
        ],
    }

    if where_filter is not None:
        search_arguments["where"] = where_filter

    results = collection.query(**search_arguments)

    retrieved_clauses = []

    result_ids = results["ids"][0]
    result_documents = results["documents"][0]
    result_metadatas = results["metadatas"][0]
    result_distances = results["distances"][0]

    for index, retrieval_id in enumerate(result_ids):

        metadata = result_metadatas[index]

        page_numbers = metadata.get("page_numbers", "[]")

        try:
            page_numbers = json.loads(page_numbers)
        except (TypeError, json.JSONDecodeError):
            page_numbers = []

        clause = {
            "retrieval_id": retrieval_id,

            "text": result_documents[index],

            "category": metadata.get("category"),

            "record_type": metadata.get("record_type"),

            "document_type": metadata.get("document_type"),

            "document_reference": metadata.get(
                "document_reference"
            ),

            "source_file": metadata.get("source_file"),

            "main_clause_number": metadata.get(
                "main_clause_number"
            ),

            "immediate_parent_number": metadata.get(
                "immediate_parent_number"
            ),

            "clause_number": metadata.get("clause_number"),

            "clause_title": metadata.get("clause_title"),

            "start_page": metadata.get("start_page"),

            "end_page": metadata.get("end_page"),

            "page_numbers": page_numbers,

            "distance": result_distances[index],
        }

        retrieved_clauses.append(clause)

    return retrieved_clauses


# --------------------------------------------------
# DISPLAY SEARCH RESULTS
# --------------------------------------------------

def display_results(query, clauses):

    print("\n" + "=" * 70)
    print("PROCUREMENT DOCUMENT RETRIEVAL RESULTS")
    print("=" * 70)

    print("\nQUESTION:")
    print(query)

    print("\nRESULTS FOUND:", len(clauses))

    for index, clause in enumerate(clauses, start=1):

        print("\n" + "-" * 70)

        print(f"RESULT {index}")

        print("Category:", clause["category"])

        print("Document:", clause["source_file"])

        print("Document reference:", clause["document_reference"])

        print("Clause:", clause["clause_number"])

        print("Title:", clause["clause_title"])

        print("Pages:", clause["page_numbers"])

        print("Cosine distance:", clause["distance"])

        print("\nRETRIEVED TEXT:")

        print(clause["text"])

    print("\n" + "=" * 70)


# --------------------------------------------------
# INTERACTIVE RETRIEVAL TEST
# --------------------------------------------------

def main():

    print("=" * 70)
    print("SLT PROCUREMENT RAG - RETRIEVAL TEST")
    print("=" * 70)

    print("\nConnecting to Vertex AI and ChromaDB...")

    vertex_client = create_vertex_client()

    _, collection = get_procurement_collection()

    print("ChromaDB collection:", COLLECTION_NAME)
    print("Stored procurement records:", collection.count())

    print("\nEnter a procurement question to test retrieval.")

    query = input("\nQuestion: ").strip()

    print("\nSearching procurement references...")

    clauses = retrieve_clauses(
        query=query,
        top_k=5,
        vertex_client=vertex_client,
    )

    display_results(query, clauses)


if __name__ == "__main__":
    main()