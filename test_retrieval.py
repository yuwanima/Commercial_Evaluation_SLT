
import json
import math
import re
from collections import Counter
from pathlib import Path

CHUNKS_FILE = Path("Backend") / "storage" / "rag_chunks.jsonl"
TOP_K = 5


def tokenize(text):
    return re.findall(r"\b[a-z0-9]+\b", text.lower())


def main():
    if not CHUNKS_FILE.exists():
        print(f"Chunk file not found: {CHUNKS_FILE}")
        print("Run py prepare_documents.py first.")
        return

    with CHUNKS_FILE.open(encoding="utf-8") as file:
        chunks = [
            json.loads(line)
            for line in file
            if line.strip()
        ]

    if not chunks:
        print("No chunks found in the JSONL file.")
        return

    print(f"Loaded {len(chunks)} chunks.")
    query = input("\nEnter a procurement question: ").strip()

    query_terms = set(tokenize(query))
    if not query_terms:
        print("Please enter a question containing searchable words.")
        return

    # Calculate how many chunks contain each query term.
    document_frequency = Counter()

    for chunk in chunks:
        terms = set(tokenize(chunk["text"]))
        document_frequency.update(terms & query_terms)

    ranked = []

    for chunk in chunks:
        text = chunk["text"]
        source = chunk["source"]
        terms = tokenize(text)
        term_counts = Counter(terms)

        # Include source metadata when calculating relevance.
        metadata = tokenize(source.replace(".pdf", ""))
        all_terms = terms + metadata
        all_counts = Counter(all_terms)

        
        score = 0.0

        for term in query_terms:
            count = all_counts[term]

            if count == 0:
                continue

            frequency = document_frequency[term]
            idf = math.log(
                (len(chunks) + 1) / (frequency + 1)
            ) + 1

            score += idf * (1 + math.log(count))

        # Prefer Part II for questions explicitly asking about
        # the "guidelines for procurement processes".
        query_lower = query.lower()
        source_lower = source.lower()

        if (
            "guidelines" in query_lower
            and "procurement processes" in query_lower
            and "part ii" in source_lower
        ):
            score += 15

        # Prefer Part I when the question explicitly asks about
        # principles and policies.
        if (
            "principles" in query_lower
            and "policies" in query_lower
            and "part-i" in source_lower
        ):
            score += 15

        if score > 0:
            ranked.append((score, chunk))


        for term in query_terms:
            count = all_counts[term]

            if count == 0:
                continue

            frequency = document_frequency[term]
            idf = math.log(
                (len(chunks) + 1) / (frequency + 1)
            ) + 1

            # Avoid giving excessive weight to repeated words.
            score += idf * (1 + math.log(count))

        # Small boost for exact phrase matches in the text.
        normalized_query = " ".join(tokenize(query))
        normalized_text = " ".join(terms)

        if normalized_query and normalized_query in normalized_text:
            score += 5

        if score > 0:
            ranked.append((score, chunk))

    ranked.sort(key=lambda item: item[0], reverse=True)

    if not ranked:
        print("\nNo matching chunks found.")
        print("Try using different words from the question.")
        return

    print(f"\nTop {min(TOP_K, len(ranked))} matching chunks:")

    for rank, (score, chunk) in enumerate(
        ranked[:TOP_K], start=1
    ):
        print("\n" + "=" * 65)
        print(f"Rank: {rank} | Score: {score:.3f}")
        print(f"Manual: {chunk['source']}")
        print(f"Page: {chunk['page']}")
        print(f"Chunk: {chunk['chunk']}")
        print(f"Text: {chunk['text'][:900]}")


if __name__ == "__main__":
    main()
