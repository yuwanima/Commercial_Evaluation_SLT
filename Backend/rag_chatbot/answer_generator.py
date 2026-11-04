import os

from google.genai import types

import rag_config
from embedding_client import create_vertex_client
from retrieval_engine import retrieve_clauses


# --------------------------------------------------
# GEMINI ANSWER GENERATION SETTINGS
# --------------------------------------------------

GENERATION_MODEL = os.getenv(
    "GENERATION_MODEL",
    "gemini-3.8-flash"
)

DEFAULT_TOP_K = 5


# --------------------------------------------------
# BUILD PROCUREMENT EVIDENCE
# --------------------------------------------------

def build_evidence(clauses):

    evidence_parts = []

    for index, clause in enumerate(clauses, start=1):

        evidence = f"""
[SOURCE {index}]
Category: {clause.get("category")}
Document: {clause.get("source_file")}
Document Reference: {clause.get("document_reference")}
Clause Number: {clause.get("clause_number")}
Clause Title: {clause.get("clause_title")}
Pages: {clause.get("page_numbers")}

Clause Text:
{clause.get("text")}
"""

        evidence_parts.append(evidence)

    return "\n".join(evidence_parts)


# --------------------------------------------------
# BUILD SOURCE REFERENCES
# --------------------------------------------------

def build_source_references(clauses):

    references = []

    for index, clause in enumerate(clauses, start=1):

        reference = (
            f"[{index}] "
            f"{clause.get('document_reference') or 'Reference unavailable'} | "
            f"{clause.get('source_file') or 'Document unavailable'} | "
            f"Clause {clause.get('clause_number') or 'N/A'} "
            f"({clause.get('clause_title') or 'Untitled'}) | "
            f"Pages {clause.get('page_numbers') or 'N/A'}"
        )

        references.append(reference)

    return references


# --------------------------------------------------
# GENERATE GROUNDED PROCUREMENT ANSWER
# --------------------------------------------------

def answer_question(
    question,
    category=None,
    top_k=DEFAULT_TOP_K,
    vertex_client=None,
):

    question = question.strip() if isinstance(question, str) else ""

    if not question:
        raise ValueError("Please enter a procurement question.")

    # Create one Vertex AI client for retrieval and generation.
    if vertex_client is None:
        vertex_client = create_vertex_client()

    # Retrieve relevant clauses from the existing ChromaDB.
    clauses = retrieve_clauses(
        query=question,
        category=category,
        top_k=top_k,
        vertex_client=vertex_client,
    )

    if not clauses:
        return {
            "question": question,
            "answer": (
                "I could not find relevant clauses in the available "
                "procurement contract templates. Please check the "
                "procurement category or rephrase your question."
            ),
            "sources": [],
            "retrieved_count": 0,
        }

    evidence = build_evidence(clauses)

    prompt = f"""
You are an SLT Procurement Contract Assistant.

Answer the user's procurement question using ONLY the supplied
procurement contract evidence.

STRICT RULES:

1. Do not use outside knowledge or invent procurement requirements.
2. Do not assume that a requirement exists if it is not stated in
   the supplied clauses.
3. If the evidence does not sufficiently answer the question,
   clearly say that the available retrieved clauses do not provide
   enough information to answer.
4. Distinguish between different procurement categories when their
   contract terms differ.
5. Preserve important figures, percentages, dates, time limits,
   conditions and exceptions exactly as stated.
6. Explain the answer clearly and professionally.
7. Add inline source references such as [1] or [2] to factual
   statements. Only cite the supplied source numbers.
8. Do not create clause numbers or page numbers.
9. If sources provide different requirements, explain the
   differences rather than combining them into one requirement.

USER QUESTION:
{question}

RETRIEVED PROCUREMENT EVIDENCE:
{evidence}

Provide a clear answer with supporting source references.
"""

    response = vertex_client.models.generate_content(
        model=GENERATION_MODEL,
        contents=prompt,
        config=types.GenerateContentConfig(
            temperature=0.2,
            max_output_tokens=2048,
        ),
    )

    answer = response.text

    if not answer:
        raise RuntimeError(
            "Gemini returned an empty answer. Please try again."
        )

    return {
        "question": question,
        "answer": answer,
        "sources": build_source_references(clauses),
        "retrieved_count": len(clauses),
    }


# --------------------------------------------------
# DISPLAY ANSWER
# --------------------------------------------------

def display_answer(result):

    print("\n" + "=" * 75)
    print("SLT PROCUREMENT RAG CHATBOT")
    print("=" * 75)

    print("\nQUESTION:")
    print(result["question"])

    print("\nGEMINI ANSWER:")
    print(result["answer"])

    print("\nSUPPORTING DOCUMENT REFERENCES:")

    if result["sources"]:
        for source in result["sources"]:
            print(source)
    else:
        print("No supporting references found.")

    print("\nRetrieved clauses:", result["retrieved_count"])

    print("\n" + "=" * 75)


# --------------------------------------------------
# INTERACTIVE ANSWER GENERATION TEST
# --------------------------------------------------

def main():

    print("=" * 75)
    print("SLT PROCUREMENT RAG - ANSWER GENERATION TEST")
    print("=" * 75)

    rag_config.validate_vertex_settings()

    print("\nGeneration model:", GENERATION_MODEL)
    print("Connecting to Vertex AI...")

    vertex_client = create_vertex_client()

    print("Vertex AI connection initialized.")

    question = input("\nEnter your procurement question: ").strip()

    print("\nRetrieving clauses and generating an answer...")
    print("Please wait...\n")

    result = answer_question(
        question=question,
        top_k=5,
        vertex_client=vertex_client,
    )

    display_answer(result)


if __name__ == "__main__":
    main()
