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
# BUILD VENDOR DOCUMENT EVIDENCE
# --------------------------------------------------

def build_vendor_evidence(vendor_evidence):

    evidence_parts = []

    if not isinstance(vendor_evidence, list):
        return ""

    for index, item in enumerate(vendor_evidence, start=1):

        if not isinstance(item, dict):
            continue

        file_name = item.get("file_name", "")
        location = item.get("location", "")
        text = item.get("text", "")

        if not isinstance(text, str) or not text.strip():
            continue

        if not isinstance(file_name, str):
            file_name = "Unknown document"

        if not isinstance(location, str):
            location = "Location unavailable"

        evidence = f"""
[V{index}]
Vendor Document: {file_name}
Original Document Location: {location}

Extracted Content:
{text}
"""

        evidence_parts.append(evidence)

    return "\n".join(evidence_parts)


# --------------------------------------------------
# BUILD VENDOR DOCUMENT REFERENCES
# --------------------------------------------------

def build_vendor_source_references(vendor_evidence):

    references = []

    if not isinstance(vendor_evidence, list):
        return references

    for index, item in enumerate(vendor_evidence, start=1):

        if not isinstance(item, dict):
            continue

        text = item.get("text", "")

        if not isinstance(text, str) or not text.strip():
            continue

        file_name = item.get("file_name") or "Unknown document"
        location = item.get("location") or "Location unavailable"

        references.append(
            f"[V{index}] Vendor document: {file_name} | "
            f"Location: {location}"
        )

    return references


# --------------------------------------------------
# GENERATE GROUNDED PROCUREMENT ANSWER
# --------------------------------------------------

def answer_question(
    question,
    category=None,
    top_k=DEFAULT_TOP_K,
    vertex_client=None,
    vendor_evidence=None,
):

    question = question.strip() if isinstance(question, str) else ""

    if not question:
        raise ValueError("Please enter a procurement question.")

    if vendor_evidence is None:
        vendor_evidence = []

    if not isinstance(vendor_evidence, list):
        vendor_evidence = []

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

    # Build both types of evidence.
    procurement_evidence = build_evidence(clauses)

    vendor_document_evidence = build_vendor_evidence(
        vendor_evidence
    )

    vendor_references = build_vendor_source_references(
        vendor_evidence
    )

    # Do not generate an unsupported answer when no evidence exists.
    if not clauses and not vendor_references:
        return {
            "question": question,
            "answer": (
                "I could not find relevant information in the available "
                "procurement manuals or uploaded vendor documents. "
                "Please check the procurement category, upload the "
                "relevant document, or rephrase your question."
            ),
            "sources": [],
            "retrieved_count": 0,
            "vendor_evidence_count": 0,
        }

    prompt = f"""
You are ContractWise AI, an SLT Procurement Contract Assistant.

Your task is to answer procurement questions using ONLY the evidence
provided below.

There are two possible evidence sources:

1. SLT PROCUREMENT MANUALS
   These contain procurement requirements and contract clauses.
   Their citations use [1], [2], [3], etc.

2. UPLOADED VENDOR DOCUMENTS
   These contain information, declarations, responses and evidence
   extracted from documents uploaded by the user.
   Their citations use [V1], [V2], [V3], etc.

STRICT ACCURACY RULES:

1. Never use outside knowledge or invent procurement requirements.

2. Distinguish clearly between an SLT procurement requirement and
   something a vendor has stated or submitted.

3. When comparing a vendor document against a procurement clause,
   use the actual requirement from the manual and the actual
   extracted vendor evidence.

4. A vendor declaration saying "Complied" or "Yes" is a declaration.
   It does not automatically prove substantive compliance.

5. A blank field means that no value was extracted from that field.
   Do not automatically conclude that the entire tender is incomplete
   or that the vendor is non-compliant.

6. If information is missing, explicitly identify what information
   could not be found in the supplied evidence.

7. Do not claim that a vendor is compliant or non-compliant unless
   the supplied evidence supports that conclusion.

8. Do not claim that a document has been signed, stamped or visually
   verified. Text extraction alone cannot establish the presence
   or validity of a visual signature or stamp.

9. Preserve important figures, percentages, dates, time limits,
   conditions and exceptions exactly as stated.

10. If the manual and vendor document provide different information,
    explain the difference clearly. Do not silently resolve conflicts.

11. If the available evidence is insufficient, explicitly say so.

12. Do not create clause numbers, page numbers, paragraph numbers,
    sheet names or cell references.

13. Cite factual statements using only the supplied references.
    Use [1], [2] for manual evidence and [V1], [V2] for vendor evidence.

14. Do not confuse a vendor's statement with an independently
    verified fact.

15. Explain answers clearly and professionally.

USER QUESTION:
{question}

RETRIEVED SLT PROCUREMENT MANUAL EVIDENCE:
{procurement_evidence or "No relevant manual clauses were retrieved."}

UPLOADED VENDOR DOCUMENT EVIDENCE:
{vendor_document_evidence or "No vendor document evidence was supplied."}

Provide a clear, structured answer with appropriate inline citations.

When relevant, organize the answer into:
- Relevant procurement requirement
- Information found in the vendor document
- Comparison or assessment
- Missing information or limitations

Do not invent information to fill gaps.
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

    # Return manual and vendor references together.
    all_sources = (
        build_source_references(clauses)
        + vendor_references
    )

    return {
        "question": question,
        "answer": answer,
        "sources": all_sources,
        "retrieved_count": len(clauses),
        "vendor_evidence_count": len(vendor_references),
    }


# --------------------------------------------------
# DISPLAY ANSWER
# --------------------------------------------------

def display_answer(result):

    print("\n" + "=" * 75)
    print("CONTRACTWISE AI")
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

    print("\nRetrieved procurement clauses:", result["retrieved_count"])

    print(
        "Vendor evidence sections:",
        result.get("vendor_evidence_count", 0)
    )

    print("\n" + "=" * 75)


# --------------------------------------------------
# INTERACTIVE ANSWER GENERATION TEST
# --------------------------------------------------

def main():

    print("=" * 75)
    print("CONTRACTWISE AI - ANSWER GENERATION TEST")
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