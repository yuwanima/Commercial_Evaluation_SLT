import json
import re
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]

STORAGE_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

MAIN_CLAUSE_FILE = STORAGE_DIR / "gcc_clauses.jsonl"


# Recognise a definition beginning with a quoted term.
# Also allow the missing opening quotation mark before Specifications.
DEFINITION_START = re.compile(
    r"^\s*(?P<term>.+?)\s+means\b",
    re.IGNORECASE
)


def read_jsonl(path):

    records = []

    with path.open("r", encoding="utf-8") as file:
        for line in file:
            if line.strip():
                records.append(json.loads(line))

    return records


def extract_definitions(clause):

    lines = clause["text"].splitlines()

    definitions = []
    current_definition = None
    current_page = None

    for line in lines:

        page_match = re.match(
            r"\[Physical PDF page (\d+)\]",
            line.strip()
        )

        if page_match:
            current_page = int(page_match.group(1))
            continue

        # Skip PDF headers and footers.
        # These contain the category name and document reference.
        if (
            clause["document_reference"] in line
            or re.match(r"^\s*Page\s+\d+\s+of\s+\d+", line, re.IGNORECASE)
        ):
            continue

        match = DEFINITION_START.match(line)

        if match:

            if current_definition:
                definitions.append(current_definition)

            term = match.group("term").strip()

            current_definition = {
                "term": term,
                "text_lines": [line.strip()],
                "start_page": current_page,
                "end_page": current_page,
                "page_numbers": [current_page]
            }

        elif current_definition and line.strip():

            current_definition["text_lines"].append(
                line.strip()
            )

            if current_page is not None:

                current_definition["end_page"] = current_page

                if current_page not in current_definition["page_numbers"]:
                    current_definition["page_numbers"].append(
                        current_page
                    )

    if current_definition:
        definitions.append(current_definition)

    for definition in definitions:

        definition["text"] = " ".join(
            line.strip()
            for line in definition.pop("text_lines")
            if line.strip()
        )

        definition["text"] = re.sub(
            r"\s+",
            " ",
            definition["text"]
        ).strip()

    return definitions


def main():

    records = read_jsonl(MAIN_CLAUSE_FILE)

    print("=" * 80)
    print("DEFINITION EXTRACTION TEST")
    print("=" * 80)

    total_definitions = 0

    for category in sorted({r["category"] for r in records}):

        clause = next(
            r for r in records
            if r["category"] == category
            and r["clause_number"] == "1"
        )

        definitions = extract_definitions(clause)

        total_definitions += len(definitions)

        print("\nCATEGORY:", category)
        print("REFERENCE:", clause["document_reference"])
        print("Definitions detected:", len(definitions))

        for definition in definitions:

            print(
                f"  {definition['term']} "
                f"| pages {definition['start_page']}-{definition['end_page']} "
                f"| {len(definition['text'])} characters"
            )

            if not definition["text"].strip():
                raise ValueError(
                    f"Empty definition detected: {definition['term']}"
                )

    print("\n" + "=" * 80)
    print("TOTAL DEFINITIONS:", total_definitions)
    print("TEST COMPLETED")
    print("=" * 80)


if __name__ == "__main__":
    main()