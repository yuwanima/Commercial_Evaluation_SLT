import json
import re
from pathlib import Path
from collections import Counter


# ------------------------------------------------------------
# PATHS
# ------------------------------------------------------------

PROJECT_ROOT = Path(__file__).resolve().parents[2]

STORAGE_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

MAIN_CLAUSE_FILE = STORAGE_DIR / "gcc_clauses.jsonl"
SUBCLAUSE_FILE = STORAGE_DIR / "gcc_subclauses.jsonl"

OUTPUT_FILE = STORAGE_DIR / "retrieval_units.jsonl"


# ------------------------------------------------------------
# DEFINITION EXTRACTION
# ------------------------------------------------------------

DEFINITION_START = re.compile(
    r"^\s*(?P<term>.+?)\s+means\b",
    re.IGNORECASE
)


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
        if (
            clause["document_reference"] in line
            or re.match(
                r"^\s*Page\s+\d+\s+of\s+\d+",
                line,
                re.IGNORECASE
            )
        ):
            continue

        match = DEFINITION_START.match(line)

        if match:

            if current_definition:
                definitions.append(current_definition)

            current_definition = {
                "term": match.group("term").strip(),
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


# ------------------------------------------------------------
# JSONL READER
# ------------------------------------------------------------

def read_jsonl(path):

    if not path.exists():
        raise FileNotFoundError(
            f"Required input file not found: {path}"
        )

    records = []

    with path.open("r", encoding="utf-8") as file:

        for line_number, line in enumerate(file, start=1):

            if not line.strip():
                continue

            try:
                records.append(json.loads(line))

            except json.JSONDecodeError as error:
                raise ValueError(
                    f"Invalid JSON in {path.name}, "
                    f"line {line_number}: {error}"
                )

    return records


# ------------------------------------------------------------
# PREPARE MAIN CLAUSE LOOKUP
# ------------------------------------------------------------

def build_main_clause_lookup(main_clauses):

    lookup = {}

    for clause in main_clauses:

        key = (
            clause["category"],
            clause["document_reference"],
            clause["clause_number"]
        )

        if key in lookup:
            raise ValueError(
                f"Duplicate main clause key: {key}"
            )

        lookup[key] = clause

    return lookup


# ------------------------------------------------------------
# BUILD RETRIEVAL RECORDS
# ------------------------------------------------------------

def prepare_retrieval_units(main_clauses, subclauses):

    main_lookup = build_main_clause_lookup(main_clauses)

    retrieval_units = []

    # Track which main clauses contain extracted subclauses.
    clauses_with_subclauses = set()

    # --------------------------------------------------------
    # Subclause retrieval units
    # --------------------------------------------------------

    for subclause in subclauses:

        main_key = (
            subclause["category"],
            subclause["document_reference"],
            subclause["main_clause_number"]
        )

        parent_clause = main_lookup.get(main_key)

        if parent_clause is None:
            raise ValueError(
                "Main clause not found for subclause: "
                f"{subclause.get('record_id')}"
            )

        clauses_with_subclauses.add(main_key)

        parent_title = parent_clause.get("clause_title", "").strip()
        parent_context = (
            f"Main clause {subclause['main_clause_number']}: "
            f"{parent_title}"
        )

        subclause_text = subclause.get("text", "").strip()

        retrieval_text = (
            f"Procurement category: {subclause['category']}\n"
            f"Document reference: {subclause['document_reference']}\n"
            f"{parent_context}\n"
            f"Subclause: {subclause['clause_number']} "
            f"- {subclause.get('clause_title', '')}\n\n"
            f"Subclause content:\n{subclause_text}"
        )

        retrieval_units.append({

            "retrieval_id": subclause["record_id"],

            "record_type": "subclause",

            "category": subclause["category"],

            "document_type": subclause["document_type"],

            "document_reference": subclause["document_reference"],

            "source_file": subclause["source_file"],

            "main_clause_number": subclause["main_clause_number"],

            "immediate_parent_number": (
                subclause["immediate_parent_number"]
            ),

            "clause_number": subclause["clause_number"],

            "clause_title": subclause["clause_title"],

            "hierarchy_level": subclause["hierarchy_level"],

            "start_page": subclause["start_page"],

            "end_page": subclause["end_page"],

            "page_numbers": subclause["page_numbers"],

            "source_numbering_warning": subclause.get(
                "source_numbering_warning"
            ),

            "text": retrieval_text,

            "original_text": subclause_text,

            "parent_context": parent_context

        })

    # --------------------------------------------------------
    # Standalone main clauses
    # --------------------------------------------------------
    # Main clauses without extracted subclauses must remain
    # searchable as complete records.

    for clause in main_clauses:

        key = (
            clause["category"],
            clause["document_reference"],
            clause["clause_number"]
        )

        if key in clauses_with_subclauses:
            continue

        original_text = clause.get("text", "").strip()

        retrieval_text = (
            f"Procurement category: {clause['category']}\n"
            f"Document reference: {clause['document_reference']}\n"
            f"Main clause: {clause['clause_number']} "
            f"- {clause.get('clause_title', '')}\n\n"
            f"Clause content:\n{original_text}"
        )

        retrieval_units.append({

            "retrieval_id": (
                f"{clause['document_reference'].replace('/', '-')}-"
                f"{clause['clause_number'].replace('.', '-')}-MAIN"
            ),

            "record_type": "main_clause",

            "category": clause["category"],

            "document_type": clause["document_type"],

            "document_reference": clause["document_reference"],

            "source_file": clause["source_file"],

            "main_clause_number": clause["clause_number"],

            "immediate_parent_number": None,

            "clause_number": clause["clause_number"],

            "clause_title": clause["clause_title"],

            "hierarchy_level": 0,

            "start_page": clause["start_page"],

            "end_page": clause["end_page"],

            "page_numbers": clause.get(
                "page_numbers",
                list(range(
                    clause["start_page"],
                    clause["end_page"] + 1
                ))
            ),

            "source_numbering_warning": None,

            "text": retrieval_text,

            "original_text": original_text,

            "parent_context": None

        })

    # --------------------------------------------------------
    # Definition retrieval units
    # --------------------------------------------------------
    # Each definition receives its own searchable record.
    # Original GCC clause numbering is preserved.
    # Definition IDs are generated separately.

    for clause in main_clauses:

        if clause["clause_number"] != "1":
            continue

        definitions = extract_definitions(clause)

        for definition_index, definition in enumerate(
            definitions,
            start=1
        ):

            reference_id = (
                clause["document_reference"].replace("/", "-")
            )

            retrieval_id = (
                f"{reference_id}-DEF-{definition_index:03d}"
            )

            parent_context = "Main clause 1: Definitions"

            original_text = definition["text"]

            retrieval_text = (
                f"Procurement category: {clause['category']}\n"
                f"Document reference: {clause['document_reference']}\n"
                f"{parent_context}\n"
                f"Definition: {definition['term']}\n\n"
                f"Definition content:\n{original_text}"
            )

            retrieval_units.append({

                "retrieval_id": retrieval_id,

                "record_type": "definition",

                "category": clause["category"],

                "document_type": clause["document_type"],

                "document_reference": clause["document_reference"],

                "source_file": clause["source_file"],

                "main_clause_number": "1",

                "immediate_parent_number": "1.1",

                "clause_number": "1.1",

                "clause_title": "Definitions",

                "definition_term": definition["term"],

                "hierarchy_level": 1,

                "start_page": definition["start_page"],

                "end_page": definition["end_page"],

                "page_numbers": definition["page_numbers"],

                "source_numbering_warning": None,

                "text": retrieval_text,

                "original_text": original_text,

                "parent_context": parent_context

            })

    return retrieval_units


# ------------------------------------------------------------
# VALIDATE RETRIEVAL RECORDS
# ------------------------------------------------------------

def validate_retrieval_units(units, main_clauses, subclauses):

    errors = []

    # --------------------------------------------------------
    # 1. Validate unique retrieval IDs
    # --------------------------------------------------------

    retrieval_ids = [
        unit.get("retrieval_id")
        for unit in units
    ]

    duplicate_ids = [
        value
        for value, count in Counter(retrieval_ids).items()
        if value and count > 1
    ]

    if duplicate_ids:
        errors.append(
            f"Duplicate retrieval IDs: {duplicate_ids}"
        )

    if any(not value for value in retrieval_ids):
        errors.append("Missing retrieval IDs detected.")

    # --------------------------------------------------------
    # 2. Validate required fields
    # --------------------------------------------------------

    required_fields = [
        "retrieval_id",
        "record_type",
        "category",
        "document_reference",
        "source_file",
        "clause_number",
        "start_page",
        "end_page",
        "page_numbers",
        "text",
        "original_text"
    ]

    for index, unit in enumerate(units, start=1):

        missing = [
            field
            for field in required_fields
            if field not in unit or unit[field] is None
        ]

        if missing:
            errors.append(
                f"Record {index} missing fields: {missing}"
            )

        if not isinstance(unit.get("text"), str) or not unit["text"].strip():
            errors.append(
                f"Record {index} has empty or invalid retrieval text."
            )

        if not isinstance(unit.get("original_text"), str) or not unit["original_text"].strip():
            errors.append(
                f"Record {index} has empty or invalid original text."
            )

        if (
            not isinstance(unit.get("start_page"), int)
            or not isinstance(unit.get("end_page"), int)
            or unit["start_page"] < 1
            or unit["end_page"] < unit["start_page"]
        ):
            errors.append(
                f"Record {index} has invalid page range."
            )

        page_numbers = unit.get("page_numbers")

        if (
            not isinstance(page_numbers, list)
            or not page_numbers
            or any(
                not isinstance(page, int)
                or isinstance(page, bool)
                or page < 1
                for page in page_numbers
            )
        ):
            errors.append(
                f"Record {index} has invalid page numbers."
            )

        elif (
            isinstance(unit.get("start_page"), int)
            and isinstance(unit.get("end_page"), int)
            and (
                min(page_numbers) < unit["start_page"]
                or max(page_numbers) > unit["end_page"]
            )
        ):
            errors.append(
                f"Record {index} page numbers fall outside its page range."
            )

        if unit.get("record_type") == "definition":

            if not isinstance(unit.get("definition_term"), str) or not unit["definition_term"].strip():
                errors.append(
                    f"Definition record {index} has an empty or invalid term."
                )

    # --------------------------------------------------------
    # 3. Validate record types and expected counts
    # --------------------------------------------------------

    actual_counts = Counter(
        unit.get("record_type")
        for unit in units
    )

    # Calculate expected standalone main clauses from the
    # extracted hierarchy instead of hardcoding an old count.
    main_clause_keys = {
        (
            clause["category"],
            clause["document_reference"],
            clause["clause_number"]
        )
        for clause in main_clauses
    }

    subclause_parent_keys = {
        (
            subclause["category"],
            subclause["document_reference"],
            subclause["main_clause_number"]
        )
        for subclause in subclauses
    }

    expected_main_clauses = len(
        main_clause_keys - subclause_parent_keys
    )

    expected_counts = {
        "main_clause": expected_main_clauses,
        "subclause": len(subclauses),
        "definition": 103
    }

    for record_type, expected_count in expected_counts.items():

        actual_count = actual_counts.get(record_type, 0)

        if actual_count != expected_count:
            errors.append(
                f"{record_type} count mismatch: "
                f"expected {expected_count}, "
                f"found {actual_count}."
            )

    unexpected_types = set(actual_counts) - set(expected_counts)

    if unexpected_types:
        errors.append(
            f"Unexpected record types: {sorted(unexpected_types)}"
        )

    # --------------------------------------------------------
    # 4. Validate source main clauses
    # --------------------------------------------------------

    if len(main_clauses) != 210:
        errors.append(
            f"Expected 210 source main clauses, "
            f"found {len(main_clauses)}."
        )

    # --------------------------------------------------------
    # 5. Validate procurement categories
    # --------------------------------------------------------

    expected_categories = {
        clause["category"]
        for clause in main_clauses
    }

    actual_categories = {
        unit.get("category")
        for unit in units
    }

    if actual_categories != expected_categories:
        errors.append(
            "Procurement category mismatch. "
            f"Expected: {sorted(expected_categories)}; "
            f"Found: {sorted(actual_categories)}."
        )

    # --------------------------------------------------------
    # 6. Validate definition records per category
    # --------------------------------------------------------

    expected_definitions_by_category = {}

    for clause in main_clauses:

        if clause["clause_number"] != "1":
            continue

        expected_definitions_by_category[clause["category"]] = len(
            extract_definitions(clause)
        )

    actual_definitions_by_category = Counter(
        unit["category"]
        for unit in units
        if unit.get("record_type") == "definition"
    )

    if actual_definitions_by_category != Counter(
        expected_definitions_by_category
    ):
        errors.append(
            "Definition count by category mismatch. "
            f"Expected: {dict(expected_definitions_by_category)}; "
            f"Found: {dict(actual_definitions_by_category)}."
        )

    # --------------------------------------------------------
    # 7. Validate total retrieval count
    # --------------------------------------------------------

    expected_total = sum(expected_counts.values())

    if len(units) != expected_total:
        errors.append(
            f"Total retrieval count mismatch: "
            f"expected {expected_total}, "
            f"found {len(units)}."
        )

    if errors:
        raise ValueError(
            "\n".join(errors)
        )


# ------------------------------------------------------------
# MAIN
# ------------------------------------------------------------

def main():

    print("=" * 75)
    print("SLT PROCUREMENT RETRIEVAL UNIT PREPARATION")
    print("=" * 75)

    main_clauses = read_jsonl(MAIN_CLAUSE_FILE)
    subclauses = read_jsonl(SUBCLAUSE_FILE)

    print("Main clauses loaded:", len(main_clauses))
    print("Subclauses loaded:", len(subclauses))

    retrieval_units = prepare_retrieval_units(
        main_clauses,
        subclauses
    )

    validate_retrieval_units(
        retrieval_units,
        main_clauses,
        subclauses
    )

    # Do not overwrite an existing output without confirmation.
    if OUTPUT_FILE.exists():
        raise FileExistsError(
            f"Output already exists: {OUTPUT_FILE}\n"
            "Move or back up the existing file before rerunning."
        )

    OUTPUT_FILE.parent.mkdir(
        parents=True,
        exist_ok=True
    )

    with OUTPUT_FILE.open("w", encoding="utf-8") as file:

        for unit in retrieval_units:

            file.write(
                json.dumps(unit, ensure_ascii=False) + "\n"
            )

    print("\nRETRIEVAL PREPARATION SUMMARY")
    print("-" * 75)

    print("Total retrieval units:", len(retrieval_units))

    print(
        "Subclause units:",
        sum(
            unit["record_type"] == "subclause"
            for unit in retrieval_units
        )
    )

    print(
        "Definition units:",
        sum(
            unit["record_type"] == "definition"
            for unit in retrieval_units
        )
    )

    print(
        "Standalone main clause units:",
        sum(
            unit["record_type"] == "main_clause"
            for unit in retrieval_units
        )
    )

    for category in sorted({
        unit["category"]
        for unit in retrieval_units
    }):

        category_count = sum(
            unit["category"] == category
            for unit in retrieval_units
        )

        print(f"{category}: {category_count} retrieval units")

    print("\nValidation: PASSED")
    print("Output:", OUTPUT_FILE)
    print("\nRetrieval unit preparation completed.")


if __name__ == "__main__":
    main()
