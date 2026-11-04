
import json
import re
from pathlib import Path
from datetime import datetime, timezone

from openpyxl import load_workbook


# ============================================================
# PATHS
# ============================================================

PROJECT_ROOT = Path(__file__).resolve().parents[2]

SOURCE_ROOT = (
    PROJECT_ROOT
    / "Backend"
    / "documents"
    / "rag_sources"
    / "GCC & SCC"
)

INPUT_FILE = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
    / "extracted_documents.jsonl"
)

OUTPUT_DIR = (
    PROJECT_ROOT
    / "Backend"
    / "storage"
    / "rag_chatbot"
)

CLAUSE_OUTPUT = OUTPUT_DIR / "gcc_clauses.jsonl"
SUBCLAUSE_OUTPUT = OUTPUT_DIR / "gcc_subclauses.jsonl"
REPORT_OUTPUT = OUTPUT_DIR / "gcc_clause_validation.txt"


# ============================================================
# HEADING PATTERNS
# ============================================================

RANGE_HEADING = re.compile(
    r"^\s*8\s*(?:to|[-–])\s*12\s+(.+?)\s*$",
    re.IGNORECASE
)

NORMAL_HEADING = re.compile(
    r"^\s*(\d{1,2})\.?\s+([A-Z].*?)\s*$"
)

SUBCLAUSE_HEADING = re.compile(
    r"^\s*(\d{1,2}(?:\.\d+)+)\.?\s+([A-Z].*?)\s*$"
)

# ============================================================
# NORMALIZE CLAUSE IDENTIFIERS
# ============================================================

def normalize_clause_id(value):

    if value is None:
        return None

    text = str(value).strip().lower()
    text = re.sub(r"\s+", "", text)

    if re.fullmatch(r"8(?:to|-|–)12", text):
        return "8to12"

    if re.fullmatch(r"\d+", text):
        return text

    return None


# ============================================================
# READ EXPECTED CLAUSES FROM COMPLIANCE WORKBOOK
# ============================================================

def read_workbook_clause_ids(category):

    category_folder = SOURCE_ROOT / category
    workbook_files = list(category_folder.glob("*.xlsx"))

    if len(workbook_files) != 1:
        raise ValueError(
            f"Expected one Excel workbook in {category}, "
            f"found {len(workbook_files)}"
        )

    workbook = load_workbook(
        workbook_files[0],
        read_only=True,
        data_only=True
    )

    sheet = workbook[workbook.sheetnames[0]]

    header_row = None
    clause_column = None

    for row_number in range(1, sheet.max_row + 1):

        for column_number in range(1, sheet.max_column + 1):

            value = sheet.cell(row_number, column_number).value

            if (
                value is not None
                and str(value).strip().lower() == "clause no"
            ):
                header_row = row_number
                clause_column = column_number
                break

        if header_row is not None:
            break

    if header_row is None:
        workbook.close()
        raise ValueError(
            f"Could not find 'Clause No' in {workbook_files[0].name}"
        )

    expected = set()

    for row_number in range(header_row + 1, sheet.max_row + 1):

        value = sheet.cell(row_number, clause_column).value
        clause_id = normalize_clause_id(value)

        if clause_id:
            expected.add(clause_id)

    workbook.close()

    return expected


# ============================================================
# IDENTIFY MAIN GCC HEADINGS
# ============================================================

def detect_main_heading(line):

    range_match = RANGE_HEADING.match(line)

    if range_match:
        return "8to12", range_match.group(1).strip()

    normal_match = NORMAL_HEADING.match(line)

    if normal_match:
        clause_number = normal_match.group(1)
        title = normal_match.group(2).strip()

        return clause_number, title

    return None


# ============================================================
# IDENTIFY SUBCLAUSE HEADINGS
# ============================================================

def detect_subclause_heading(line):

    match = SUBCLAUSE_HEADING.match(line)

    if not match:
        return None

    return match.group(1), match.group(2).strip()


# ============================================================
# JOIN WRAPPED HEADING LINES
# ============================================================

def get_heading_title(lines, index, initial_title):

    title = initial_title.strip()

    if not title:
        return title

    continuation_endings = (
        "SLT",
        "third party",
        "and",
        "of",
        "for",
        "to",
        "with",
        "/"
    )

    lower_title = title.lower()

    needs_continuation = (
        len(title) >= 65
        or lower_title.endswith(continuation_endings)
    )

    if not needs_continuation:
        return title

    next_index = index + 1

    while next_index < len(lines) and not lines[next_index].strip():
        next_index += 1

    if next_index >= len(lines):
        return title

    next_line = lines[next_index].strip()

    if (
        len(next_line) <= 80
        and not detect_main_heading(next_line)
        and not detect_subclause_heading(next_line)
        and not next_line.endswith((".", ";", ":"))
    ):
        title = title + " " + next_line

    return title


# ============================================================
# FORMAT TEXT WHILE RETAINING PHYSICAL PAGE MARKERS
# ============================================================

def build_text(flat_lines, start, end):

    output = []
    previous_page = None

    for page_number, line in flat_lines[start:end]:

        if page_number != previous_page:

            if output:
                output.append("")

            output.append(
                f"[Physical PDF page {page_number}]"
            )

            previous_page = page_number

        output.append(line)

    return "\n".join(output).strip()


# ============================================================
# EXTRACT CLAUSES FROM ONE CATEGORY
# ============================================================

def extract_category(category, page_records):

    expected_ids = read_workbook_clause_ids(category)

    flat_lines = []

    for record in sorted(
        page_records,
        key=lambda item: item["page_number"]
    ):

        page_number = record["page_number"]

        # Page 1 is the table of contents.
        # Actual clause extraction begins on physical page 2.
        if page_number < 2:
            continue

        for line in record["text"].splitlines():
            flat_lines.append((page_number, line))

    headings = []

    for index, (page_number, line) in enumerate(flat_lines):

        detected = detect_main_heading(line)

        if not detected:
            continue

        clause_number, initial_title = detected

        title = get_heading_title(
            [item[1] for item in flat_lines],
            index,
            initial_title
        )

        headings.append({
            "clause_number": clause_number,
            "clause_title": title,
            "page_number": page_number,
            "line_index": index
        })

    clauses = []

    for index, heading in enumerate(headings):

        start = heading["line_index"]

        if index + 1 < len(headings):
            end = headings[index + 1]["line_index"]
        else:
            end = len(flat_lines)

        clause_text = build_text(
            flat_lines,
            start,
            end
        )

        page_numbers = sorted({
            page
            for page, _ in flat_lines[start:end]
        })

        clauses.append({
            "record_type": "main_clause",
            "category": category,
            "document_type": "PDF",
            "clause_number": heading["clause_number"],
            "clause_title": heading["clause_title"],
            "source_file": page_records[0]["source_file"],
            "document_reference": page_records[0].get(
                "document_reference"
            ),
            "start_page": min(page_numbers),
            "end_page": max(page_numbers),
            "page_numbers": page_numbers,
            "text": clause_text
        })

    actual_ids = {
        clause["clause_number"]
        for clause in clauses
    }

    missing_ids = sorted(
        expected_ids - actual_ids,
        key=lambda value: (value.isdigit(), value)
    )

    extra_ids = sorted(
        actual_ids - expected_ids,
        key=lambda value: (value.isdigit(), value)
    )

    # --------------------------------------------------------
    # Extract individual subclauses while preserving their hierarchy
    # --------------------------------------------------------

    subclauses = []

    for clause in clauses:

        clause_lines = clause["text"].splitlines()
        subheadings = []

        for line_index, line in enumerate(clause_lines):

            detected = detect_subclause_heading(line)

            if not detected:
                continue

            subclause_number, subclause_title = detected

            main_parent = subclause_number.split(".")[0]

            if main_parent != clause["clause_number"]:
                continue

            subheadings.append({
                "number": subclause_number,
                "title": subclause_title,
                "line_index": line_index,
                "depth": subclause_number.count(".")
            })

        # Map every line to its physical PDF page.
        # A page marker applies to all following lines.
        line_pages = []
        current_page = None

        for line in clause_lines:
            page_match = re.match(
                r"^\[Physical PDF page\s+(\d+)\]\s*$",
                line.strip()
            )

            if page_match:
                current_page = int(page_match.group(1))

            line_pages.append(current_page)

        for sub_index, subheading in enumerate(subheadings):

            start_line = subheading["line_index"]
            current_depth = subheading["depth"]

            # A parent includes its nested children.
            # It ends at the next heading at the same or a higher level.
            end_line = len(clause_lines)

            for next_heading in subheadings[sub_index + 1:]:

                if next_heading["depth"] <= current_depth:
                    end_line = next_heading["line_index"]
                    break

            subclause_text = "\n".join(
                clause_lines[start_line:end_line]
            ).strip()

            # Calculate pages using the physical page assigned to
            # each line, including the starting page of this subclause.
            page_numbers = sorted({
                page_number
                for page_number in line_pages[start_line:end_line]
                if page_number is not None
            })

            # Fall back to the main clause's page metadata only if
            # this subclause contains no physical page marker.
            if not page_numbers:
                page_numbers = clause.get(
                    "page_numbers",
                    list(range(
                        clause["start_page"],
                        clause["end_page"] + 1
                    ))
                )

            subclause_number = subheading["number"]

            immediate_parent = subclause_number.rsplit(".", 1)[0]

            subclauses.append({
                "record_type": "subclause",
                "category": category,
                "document_type": "PDF",

                # Retain the original main-clause parent field.
                "parent_clause_number": clause["clause_number"],

                # Additional hierarchy metadata.
                "main_clause_number": clause["clause_number"],
                "immediate_parent_number": immediate_parent,
                "hierarchy_level": subheading["depth"],

                "clause_number": subclause_number,
                "clause_title": subheading["title"],

                "source_file": clause["source_file"],
                "document_reference": clause["document_reference"],

                "start_page": min(page_numbers),
                "end_page": max(page_numbers),
                "page_numbers": page_numbers,

                "text": subclause_text
            })

    validation = {
        "category": category,
        "expected_count": len(expected_ids),
        "actual_count": len(clauses),
        "missing_ids": missing_ids,
        "extra_ids": extra_ids,
        "subclause_count": len(subclauses)
    }

    return clauses, subclauses, validation


# ============================================================
# MAIN
# ============================================================

def main():

    print("=" * 75)
    print("SLT GCC CLAUSE EXTRACTION")
    print("=" * 75)

    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Input file not found: {INPUT_FILE}"
        )

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    with INPUT_FILE.open(encoding="utf-8") as file:
        records = [
            json.loads(line)
            for line in file
            if line.strip()
        ]

    categories = sorted({
        record["category"]
        for record in records
        if record.get("document_type") == "PDF"
    })

    all_clauses = []
    all_subclauses = []
    validations = []

    for category in categories:

        print(f"\nCATEGORY: {category}")

        page_records = [
            record
            for record in records
            if (
                record.get("category") == category
                and record.get("document_type") == "PDF"
                and "General Guidelines" not in record["source_file"]
            )
        ]

        if not page_records:
            print("WARNING: No main GCC PDF records found.")
            continue

        references = set()

        for record in page_records:
            match = re.search(
                r"LGLS/FCC/\d+",
                record["text"]
            )

            if match:
                references.add(match.group(0))

        document_reference = (
            sorted(references)[0]
            if references
            else None
        )

        for record in page_records:
            record["document_reference"] = document_reference

        clauses, subclauses, validation = extract_category(
            category,
            page_records
        )

        all_clauses.extend(clauses)
        all_subclauses.extend(subclauses)
        validations.append(validation)

        print("Document reference:", document_reference)
        print("Main clauses extracted:", len(clauses))
        print("Subclauses extracted:", len(subclauses))
        print("Expected workbook clauses:", validation["expected_count"])

        if validation["missing_ids"]:
            print("Missing from PDF:", validation["missing_ids"])

        if validation["extra_ids"]:
            print("Extra in PDF:", validation["extra_ids"])

    # --------------------------------------------------------
    # Write main clause records
    # --------------------------------------------------------

    with CLAUSE_OUTPUT.open("w", encoding="utf-8") as file:

        for record in all_clauses:

            record["extracted_at"] = datetime.now(
                timezone.utc
            ).isoformat()

            file.write(
                json.dumps(record, ensure_ascii=False) + "\n"
            )

    # --------------------------------------------------------
    # Assign deterministic IDs to subclause records
    # --------------------------------------------------------

    duplicate_groups = {}

    for record in all_subclauses:
        key = (
            record["document_reference"],
            record["clause_number"]
        )
        duplicate_groups[key] = duplicate_groups.get(key, 0) + 1

    occurrence_counts = {}

    for record in all_subclauses:
        reference = (record["document_reference"] or "UNKNOWN")
        reference_id = reference.replace("/", "-")

        clause_id = record["clause_number"].replace(".", "-")

        key = (
            record["document_reference"],
            record["clause_number"]
        )

        occurrence_counts[key] = occurrence_counts.get(key, 0) + 1
        occurrence = occurrence_counts[key]

        record["record_id"] = (
            f"{reference_id}-{clause_id}-{occurrence:02d}"
        )

        if duplicate_groups[key] > 1:
            record["source_numbering_warning"] = (
                "Duplicate clause number in original GCC PDF"
            )

    # --------------------------------------------------------
    # Write subclause records
    # --------------------------------------------------------

    with SUBCLAUSE_OUTPUT.open("w", encoding="utf-8") as file:

        for record in all_subclauses:

            record["extracted_at"] = datetime.now(
                timezone.utc
            ).isoformat()

            file.write(
                json.dumps(record, ensure_ascii=False) + "\n"
            )

    # --------------------------------------------------------
    # Write validation report
    # --------------------------------------------------------

    report_lines = [
        "SLT GCC CLAUSE VALIDATION REPORT",
        "=" * 75,
        f"Generated: {datetime.now(timezone.utc).isoformat()}",
        "",
        f"Total main clauses: {len(all_clauses)}",
        f"Total subclauses: {len(all_subclauses)}",
        ""
    ]

    for validation in validations:

        report_lines.extend([
            "-" * 75,
            f"CATEGORY: {validation['category']}",
            f"Workbook clauses: {validation['expected_count']}",
            f"PDF clauses detected: {validation['actual_count']}",
            f"Subclauses detected: {validation['subclause_count']}",
            f"Missing from PDF: {validation['missing_ids']}",
            f"Extra in PDF: {validation['extra_ids']}",
            ""
        ])

    report_lines.extend([
        "-" * 75,
        "IMPORTANT:",
        "Differences between GCC PDFs and compliance workbooks",
        "must be manually reviewed. Do not automatically renumber clauses.",
        ""
    ])

    REPORT_OUTPUT.write_text(
        "\n".join(report_lines),
        encoding="utf-8"
    )

    print("\n" + "=" * 75)
    print("EXTRACTION SUMMARY")
    print("=" * 75)

    print("Total main clauses:", len(all_clauses))
    print("Total subclauses:", len(all_subclauses))
    print("Main clause output:", CLAUSE_OUTPUT)
    print("Subclause output:", SUBCLAUSE_OUTPUT)
    print("Validation report:", REPORT_OUTPUT)

    print("\nExtraction completed.")


if __name__ == "__main__":
    main()
