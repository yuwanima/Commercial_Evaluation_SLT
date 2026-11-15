
"""
Vendor compliance schedule analyzer.

This module analyses extracted vendor-document evidence and produces
structured, vendor-declared compliance results.

Supported declaration statuses:
    - Complied
    - Not Complied
    - Partially Complied

Important:
    This module reports vendor-declared statuses. It does not independently
    establish whether the vendor has substantively complied with a contract.
"""

import re
from collections import Counter
from typing import Any


# ---------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------

SUPPORTED_STATUSES = (
    "Partially Complied",
    "Not Complied",
    "Complied",
)

STATUS_PATTERN = re.compile(
    r"\b(Partially\s+Complied|Not\s+Complied|Complied)\b",
    re.IGNORECASE,
)

# Matches clause headings such as:
# 1 Definitions and Contract Documents
# 1Definitions and Contract Documents
# 36 Counterpart signing and Digital Signature
CLAUSE_PATTERN = re.compile(
    r"(?m)^\s*(\d{1,3})\s*([A-Z][^\r\n]*)"
)

# Lines containing these phrases are treated as headings or metadata,
# not as individual compliance declarations.
IGNORED_STATUS_LINE_PATTERN = re.compile(
    r"\b("
    r"compliance\s+for|"
    r"complied\s+or|"
    r"not\s+complied\s+or|"
    r"explanatory\s+note|"
    r"clause\s+no|"
    r"description|"
    r"signature\s+of|"
    r"name\s+of\s+project|"
    r"tender\s+no|"
    r"official\s+frank|"
    r"page\s+\d+"
    r")\b",
    re.IGNORECASE,
)


# ---------------------------------------------------------------------
# Text helpers
# ---------------------------------------------------------------------

def clean_text(value: Any) -> str:
    """Convert extracted content into clean, safe text."""

    if value is None:
        return ""

    text = str(value)

    text = text.replace("\x00", " ")
    text = text.replace("\u00a0", " ")
    text = text.replace("\r\n", "\n")
    text = text.replace("\r", "\n")

    return text.strip()


def normalise_status(value: str) -> str | None:
    """Convert a recognised status into its standard form."""

    if not value:
        return None

    value = re.sub(r"\s+", " ", value.strip().lower())

    if value == "partially complied":
        return "Partially Complied"

    if value == "not complied":
        return "Not Complied"

    if value == "complied":
        return "Complied"

    return None


def get_source_details(section: dict) -> dict:
    """Return source information for an extracted evidence section."""

    return {
        "file_name": section.get("file_name") or "Unknown file",
        "location": section.get("location") or "Unknown location",
    }


# ---------------------------------------------------------------------
# Clause extraction
# ---------------------------------------------------------------------

def extract_clauses(evidence: list[dict]) -> list[dict]:
    """
    Extract numbered clause headings from all evidence sections.

    The first occurrence of a clause number is retained. Duplicate clause
    numbers are reported separately by the analysis function.
    """

    clauses_by_number = {}
    duplicate_numbers = []

    for section in evidence:
        text = clean_text(section.get("text"))

        if not text:
            continue

        source = get_source_details(section)

        for match in CLAUSE_PATTERN.finditer(text):
            clause_number = int(match.group(1))
            description = match.group(2).strip()

            # Ignore obvious page numbers and invalid headings.
            if not description:
                continue

            if clause_number < 1 or clause_number > 999:
                continue

            # Remove trailing page/header artefacts.
            description = re.sub(
                r"\s+COMPLIANCE FOR GENERAL CONDITIONS OF CONTRACT.*$",
                "",
                description,
                flags=re.IGNORECASE,
            ).strip()

            if not description:
                continue

            clause_record = {
                "clause_number": clause_number,
                "description": description,
                "vendor_status": "Not Specified",
                "explanatory_note": None,
                "source": source,
                "status_source": None,
                "verification": "Vendor declaration not yet matched",
            }

            if clause_number in clauses_by_number:
                duplicate_numbers.append(clause_number)
                continue

            clauses_by_number[clause_number] = clause_record

    clauses = list(clauses_by_number.values())
    clauses.sort(key=lambda item: item["clause_number"])

    return clauses, sorted(set(duplicate_numbers))


# ---------------------------------------------------------------------
# Status extraction
# ---------------------------------------------------------------------


def extract_status_declarations(evidence: list[dict]) -> list[dict]:
    """
    Extract compliance declarations after the clause list on each page.

    This avoids counting column headings such as "Complied or Not
    Complied" as actual vendor declarations.
    """

    declarations = []

    for section in evidence:
        text = clean_text(section.get("text"))

        if not text:
            continue

        source = get_source_details(section)
        lines = text.splitlines()

        # Find all clause-heading lines on this page.
        clause_line_indexes = [
            index
            for index, line in enumerate(lines)
            if CLAUSE_PATTERN.match(line)
        ]

        # If no clause heading exists, do not guess where declarations
        # begin on this page.
        if not clause_line_indexes:
            continue

        # Compliance declarations are expected after the final
        # clause heading on the page.
        first_status_line = clause_line_indexes[-1] + 1

        for line_index in range(first_status_line, len(lines)):
            line = lines[line_index].strip()

            if not line:
                continue

            if IGNORED_STATUS_LINE_PATTERN.search(line):
                continue

            matches = STATUS_PATTERN.finditer(line)

            for match in matches:
                status = normalise_status(match.group(0))

                if status is None:
                    continue

                declarations.append({
                    "status": status,
                    "source": source,
                    "line_number": line_index + 1,
                    "original_text": line,
                })

    return declarations



# ---------------------------------------------------------------------
# Status matching
# ---------------------------------------------------------------------

def match_statuses_to_clauses(
    clauses: list[dict],
    declarations: list[dict],
) -> tuple[list[dict], list[str]]:
    """
    Match declarations to clauses by document order.

    Matching is performed only when the number of extracted clauses
    exactly equals the number of declarations.

    This avoids silently assigning incorrect statuses when extraction
    is incomplete or ambiguous.
    """

    warnings = []

    if not clauses:
        warnings.append(
            "No numbered GCC clauses could be extracted."
        )
        return clauses, warnings

    if len(clauses) != len(declarations):
        warnings.append(
            "The number of extracted clauses does not match the number "
            "of compliance declarations. Automatic status matching was "
            "not performed."
        )

        warnings.append(
            f"Extracted clauses: {len(clauses)}; "
            f"compliance declarations: {len(declarations)}."
        )

        return clauses, warnings

    for clause, declaration in zip(clauses, declarations):
        clause["vendor_status"] = declaration["status"]
        clause["status_source"] = declaration["source"]
        clause["verification"] = (
            "Vendor-declared status; independent verification required"
        )

    return clauses, warnings


# ---------------------------------------------------------------------
# Summary calculations
# ---------------------------------------------------------------------

def calculate_summary(clauses: list[dict]) -> dict:
    """Calculate status counts and percentages deterministically."""

    counts = Counter(
        clause.get("vendor_status", "Not Specified")
        for clause in clauses
    )

    total = len(clauses)

    status_order = [
        "Complied",
        "Not Complied",
        "Partially Complied",
        "Not Specified",
        "Unclear",
        "OCR Required",
    ]

    breakdown = {}

    for status in status_order:
        count = counts.get(status, 0)

        breakdown[status] = {
            "count": count,
            "percentage": round(
                (count / total) * 100, 2
            ) if total else 0,
        }

    return {
        "total_clauses": total,
        "complied": counts.get("Complied", 0),
        "not_complied": counts.get("Not Complied", 0),
        "partially_complied": counts.get("Partially Complied", 0),
        "not_specified": counts.get("Not Specified", 0),
        "unclear": counts.get("Unclear", 0),
        "ocr_required": counts.get("OCR Required", 0),
        "breakdown": breakdown,
    }


# ---------------------------------------------------------------------
# Main analysis function
# ---------------------------------------------------------------------

def analyze_compliance(evidence: list[dict]) -> dict:
    """
    Analyse all extracted evidence for a vendor submission.

    Expected evidence format:

        [
            {
                "file_name": "compliance_schedule.pdf",
                "location": "Page 1",
                "text": "1 Definitions ...",
                "requires_ocr": False
            }
        ]

    Returns a structured result suitable for:
        - JSON storage
        - Flask API responses
        - Compliance charts
        - Clause-by-clause frontend tables
    """

    if not isinstance(evidence, list):
        raise ValueError("Evidence must be provided as a list.")

    valid_evidence = [
        section
        for section in evidence
        if isinstance(section, dict)
    ]

    if not valid_evidence:
        raise ValueError("No valid extracted evidence was provided.")

    # Identify documents/pages requiring OCR.
    ocr_required_sources = []

    for section in valid_evidence:
        if section.get("requires_ocr"):
            ocr_required_sources.append(
                get_source_details(section)
            )

    # Extract clause headings and declarations.
    clauses, duplicate_numbers = extract_clauses(valid_evidence)

    declarations = extract_status_declarations(valid_evidence)

    # Match statuses only if the counts agree.
    clauses, warnings = match_statuses_to_clauses(
        clauses,
        declarations,
    )

    if duplicate_numbers:
        warnings.append(
            "Duplicate clause numbers were found: "
            + ", ".join(map(str, duplicate_numbers))
        )

    if ocr_required_sources:
        warnings.append(
            "Some document sections require OCR. "
            "Compliance analysis may be incomplete."
        )

    if not declarations:
        warnings.append(
            "No compliance declarations were identified."
        )

    summary = calculate_summary(clauses)

    # Determine whether the analysis can be considered complete.
    analysis_complete = (
        bool(clauses)
        and len(clauses) == len(declarations)
        and not duplicate_numbers
        and not ocr_required_sources
        and all(
          clause.get("vendor_status", "Not Specified") != "Not Specified"
          for clause in clauses
        )
    )
    

    return {
        "analysis_type": "vendor_declared_compliance",
        "analysis_complete": analysis_complete,
        "summary": summary,
        "clauses": clauses,
        "declaration_count": len(declarations),
        "duplicate_clause_numbers": duplicate_numbers,
        "ocr_required_sources": ocr_required_sources,
        "warnings": warnings,
        "notice": (
            "These results represent compliance statuses declared "
            "by the vendor. They do not constitute independent "
            "verification of contractual compliance."
        ),
    }


# ---------------------------------------------------------------------
# Optional helper for chart data
# ---------------------------------------------------------------------

def get_chart_data(analysis_result: dict) -> dict:
    """
    Return chart-friendly data for the frontend.

    The frontend can use these values with Chart.js.
    """

    summary = analysis_result.get("summary", {})
    breakdown = summary.get("breakdown", {})

    chart_statuses = [
        "Complied",
        "Not Complied",
        "Partially Complied",
        "Not Specified",
        "Unclear",
        "OCR Required",
    ]

    labels = []
    counts = []
    percentages = []

    for status in chart_statuses:
        count = breakdown.get(status, {}).get("count", 0)

        # Do not include zero-count categories in the graph.
        if count == 0:
            continue

        labels.append(status)
        counts.append(count)
        percentages.append(
            breakdown[status].get("percentage", 0)
        )

    return {
        "labels": labels,
        "counts": counts,
        "percentages": percentages,
    }
