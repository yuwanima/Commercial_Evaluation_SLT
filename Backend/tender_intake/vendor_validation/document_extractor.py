
from pathlib import Path

from docx import Document
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from pypdf import PdfReader


SUPPORTED_EXTENSIONS = {
    ".pdf",
    ".docx",
    ".xlsx",
    ".xlsm",
}


def extract_pdf(file_path):
    """Extract PDF text and preserve pages requiring OCR."""

    reader = PdfReader(str(file_path))
    sections = []

    if reader.is_encrypted:
        raise ValueError(
            f"The PDF is password-protected: {Path(file_path).name}"
        )

    for page_number, page in enumerate(reader.pages, start=1):
        text = page.extract_text() or ""

        if text.strip():
            sections.append({
                "location": f"Page {page_number}",
                "text": text.strip(),
                "requires_ocr": False,
            })
        else:
            sections.append({
                "location": f"Page {page_number}",
                "text": "",
                "requires_ocr": True,
            })

    if not sections:
        raise ValueError(
            f"No pages were found in PDF: {Path(file_path).name}"
        )

    return sections


def extract_docx(file_path):
    """Extract paragraphs and tables from a Word document."""

    document = Document(str(file_path))
    sections = []

    for index, paragraph in enumerate(document.paragraphs, start=1):
        text = paragraph.text.strip()

        if text:
            sections.append({
                "location": f"Paragraph {index}",
                "text": text,
            })

    for table_number, table in enumerate(document.tables, start=1):
        for row_number, row in enumerate(table.rows, start=1):
            values = [
                cell.text.strip().replace("\n", " | ")
                for cell in row.cells
            ]

            if any(values):
                sections.append({
                    "location": f"Table {table_number}, Row {row_number}",
                    "text": " | ".join(values),
                })

    return sections


def extract_excel(file_path):
    """
    Extract Excel data while preserving:
    - Worksheet and row numbers
    - Column headings
    - Empty cells in non-empty rows
    - Formulas and available cached results

    Formulas are never executed.
    """

    formula_workbook = load_workbook(
        filename=str(file_path),
        read_only=True,
        data_only=False,
        keep_links=False,
    )

    value_workbook = load_workbook(
        filename=str(file_path),
        read_only=True,
        data_only=True,
        keep_links=False,
    )

    sections = []

    header_keywords = (
        "clause",
        "description",
        "complied",
        "explanatory",
        "signature",
        "date",
    )

    try:
        for formula_sheet in formula_workbook.worksheets:
            formula_rows = formula_sheet.iter_rows()
            value_rows = value_workbook[
                formula_sheet.title
            ].iter_rows()

            rows = list(zip(formula_rows, value_rows))

            header_row_number = None
            headers = {}

            # Find a likely column-heading row.
            for row_number, (formula_row, _) in enumerate(
                rows,
                start=1,
            ):
                nonempty = [
                    str(cell.value).strip()
                    for cell in formula_row
                    if cell.value is not None
                    and str(cell.value).strip()
                ]

                matches = sum(
                    1
                    for keyword in header_keywords
                    if any(
                        keyword in value.lower()
                        for value in nonempty
                    )
                )

                has_description = any(
                    "description" in value.lower()
                    for value in nonempty
                )

                has_response_heading = any(
                    "complied" in value.lower()
                    or "explanatory" in value.lower()
                    for value in nonempty
                )

                if (
                    len(nonempty) >= 2
                    and (
                        (has_description and has_response_heading)
                        or matches >= 3
                    )
                ):
                    header_row_number = row_number

                    for column_number, cell in enumerate(
                        formula_row,
                        start=1,
                    ):
                        if cell.value is not None:
                            headers[column_number] = str(
                                cell.value
                            ).strip()

                    break

            # Extract worksheet data.
            for row_number, (formula_row, value_row) in enumerate(
                rows,
                start=1,
            ):
                if row_number == header_row_number:
                    continue

                # Ignore completely empty rows.
                has_content = any(
                    cell.value is not None
                    and str(cell.value).strip()
                    for cell in formula_row
                )

                if not has_content:
                    continue

                values = []

                for column_number, (
                    formula_cell,
                    value_cell,
                ) in enumerate(
                    zip(formula_row, value_row),
                    start=1,
                ):
                    coordinate = (
                        f"{get_column_letter(column_number)}{row_number}"
                    )

                    heading = headers.get(
                        column_number,
                        get_column_letter(column_number),
                    )

                    formula = formula_cell.value
                    cached_value = value_cell.value

                    if formula_cell.data_type == "f":
                        if cached_value is None:
                            cell_value = (
                                f"formula = {formula}; "
                                "cached result unavailable"
                            )
                        else:
                            cell_value = (
                                f"formula = {formula}; "
                                f"cached result: {cached_value}"
                            )

                    elif formula is None or not str(formula).strip():
                        cell_value = "[BLANK]"

                    else:
                        cell_value = str(formula).strip()

                    values.append(
                        f"{coordinate} [{heading}]: {cell_value}"
                    )

                sections.append({
                    "location": (
                        f"Sheet '{formula_sheet.title}', "
                        f"Row {row_number}"
                    ),
                    "text": " | ".join(values),
                })

    finally:
        formula_workbook.close()
        value_workbook.close()

    return sections


def extract_document(file_path):
    """
    Extract a supported vendor document.

    Returns filename, extension and extracted sections.
    Each section includes its original document location.
    """

    path = Path(file_path)

    if not path.is_file():
        raise FileNotFoundError(
            f"Document not found: {path}"
        )

    extension = path.suffix.lower()

    if extension not in SUPPORTED_EXTENSIONS:
        raise ValueError(
            f"Unsupported file type: {extension}"
        )

    if extension == ".pdf":
        sections = extract_pdf(path)

    elif extension == ".docx":
        sections = extract_docx(path)

    else:
        sections = extract_excel(path)

    if not sections:
        raise ValueError(
            f"No readable text was extracted from {path.name}. "
            "The document may be empty, scanned or image-based."
        )

    return {
        "file_name": path.name,
        "extension": extension,
        "sections": sections,
        "section_count": len(sections),
    }
