
import io, re, shutil, sys, tempfile, time, uuid, zipfile
import datetime as dt
from pathlib import Path
import os

BASE = Path(__file__).resolve().parent
BACKEND_DIR = BASE.parent
FRONTEND_DIR = BACKEND_DIR.parent / "Frontend" / "tender_intake"
MAIN_FRONTEND_DIR = BACKEND_DIR.parent / "Frontend" / "Main"
MAIN_IMAGES_DIR = MAIN_FRONTEND_DIR / "images"
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from flask import Flask, jsonify, request, send_file, send_from_directory, abort, render_template
from goods_services import consolidate
from Price_Doc import pricesched
from database import tender_data
from Technical.evalreport import build_eval_workbook, eval_to_json, parse_eval
from Technical import award_recommender
from chatbot_routes import chatbot_bp

app = Flask(
    __name__,
    template_folder=str(FRONTEND_DIR),
    static_folder=str(FRONTEND_DIR),
    static_url_path="",
)
SECRET_KEY = os.environ.get("SECRET_KEY")

if not SECRET_KEY:
    raise RuntimeError("SECRET_KEY environment variable is required.")

app.secret_key = SECRET_KEY

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        os.environ.get("SESSION_COOKIE_SECURE", "false").lower()
        in ("1", "true", "yes")
    ),
)
app.register_blueprint(chatbot_bp)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
WORK = Path(tempfile.gettempdir()) / "store"
WORK.mkdir(exist_ok=True)
FILES = {"gcc": "GCC_Compliance_Comparison.xlsx", "gct": "GCT_Compliance_Comparison.xlsx",
         "scc": "SCC_Compliance_Comparison.xlsx",
         "eval": "Technical_Evaluation_Compliance.xlsx", "price": "Price_Schedules.xlsx",
         "award": "Award_Recommendation.xlsx"}
OUTPUT_PREFIXES = {
    "gcc": "GCC_Compliance_Comparison",
    "gct": "GCT_Compliance_Comparison",
    "scc": "SCC_Compliance_Comparison",
    "eval": "Technical_Evaluation_Compliance",
    "price": "Price_Schedules",
    "award": "Award_Recommendation",
}
PURCHASE_ORDER_EXTENSIONS = {".pdf", ".doc", ".docx", ".xls", ".xlsx", ".xlsm", ".zip"}
PURCHASE_ORDER_DATE_RE = re.compile(
    r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}|"
    r"\d{1,2}\s+(?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
    r"Jul(?:y)?|Aug(?:ust)?|Sep(?:t(?:ember)?)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)\s+\d{4})\b",
    re.IGNORECASE,
)

# Uploaded documents are kept in the tenderdb database, not only in the job folder (which
# is OS temp space and disappears on cleanup/reboot).
_db_ready, _db_retry_at = False, 0.0

def database_ready():
    """True when the uploaded-document table can be reached. An unreachable server is
    retried at most once a minute so a dead database cannot stall every upload."""
    global _db_ready, _db_retry_at
    if _db_ready:
        return True
    now = time.monotonic()
    if now < _db_retry_at:
        return False
    try:
        tender_data.init()
        _db_ready = True
    except Exception:
        _db_retry_at = now + 60
    return _db_ready

def store_upload(job, mode, bidder, path, original_name, content_type):
    """Insert one uploaded document into the database. Raises when the database refuses."""
    with path.open("rb") as fh:
        tender_data.store_document(job, original_name, fh.read(), mode=mode, bidder=bidder or None,
                                   content_type=content_type)


def _purchase_order_text(filename, content, depth=0):
    """Extract searchable text without unpacking uploaded archives to disk."""
    extension = Path(filename).suffix.lower()
    if extension == ".zip" and depth < 2:
        parts = []
        try:
            with zipfile.ZipFile(io.BytesIO(content)) as archive:
                for member in archive.infolist():
                    if member.is_dir() or member.file_size > 25 * 1024 * 1024:
                        continue
                    member_name = Path(member.filename).name
                    if Path(member_name).suffix.lower() in PURCHASE_ORDER_EXTENSIONS:
                        parts.append(_purchase_order_text(member_name, archive.read(member), depth + 1))
        except (OSError, zipfile.BadZipFile):
            return ""
        return "\n".join(parts)
    try:
        if extension == ".pdf":
            from pypdf import PdfReader
            reader = PdfReader(io.BytesIO(content), strict=False)
            text = "\n".join(page.extract_text(extraction_mode="layout") or "" for page in reader.pages)
            if text.strip():
                return text
            # Image-only purchase orders need OCR; keep this optional so text
            # extraction still works where the local OCR engine is unavailable.
            try:
                from pdf2image import convert_from_bytes
                import pytesseract
                tesseract = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
                if tesseract.exists():
                    pytesseract.pytesseract.tesseract_cmd = str(tesseract)
                return "\n".join(pytesseract.image_to_string(page) for page in convert_from_bytes(content))
            except Exception:
                return ""
        if extension in (".docx", ".doc"):
            import docx
            document = docx.Document(io.BytesIO(content))
            values = [paragraph.text for paragraph in document.paragraphs]
            values.extend(" | ".join(cell.text for cell in row.cells) for table in document.tables for row in table.rows)
            return "\n".join(values)
        if extension in (".xlsx", ".xlsm"):
            import openpyxl
            workbook = openpyxl.load_workbook(io.BytesIO(content), read_only=True, data_only=True)
            try:
                return "\n".join(
                    str(cell) for sheet in workbook.worksheets
                    for row in sheet.iter_rows(values_only=True)
                    for cell in (" | ".join(str(value) for value in row if value is not None),)
                    if cell
                )
            finally:
                workbook.close()
        if extension == ".xls":
            import xlrd
            workbook = xlrd.open_workbook(file_contents=content)
            return "\n".join(
                " | ".join(str(cell.value) for cell in sheet.row(row) if cell.value not in (None, ""))
                for sheet in workbook.sheets() for row in range(sheet.nrows)
            )
    except Exception:
        app.logger.exception("Could not extract text from purchase-order document %s", filename)
    return content.decode("utf-8", errors="ignore") if extension not in (".zip",) else ""


def _normalise_purchase_order_date(value):
    value = re.sub(r"\s+", " ", value.strip())
    for pattern in ("%d/%m/%Y", "%d-%m-%Y", "%d/%m/%y", "%d-%m-%y", "%Y-%m-%d",
                    "%d %B %Y", "%d %b %Y"):
        try:
            return dt.datetime.strptime(value, pattern).date().isoformat()
        except ValueError:
            continue
    return value


def analyse_purchase_order(filename, content):
    """Extract purchase-order metadata, including supplier and supply description."""
    # Preserve PDF layout spacing: in purchase-order tables, columns are often
    # represented by whitespace rather than punctuation or actual table cells.
    text = _purchase_order_text(filename, content)
    lines = [line.strip() for line in text.splitlines() if line.strip()]

    tender_no = ""
    # Require an explicit number/reference label so headings such as
    # "Tender for Purchase" are never mistaken for a tender number.
    tender_match = re.search(
        r"(?im)\b(?:tender|rfq|bid)\s*(?:no\.?|number|ref(?:erence)?|id)\s*"
        r"(?:[:#.\-–—]|\s)*"
        r"([A-Za-z0-9][A-Za-z0-9./_-]{2,})", text,
    )
    if tender_match:
        tender_no = tender_match.group(1).strip(".,;:")

    bidder_name = ""
    bidder_label = (
        r"(?:awarded\s+(?:company|supplier|bidder|tenderer)"
        r"(?:\s*/\s*(?:company|supplier|bidder|tenderer))*\s*(?:name)?|"
        r"successful\s+(?:bidder|supplier|tenderer)|"
        r"awarded\s+to|"
        r"(?:company|bidder|supplier|vendor)(?:\s*/\s*(?:company|bidder|supplier|vendor))+\s*(?:name)?|"
        r"(?:bidder|tenderer|supplier|vendor|company|contractor)\s*(?:name)?|"
        r"name\s+of\s+(?:the\s+)?(?:bidder|tenderer|supplier|vendor|company))"
    )
    bidder_match = re.search(
        rf"(?im)^\s*{bidder_label}\s*(?:[:#.\-–—]+\s*(.*\S)|\s*)$", text
    )
    if bidder_match and bidder_match.group(1):
        bidder_name = bidder_match.group(1).strip(" .,:;-–—")
    elif bidder_match:
        label_index = next((i for i, line in enumerate(lines) if re.fullmatch(
            rf"\s*{bidder_label}\s*", line, re.IGNORECASE)), -1)
        if label_index >= 0:
            for candidate in lines[label_index + 1:label_index + 3]:
                candidate = candidate.strip(" .,:;-–—")
                if candidate and not re.search(r"(?i)^(?:name|address|date|no\.?|number)$", candidate):
                    bidder_name = candidate
                    break

    # Award-specific labels take precedence over an earlier generic supplier
    # field, because purchase-order forms often contain both.
    award_label = r"(?:awarded\s+(?:company|supplier|bidder|tenderer)(?:\s*/\s*(?:company|supplier|bidder|tenderer))*\s*(?:name)?|successful\s+(?:bidder|supplier|tenderer)|awarded\s+to)"
    award_match = re.search(rf"(?im)^\s*{award_label}\s*[:#.-]+\s*(.+?\S)\s*$", text)
    if not award_match:
        award_match = re.search(rf"(?im)^\s*{award_label}\s*$\n\s*(.+?\S)\s*$", text)
    if award_match:
        bidder_name = award_match.group(1).strip(" .,:;-")

    # Supplier Details can occur anywhere in the form and may span multiple lines.
    supplier_section = re.search(
        r"(?is)\b(?:supplier\s+details|supplier\s+information|vendor\s+details)\b\s*[:\-]?"
        r"(.*?)(?=\b(?:delivery\s+details|billing\s+details|payment\s+terms|order\s+details|"
        r"goods\s*/\s*services\s+details|description\s+of\s+(?:goods|services|supply))\b|\Z)", text,
    )
    supplier_text = supplier_section.group(1) if supplier_section else text

    def labelled_value(source, labels):
        label_pattern = "|".join(labels)
        for line_index, line in enumerate(source.splitlines()):
            # Works with conventional "Label: value", DOCX/XLSX row separators,
            # and PDF layout output where separate columns are spaced apart.
            match = re.match(
                rf"^\s*(?:{label_pattern})\s*(?::|#|=|\||\s{{2,}})\s*(.*?)\s*$",
                line, re.IGNORECASE,
            )
            value = match.group(1).strip(" .,:;-|") if match else ""
            if value:
                return value
            if match:
                for candidate_line in source.splitlines()[line_index + 1:line_index + 3]:
                    candidate = candidate_line.strip(" .,:;-|")
                    if candidate and not re.match(
                        r"(?i)^(?:address|telephone|phone|email|contact|name|field|supplier information)$",
                        candidate,
                    ):
                        return candidate
        return ""

    supplier_name = labelled_value(supplier_text, [
        r"awarded\s+supplier\s*/\s*company", r"awarded\s+supplier", r"awarded\s+company",
        r"name\s+of\s+(?:the\s+)?supplier", r"supplier\s*(?:/\s*company)?\s*name",
        r"supplier\s+company", r"company\s*name", r"supplier", r"vendor\s*name"
    ])
    if not supplier_name:
        supplier_name = labelled_value(text, [r"name\s+of\s+(?:the\s+)?supplier",
                                              r"supplier\s*(?:/\s*company)?\s*name", r"vendor\s*name"])
    supply_description = labelled_value(text, [
        r"description\s+of\s+(?:goods|services|supply)(?:\s+to\s+be\s+supplied)?",
        r"goods\s*/\s*services\s*(?:description)?",
        r"scope\s+of\s+supply", r"item\s+description", r"supply\s+description", r"particulars"
    ])
    if not supply_description:
        item_section = re.search(
            r"(?is)\bordered\s+items\s*/\s*services\b(.*?)(?=\bfinancial\s+summary\b|\Z)", text
        )
        if item_section:
            descriptions = []
            for line in item_section.group(1).splitlines():
                cells = [cell.strip() for cell in re.split(r"\s{2,}|\s*\|\s*", line.strip()) if cell.strip()]
                if len(cells) >= 2 and re.fullmatch(r"\d+[.)]?", cells[0]):
                    description = cells[1].strip(" .,:;-|")
                    if description and not re.match(r"(?i)^description(?:\s*/.*)?$", description):
                        descriptions.append(description)
            supply_description = "; ".join(descriptions)

    po_date = ""
    po_date_match = re.search(
        rf"(?im)\b(?:purchase\s*order|po)\s*date\s*[:#-]?\s*({PURCHASE_ORDER_DATE_RE.pattern})",
        text,
    )
    if po_date_match:
        po_date = _normalise_purchase_order_date(po_date_match.group(1))
    else:
        for line in lines:
            if re.search(r"(?i)\b(?:purchase\s*order|po)\b", line):
                date_match = PURCHASE_ORDER_DATE_RE.search(line)
                if date_match:
                    po_date = _normalise_purchase_order_date(date_match.group(0))
                    break

    return {"tender_no": tender_no or None, "bidder_name": bidder_name or supplier_name or None,
            "supplier_name": supplier_name or None,
            "supply_description": supply_description or None, "po_date": po_date or None}


def technical_closing_date(evals):
    """Get the tender closing date extracted from a technical evaluation report."""
    for evaluation in evals or []:
        for key, value in (evaluation.get("meta") or {}).items():
            key_text = str(key).casefold()
            if ("closing" in key_text or "closed" in key_text) and "date" in key_text and value:
                return _normalise_purchase_order_date(str(value))
    return None

# Excel summaries kept in the excel_file table (View Evaluation). Order = display order.
SAVED_KINDS = ("price", "eval", "gcc", "scc", "gct", "award")
KIND_LABELS = {
    "price": "Price Schedule summary",
    "eval": "Technical evaluation summary",
    "gcc": "GCC Document summary",
    "scc": "SCC Document summary",
    "gct": "GCT document summary",
    "award": "Award recommendation summary",
}
EXCEL_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

def save_excel_outputs(job, detected):
    """Store every Excel summary this job produced (price, technical, GCC, SCC, GCT) in the
    excel_file table under its tender number. Returns (tender_no, saved names, problems)."""
    numbers = list(dict.fromkeys(
        str(n).strip() for n in (detected or [])
        if str(n).strip() and str(n).strip().casefold() != "not found"))
    tender_no = ", ".join(numbers) or "UNASSIGNED"
    problems, saved = [], []
    out = WORK / job / "out"
    if not out.is_dir():
        return tender_no, saved, problems
    if not database_ready():
        problems.append("Excel files were not saved: the tenderdb database could not be reached.")
        return tender_no, saved, problems
    for kind in SAVED_KINDS:
        for path in sorted(out.glob(f"{OUTPUT_PREFIXES[kind]}*.xlsx")):
            try:
                tender_data.store_excel(tender_no, path.name, path.read_bytes())
                saved.append(path.name)
            except Exception as exc:
                problems.append(f"{path.name}: could not be saved to the database ({exc})")
    return tender_no, saved, problems


def persist_evaluation_scores(tender_numbers, evaluations=None, price_rows=None, weights=None):
    """Save one complete score row per bidder for one unambiguous tender.

    Technical scores are taken from the technical report recommendations.  The
    commercial score is taken from the price parser's normalized price score;
    an explicitly labelled commercial score in an evaluation report is used as
    a fallback when a price schedule did not provide one.  The database layer
    applies the approved rates and calculates the two weighted scores and the
    overall score in the same transaction as the upsert.
    """
    numbers = list(dict.fromkeys(str(value).strip() for value in (tender_numbers or [])
                                 if str(value).strip() and str(value).strip().casefold() != "not found"))
    if len(numbers) != 1:
        return ["Bidder scores were not saved because the uploaded results do not identify exactly one tender number."]
    tender_no = numbers[0]
    if not database_ready():
        return ["Bidder scores were not saved because the tenderdb database could not be reached."]

    records = {}
    for evaluation in evaluations or []:
        for recommendation in evaluation.get("recommendations", []):
            key = award_recommender.bidder_key(recommendation.get("bidder"))
            if not key:
                continue
            record = records.setdefault(key, {"bidder": recommendation.get("bidder")})
            score = recommendation.get("technical_score")
            if score is not None:
                record["technical_score"] = score
            commercial_score = recommendation.get("commercial_score")
            if commercial_score is not None:
                record["commercial_score"] = commercial_score

    price_rows = price_rows or []
    currencies = {str(row.get("currency") or "").upper() for row in price_rows
                  if row.get("price_score") is not None}
    if len(currencies) <= 1:
        for price_row in price_rows:
            score = price_row.get("price_score")
            key = award_recommender.bidder_key(price_row.get("bidder"))
            if score is None or not key:
                continue
            record = records.setdefault(key, {"bidder": price_row.get("bidder")})
            record["commercial_score"] = score

    if weights is None:
        weights = tender_data.score_weightings(tender_no)
    tender_data.store_bidder_score_rows(
        tender_no, list(records.values()),
        (weights or {}).get("technical"), (weights or {}).get("commercial"),
    )
    return []

def kind_of_filename(filename):
    for kind in SAVED_KINDS:
        if filename.startswith(OUTPUT_PREFIXES[kind]):
            return kind
    return None

@app.get("/")
def index():
    return send_from_directory(MAIN_FRONTEND_DIR, "main.html")

@app.get("/main.css")
def main_stylesheet():
    return send_from_directory(MAIN_FRONTEND_DIR, "main.css")

@app.get("/main-images/<path:filename>")
def main_image(filename):
    if filename not in {"logo.png", "background.png"}:
        abort(404)
    if filename == "background.png":
        image_path = MAIN_IMAGES_DIR / filename
        if not image_path.is_file():
            image_path = MAIN_IMAGES_DIR / "background"
        if not image_path.is_file():
            abort(404)
        return send_file(image_path, mimetype="image/png")
    return send_from_directory(MAIN_IMAGES_DIR, filename)

@app.get("/commercial-evaluation")
def commercial_evaluation():
    return send_from_directory(FRONTEND_DIR, "index.html")


@app.get("/commercial-use-cases")
def commercial_use_cases():
    return send_from_directory(MAIN_FRONTEND_DIR, "use_cases.html")


@app.get("/use-cases.css")
def use_cases_stylesheet():
    return send_from_directory(MAIN_FRONTEND_DIR, "use-cases.css")


@app.get("/store-documents")
def store_documents_page():
    return send_from_directory(MAIN_FRONTEND_DIR, "store_documents.html")


@app.get("/store-documents.css")
def store_documents_stylesheet():
    return send_from_directory(MAIN_FRONTEND_DIR, "store-documents.css")


@app.get("/store-documents.js")
def store_documents_script():
    return send_from_directory(MAIN_FRONTEND_DIR, "store-documents.js")


@app.get("/api/purchase-orders")
def stored_purchase_orders():
    if not database_ready():
        return jsonify(error="The tender database could not be reached."), 503
    return jsonify(documents=tender_data.purchase_orders())


@app.post("/api/purchase-orders")
def upload_purchase_orders():
    if not database_ready():
        return jsonify(error="The tender database could not be reached."), 503
    uploaded = request.files.getlist("files")
    if not uploaded:
        return jsonify(error="Please select at least one purchase-order document."), 400

    saved, saved_documents, rejected = [], [], []
    for document in uploaded:
        original_name = Path(document.filename or "").name
        extension = Path(original_name).suffix.lower()
        if not original_name or extension not in PURCHASE_ORDER_EXTENSIONS:
            rejected.append(original_name or "unnamed file")
            continue
        content = document.read()
        if not content:
            rejected.append(f"{original_name} (empty file)")
            continue
        try:
            metadata = analyse_purchase_order(original_name, content)
            tender_data.store_purchase_order(
                tender_no=metadata["tender_no"],
                bidder_name=metadata["bidder_name"],
                supplier_name=metadata["supplier_name"],
                supply_description=metadata["supply_description"],
                po_date=metadata["po_date"],
            )
            saved.append(original_name)
            saved_documents.append(metadata)
        except Exception:
            app.logger.exception("Could not store purchase-order document %s", original_name)
            return jsonify(error=f"Could not store extracted metadata from {original_name} in the purchase_order table."), 503

    if not saved:
        return jsonify(error="No supported purchase-order documents were uploaded.", rejected=rejected), 422
    return jsonify(saved=saved, documents=saved_documents, rejected=rejected,
                   message="Purchase-order documents analyzed; extracted metadata was stored.")


def extract_approval_scores(pdf_bytes):
    """Read the tender number and the two approved weighting percentages from a PDF."""
    from pypdf import PdfReader

    reader = PdfReader(io.BytesIO(pdf_bytes), strict=False)
    text = "\n".join(page.extract_text() or "" for page in reader.pages)
    if not text.strip():
        raise ValueError("No readable text was found in the PDF. Please use a text-based approval PDF.")
    if not re.search(r"approval|tender\s+strategy|evaluation\s+criteria", text, re.I):
        raise ValueError("The PDF does not appear to be an approved tender criteria or strategy document.")
    tender_matches = re.findall(
        r"(?im)^\s*tender\s*(?:number|no\.?)\s*[:\-]?\s*([A-Z0-9][A-Z0-9./_-]*)\b", text
    )
    tender_matches = list(dict.fromkeys(value.strip(".,;:") for value in tender_matches))
    if len(tender_matches) != 1:
        raise ValueError("Could not identify one unambiguous Tender Number in the PDF.")

    lines = text.splitlines()
    section_start = next((i for i, line in enumerate(lines)
                          if re.search(r"overall\s+technical\s+and\s+commercial\s+weightings", line, re.I)), None)
    if section_start is None:
        section_start = next((i for i, line in enumerate(lines)
                              if re.search(r"technical.{0,50}commercial|commercial.{0,50}technical", line, re.I)
                              and re.search(r"weight(?:ing|age)?", line, re.I)), None)
    weighting_lines = []
    source_lines = lines[section_start + 1:] if section_start is not None else lines
    for line in source_lines:
        if section_start is not None and re.search(r"^\s*TOTAL\b|overall\s+score\s+calculation", line, re.I):
            break
        if re.search(r"illustrative\s+example|replace\s+these\s+figures", line, re.I):
            continue
        if section_start is None and re.search(r"weighted\s+(?:technical|commercial)|score\s+calculation", line, re.I):
            continue
        weighting_lines.append(re.sub(r"\s+", " ", line).strip())

    technical, commercial = [], []
    for index, line in enumerate(weighting_lines):
        category = ("technical" if re.search(r"\btechnical(?:\s+evaluation)?\b", line, re.I) else
                    "commercial" if re.search(r"\bcommercial(?:\s+evaluation)?\b", line, re.I) else None)
        if not category:
            continue
        match = re.search(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", line)
        if not match:
            for following in weighting_lines[index + 1:index + 4]:
                if re.search(r"\btechnical(?:\s+evaluation)?\b|\bcommercial(?:\s+evaluation)?\b", following, re.I):
                    break
                match = re.search(r"(?<!\d)(\d+(?:\.\d+)?)\s*%", following)
                if match:
                    break
        if match:
            (technical if category == "technical" else commercial).append(float(match.group(1)))
    if len(technical) != 1 or len(commercial) != 1:
        raise ValueError("Could not identify exactly one technical and one commercial approved weighting in the PDF.")
    if not (0 <= technical[0] <= 100 and 0 <= commercial[0] <= 100):
        raise ValueError("Weightings must each be between 0% and 100%.")
    if abs(technical[0] + commercial[0] - 100) > 0.01:
        raise ValueError("Technical and commercial weightings must total 100%.")
    return tender_matches[0], technical[0], commercial[0]


@app.post("/api/approval-criteria")
def upload_approval_criteria():
    uploaded = request.files.get("file")
    if not uploaded or not uploaded.filename:
        return jsonify(error="Choose an approval criteria PDF to upload."), 400
    if Path(uploaded.filename).suffix.lower() != ".pdf":
        return jsonify(error="Only PDF files are accepted."), 400
    content = uploaded.read()
    if not content.startswith(b"%PDF-"):
        return jsonify(error="The selected file is not a valid PDF."), 400
    try:
        tender_no, technical, commercial = extract_approval_scores(content)
    except ValueError as exc:
        return jsonify(error=str(exc)), 422
    except Exception:
        app.logger.exception("Could not read approval criteria PDF")
        return jsonify(error="The PDF could not be read. Please check the file and try again."), 422
    if not database_ready():
        return jsonify(error="The tender database could not be reached."), 503
    try:
        tender_data.store_scores(tender_no, technical, commercial)
    except Exception:
        app.logger.exception("Could not save approval criteria scores")
        return jsonify(error="The extracted scores could not be saved to the database."), 503
    return jsonify(tender_no=tender_no, technical_weighting=technical,
                   commercial_weighting=commercial, saved=True)

@app.post("/api/process")
def process():
    job = uuid.uuid4().hex[:12]
    inp = WORK / job / "in"
    inp.mkdir(parents=True)
    specs, bad, store_problems = [], [], []
    approval_weights = {}
    approval_documents = []
    names = request.form.getlist("names")
    mode = request.form.get("mode") if request.form.get("mode") in ("all", "bids", "eval", "price") else None
    store = database_ready()
    for i, f in enumerate(request.files.getlist("files")):
        ext = Path(f.filename or "").suffix.lower()
        if ext not in (".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf", ".zip"):
            bad.append(f.filename)
            continue
        d = inp / f"f{i}"
        d.mkdir()
        original_name = Path(f.filename or "").name
        safe_stem = re.sub(r'[\\/:*?"<>|=]+', "-", Path(original_name).stem).strip() or f"document-{i + 1}"
        path = d / ("bundle.zip" if ext == ".zip" else safe_stem + ext)
        f.save(path)
        name = re.sub(r'[\\/:*?"<>|=]+', "-", names[i].strip() if i < len(names) else "")
        if ext == ".zip" or not name:
            name = "" if ext == ".zip" else consolidate.derive_bidder_name(Path(f.filename), None)
        if ext == ".zip":
            name = ""
        specs.append(f"{name}={path}" if name else str(path))
        if store:
            try:
                store_upload(job, mode, name, path, original_name, f.mimetype)
            except Exception as exc:
                store_problems.append(f"{original_name}: could not be stored in the database ({exc})")
        if ext == ".pdf":
            try:
                approval_tender, technical_weight, commercial_weight = extract_approval_scores(path.read_bytes())
            except ValueError as exc:
                if re.search(r"approval|criteria|strategy", original_name, re.I):
                    store_problems.append(f"{original_name}: approval criteria could not be extracted ({exc})")
            except Exception:
                app.logger.exception("Could not inspect uploaded PDF for approval weightings")
            else:
                approval_document = {"tender_no": approval_tender,
                                     "technical_weighting": technical_weight,
                                     "commercial_weighting": commercial_weight,
                                     "saved": False}
                approval_documents.append(approval_document)
                approval_weights[approval_tender] = {"technical": technical_weight,
                                                     "commercial": commercial_weight}
                if store:
                    try:
                        tender_data.store_scores(approval_tender, technical_weight, commercial_weight)
                        approval_document["saved"] = True
                    except Exception as exc:
                        store_problems.append(f"{original_name}: approval weightings could not be stored ({exc})")
                # Approval criteria PDFs are metadata inputs, not evaluation reports or price schedules.
                specs.pop()
                continue
    if not specs and approval_documents:
        if not store:
            return jsonify(error="The approval criteria were read, but the tenderdb database could not be reached.",
                           approval_criteria=approval_documents,
                           problems=store_problems), 503
        if store_problems:
            return jsonify(error="The approval criteria were read, but their weightings could not be saved.",
                           approval_criteria=approval_documents,
                           problems=store_problems), 503
        return jsonify(job=job, approval_criteria=approval_documents,
                       tender_no=approval_documents[0]["tender_no"], saved=[])
    if not specs:
        shutil.rmtree(WORK / job, ignore_errors=True)
        return jsonify(error="No usable files received (.xlsx/.xls/.pdf for GCC, .docx/.doc/.pdf for SCC or evaluation reports, or .zip)."), 400
    extra_problems = [f"{x}: unsupported file type, skipped" for x in bad] + store_problems
    if not store:
        extra_problems.append("documents were not stored: the tenderdb database could not be reached")
    if mode == "eval":
        evals, problems = [], []
        for spec in specs:
            label, raw_path = spec.split("=", 1) if "=" in spec and not Path(spec).exists() else ("", spec)
            path = Path(raw_path)
            try:
                parsed = parse_eval(path)
                evals.append((label or path.name, parsed))
            except Exception as exc:
                problems.append(f"{path.name}: {exc}")
        if not evals:
            return jsonify(error="None of the evaluation reports could be read.", problems=problems), 422
        out = WORK / job / "out"
        out.mkdir(parents=True, exist_ok=True)
        tenders = [ev["meta"].get("Tender No.", "") for _, ev in evals]
        eval_path = out / f"Technical_Evaluation_Compliance{consolidate._tender_suffix(tenders)}.xlsx"
        build_eval_workbook(evals, eval_path)
        tender_no, saved, save_problems = save_excel_outputs(job, tenders)
        eval_json = [eval_to_json(label, parsed) for label, parsed in evals]
        score_problems = persist_evaluation_scores(tenders, evaluations=eval_json,
                                                   weights=approval_weights.get(tenders[0]) if len(tenders) == 1 else None)
        return jsonify(job=job, eval=True, evals=eval_json,
                       problems=problems + extra_problems + save_problems + score_problems,
                       tender_no=tender_no, saved=saved)

    if mode == "price":
        detected_price_numbers = pricesched.tender_numbers_from_inputs(specs)
        price_numbers = list(dict.fromkeys(list(approval_weights.keys()) + detected_price_numbers))
        price_weight = None
        if len(price_numbers) == 1:
            existing_weights = approval_weights.get(price_numbers[0])
            if existing_weights is None and store:
                existing_weights = tender_data.score_weightings(price_numbers[0])
            price_weight = (existing_weights or {}).get("commercial")
        res = pricesched.run(specs, WORK / job / "out",
                             commercial_weight=price_weight if price_weight is not None else 100)
        if not res["file"]:
            return jsonify(error="None of the price schedules could be read.", problems=res["problems"]), 422
        tender_no, saved, save_problems = save_excel_outputs(job, res.get("tenders"))
        score_problems = persist_evaluation_scores(
            res.get("tenders"), price_rows=res.get("rows"),
            weights=approval_weights.get(res["tenders"][0]) if len(res.get("tenders", [])) == 1 else None)
        if res["errors"]:
            return jsonify(error=res["errors"][0], job=job, price=True, rows=res["rows"], errors=res["errors"],
                           problems=res["problems"] + extra_problems + save_problems + score_problems,
                           tender_no=tender_no, saved=saved), 422
        return jsonify(job=job, price=True, rows=res["rows"], notes=res["notes"],
                       problems=res["problems"] + extra_problems + save_problems + score_problems,
                       tender_no=tender_no, saved=saved)

    res = consolidate.run(specs, WORK / job / "out")
    if mode == "all":
        detected_price_numbers = pricesched.tender_numbers_from_inputs(specs)
        tender_hints = list(dict.fromkeys(list(approval_weights.keys()) + detected_price_numbers))
        existing_weightings = (approval_weights.get(tender_hints[0]) if len(tender_hints) == 1 else None)
        if existing_weightings is None and len(tender_hints) == 1 and store:
            existing_weightings = tender_data.score_weightings(tender_hints[0])
        existing_price_weight = (existing_weightings or {}).get("commercial")
        price_res = pricesched.run(
            specs,
            WORK / job / "out",
            closing_date=technical_closing_date(res.get("evals")),
            commercial_weight=existing_price_weight if existing_price_weight is not None else 100,
        )
        award_assessment = None
        award_path = None
        if res.get("evals"):
            scan_dir = WORK / job / "award_scan"
            source_metadata = award_recommender.inspect_input_documents(specs, work_dir=scan_dir)
            award_tender_numbers = list(dict.fromkeys((res.get("tenders", []) or []) + (price_res.get("tenders", []) or [])))
            if len(award_tender_numbers) == 1:
                approved = approval_weights.get(award_tender_numbers[0])
                if approved is None and store:
                    approved = tender_data.score_weightings(award_tender_numbers[0])
                if approved and approved.get("technical") is not None and approved.get("commercial") is not None:
                    source_metadata["weights"] = {"technical": approved["technical"],
                                                   "commercial": approved["commercial"],
                                                   "valid": abs(approved["technical"] + approved["commercial"] - 100) <= 0.01,
                                                   "approved": True, "source": "Approval Criteria PDF"}
            award_assessment = award_recommender.build_award_assessment(
                res.get("evals", []), price_res.get("rows", []), source_metadata=source_metadata
            )
            tender_numbers = list(dict.fromkeys((res.get("tenders", []) or []) + (price_res.get("tenders", []) or [])))
            award_path = WORK / job / "out" / (
                "Award_Recommendation" + consolidate._tender_suffix(tender_numbers) + ".xlsx"
            )
            award_recommender.build_award_workbook(
                award_assessment, award_path, tender=", ".join(tender_numbers)
            )
        source_names = set()
        for spec in specs:
            _, raw_path = spec.split("=", 1) if "=" in spec and not Path(spec).exists() else ("", spec)
            source_names.add(Path(raw_path).name.casefold())
        price_sources = {Path(row["source"]).name.casefold() for row in price_res["rows"]}
        comparison_sources = {summary["document"].casefold() for summary in res["document_summaries"]}
        eval_sources = set()
        eval_labels = {ev["label"].casefold() for ev in res["evals"]}
        for source in source_names:
            if Path(source).stem.casefold() in eval_labels:
                eval_sources.add(source)

        def is_cross_parser_warning(problem):
            source = problem.split(" (", 1)[0].casefold()
            if source in price_sources and "no scc table 1 / table 2 found" in problem.casefold():
                return True
            if source in comparison_sources and "no filled-in price schedule found" in problem.casefold():
                return True
            if source in eval_sources and "no filled-in price schedule found" in problem.casefold():
                return True
            if price_res["file"] and source in source_names and "no scc table 1 / table 2 found" in problem.casefold():
                return True
            if (res["gcc"] or res["gct"] or res["scc"]) and source in source_names and "no filled-in price schedule found" in problem.casefold():
                return True
            if source in source_names and source not in price_sources and "no filled-in price schedule found" in problem.casefold():
                return True
            return False

        score_problems = persist_evaluation_scores(
            list(dict.fromkeys((res.get("tenders", []) or []) + (price_res.get("tenders", []) or []))),
            evaluations=res.get("evals", []), price_rows=price_res.get("rows", []),
            weights=(approval_weights.get(tender_hints[0]) if len(tender_hints) == 1 else None),
        )
        ranked_history_problems = []
        if award_assessment and len(tender_numbers) == 1 and store:
            tender_no_for_history = tender_numbers[0]
            closing_date = technical_closing_date(res.get("evals"))
            if closing_date:
                try:
                    tender_data.set_tender(tender_no_for_history, closing_date=closing_date)
                except Exception as exc:
                    ranked_history_problems.append(
                        f"{tender_no_for_history}: tender closing date could not be stored ({exc})"
                    )
            leaders = []
            if award_assessment.get("decision") == "recommendation_ready":
                leaders = sorted(
                    (bidder for bidder in award_assessment.get("bidders", [])
                     if bidder.get("eligibility_status") == "eligible" and bidder.get("overall_score") is not None),
                    key=lambda bidder: (-float(bidder["overall_score"]),
                                        award_recommender.bidder_key(bidder.get("bidder"))),
                )[:3]
            try:
                _stored_count, history_warnings = tender_data.store_ranked_bidder_history(
                    tender_no_for_history, leaders, closing_date=closing_date,
                )
                ranked_history_problems.extend(history_warnings)
            except Exception as exc:
                app.logger.exception("Could not store ranked bidder history for %s", tender_no_for_history)
                ranked_history_problems.append(
                    f"{tender_no_for_history}: ranked bidders could not be saved to past_award ({exc})"
                )
        elif award_assessment and len(tender_numbers) == 1:
            ranked_history_problems.append(
                "Top-ranked bidder details were not saved because the tender database is unavailable."
            )
        problems = [
            problem for problem in res["problems"] + price_res["problems"]
            if not is_cross_parser_warning(problem)
        ] + price_res["errors"] + extra_problems + score_problems + ranked_history_problems
        if not (res["gcc"] or res["gct"] or res["scc"] or res["evals"] or price_res["file"]):
            return jsonify(error="None of the uploaded documents could be read.", problems=problems), 422
        tender_no, saved, save_problems = save_excel_outputs(
            job, (res.get("tenders") or []) + (price_res.get("tenders") or []))
        return jsonify(
            job=job, all=True, uploaded_count=len(specs), gcc=bool(res["gcc"]), gct=bool(res["gct"]),
            scc=bool(res["scc"]),
            eval=bool(res["evals"]), evals=res["evals"], price=bool(price_res["file"]),
            rows=price_res["rows"], notes=price_res["notes"],
            approval_criteria=approval_documents,
            award=bool(award_path and award_path.exists()), award_assessment=award_assessment,
            gcc_bidders=res["gcc_bidders"], gct_bidders=res["gct_bidders"], scc_bidders=res["scc_bidders"],
            document_summaries=res["document_summaries"], problems=problems + save_problems,
            tender_no=tender_no, saved=saved,
        )
    if not res["gcc"] and not res["gct"] and not res["scc"]:
        return jsonify(error="None of the files could be read.", problems=res["problems"]), 422
    tender_no, saved, save_problems = save_excel_outputs(job, res.get("tenders"))
    return jsonify(job=job, gcc=bool(res["gcc"]), gct=bool(res["gct"]), scc=bool(res["scc"]),
                   eval=False, evals=[], gcc_bidders=res["gcc_bidders"],
                   gct_bidders=res["gct_bidders"], scc_bidders=res["scc_bidders"],
                   document_summaries=res["document_summaries"],
                   problems=res["problems"] + extra_problems + save_problems,
                   tender_no=tender_no, saved=saved)

@app.get("/api/download/<job>/<kind>")
def download(job, kind):
    if kind not in FILES or not re.fullmatch(r"[0-9a-f]{12}", job):
        abort(404)
    output_dir = WORK / job / "out"
    output_files = list(output_dir.glob(f"{OUTPUT_PREFIXES[kind]}*.xlsx"))
    if len(output_files) != 1:
        abort(404)
    p = output_files[0]
    if not p.exists():
        abort(404)
    return send_file(p, as_attachment=True, download_name=p.name)

@app.get("/api/tenders")
def tenders_with_saved_files():
    if not database_ready():
        return jsonify(error="The tenderdb database could not be reached.", tenders=[]), 503
    return jsonify(tenders=tender_data.tenders_with_excel())

@app.get("/api/tender-files")
def tender_saved_files():
    tender_no = (request.args.get("tender_no") or "").strip()
    if not tender_no:
        return jsonify(error="tender_no is required."), 400
    if not database_ready():
        return jsonify(error="The tenderdb database could not be reached."), 503
    latest = {}
    for row in tender_data.excel_files_for_tender(tender_no):   # newest first
        kind = kind_of_filename(row["filename"])
        if kind and kind not in latest:
            latest[kind] = row
    files = []
    for kind in SAVED_KINDS:
        row = latest.get(kind)
        files.append({
            "kind": kind, "label": KIND_LABELS[kind],
            "filename": row["filename"] if row else None,
            "date": row["date"] if row else None, "time": row["time"] if row else None,
            "size_bytes": row["size_bytes"] if row else None,
            "url": f"/api/excel/{row['id']}" if row else None,
        })
    return jsonify(tender_no=tender_no, files=files)

@app.get("/api/excel/<int:file_id>")
def saved_excel(file_id):
    if not database_ready():
        abort(503)
    found = tender_data.excel_file_content(file_id)
    if not found:
        abort(404)
    name, content = found
    return send_file(io.BytesIO(content), as_attachment=True, download_name=name, mimetype=EXCEL_MIME)

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
