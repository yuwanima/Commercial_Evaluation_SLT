
import io, re, shutil, sys, tempfile, time, uuid
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


def technical_closing_date(evals):
    """Get the tender closing date extracted from a technical evaluation report."""
    for evaluation in evals or []:
        for key, value in (evaluation.get("meta") or {}).items():
            key_text = str(key).casefold()
            if ("closing" in key_text or "closed" in key_text) and "date" in key_text and value:
                return value
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
            except ValueError:
                pass
            except Exception:
                app.logger.exception("Could not inspect uploaded PDF for approval weightings")
            else:
                approval_documents.append({"tender_no": approval_tender,
                                           "technical_weighting": technical_weight,
                                           "commercial_weighting": commercial_weight})
                approval_weights[approval_tender] = {"technical": technical_weight,
                                                     "commercial": commercial_weight}
                if store:
                    try:
                        tender_data.store_scores(approval_tender, technical_weight, commercial_weight)
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
        problems = [
            problem for problem in res["problems"] + price_res["problems"]
            if not is_cross_parser_warning(problem)
        ] + price_res["errors"] + extra_problems + score_problems
        if not (res["gcc"] or res["gct"] or res["scc"] or res["evals"] or price_res["file"]):
            return jsonify(error="None of the uploaded documents could be read.", problems=problems), 422
        tender_no, saved, save_problems = save_excel_outputs(
            job, (res.get("tenders") or []) + (price_res.get("tenders") or []))
        return jsonify(
            job=job, all=True, uploaded_count=len(specs), gcc=bool(res["gcc"]), gct=bool(res["gct"]),
            scc=bool(res["scc"]),
            eval=bool(res["evals"]), evals=res["evals"], price=bool(price_res["file"]),
            rows=price_res["rows"], notes=price_res["notes"],
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
