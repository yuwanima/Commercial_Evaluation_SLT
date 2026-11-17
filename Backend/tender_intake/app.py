
import re, shutil, sys, tempfile, time, uuid
from pathlib import Path
import os

BASE = Path(__file__).resolve().parent
BACKEND_DIR = BASE.parent
FRONTEND_DIR = BACKEND_DIR.parent / "Frontend" / "tender_intake"
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

@app.get("/")
def index():
    return send_from_directory(FRONTEND_DIR, "main.html")

@app.get("/commercial-evaluation")
def commercial_evaluation():
    return send_from_directory(FRONTEND_DIR, "index.html")

@app.post("/api/process")
def process():
    job = uuid.uuid4().hex[:12]
    inp = WORK / job / "in"
    inp.mkdir(parents=True)
    specs, bad, store_problems = [], [], []
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
        return jsonify(job=job, eval=True, evals=[eval_to_json(label, parsed) for label, parsed in evals],
                       problems=problems + extra_problems)

    if mode == "price":
        res = pricesched.run(specs, WORK / job / "out")
        if not res["file"]:
            return jsonify(error="None of the price schedules could be read.", problems=res["problems"]), 422
        if res["errors"]:
            return jsonify(error=res["errors"][0], job=job, price=True, rows=res["rows"], errors=res["errors"],
                           problems=res["problems"] + extra_problems), 422
        return jsonify(job=job, price=True, rows=res["rows"], notes=res["notes"],
                       problems=res["problems"] + extra_problems)

    res = consolidate.run(specs, WORK / job / "out")
    if mode == "all":
        price_res = pricesched.run(
            specs,
            WORK / job / "out",
            closing_date=technical_closing_date(res.get("evals")),
        )
        award_assessment = None
        award_path = None
        if res.get("evals"):
            scan_dir = WORK / job / "award_scan"
            source_metadata = award_recommender.inspect_input_documents(specs, work_dir=scan_dir)
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

        problems = [
            problem for problem in res["problems"] + price_res["problems"]
            if not is_cross_parser_warning(problem)
        ] + price_res["errors"] + extra_problems
        if not (res["gcc"] or res["gct"] or res["scc"] or res["evals"] or price_res["file"]):
            return jsonify(error="None of the uploaded documents could be read.", problems=problems), 422
        return jsonify(
            job=job, all=True, uploaded_count=len(specs), gcc=bool(res["gcc"]), gct=bool(res["gct"]),
            scc=bool(res["scc"]),
            eval=bool(res["evals"]), evals=res["evals"], price=bool(price_res["file"]),
            rows=price_res["rows"], notes=price_res["notes"],
            award=bool(award_path and award_path.exists()), award_assessment=award_assessment,
            gcc_bidders=res["gcc_bidders"], gct_bidders=res["gct_bidders"], scc_bidders=res["scc_bidders"],
            document_summaries=res["document_summaries"], problems=problems,
        )
    if not res["gcc"] and not res["gct"] and not res["scc"]:
        return jsonify(error="None of the files could be read.", problems=res["problems"]), 422
    return jsonify(job=job, gcc=bool(res["gcc"]), gct=bool(res["gct"]), scc=bool(res["scc"]),
                   eval=False, evals=[], gcc_bidders=res["gcc_bidders"],
                   gct_bidders=res["gct_bidders"], scc_bidders=res["scc_bidders"],
                   document_summaries=res["document_summaries"],
                   problems=res["problems"] + extra_problems)

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

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
