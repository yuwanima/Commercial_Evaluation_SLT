
import re, shutil, sys, tempfile, uuid
from pathlib import Path

BASE = Path(__file__).resolve().parent
BACKEND_DIR = BASE.parent
FRONTEND_DIR = BACKEND_DIR.parent / "Frontend" / "tender_intake"
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from flask import Flask, jsonify, request, send_file, send_from_directory, abort

from goods_services import consolidate
from Price_Doc import pricesched
from Technical.evalreport import build_eval_workbook, eval_to_json, parse_eval

app = Flask(__name__, static_folder=str(FRONTEND_DIR), static_url_path="")
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
WORK = Path(tempfile.gettempdir()) / "store"
WORK.mkdir(exist_ok=True)
FILES = {"gcc": "GCC_Compliance_Comparison.xlsx", "scc": "SCC_Compliance_Comparison.xlsx",
         "eval": "Technical_Evaluation_Compliance.xlsx", "price": "Price_Schedules.xlsx"}

@app.get("/")
def index():
    return send_from_directory(FRONTEND_DIR, "index.html")

@app.post("/api/process")
def process():
    job = uuid.uuid4().hex[:12]
    inp = WORK / job / "in"
    inp.mkdir(parents=True)
    specs, bad = [], []
    names = request.form.getlist("names")
    for i, f in enumerate(request.files.getlist("files")):
        ext = Path(f.filename or "").suffix.lower()
        if ext not in (".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf", ".zip"):
            bad.append(f.filename)
            continue
        d = inp / f"f{i}"
        d.mkdir()
        path = d / ("GCC" + ext if ext in (".xlsx", ".xlsm", ".xls") else "bundle.zip" if ext == ".zip" else "SCC" + ext)
        f.save(path)
        name = re.sub(r'[\\/:*?"<>|=]+', "-", names[i].strip() if i < len(names) else "")
        if ext == ".zip" or not name:
            name = "" if ext == ".zip" else consolidate.derive_bidder_name(Path(f.filename), None)
        if ext == ".zip":
            name = ""
        specs.append(f"{name}={path}" if name else str(path))
    if not specs:
        shutil.rmtree(WORK / job, ignore_errors=True)
        return jsonify(error="No usable files received (.xlsx/.xls/.pdf for GCC, .docx/.doc/.pdf for SCC or evaluation reports, or .zip)."), 400
    mode = request.form.get("mode") if request.form.get("mode") in ("bids", "eval", "price") else None
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
        build_eval_workbook(evals, out / FILES["eval"])
        return jsonify(job=job, eval=True, evals=[eval_to_json(label, parsed) for label, parsed in evals],
                       problems=problems + [f"{x}: unsupported file type, skipped" for x in bad])

    if mode == "price":
        res = pricesched.run(specs, WORK / job / "out")
        if not res["file"]:
            return jsonify(error="None of the price schedules could be read.", problems=res["problems"]), 422
        return jsonify(job=job, price=True, rows=res["rows"], notes=res["notes"],
                       problems=res["problems"] + [f"{x}: unsupported file type, skipped" for x in bad])

    res = consolidate.run(specs, WORK / job / "out")
    if not res["gcc"] and not res["scc"]:
        return jsonify(error="None of the files could be read.", problems=res["problems"]), 422
    return jsonify(job=job, gcc=bool(res["gcc"]), scc=bool(res["scc"]), eval=False, evals=[],
                   gcc_bidders=res["gcc_bidders"], scc_bidders=res["scc_bidders"],
                   problems=res["problems"] + [f"{x}: unsupported file type, skipped" for x in bad])

@app.get("/api/download/<job>/<kind>")
def download(job, kind):
    if kind not in FILES or not re.fullmatch(r"[0-9a-f]{12}", job):
        abort(404)
    p = WORK / job / "out" / FILES[kind]
    if not p.exists():
        abort(404)
    return send_file(p, as_attachment=True, download_name=FILES[kind])

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=False)
