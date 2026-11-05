#!/usr/bin/env python3
"""Web front-end for the GCC / SCC bidder compliance consolidator.

    pip install flask openpyxl python-docx
    python app.py            ->  open http://127.0.0.1:5000
"""
import re, shutil, sys, tempfile, uuid
from pathlib import Path
from flask import Flask, jsonify, request, send_file, send_from_directory, abort

BASE = Path(__file__).resolve().parent
BACKEND_DIR = BASE.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))



from goods_services import consolidate

FRONTEND_DIR = BASE.parent.parent / "Frontend" / "tender_intake"

app = Flask(__name__, static_folder=str(FRONTEND_DIR))
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024
WORK = Path(tempfile.gettempdir()) / "bid_compliance_jobs"
WORK.mkdir(exist_ok=True)
FILES = {"gcc": "GCC_Compliance_Comparison.xlsx", "scc": "SCC_Compliance_Comparison.xlsx"}

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
        if ext not in (".xlsx", ".xlsm", ".docx", ".zip"):
            bad.append(f.filename)
            continue
        d = inp / f"f{i}"
        d.mkdir()
        path = d / ("GCC" + ext if ext in (".xlsx", ".xlsm") else "SCC" + ext if ext == ".docx" else "bundle.zip")
        f.save(path)
        name = re.sub(r'[\\/:*?"<>|=]+', "-", names[i].strip() if i < len(names) else "")
        if ext == ".zip" or not name:
            name = "" if ext == ".zip" else consolidate.derive_bidder_name(Path(f.filename), None)
        specs.append(f"{name}={path}" if name else str(path))
    if not specs:
        shutil.rmtree(WORK / job, ignore_errors=True)
        return jsonify(error="No usable files received (.xlsx for GCC, .docx for SCC, or .zip)."), 400
    res = consolidate.run(specs, WORK / job / "out")
    if not res["gcc"] and not res["scc"]:
        return jsonify(error="None of the files could be read.", problems=res["problems"]), 422
    return jsonify(job=job, gcc=bool(res["gcc"]), scc=bool(res["scc"]),
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
