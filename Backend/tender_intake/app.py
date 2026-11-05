
import os, re, shutil, sys, tempfile, uuid
from pathlib import Path

# Ensure Backend directory is on sys.path so goods_services can be imported
BASE = Path(__file__).resolve().parent
BACKEND_DIR = BASE.parent
if str(BACKEND_DIR) not in sys.path:
    sys.path.insert(0, str(BACKEND_DIR))

from flask import Flask, jsonify, request, send_file, send_from_directory, abort
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
    kinds = request.form.getlist("kinds")
    for i, f in enumerate(request.files.getlist("files")):
        orig_name = f.filename or ""
        ext = Path(orig_name).suffix.lower()
        if ext not in (".xlsx", ".xlsm", ".docx", ".doc", ".pdf", ".zip"):
            bad.append(orig_name)
            continue
        d = inp / f"f{i}"
        d.mkdir()

        k = kinds[i].strip() if i < len(kinds) else ""
        raw_name = Path(orig_name).name
        clean_fname = re.sub(r'[\\/:*?"<>|]+', "_", raw_name) or f"upload_{i}{ext}"
        if k in ("GCC", "SCC") and not clean_fname.lower().startswith(k.lower()):
            clean_fname = f"{k}_{clean_fname}"
        path = d / clean_fname
        f.save(path)

        name = re.sub(r'[\\/:*?"<>|=]+', "-", names[i].strip() if i < len(names) else "")
        if ext == ".zip":
            name = ""
        elif not name:
            parts = [p for p in re.split(r"[\\/]", orig_name) if p]
            if len(parts) > 1:
                name = parts[0]
            else:
                name = consolidate.derive_bidder_name(path, None)
        specs.append(f"{name}={path}" if name else str(path))

    if not specs:
        shutil.rmtree(WORK / job, ignore_errors=True)
        return jsonify(error="No usable files received (.xlsx for GCC; .docx/.doc/.pdf for SCC; or .zip)."), 400

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
