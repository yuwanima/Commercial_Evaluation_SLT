import argparse
import json
import os
import subprocess
import sys
import time
import uuid

from flask import Flask, abort, flash, jsonify, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from process import REPORT_TYPES, process_files
from chatbot_routes import chatbot_bp

BASE = os.path.dirname(os.path.abspath(__file__))
JOBS = os.path.join(BASE, "jobs")

FRONTEND_DIR = os.path.abspath(os.path.join(BASE, "..", "..", "Frontend", "tender_intake"))

app = Flask(
    __name__,
    template_folder=FRONTEND_DIR,
    static_folder=FRONTEND_DIR,
)
SECRET_KEY = os.environ.get("SECRET_KEY")

if not SECRET_KEY:
    raise RuntimeError(
        "SECRET_KEY environment variable is required."
    )

app.secret_key = SECRET_KEY

app.config.update(
    SESSION_COOKIE_HTTPONLY=True,
    SESSION_COOKIE_SAMESITE="Lax",
    SESSION_COOKIE_SECURE=(
        os.environ.get("SESSION_COOKIE_SECURE", "false").lower()
        in ("1", "true", "yes")
    ),
)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
OPEN_LOCAL = False
app.register_blueprint(chatbot_bp)


@app.errorhandler(413)
def handle_request_too_large(error):
    return jsonify({
        "error": "The total upload size cannot exceed 50 MB."
    }), 413


def _kind(kind):
    if kind not in REPORT_TYPES:
        abort(404)
    return kind


def _jobs_dir(kind):
    d = os.path.join(JOBS, kind)
    os.makedirs(d, exist_ok=True)
    return d


def _job_dir(kind, job_id):
    d = os.path.join(_jobs_dir(kind), job_id)
    if not job_id.isalnum() or not os.path.isdir(d):
        abort(404)
    return d


def _list_jobs(kind):
    out = []
    base = _jobs_dir(kind)
    for jid in sorted(os.listdir(base), reverse=True):
        meta = os.path.join(base, jid, "meta.json")
        if os.path.exists(meta):
            with open(meta, encoding="utf-8") as f:
                out.append(json.load(f))
    return out


@app.get("/")
def home():
    return redirect(url_for("page", kind=next(iter(REPORT_TYPES))))


@app.get("/<kind>")
def page(kind):
    _kind(kind)
    return render_template("index.html", kind=kind, types=REPORT_TYPES, spec=REPORT_TYPES[kind], jobs=_list_jobs(kind))


@app.post("/<kind>/process")
def process(kind):
    _kind(kind)
    files = [f for f in request.files.getlist("reports") if f.filename and f.filename.lower().endswith(".docx")]
    if not files:
        flash("Please choose at least one report (.docx).")
        return redirect(url_for("page", kind=kind))
    job_id = uuid.uuid4().hex[:12]
    d = os.path.join(_jobs_dir(kind), job_id)
    os.makedirs(os.path.join(d, "in"))
    paths = []
    for f in files:
        p = os.path.join(d, "in", secure_filename(f.filename))
        f.save(p)
        paths.append(p)
    try:
        reports = process_files(kind, paths, os.path.join(d, "tables.xlsx"))
    except Exception as e:  # corrupt / unsupported file
        flash(f"Could not process the file(s): {e}")
        return redirect(url_for("page", kind=kind))
    meta = {"id": job_id, "time": time.strftime("%Y-%m-%d %H:%M"),
            "reports": [{"file": r["source"], "fields": len(r["fields"]), "bidders": len(r["bidders"]),
                         "notes": r["notes"]} for r in reports]}
    with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f)
    flash("Tables generated. Click View to open them in Excel.")
    return redirect(url_for("page", kind=kind))


@app.get("/<kind>/view/<job_id>")
def view(kind, job_id):
    _kind(kind)
    path = os.path.join(_job_dir(kind, job_id), "tables.xlsx")
    if OPEN_LOCAL and request.remote_addr in ("127.0.0.1", "::1"):
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        flash("Opened in Excel. Save in Excel to keep your edits.")
        return redirect(url_for("page", kind=kind))
    return send_file(path, as_attachment=True, download_name=f"{kind.upper()}_tables_{job_id}.xlsx")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--open-local", action="store_true")
    ap.add_argument("--port", type=int, default=5000)
    a = ap.parse_args()
    OPEN_LOCAL = a.open_local
    app.run(host="127.0.0.1", port=a.port)

