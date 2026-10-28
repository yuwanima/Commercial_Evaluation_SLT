
import argparse
import json
import os
import subprocess
import sys
import time
import uuid

from flask import Flask, abort, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

from process import process_files

BASE = os.path.dirname(os.path.abspath(__file__))
FRONTEND_DIR = os.path.abspath(os.path.join(BASE, "..", "..", "Frontend", "tender_intake"))
JOBS = os.path.join(BASE, "jobs")
os.makedirs(JOBS, exist_ok=True)

app = Flask(
    __name__,
    template_folder=FRONTEND_DIR,
    static_folder=FRONTEND_DIR,
    static_url_path="",
)
app.secret_key = os.environ.get("SECRET_KEY", "change-me")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
OPEN_LOCAL = False


def _job_dir(job_id):
    if not job_id.isalnum():
        abort(404)
    d = os.path.join(JOBS, job_id)
    if not os.path.isdir(d):
        abort(404)
    return d


def _list_jobs():
    if not os.path.isdir(JOBS):
        os.makedirs(JOBS, exist_ok=True)
        return []
    out = []
    for jid in sorted(os.listdir(JOBS), reverse=True):
        meta = os.path.join(JOBS, jid, "meta.json")
        if os.path.exists(meta):
            with open(meta, encoding="utf-8") as f:
                out.append(json.load(f))
    return out


@app.get("/")
def index():
    return render_template("index.html", jobs=_list_jobs())


@app.post("/process")
def process():
    files = [f for f in request.files.getlist("reports") if f.filename]
    files = [f for f in files if f.filename.lower().endswith(".docx")]
    if not files:
        flash("Please choose at least one SSET evaluation report (.docx).")
        return redirect(url_for("index"))
    job_id = uuid.uuid4().hex[:12]
    d = os.path.join(JOBS, job_id)
    os.makedirs(os.path.join(d, "in"))
    paths = []
    for f in files:
        p = os.path.join(d, "in", secure_filename(f.filename))
        f.save(p)
        paths.append(p)
    xlsx = os.path.join(d, "SSET_tables.xlsx")
    try:
        reports = process_files(paths, xlsx)
    except Exception as e:  # corrupt / unsupported file
        flash(f"Could not process the file(s): {e}")
        return redirect(url_for("index"))
    meta = {"id": job_id, "time": time.strftime("%Y-%m-%d %H:%M"),
            "reports": [{"file": r["source"], "fields": len(r["fields"]), "bidders": len(r["bidders"]),
                         "notes": r["notes"]} for r in reports]}
    with open(os.path.join(d, "meta.json"), "w", encoding="utf-8") as f:
        json.dump(meta, f)
    flash("Tables generated. Click View to open them in Excel.")
    return redirect(url_for("index"))


@app.get("/view/<job_id>")
def view(job_id):
    path = os.path.join(_job_dir(job_id), "SSET_tables.xlsx")
    if OPEN_LOCAL and request.remote_addr in ("127.0.0.1", "::1"):
        # opens in the user's own Excel; saving in Excel updates this same file
        if sys.platform.startswith("win"):
            os.startfile(path)  # noqa
        elif sys.platform == "darwin":
            subprocess.Popen(["open", path])
        else:
            subprocess.Popen(["xdg-open", path])
        flash("Opened in Excel. Save in Excel to keep your edits.")
        return redirect(url_for("index"))
    return send_file(path, as_attachment=True, download_name=f"SSET_tables_{job_id}.xlsx")


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--open-local", action="store_true")
    ap.add_argument("--port", type=int, default=5000)
    a = ap.parse_args()
    OPEN_LOCAL = a.open_local
    app.run(host="127.0.0.1", port=a.port)
