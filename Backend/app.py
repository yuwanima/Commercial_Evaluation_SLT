
import json, os, shutil, uuid
from collections import Counter
from datetime import datetime

from flask import Flask, redirect, render_template, request, send_file, url_for, abort, jsonify
from sqlalchemy import func, select
from werkzeug.utils import secure_filename

import checkers
import config
import extract
from config import STORE, SEED
from database import Session, migrate
from models import CheckResult, ReferenceDoc, Tender, Upload

ALLOWED = {".docx", ".doc", ".xlsx", ".xlsm", ".pdf"}

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024


@app.teardown_appcontext
def shutdown_session(_):
    Session.remove()


TENDER_FIELDS = [
    ("tender_no", "Tender No", "text", ""),
    ("title", "Tender title", "text", ""),
    ("envelope", "Envelope type", "select", "single|two"),
    ("open_date", "Technical bid opening date", "date", ""),
    ("bond_validity_days", "Tender validity (days) – default 150", "number", ""),
    ("bond_amount_required", "Required bid bond amount (LKR)", "number", ""),
    ("budget_type", "Budget type", "select", "|CAPEX|OPEX"),
    ("budget_amount", "Confirmed budget amount", "number", ""),
    ("contract_price", "Contract price (for performance bond)", "number", ""),
    ("contract_currency", "Contract currency", "select", "LKR|USD|EUR|GBP"),
    ("loi_date", "Letter of Intent date", "date", ""),
    ("warranty_end", "Warranty end date", "date", ""),
]

REF_FILES = [
    ("tet_template", "Technical Evaluation Report – template", "TECHNICAL_EVALUATION_REPORT.docx"),
    ("sset_template", "SSET Evaluation Report – template (SCMD-RSS-001)", "SSET_EVALUATION_REPORT.docx"),
    ("price_schedule", "Price Schedule – template", "PRICE_SCHEDULE.docx"),
    ("gct_single", "General Conditions of Tender – Single envelope (V3) incl. Bid Bond & Performance Bond forms", "GCT_Single_Envelope_V3.pdf"),
    ("gct_two", "General Conditions of Tender – Two envelope (V2) incl. Bid Bond & Performance Bond forms", "GCT_Two_Envelope_V2.pdf"),
    ("manual_1", "Procurement Manual Part I (Annex II = TET report format)", "Procurement_Manual_Part_I.txt"),
    ("manual_2", "Procurement Manual Part II", "Procurement_Manual_Part_II.txt"),
]


def now():
    return datetime.now().isoformat(timespec="seconds")


# ------------------------------------------------------------------ startup
def init():
    for d in ("uploads", "editable", "reference"):
        os.makedirs(os.path.join(STORE, d), exist_ok=True)
    migrate()                                   # create / upgrade tables (Alembic)
    have = set(Session.scalars(select(ReferenceDoc.role)))
    for role, name, fn in REF_FILES:
        src = os.path.join(SEED, fn)
        if role in have or not os.path.exists(src):
            continue
        dst = os.path.join(STORE, "reference", fn)
        shutil.copy(src, dst)
        text = open(dst, errors="ignore").read() if fn.endswith(".txt") else extract.read_any(dst)["text"]
        Session.add(ReferenceDoc(role=role, name=name, file=fn, text=text.replace("\x00", "")))
    Session.commit()
    Session.remove()


# ------------------------------------------------------------------ reference data
def ref_text(role):
    r = Session.scalar(select(ReferenceDoc).where(ReferenceDoc.role == role))
    return r.text if r else ""


def refs_for(tender):
    gct = ref_text("gct_two" if tender.get("envelope") == "two" else "gct_single")
    # the forms are at the end of the GCT: use the LAST occurrence (first is in the table of contents)
    bi = gct.rfind("TENDER SECURITY FORM (BID BOND)")
    pi = gct.rfind("PERFORMANCE SECURITY FORM (PERFORMANCE BOND)")
    r = {"bid_bond_text": gct[bi:pi] if 0 <= bi < pi else "", "perf_bond_text": gct[pi:] if pi >= 0 else ""}
    for role, key in (("tet_template", "tet_tpl"), ("sset_template", "sset_tpl")):
        row = Session.scalar(select(ReferenceDoc).where(ReferenceDoc.role == role))
        if row:
            r[key] = extract.read_any(os.path.join(STORE, "reference", row.file))
    return r


# ------------------------------------------------------------------ core
def run_upload_checks(up_id):
    up = Session.get(Upload, up_id)
    tender = up.tender.as_dict()
    doc = extract.read_any(os.path.join(STORE, "uploads", up.stored))
    fields = json.loads(up.fields or "{}")
    results = checkers.run_checks(up.doc_type, doc, tender, fields, refs_for(tender))
    up.results.clear()
    stamp = now()
    for r in results:
        up.results.append(CheckResult(rule=r["rule"], title=r["title"], status=r["status"],
                                      detail=r["detail"], source=r["source"], checked=stamp))
    up.overall = checkers.overall(results)
    Session.commit()
    return up.overall


def save_upload(tender_id, doc_type, file, fields):
    ext = os.path.splitext(file.filename)[1].lower()
    if ext not in ALLOWED:
        raise ValueError(f"File type {ext} not allowed. Use: {', '.join(sorted(ALLOWED))}")
    uid = uuid.uuid4().hex[:10]
    stored = f"{uid}_{secure_filename(file.filename)}"
    path = os.path.join(STORE, "uploads", stored)
    file.save(path)
    ed = extract.make_editable(path, os.path.join(STORE, "editable"), uid + "_" + os.path.splitext(secure_filename(file.filename))[0])
    ver = Session.scalar(select(func.count()).select_from(Upload).where(Upload.tender_id == tender_id, Upload.doc_type == doc_type)) + 1
    up = Upload(tender_id=tender_id, doc_type=doc_type, orig_name=file.filename, stored=stored,
                editable=os.path.basename(ed) if ed else None, fields=json.dumps(fields), uploaded=now(), version=ver, overall="")
    Session.add(up)
    Session.commit()
    run_upload_checks(up.id)
    return up.id


# ------------------------------------------------------------------ views
@app.template_filter("pretty")
def pretty(s):
    return checkers.DOC_TYPES.get(s, s)


@app.route("/")
def index():
    return render_template("index.html", tenders=Session.scalars(select(Tender).order_by(Tender.id.desc())).all())


@app.route("/tender/new", methods=["POST"])
def tender_new():
    t = Tender(created=now(), tender_no=request.form["tender_no"].strip(), title=request.form.get("title", "").strip(),
               envelope=request.form.get("envelope", "single"))
    Session.add(t)
    Session.commit()
    return redirect(url_for("tender", tid=t.id))


@app.route("/tender/<int:tid>", methods=["GET", "POST"])
def tender(tid):
    t = Session.get(Tender, tid) or abort(404)
    if request.method == "POST":
        for k, *_ in TENDER_FIELDS:
            setattr(t, k, request.form.get(k, "").strip())
        Session.commit()
        for u in list(t.uploads):                  # tender data changed -> re-run all checks
            run_upload_checks(u.id)
        return redirect(url_for("tender", tid=tid))
    ups = Session.scalars(select(Upload).where(Upload.tender_id == tid).order_by(Upload.doc_type, Upload.version.desc())).all()
    counts = {u.id: Counter(r.status for r in u.results) for u in ups}
    return render_template("tender.html", t=t, fields=TENDER_FIELDS, ups=ups, counts=counts, types=checkers.DOC_TYPES,
                           err=request.args.get("err"))


@app.route("/tender/<int:tid>/upload", methods=["POST"])
def upload(tid):
    Session.get(Tender, tid) or abort(404)
    f = request.files.get("file")
    dt = request.form.get("doc_type")
    if not f or not f.filename or dt not in checkers.DOC_TYPES:
        return redirect(url_for("tender", tid=tid, err="Choose a document type and a file."))
    fields = {k: request.form.get(k, "").strip() for k in ("expiry_date", "received_date", "bond_currency")}
    try:
        uid = save_upload(tid, dt, f, fields)
    except ValueError as e:
        return redirect(url_for("tender", tid=tid, err=str(e)))
    return redirect(url_for("doc", uid=uid))


@app.route("/doc/<int:uid>")
def doc(uid):
    u = Session.get(Upload, uid) or abort(404)
    t = u.tender
    two = t.envelope == "two"
    tpl_role = {"TET": "tet_template", "SSET": "sset_template", "PRICE_SCHEDULE": "price_schedule",
                "BID_BOND": "gct_two" if two else "gct_single", "PERF_BOND": "gct_two" if two else "gct_single"}[u.doc_type]
    tpl = Session.scalar(select(ReferenceDoc).where(ReferenceDoc.role == tpl_role))
    return render_template("doc.html", u=u, t=t, res=u.results, tpl=tpl, fields=json.loads(u.fields or "{}"))


@app.route("/doc/<int:uid>/recheck", methods=["POST"])
def recheck(uid):
    Session.get(Upload, uid) or abort(404)
    run_upload_checks(uid)
    return redirect(url_for("doc", uid=uid))


@app.route("/doc/<int:uid>/fields", methods=["POST"])
def doc_fields(uid):
    u = Session.get(Upload, uid) or abort(404)
    u.fields = json.dumps({k: request.form.get(k, "").strip() for k in ("expiry_date", "received_date", "bond_currency")})
    Session.commit()
    run_upload_checks(uid)
    return redirect(url_for("doc", uid=uid))


@app.route("/download/<kind>/<int:xid>")
def download(kind, xid):
    if kind in ("original", "editable"):
        u = Session.get(Upload, xid) or abort(404)
        if kind == "original":
            return send_file(os.path.join(STORE, "uploads", u.stored), as_attachment=True, download_name=u.orig_name)
        if not u.editable:
            abort(404)
        return send_file(os.path.join(STORE, "editable", u.editable), as_attachment=True)
    r = Session.get(ReferenceDoc, xid) or abort(404)
    path = os.path.join(STORE, "reference", r.file)
    if kind == "ref":
        return send_file(path, as_attachment=True, download_name=r.file)
    if kind == "refeditable":                      # editable Word copy of a reference PDF (e.g. bond forms)
        out = extract.make_editable(path, os.path.join(STORE, "reference_editable"), os.path.splitext(r.file)[0])
        if not out:
            abort(404)
        return send_file(out, as_attachment=True)
    abort(404)


@app.route("/reference")
def reference():
    refs = Session.execute(select(ReferenceDoc.id, ReferenceDoc.role, ReferenceDoc.name, ReferenceDoc.file,
                                  func.length(ReferenceDoc.text).label("n"))).all()
    return render_template("reference.html", refs=refs)


@app.route("/rules")
def rules():
    return render_template("rules.html", rules=checkers.RULES.values())


@app.route("/api/doc/<int:uid>")
def api_doc(uid):
    u = Session.get(Upload, uid) or abort(404)
    return jsonify({"id": uid, "type": u.doc_type, "overall": u.overall,
                    "results": [{"rule": r.rule, "title": r.title, "status": r.status, "detail": r.detail, "source": r.source}
                                for r in u.results]})


init()
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
