# pyrefly: ignore[missing-import]
"""
extractor.py - reads an SSET evaluation report (.docx) and pulls out the fields
needed for the five review tables.

How it works
------------
1. The document is flattened, in reading order, into text lines and table rows.
2. Every line / cell is checked against a list of known LABELS
   ("Tender No", "Compliance to General Conditions", "AMC Status", ...).
3. When a label is found, the text after it (same line, next cell, or the
   following lines) becomes the value.
4. The bidder table (Bidder's Name | Local Agent | ...) is read structurally.

To support new wording, add a pattern to LABELS below. Nothing else changes.
"""
import re
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

# --------------------------------------------------------------------------
# Label dictionary: (key, regex on the lower-cased line start)
# Order matters - more specific patterns first.
# --------------------------------------------------------------------------
_L = [
    ("subject", r"subject"),
    ("tender_no", r"tender\s*(?:no|number|ref(?:erence)?)\b\.?"),
    ("user_department", r"user\s*department"),
    ("forwarded_by", r"forwarded\s*by"),
    ("budget_hdr", r"budget\s*confirmed"),
    ("capex", r"capex"),
    ("opex", r"opex"),
    ("jtb_hdr", r"previous\s*jtb\s*ref(?:erence)?(?:\s*\(if\s*any\))?"),
    ("jtb_ref", r"jtb\s*ref(?:erence)?\s*no\.?"),
    ("jtb_date", r"date"),  # only valid right after a JTB reference
    ("published", r"tender\s*(?:published|floating)\s*date"),
    ("closed", r"tender\s*clos(?:ed|ing)\s*date"),
    ("opened", r"(?:technical\s*)?bid\s*open(?:ed|ing)\s*date"),
    ("submitted", r"sset\s*report\s*submission\s*date"),
    ("bond_expiry", r"bid\s*bond\s*expiry\s*date(?:\s*\(s\))?"),
    ("purpose", r"purpose\s*of\s*(?:the\s*)?tender(?:\s*\(sec\s*2a\))?"),
    ("description", r"short\s*description(?:\s*of\s*(?:the\s*)?(?:project\s*/\s*goods\s*/\s*services|goods\s*/\s*services|goods|services|project))?(?:\s*\(sec\s*2b\))?"),
    ("need", r"how\s*the\s*need\s*(?:arose|arise[sd]?)"),
    ("approval", r"relevant\s*approval(?:\s*to\s*(?:proceed|precede)(?:\s*(?:with\s*)?(?:the\s*)?procurement(?:\s*process)?)?)?"),
    ("project", r"project\s*(?:no\.?|number)(?:\s*(?:&|and)\s*name)?"),
    ("account", r"account\s*code(?:\s*(?:&|and)\s*cost\s*cent(?:er|re))?"),
    ("budget_res", r"budget\s*(?:reserved|available)(?:\s*(?:out\s*of|/)\s*(?:the\s*)?total\s*available)?"),
    ("method", r"method\s*of\s*tendering"),
    ("process", r"procurement\s*process(?:\s*up\s*to\s*bid\s*opening)?"),
    ("changes", r"(?:any\s*)?changes\s*to\s*(?:the\s*)?(?:original\s*)?tender\s*condition(?:s)?"),
    ("c_general", r"compliance\s*to\s*general(?:\s*conditions)?"),
    ("c_technical", r"compliance\s*to\s*technical(?:\s*conditions)?"),
    ("c_special", r"compliance\s*to\s*special(?:\s*conditions)?"),
    ("c_contract", r"compliance\s*to\s*contract(?:\s*conditions)?"),
    ("major", r"major\s*non[\s-]*complian\w*"),
    ("minor", r"(?:minor(?:\s*/\s*acceptable)?|acceptable)\s*non[\s-]*complian\w*"),
    ("impact", r"impact(?:\s*of\s*(?:the\s*)?(?:major\s*)?non[\s-]*complian\w*)?"),
    ("negotiate", r"(?:any\s*)?commercial\s*terms\s*to\s*be\s*negotiated(?:\s*with\s*(?:the\s*)?bidder)?"),
    ("recommended", r"recommended\s*(?:bids?\s*)?for\s*commercial\s*evaluation\??"),
    ("bom", r"evaluated\s*bom(?:\s*provided)?\??"),
    ("breakdown", r"breakdown\s*of\s*services"),
    ("amc", r"amc(?:\s*status)?"),
    ("warranty", r"warranty(?:\s*period)?(?:\s*included)?"),
    ("missing", r"missing\s*prices?(?:\s*for\s*items)?\??"),
    ("prev_prices", r"previous\s*purchase\s*prices?"),
    ("other", r"other\s*commercial\s*items"),
]
LABELS = [(k, re.compile(p + r"(?![a-z])")) for k, p in _L]

# Labels that only count when followed by ':' / '-' (or nothing), so ordinary
# sentences such as "AMC is not required..." are not mistaken for labels.
NEEDS_SEPARATOR = {"amc", "warranty", "impact", "jtb_date", "subject", "major", "minor"}
# Labels that are headers only (value comes from sub-labels)
HEADER_ONLY = {"budget_hdr", "jtb_hdr", "subject"}
# Single-value labels: a following line is taken as the value only if the value is still empty
SHORT = {"tender_no", "user_department", "forwarded_by", "capex", "opex", "jtb_ref",
         "jtb_date", "published", "closed", "opened", "submitted", "bond_expiry",
         "c_general", "c_technical", "c_special", "c_contract", "recommended", "bom",
         "amc", "warranty", "account", "budget_res"}

PREFIX = re.compile(r"^\s*(?:[•\-\*]|\(?(?:\d{1,2}|[a-z]|[ivx]{1,4})[\.\)])\s+", re.I)
SECTION = re.compile(r"^\s*(?:\d\.\s*)?(heading|introduction|background|observations|recommendations)\b\s*:?", re.I)
GUIDANCE = re.compile(r"^(?:\*?\s*)(?:eg\.?|e\.g\.)|essential parts that sset|^field$|^details$|^status\s*/\s*details$|^requirement$", re.I)
NUMBERING_ONLY = re.compile(r"^\(?(?:\d{1,2}|[a-z]|[ivx]{1,4})[\.\)]?$", re.I)
SEP = re.compile(r"^\s*[:\-–—=]")


def _clean(s):
    return re.sub(r"[ \t\u00a0]+", " ", s or "").strip()


def match_label(line, last_key=None, labels=None):
    """Return (key, remainder) if the line starts with a known label."""
    s = PREFIX.sub("", line, count=1).strip()
    low = s.lower()
    if labels is None:
        labels = LABELS
    for key, rx in labels:
        m = rx.match(low)
        if not m:
            continue
        if key == "jtb_date" and last_key not in ("jtb_ref", "jtb_date"):
            continue
        rest = s[m.end():]
        has_sep = bool(SEP.match(rest))
        if key in NEEDS_SEPARATOR and rest.strip() and not has_sep:
            continue
        rest = re.sub(r"^\s*(?:\(s\)|\(if any\))", "", rest)
        rest = re.sub(r"^[\s:\-–—=]+", "", rest).strip()
        return key, rest
    return None, line


def _is_placeholder(v):
    v = v.strip()
    return (not v) or v in {"-", "–", "—", "N/A?"} or (v.startswith("(") and v.endswith(")") and ("/" in v or "if any" in v.lower()))


def is_template_text(line, labels=None):
    """Check if a line represents template guidance, boilerplate or placeholder text."""
    s = _clean(line)
    if not s:
        return True
    if _is_placeholder(s) or GUIDANCE.search(s) or NUMBERING_ONLY.match(s):
        return True
    if s.startswith("(") and s.endswith(")") and ("/" in s or "if any" in s.lower() or "e.g." in s.lower()):
        return True
    return False


# --------------------------------------------------------------------------
def _linearize(doc):
    items = []
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            for ln in Paragraph(el, doc).text.split("\n"):
                items.append(("p", ln))
        elif tag == "tbl":
            t = Table(el, doc)
            items.append(("tbl", t))
            for row in t.rows:
                seen, cells = set(), []
                for c in row.cells:
                    if id(c._tc) in seen:
                        continue
                    seen.add(id(c._tc))
                    cells.append(c.text)
                items.append(("row", cells))
    return items


def _hdr_map(cells):
    """Map bidder-table columns by header text."""
    rx = {
        "name": r"bidder", "agent": r"local\s*agent", "loa": r"letter\s*of\s*authority",
        "brand": r"brand|model", "validity": r"validity", "bond": r"bond",
    }
    out = {}
    for i, c in enumerate(cells):
        t = _clean(c).lower().replace("\n", " ")
        for k, p in rx.items():
            if k not in out and re.search(p, t):
                if k == "bond" and "validity" in t:
                    continue
                out[k] = i
                break
    return out


def extract_report(path):
    doc = Document(path)
    items = _linearize(doc)
    vals, bidders, notes = {}, [], []
    signers = []
    cur = last = None
    section = None
    section_text = {"introduction": [], "background": []}
    skip_rows = None  # None | 'bidders' | 'sign'
    bmap = {}

    def add(key, text):
        text = _clean(text)
        if not text or key in HEADER_ONLY:
            return
        vals.setdefault(key, []).append(text)

    for kind, obj in items:
        if kind == "tbl":
            skip_rows, bmap = None, {}
            hdr = [_clean(c.text).lower() for c in obj.rows[0].cells]
            joined = " ".join(hdr)
            if "bidder" in joined and ("local" in joined or "brand" in joined):
                skip_rows, bmap = "bidders", _hdr_map([c.text for c in obj.rows[0].cells])
                cur = None
            elif "name" in hdr and "designation" in joined:
                skip_rows = "sign"
                bmap = {"name": hdr.index("name"), "des": next(i for i, h in enumerate(hdr) if "designation" in h)}
                cur = None
            continue

        if kind == "row" and skip_rows:
            cells = [_clean(c) for c in obj]
            if skip_rows == "bidders" and cells and any(cells):
                if bmap.get("name") is not None and not re.search(r"bidder", cells[bmap["name"]], re.I) and cells[bmap["name"]]:
                    g = lambda k: cells[bmap[k]] if k in bmap and bmap[k] < len(cells) else ""
                    bidders.append({k: g(k) for k in ("name", "agent", "loa", "brand", "validity", "bond")})
            elif skip_rows == "sign" and cells:
                n = cells[bmap["name"]] if bmap["name"] < len(cells) else ""
                d = cells[bmap["des"]] if bmap["des"] < len(cells) else ""
                if n and n.lower() != "name":
                    signers.append(f"{n} ({d})" if d else n)
            continue

        lines = []
        if kind == "p":
            lines = [obj]
        else:  # table row: process every cell, line by line
            for c in obj:
                lines.extend(c.split("\n"))

        for raw in lines:
            ln = _clean(raw)
            if not ln or NUMBERING_ONLY.match(ln) or GUIDANCE.search(ln):
                continue
            if ln.startswith("(") and ln.endswith(")") and "/" in ln:
                continue
            sm = SECTION.match(ln) if kind == "p" else None
            if sm and len(ln) < 60:
                section, cur = sm.group(1).lower(), None
                continue
            key, rest = match_label(ln, last)
            if key:
                last = key
                cur = None if key in HEADER_ONLY else key
                if rest:
                    add(key, rest)
                continue
            # un-labelled text -> continuation of the current label
            if cur and (cur not in SHORT or cur not in vals):
                add(cur, ln)
            elif section in section_text and kind == "p":
                section_text[section].append(ln)

    # ---------- assemble result ----------
    d = {k: "\n".join(v) for k, v in vals.items()}
    d = {k: v for k, v in d.items() if not _is_placeholder(v)}

    if "forwarded_by" not in d and signers:
        d["forwarded_by"] = "; ".join(signers)
        notes.append("Forwarded by: taken from the SSET signature table.")
    if "description" not in d and d.get("subject"):
        d["description"] = d["subject"]
        notes.append("Short description: taken from the Subject line - please verify.")
    for sec, key, nice in (("introduction", "purpose", "Purpose of Tender"), ("background", "need", "How the need arose")):
        if key not in d and section_text[sec]:
            d[key] = "\n".join(section_text[sec])
            notes.append(f"{nice}: no label found; copied the whole '{sec.title()}' section - please trim.")
    # compliance stated in one sentence for all conditions
    if not any(k in d for k in ("c_general", "c_technical", "c_special", "c_contract")):
        body = " ".join(_clean(x) for kk, x in items if kk == "p")
        if re.search(r"compl(?:ied|iant)\s+fully|fully\s+compl", body, re.I):
            for k in ("c_general", "c_technical", "c_special", "c_contract"):
                d[k] = "Compliant"
            notes.append("Compliance: report states full compliance in one sentence; all four set to 'Compliant' - please verify.")
    # compose JTB / Budget text
    d["_budget"] = "\n".join([f"CAPEX: {d.get('capex', '-')}", f"OPEX: {d.get('opex', '-')}"])
    d["_jtb"] = "\n".join([f"JTB Ref No: {d.get('jtb_ref', '')}".rstrip(), f"Date: {d.get('jtb_date', '')}".rstrip()])

    bidders = [b for b in bidders if any(b.values())]
    return {"fields": d, "bidders": bidders, "notes": notes}
