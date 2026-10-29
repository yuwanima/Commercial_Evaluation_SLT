import re
from docx import Document
from extractor import (_L, GUIDANCE, NUMBERING_ONLY, SECTION, SHORT, _clean, _hdr_map,
                       _is_placeholder, _linearize, is_template_text, match_label)

import extractor as _base

# ---- TET-specific label patterns (tried BEFORE the shared ones) ----------
_EXTRA = [
    ("status", r"qualified\s*/\s*disqualified"),
    ("rec_fin", r"recommended\s*(?:bids?\s*)?for\s*(?:financial\s*bid\s*opening|commercial\s*evaluation|opening\s*of\s*financial\s*bids)\??"),
    ("bom", r"evaluated\s*bom(?:\s*(?:provided|included))?\??"),
    ("missing", r"missing\s*price(?:s)?(?:\s*items)?(?:\s*needing\s*(?:a\s*)?(?:separate\s*)?sealed\s*envelope)?\??"),
    ("ref_price", r"reference\s*price(?:\s*of\s*similar\s*items)?"),
    ("suitable", r"more\s*suitable"),
]
_OVERRIDE = {k for k, _ in _EXTRA}
TET_LABELS = ([(k, re.compile(p + r"(?![a-z])")) for k, p in _EXTRA] +
              [(k, re.compile(p + r"(?![a-z])")) for k, p in _L if k not in _OVERRIDE])

BIDDER_KEYS = ["c_general", "c_technical", "c_special", "c_contract", "major", "minor", "impact",
               "status", "negotiate", "rec_fin", "bom", "missing", "prev_prices", "ref_price", "suitable"]
BIDDER_SET = set(BIDDER_KEYS)
COMPLIANCE = ["c_general", "c_technical", "c_special", "c_contract"]
NEEDS_SEP = _base.NEEDS_SEPARATOR | {"status", "suitable"}
_base.NEEDS_SEPARATOR = NEEDS_SEP  # shared matcher reads this set

# Template instruction text that must never be mistaken for report content
NOISE = re.compile(
    r"essential parts|believe no special|if (?:sset|tet) observed|each recommended bid shall|"
    r"in case of standardi[sz]ed|selection of supplier|for renewal of service|if it an item|"
    r"project no\. and name \(if applicable\) or if|any other items that should|"
    r"if the bidders? complied for all|qualifying and disqualifying|on the other hand, some|"
    r"no qualified bid can|following points should be clearly|should provide following|"
    r"open tender; from suppliers|event and date taken|any changes to original tender|"
    r"^\(\s*-|^i confirm that|in the event of me|list of bidders recommended for opening|"
    r"^(?:i\.|ii\.)\s*$", re.I)
STOP = re.compile(r"^\s*(?:6\.\s*)?self[\s-]*declaration|^\s*(?:7\.\s*)?submitted\s*by", re.I)
BIDDER_HEAD = re.compile(r"^\s*(?:bidder|tenderer|supplier)\s*(?:no\.?\s*)?(\d{1,2})\b\s*[:\-–—\).]?\s*(.*)$", re.I)


NICE = {"c_general": "Compliance with General Conditions", "c_technical": "Compliance with Technical Conditions",
        "c_special": "Compliance with Special Conditions", "c_contract": "Compliance with Contract Conditions",
        "major": "Major Non-compliances", "minor": "Minor/Acceptable Non-compliances", "impact": "Impact of Non-compliances",
        "status": "Qualified / Disqualified", "negotiate": "Commercial terms to be negotiated",
        "rec_fin": "Recommended for Financial Bid Opening", "bom": "Evaluated BOM included?",
        "missing": "Missing price items", "prev_prices": "Previous purchase price",
        "ref_price": "Reference price of similar items", "suitable": "More Suitable"}


def _norm(s):
    return re.sub(r"[^a-z0-9]+", " ", (s or "").lower()).strip()


def _strip_num(name):
    name = _clean(name)
    if NUMBERING_ONLY.match(name):
        return ""
    return re.sub(r"^\(?\d{1,2}[\.\)]?\s+(?=[A-Za-z])", "", name)


def _norm_cmp(v):
    t = v.strip().lower()
    if re.match(r"^non[\s-]*compliant", t):
        return "Non-compliant"
    if re.match(r"^(?:compliant|complied|fully compliant)", t):
        return "Compliant"
    if t.startswith("disqualif"):
        return "Disqualified"
    if t.startswith("qualified"):
        return "Qualified"
    if t in ("yes", "no"):
        return t.capitalize()
    return v


def _find_bidder_prefix(line, bidders):
    """If line starts with a known bidder name (or 'Bidder N') return (index, remainder)."""
    ln = _clean(line)
    m = BIDDER_HEAD.match(ln)
    if m:
        i = int(m.group(1)) - 1
        if 0 <= i < len(bidders):
            rest = m.group(2)
            nm = _norm(bidders[i].get("name", ""))
            if nm and _norm(rest).startswith(nm):      # "Bidder 1: ABC Power (Pvt) Ltd" -> name is part of the heading
                words = nm.split()
                rest = re.sub(r"^\W*" + r"\W+".join(re.escape(w) for w in words) + r"\W*", "", rest, count=1, flags=re.I)
                rest = re.sub(r"^[\s:\-–—=\)]+", "", rest)
            return i, rest
    nl = _norm(ln)
    best = None
    for i, b in enumerate(bidders):
        nm = _norm(b.get("name", ""))
        if len(nm) >= 3 and nl.startswith(nm):
            if best is None or len(nm) > len(_norm(bidders[best].get("name", ""))):
                best = i
    if best is not None:
        nm_words = _norm(bidders[best]["name"]).split()
        # remove the name from the raw line, tolerant of punctuation
        pat = r"^\W*" + r"\W+".join(re.escape(w) for w in nm_words) + r"\W*"
        rest = re.sub(pat, "", ln, count=1, flags=re.I)
        rest = re.sub(r"^[\s:\-–—=\)]+", "", rest)
        return best, rest
    return None, ln


def _read_bidder_tables(items):
    """Pass 0: the 'Bidder's Name | Local Agent | ...' observation table."""
    bidders, bmap = [], {}
    in_tbl = False
    for kind, obj in items:
        if kind == "tbl":
            hdr = " ".join(_clean(c.text).lower() for c in obj.rows[0].cells)
            in_tbl = "bidder" in hdr and ("local" in hdr or "brand" in hdr)
            if in_tbl:
                bmap = _hdr_map([c.text for c in obj.rows[0].cells])
            continue
        if kind == "row" and in_tbl and "name" in bmap:
            cells = [_clean(c) for c in obj]
            if not any(cells) or re.search(r"bidder", cells[bmap["name"]] if bmap["name"] < len(cells) else "", re.I):
                continue
            g = lambda k: cells[bmap[k]] if k in bmap and bmap[k] < len(cells) else ""
            others = [g(k) for k in ("agent", "loa", "brand", "validity", "bond")]
            name = _strip_num(g("name"))
            if name or any(others):
                bidders.append({"name": name, "agent": g("agent"), "loa": g("loa"), "brand": g("brand"),
                                "validity": g("validity"), "bond": g("bond")})
    return bidders


def extract_tet(path):
    doc = Document(path)
    items = _linearize(doc)
    bidders = _read_bidder_tables(items)
    for b in bidders:
        for k in BIDDER_KEYS:
            b[k] = ""
    vals, shared, notes, signers = {}, {}, [], []
    cur = last = None
    section, stopped = None, False
    section_text = {"introduction": [], "background": []}
    matrix = None        # {col_index: bidder_index}
    in_sign = False
    sign_map = {}

    def add_tender(key, text):
        text = _clean(text).strip(" ;")
        if text and key not in _base.HEADER_ONLY:
            vals.setdefault(key, []).append(text)

    def add_bidder(i, key, text):
        text = _clean(text).strip(" ;")
        if not text or _is_placeholder(text):
            return
        b = bidders[i]
        b[key] = (b[key] + "\n" + text).strip() if b[key] else text

    def assign(key, text, ctx):
        """Put text for a per-bidder key on the right bidder (or the shared pool)."""
        text = _clean(text)
        if not text:
            return
        i, rest = _find_bidder_prefix(text, bidders)
        if i is not None:
            add_bidder(i, key, rest)
        elif ctx is not None:
            add_bidder(ctx, key, text)
        else:
            t = text.strip(" ;")
            if t and not _is_placeholder(t):
                shared.setdefault(key, []).append(t)

    ctx = None  # bidder currently "in focus" (heading style)
    for kind, obj in items:
        if kind == "tbl":
            matrix, in_sign = None, False
            first = [_clean(c.text) for c in obj.rows[0].cells]
            low = [x.lower() for x in first]
            joined = " ".join(low)
            if "name" in low and "designation" in joined:
                in_sign = True
                sign_map = {"name": low.index("name"), "des": next(i for i, h in enumerate(low) if "designation" in h)}
                continue
            # style A: header row of bidder columns
            cols, seq = {}, 0
            seen = set()
            for ci, txt in enumerate(first):
                if ci in seen:
                    continue
                seen.add(ci)
                if ci == 0 or not txt:
                    continue
                j, _ = _find_bidder_prefix(txt, bidders)
                if j is None and re.match(r"^\s*bidder\s*\d*\s*$", txt, re.I):
                    j = seq
                    seq += 1
                if j is not None and j < len(bidders):
                    cols[ci] = j
            if len(cols) >= 1 and bidders and not ("local" in joined and "agent" in joined):
                matrix = cols
            ctx = None
            continue

        if kind == "row" and in_sign:
            cells = [_clean(c) for c in obj]
            n = cells[sign_map["name"]] if sign_map["name"] < len(cells) else ""
            d = cells[sign_map["des"]] if sign_map["des"] < len(cells) else ""
            if n and n.lower() != "name":
                signers.append(f"{n} ({d})" if d else n)
            continue

        if kind == "row" and matrix is not None:
            cells = [_clean(c) for c in obj]
            key, _ = match_label(cells[0], last, TET_LABELS) if cells and cells[0] else (None, "")
            if key == "recommended":
                key = "rec_fin"
            if key in BIDDER_SET:
                for ci, bi in matrix.items():
                    if ci < len(cells):
                        add_bidder(bi, key, cells[ci])
                last = key
                continue
            # a row that is not a bidder criterion: fall through to normal handling

        if kind == "row" and "bidder" in " ".join(_clean(c).lower() for c in obj) and \
                any(re.search(r"local\s*agent|letter\s*of", _clean(c), re.I) for c in obj):
            continue  # header row of the observation table
        # bidder observation rows are consumed in pass 0
        if kind == "row" and bidders and obj and _strip_num(obj[0]) and \
                any(_norm(_strip_num(obj[0])) == _norm(b["name"]) for b in bidders) and len(obj) >= 5:
            continue

        lines = [obj] if kind == "p" else [ln for c in obj for ln in c.split("\n")]
        for raw in lines:
            ln = _clean(raw)
            if not ln or NUMBERING_ONLY.match(ln) or GUIDANCE.search(ln) or is_template_text(ln, TET_LABELS) or NOISE.search(ln):
                continue
            if STOP.match(ln):
                stopped = True
            if stopped:
                continue
            if ln.startswith("(") and ln.endswith(")") and "/" in ln:
                continue
            sm = SECTION.match(ln) if kind == "p" else None
            if sm and len(ln) < 60:
                section, cur, ctx = sm.group(1).lower(), None, None
                continue

            # B: "Bidder 2: XYZ" / "XYZ Energy" heading (a bidder name not followed by a label)
            key, rest = match_label(ln, last, TET_LABELS)
            if key is None and bidders:
                i, rem = _find_bidder_prefix(ln, bidders)
                if i is not None:
                    k2, r2 = match_label(rem, last, TET_LABELS) if rem else (None, "")
                    if k2 in BIDDER_SET or k2 == "recommended":
                        ctx, ln, key, rest = i, rem, k2, r2        # "ABC Power: Compliance..: Compliant"
                    elif cur in BIDDER_SET and rem:
                        add_bidder(i, cur, rem)                      # C: list style under a label
                        continue
                    elif not rem:
                        ctx, cur = i, None                           # pure heading
                        continue
                    else:
                        ctx = i
                        continue
            if key == "recommended":
                key = "rec_fin"
            if key:
                last = key
                if key in BIDDER_SET:
                    cur = key
                    if rest:
                        assign(key, rest, ctx)
                else:
                    cur = None if key in _base.HEADER_ONLY else key
                    if key not in {"amc", "warranty", "breakdown", "other"}:   # not part of the TET tables
                        add_tender(key, rest) if rest else None
                    else:
                        cur = None
                continue
            # continuation text
            if cur in BIDDER_SET:
                assign(cur, ln, ctx)
            elif cur and (cur not in SHORT or cur not in vals):
                add_tender(cur, ln)
            elif section in section_text and kind == "p":
                section_text[section].append(ln)

    # ------------------ assemble ------------------
    d = {k: "\n".join(v) for k, v in vals.items()}
    d = {k: v for k, v in d.items() if not _is_placeholder(v)}
    if "forwarded_by" not in d and signers:
        d["forwarded_by"] = "; ".join(signers)
        notes.append("Forwarded by: taken from the TET signature table.")
    if "description" not in d and d.get("subject"):
        d["description"] = d["subject"]
        notes.append("Short description: taken from the Subject line - please verify.")
    for sec, key, nice in (("introduction", "purpose", "Purpose of Tender"), ("background", "need", "How the need arose")):
        if key not in d and section_text[sec]:
            d[key] = "\n".join(section_text[sec])
            notes.append(f"{nice}: no label found; copied the whole '{sec.title()}' section - please trim.")
    d["_budget"] = f"CAPEX: {d.get('capex', '-')}\nOPEX: {d.get('opex', '-')}"
    d["_jtb"] = f"JTB Ref No: {d.get('jtb_ref', '')}".rstrip() + f"\nDate: {d.get('jtb_date', '')}".rstrip()

    if not bidders and shared:
        bidders.append({"name": "", "agent": "", "loa": "", "brand": "", "validity": "", "bond": "", **{k: "" for k in BIDDER_KEYS}})
        notes.append("No bidder table found - values were placed under a single 'Bidder 1' column.")
    # one sentence for everybody ("all bidders complied fully")
    body = " ".join(_clean(l) for kk, x in items if kk == "p" for l in x.split("\n") if not is_template_text(l, TET_LABELS))
    if not any(k in shared or any(b[k] for b in bidders) for k in COMPLIANCE) and \
            re.search(r"compl(?:ied|iant)\s+fully|fully\s+compl", body, re.I):
        for b in bidders:
            for k in COMPLIANCE:
                b[k] = "Compliant"
        notes.append("Compliance: report states full compliance in one sentence; all bidders set to 'Compliant' - please verify.")
    for key, texts in shared.items():
        txt = "\n".join(texts)
        targets = [b for b in bidders if not b[key]]
        for b in targets:
            b[key] = txt
        if len(bidders) > 1 and targets:
            notes.append(f"{NICE.get(key, key)}: no per-bidder breakdown found - same text copied to every bidder; please split it.")
    for b in bidders:
        for k in COMPLIANCE + ["status", "rec_fin", "bom"]:
            if b[k]:
                b[k] = _norm_cmp(b[k])
    # infer qualified / disqualified
    for i, b in enumerate(bidders):
        if not b["status"]:
            major = b["major"].strip().lower()
            if major and not re.match(r"^(?:none|nil|n/?a|-+|no)\b", major):
                b["status"] = "Disqualified"
                notes.append(f"Bidder {i + 1}: status set to 'Disqualified' because major non-compliances were listed - please verify.")
            elif b["rec_fin"] == "Yes":
                b["status"] = "Qualified"
                notes.append(f"Bidder {i + 1}: status set to 'Qualified' because it is recommended for financial opening - please verify.")
    return {"fields": d, "bidders": bidders, "notes": notes}
