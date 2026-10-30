import difflib
import os
import re
import shutil
import subprocess
import tempfile

# ----------------------------------------------------------------------------- fields
TEXT_FIELDS = ["sno", "desc", "unit", "origin", "remarks"]
NUM_FIELDS = ["qty", "unit_excl", "vat", "nbt", "other_tax", "unit_incl", "fob_unit", "fob_total",
              "insurance", "freight", "cif_total", "agent_comm", "amount"]
FIELD_LABEL = {
    "sno": "S/No", "desc": "Description", "qty": "Qty", "unit": "Unit", "origin": "Country of Origin",
    "unit_excl": "Unit Price (excl. taxes)", "vat": "VAT %", "nbt": "NBT %", "other_tax": "Other taxes %",
    "unit_incl": "Unit Price (incl. taxes)", "fob_unit": "FOB/FCA Unit Price", "fob_total": "FOB/FCA Total",
    "insurance": "Insurance", "freight": "Freight", "cif_total": "CIF/CIP Total", "agent_comm": "Agent's Commission",
    "amount": "Amount / Total", "remarks": "Remarks"}

_VOCAB = ["description", "quantity", "qty", "unit", "price", "excluding", "including", "taxes", "tax", "vat", "nbt",
          "country", "origin", "freight", "insurance", "commission", "remarks", "remark", "total", "amount", "agent",
          "specification", "particulars", "other", "cost", "rate", "item", "local"]


def _norm(t):
    t = (t or "").lower().replace("\n", " ")
    t = re.sub(r"[^a-z0-9%$]+", " ", t)
    return re.sub(r"\s+", " ", t).strip()


def _fuzzy(t):
    """Repair OCR typos in header words (oty -> qty, descnption -> description)."""
    out = []
    for w in t.split():
        if len(w) >= 3 and w not in _VOCAB:
            m = difflib.get_close_matches(w, _VOCAB, n=1, cutoff=0.78)
            w = m[0] if m else w
        elif w == "oty":
            w = "qty"
        out.append(w)
    return " ".join(out)


def classify_header(text):
    """Header cell text -> standard field name (or None)."""
    t = _fuzzy(_norm(text))
    if not t:
        return None
    if re.search(r"country.*origin|^origin\b|made in", t):
        return "origin"
    if re.search(r"agent|commission", t):
        return "agent_comm"
    if re.search(r"insurance", t):
        return "insurance"
    if re.search(r"freight", t):
        return "freight"
    if re.search(r"\b(cif|cip)\b", t):
        return "cif_total"
    if re.search(r"\b(fob|fca|exw)\b", t):
        return "fob_total" if (re.search(r"total|cost|amount", t) and "unit" not in t) else "fob_unit"
    if re.search(r"\bvat\b", t):
        return "vat"
    if re.search(r"\bnbt\b", t):
        return "nbt"
    if re.search(r"other\s*tax", t):
        return "other_tax"
    if re.search(r"unit\s*(price|rate|cost)|(price|rate)\s*(per|/)\s*unit|^(price|rate)\b", t):
        return "unit_incl" if re.search(r"including|inclusive|incl\b|with tax", t) else "unit_excl"
    if re.search(r"\b(total|amount|value|extended)\b", t):
        return "amount"
    if re.fullmatch(r"(s|sl|sr|serial)?\s*(no|number|nr|#)?\s*\.?|item\s*(no|number|nr|#)|line\s*(no|number)?|ref|sn|s n|s no", t) and t:
        return "sno"
    if re.search(r"descri|particular|specification|item name|details|goods|services?\b", t):
        return "desc"
    if t == "item":
        return "item_generic"
    if re.search(r"\bqty\b|quantity|\bno of\b|\bnos\b$", t):
        return "qty"
    if re.fullmatch(r"unit|uom|units|unit of measure(ment)?|measure", t):
        return "unit"
    if re.search(r"remark|note|comment", t):
        return "remarks"
    return None


def map_header(cells):
    """cells -> {col_index: field}; also returns list of unmapped (col, text)."""
    raw = [classify_header(c) for c in cells]
    first = True
    for i, f in enumerate(raw):
        if f == "item_generic":
            raw[i] = ("sno" if "desc" in raw else "desc") if first else None
            first = False
    mapping, seen, extra = {}, set(), []
    for i, f in enumerate(raw):
        if f and f not in seen:
            mapping[i] = f
            seen.add(f)
        elif (cells[i] or "").strip():
            extra.append((i, " ".join(cells[i].split())))
    return mapping, extra


def header_score(cells):
    mapping, _ = map_header(cells)
    f = set(mapping.values())
    return len(f) if "desc" in f and len(f) >= 3 else 0


# ----------------------------------------------------------------------------- numbers
_NONNUM = re.compile(r"^(?:[-–—.]+|n/?a|nil|none|included|incl\.?|inclusive|free|foc|free of charge|tbd|as above|exempt\w*)$", re.I)
_OCRFIX = str.maketrans({"O": "0", "o": "0", "I": "1", "l": "1", "|": "1", "S": "5", "B": "8"})


def parse_num(s, ocr=False):
    """'Rs. 1,234.50/=' -> 1234.5 ; '18%' -> 18 ; 'Included' -> None."""
    if s is None:
        return None
    if isinstance(s, (int, float)):
        return float(s)
    t = " ".join(str(s).split())
    if not t or _NONNUM.match(t):
        return None
    t = re.sub(r"(?i)\b(?:rs|lkr|usd|us|eur|gbp|inr|jpy|cny|sgd|aud)\b\.?", "", t)
    t = re.sub(r"[$€£₹]", "", t)
    t = re.sub(r"/=|/-", "", t).strip()
    neg = t.startswith("(") and t.endswith(")") or t.startswith("-")
    t = t.strip("()- ").replace("%", "").strip()
    if ocr:
        t = re.sub(r"[:;'`\"!]", "", t)                       # stray marks from the scanner
        if re.fullmatch(r"[0-9OoIl|SB,.\s]+", t) and re.search(r"\d", t):
            t = t.translate(_OCRFIX)
    t = t.replace(" ", "")
    if re.fullmatch(r"\d{1,3}(\.\d{3})+,\d+", t):                 # European 1.234.567,89
        t = t.replace(".", "").replace(",", ".")
    elif re.fullmatch(r"\d{1,3}([.,]\d{3})+[.,]\d{2}", t):        # mixed/misread separators: 14,021.796.00 -> 14021796.00
        t = re.sub(r"[.,]", "", t[:-3]) + "." + t[-2:]
    else:
        t = t.replace(",", "")
    if not re.fullmatch(r"\d+(\.\d+)?|\.\d+", t):
        return None
    v = float(t)
    return -v if neg else v


def split_qty_unit(s, ocr=False):
    s = " ".join((s or "").split())
    m = re.fullmatch(r"\s*([\d,]+(?:\.\d+)?)\s*([A-Za-z][A-Za-z\. ]{0,14})\s*", s or "")
    if m:
        return parse_num(m.group(1)), m.group(2).strip()
    return None, None


_CUR = [(r"\bUSD\b|US\s*\$", "USD"), (r"\bEUR\b|€", "EUR"), (r"\bGBP\b|£", "GBP"), (r"\bJPY\b", "JPY"), (r"\bCNY\b|\bRMB\b", "CNY"),
        (r"\bINR\b", "INR"), (r"\bSGD\b", "SGD"), (r"\bAUD\b", "AUD"), (r"\bLKR\b|\bRs\b\.?|Rupees", "LKR")]


def detect_currency(text):
    for rx, code in _CUR:
        if re.search(rx, text or "", re.I):
            return code
    return ""


def best_header(rows, search=5):
    """Find the header of a table. Tries every row alone and stacked with the next row (two-line headers),
    keeps the reading that recognises the most columns. -> (row_index, header_cells, span) or (None, None, 1)."""
    best = (0, None, None, 1)
    for i in range(min(search, len(rows))):
        cand = [(header_score(rows[i]), rows[i], 1)]
        if i + 1 < len(rows) and len(rows[i]) == len(rows[i + 1]):
            joined = [(a + " " + b).strip() for a, b in zip(rows[i], rows[i + 1])]
            cand.append((header_score(joined), joined, 2))
        for sc, cells, span in cand:
            if sc > best[0]:
                best = (sc, i, cells, span)
        if best[1] is not None:
            break
    return (best[1], best[2], best[3]) if best[1] is not None else (None, None, 1)


# ----------------------------------------------------------------------------- table -> lines
TOTAL_RX = re.compile(r"\b(grand\s*total|sub\s*-?\s*total|total|net\s*amount|balance)\b", re.I)
_PH = re.compile(r"^[.…_\-\s]*$")
_GUIDANCE = re.compile(r"^(?:\*?\s*)(?:eg\.?|e\.g\.)|^field$|^details$|^status\s*/\s*details$|^requirement$", re.I)
_NUM_ONLY = re.compile(r"^\(?(?:\d{1,2}|[a-z]|[ivx]{1,4})[\.\)]?$", re.I)


def is_template_text(line):
    """Check if a line represents template guidance, boilerplate or placeholder text."""
    s = (line or "").strip()
    if not s or s in {"-", "–", "—", "N/A?"}:
        return True
    if _PH.match(s) or _GUIDANCE.search(s) or _NUM_ONLY.match(s):
        return True
    if s.startswith("(") and s.endswith(")") and ("/" in s or "if any" in s.lower() or "e.g." in s.lower()):
        return True
    return False


def _clean(s):
    s = (s or "").replace(" ", " ")
    return "\n".join(" ".join(ln.split()) for ln in s.split("\n")).strip()


def _classify_part(fields):
    f = set(fields)
    if f & {"fob_unit", "fob_total", "cif_total", "insurance", "freight", "origin", "agent_comm"}:
        return "B Foreign"
    if f & {"vat", "nbt", "other_tax", "unit_incl"}:
        return "A Local"
    return "Other"


def table_to_lines(tbl, state, notes):
    """tbl: {"rows": [[str]], "page": int, "method": str, "conf": [[float|None]]|None}.
    Returns (lines, info). `state` carries the last header mapping so continued tables work."""
    rows = [[_clean(c) for c in r] for r in tbl["rows"]]
    conf = tbl.get("conf")
    ocr = tbl.get("method") == "ocr"
    # ---- locate header (1 row, or 2 stacked rows)
    hdr_i, hdr_cells, span = best_header(rows)
    if hdr_i is None:
        ncols = max((len(r) for r in rows), default=0)
        if state.get("mapping") and ncols == state.get("ncols"):
            mapping, extra, hdr_cells, start = state["mapping"], state["extra"], state["hdr"], 0
            notes.append(f"Page {tbl['page']}: table without header treated as continuation of the previous table.")
        else:
            return [], {"mapped": False, "part": "Other", "fields": [], "header": None}
    else:
        mapping, extra = map_header(hdr_cells)
        start = hdr_i + span
        state.update(mapping=mapping, extra=extra, hdr=hdr_cells, ncols=len(hdr_cells))
    fields = list(mapping.values())
    part = _classify_part(fields)
    htext = " ".join(hdr_cells)
    currency = detect_currency(htext) or detect_currency(" ".join(c for r in rows[start:start + 12] for c in r))
    lines, last = [], None
    skipped = 0
    for ri in range(start, len(rows)):
        r = rows[ri]
        if not any(r):
            continue
        if header_score(r):              # repeated header on a new page
            continue
        rec = {k: "" for k in TEXT_FIELDS}
        rec.update({k: None for k in NUM_FIELDS})
        low = set()
        extras = []
        for ci, cell in enumerate(r):
            if not cell:
                continue
            f = mapping.get(ci)
            lowc = bool(conf and ri < len(conf) and ci < len(conf[ri]) and conf[ri][ci] is not None and conf[ri][ci] < 85)
            if f is None:
                hdr = next((t for c, t in extra if c == ci), f"Column {ci + 1}")
                extras.append(f"{hdr}: {cell}")
                continue
            if f in NUM_FIELDS:
                v = parse_num(cell, ocr)
                if f == "qty" and v is None:
                    v, u = split_qty_unit(cell, ocr)
                    if v is not None and u and not rec["unit"]:
                        rec["unit"] = u
                if v is None:
                    # keep the supplier's own wording ("Included", "FOC", "NIL"); a bare dash means empty
                    rec[f] = None if re.fullmatch(r"[-–—.\s]+", cell) else cell
                else:
                    rec[f] = v
                if lowc:
                    low.add(f)
            else:
                rec[f] = (rec[f] + "\n" + cell).strip() if rec[f] else cell
                if lowc:
                    low.add(f)
        anynum = any(rec[f] not in (None, "") for f in NUM_FIELDS)
        has_sno, has_desc = bool(rec["sno"]), bool(rec["desc"])
        label = (rec["desc"] + " " + rec["sno"]).strip()
        # --- classify the row
        if TOTAL_RX.search(label) and (not isinstance(rec["qty"], float)):
            rtype = "Sub-total" if re.search(r"sub\s*-?\s*total", label, re.I) else "Total"
        elif not anynum and not has_sno and has_desc and last is not None and last["rowtype"] == "Item" and not extras and not rec["origin"]:
            last["desc"] += "\n" + rec["desc"]            # wrapped description line
            last["lowconf"] |= low
            continue
        elif not anynum and not has_desc and not extras and not rec["remarks"] and not rec["origin"] and not rec["unit"]:
            skipped += 1                                  # numbering only = unfilled template row
            continue
        elif not anynum and has_desc and not has_sno:
            rtype = "Note"
        else:
            rtype = "Item"
        if rtype in ("Total", "Sub-total"):
            filled = [f for f in NUM_FIELDS if isinstance(rec[f], float)]
            if len(filled) == 1 and filled[0] != "amount":
                rec["amount"], rec[filled[0]] = rec[filled[0]], None
                low = {"amount" if x == filled[0] else x for x in low}
        if rec["remarks"] and extras:
            rec["remarks"] += "; "
        rec["remarks"] = (rec["remarks"] + "; ".join(extras)).strip("; ")
        rec.update(rowtype=rtype, part=part, page=tbl["page"], currency=currency or "", ocr=ocr, lowconf=low,
                   method=tbl.get("method", ""))
        lines.append(rec)
        last = rec
    if skipped:
        notes.append(f"Page {tbl['page']}: {skipped} numbered row(s) with no content ignored (unfilled).")
    return lines, {"mapped": True, "part": part, "fields": fields, "header": hdr_cells, "extra": extra}


# ----------------------------------------------------------------------------- readers
def _soffice():
    for n in ("soffice", "libreoffice"):
        p = shutil.which(n)
        if p:
            return p
    return None


def convert_doc(path):
    """.doc / .rtf / .odt -> .docx using LibreOffice (must be installed)."""
    exe = _soffice()
    if not exe:
        raise RuntimeError("Reading .doc files needs LibreOffice installed (or save the file as .docx / .pdf first).")
    out = tempfile.mkdtemp(prefix="doc2docx_")
    prof = tempfile.mkdtemp(prefix="lo_profile_")
    cmd = [exe, f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", "docx", "--outdir", out, path]
    subprocess.run(cmd, capture_output=True, timeout=180)
    res = os.path.join(out, os.path.splitext(os.path.basename(path))[0] + ".docx")
    if not os.path.exists(res):
        raise RuntimeError("Could not convert the .doc file - please save it as .docx or PDF and upload again.")
    return res


def read_docx(path):
    from docx import Document
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    doc = Document(path)
    tables, text = [], []
    for el in doc.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            text.extend(Paragraph(el, doc).text.split("\n"))
        elif tag == "tbl":
            t = Table(el, doc)
            rows = []
            for row in t.rows:
                seen, cells = set(), []
                for c in row.cells:
                    cells.append(c.text if id(c._tc) not in seen else "")
                    seen.add(id(c._tc))
                rows.append(cells)
            tables.append({"rows": rows, "page": 1, "method": "docx", "conf": None})
    return tables, text, {"method": "Word table", "pages": 1}


def read_pdf(path):
    import pdfplumber
    from price_ocr import ocr_image, positional_tables, render_pdf_page
    tables, text, scanned_pages = [], [], 0
    with pdfplumber.open(path) as pdf:
        n = len(pdf.pages)
        for pi, page in enumerate(pdf.pages, 1):
            ptxt = page.extract_text() or ""
            if len(ptxt.strip()) < 30:                      # no text layer -> scanned page
                scanned_pages += 1
                img = render_pdf_page(path, pi)
                t, ln = ocr_image(img, pi)
                tables += t
                text += ln
                continue
            text += ptxt.split("\n")
            found = [t for t in page.extract_tables() if t and len(t) > 1 and max(len(r) for r in t) >= 3]
            for t in found:
                tables.append({"rows": [[c or "" for c in r] for r in t], "page": pi, "method": "pdf", "conf": None})
            if not any(header_score([_clean(c) for c in t["rows"][0]]) or any(header_score([_clean(c) for c in r]) for r in t["rows"][:4])
                       for t in tables if t["page"] == pi):
                words = [{"text": w["text"], "x0": w["x0"], "x1": w["x1"], "top": w["top"], "bottom": w["bottom"], "conf": None}
                         for w in page.extract_words(keep_blank_chars=False, use_text_flow=False)]
                tables += [dict(t, page=pi, method="pdf") for t in positional_tables(words)]
    meth = "PDF text" if not scanned_pages else (f"OCR ({scanned_pages}/{n} page(s))")
    return tables, text, {"method": meth, "pages": n}


def read_image(path):
    from price_ocr import load_images, ocr_image
    tables, text, n = [], [], 0
    for pi, img in enumerate(load_images(path), 1):
        t, ln = ocr_image(img, pi)
        tables += t
        text += ln
        n += 1
    return tables, text, {"method": "OCR (image)", "pages": n}


# ----------------------------------------------------------------------------- main entry
_TEND = re.compile(r"tender\s*(?:no|number|ref(?:erence)?)\.?\s*[:\-–]\s*(.+)", re.I)
_SUPP = re.compile(r"(?:name\s*of\s*(?:the\s*)?(?:tenderer|bidder|supplier|company|firm)|company(?:\s*name)?|supplier(?:\s*name)?|bidder(?:\s*name)?|quotation\s*from)\s*[:\-–]\s*(.+)", re.I)
_REMK = re.compile(r"\b(validity|valid\s+for|delivery|lead\s*time|warranty|payment|terms|remark|note|incoterm|discount|offer)\b", re.I)


def extract_price(path):

    ext = os.path.splitext(path)[1].lower()
    src = path
    if ext in (".doc", ".rtf", ".odt"):
        src = convert_doc(path)
        ext = ".docx"
    if ext == ".docx":
        tables, text, info = read_docx(src)
    elif ext == ".pdf":
        tables, text, info = read_pdf(src)
    elif ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        tables, text, info = read_image(src)
    else:
        raise ValueError(f"Unsupported file type: {ext}")
    notes, lines, raw, state = [], [], [], {}
    for ti, tb in enumerate(tables, 1):
        ls, inf = table_to_lines(tb, state, notes)
        for l in ls:
            l["table"] = ti
        lines += ls
        raw.append({"idx": ti, "page": tb["page"], "method": tb.get("method", ""), "rows": tb.get("raw_rows", tb["rows"]), "conf": tb.get("conf"),
                    "part": inf["part"] if inf["mapped"] else "unrecognised", "mapped": inf["mapped"],
                    "n_lines": len(ls)})
        if not inf["mapped"]:
            notes.append(f"Page {tb['page']}: a table was found but its columns were not recognised - kept verbatim on 'Raw Tables'.")
        elif inf.get("extra"):
            notes.append(f"Page {tb['page']}: extra column(s) kept inside Remarks: " + ", ".join(t for _, t in inf["extra"]))
    # ---- meta
    tender = supplier = ""
    remarks = []
    for ln in text:
        s = " ".join(ln.split())
        if not s or is_template_text(s):
            continue
        m = _TEND.search(s)
        if m and not tender and not _PH.match(m.group(1)):
            tender = m.group(1).strip(" .…")
        m = _SUPP.search(s)
        if m and not supplier and not _PH.match(m.group(1)):
            supplier = m.group(1).strip(" .…")
        if _REMK.search(s) and len(s) < 220 and not _PH.match(s):
            remarks.append(s)
    if not supplier:
        notes.append("Supplier name not found in the document - the file name is used; please correct it in the Summary sheet.")
    items = [l for l in lines if l["rowtype"] == "Item"]
    if not lines:
        notes.append("No filled price lines were found (the file may be a blank form).")
    if info["method"].startswith("OCR"):
        nlow = sum(len(l["lowconf"]) for l in lines)
        notes.append(f"OCR was used - {nlow} cell(s) with low confidence are highlighted orange; please verify against the original.")
    return {"fields": {"tender_no": tender, "supplier": supplier or os.path.splitext(os.path.basename(path))[0],
                       "method": info["method"], "pages": info["pages"], "doc_remarks": "\n".join(remarks[:12])},
            "bidders": [], "notes": notes, "lines": lines, "raw_tables": raw,
            "summary": f"{len(items)} item line(s), {len(raw)} table(s), {info['method']}"}
