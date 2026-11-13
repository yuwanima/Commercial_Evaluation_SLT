
import argparse, io, re, sys, tempfile, zipfile
from collections import OrderedDict, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L

try:
    import docx
except ImportError:
    docx = None

HERE = Path(__file__).resolve().parent
if str(HERE) not in sys.path:
    sys.path.insert(0, str(HERE))

# ----------------------------------------------------------------- categories
CATS = {
    1: "Goods & Services (Local or Foreign)",
    2: "Goods (Local & Foreign)",
    3: "Goods & Services (Local)",
    4: "Goods (Local)",
    5: "Services (Local or Foreign)",
}
FORM_CODE_TO_CAT = {"072": 1, "073": 2, "074": 3, "075": 4, "076": 5}

# ----------------------------------------------------------------- helpers
def clean(v):
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    return re.sub(r"[ \t ]+", " ", str(v).replace("\r", "")).strip()

def clean_multiline(v):
    s = clean(v)
    return re.sub(r"\n\s*\n+", "\n", re.sub(r" *\n *", "\n", s)).strip()

def clean_note(v):
    """Normalize note whitespace while retaining paragraph breaks from the source."""
    if v is None:
        return ""
    text = str(v).replace("\r\n", "\n").replace("\r", "\n").replace("\xa0", " ")
    text = re.sub(r"[ \t]+\n", "\n", text)
    text = re.sub(r"\n[ \t]+", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    heading = re.compile(r"\s+(?=[A-Z][A-Za-z][A-Za-z /&'()-]{1,40}\s+[–—-]\s+)")
    parts = []
    end = 0
    for match in heading.finditer(text):
        if text[:match.start()].rstrip().endswith((".", ":", ";")):
            parts.append(text[end:match.start()].rstrip())
            end = match.end()
            parts.append(text[match.start():match.end()].strip())
    if parts:
        parts.append(text[end:].lstrip())
        text = "\n\n".join(part for part in parts if part)
    return text.strip()

def norm_status(raw):
    """Return (status, raw). Status: Complied / Partially Complied / Not Complied / Not Filled / Unclear."""
    s = clean(raw)
    if not s or set(s) <= set(".-_…– "):
        return "Not Filled"
    t = s.lower()
    if re.search(r"partial", t):
        return "Partially Complied"
    if re.search(r"not\s*[- ]?compl|non[- ]?compl|\bnc\b|\bno\b|deviat|reject|not\s+agree|not\s+accept", t):
        return "Not Complied"
    if re.search(r"compl|^c$|\byes\b|✓|✔|√|\bagree|\baccept|\bok\b|confirm", t):
        return "Complied"
    return "Unclear"

def derive_bidder_name(path, root):
    if root and root != path.parent and root in path.parents:
        return path.relative_to(root).parts[0]
    stem = path.stem.replace("_", " ")             # "\b" does not see a boundary next to "_"
    stem = re.sub(r"\(\d+\)", " ", stem)
    stem = re.sub(r"(?i)\b(copy of|compliance for|compliance|general conditions of contract|"
                  r"special conditions of contract|gcc|scc|goods|services|local|foreign|"
                  r"filled|completed|bid|bidder|tender|form|lgls[/ ]?fcc[/ ]?\d+)\b", " ", stem)
    stem = re.sub(r"[_\-–.&]+", " ", stem)
    stem = re.sub(r"\s+", " ", stem).strip(" -_")
    return stem or path.stem


_OFFICIAL_FRANK_LABEL = re.compile(
    r"(?i)\bofficial\s+(?:"
    r"frank\s*(?:(?:/|or)\s*(?:stamp|seal))?"
    r"|stamp|seal"
    r")\b|\bcompany\s+stamp\b"
)
_GENERIC_OFFICIAL_FRANK_VALUES = {
    "", "n/a", "na", "nil", "none", "-", "—", "not applicable",
    "official frank", "official stamp", "official seal", "note", "date",
    "yes", "no", "complied", "updated details", "details required",
}


def extract_official_frank_name(text, fallback):
    """Return the company written in the document's official frank or stamp field."""
    lines = [clean(line) for line in text.splitlines()]
    for index, line in enumerate(lines):
        match = _OFFICIAL_FRANK_LABEL.search(line)
        if not match:
            continue
        candidates = [line[match.end():], *lines[index + 1:index + 3]]
        for candidate in candidates:
            value = re.sub(r"\s+", " ", candidate.replace("\ufffd", " ")).strip(" \t:;,-–—")
            value = re.sub(r"^[^0-9A-Za-z]+", "", value)
            if re.fullmatch(r"\d+(?:[./-]\d+)*", value):
                continue
            if value.casefold() in _GENERIC_OFFICIAL_FRANK_VALUES:
                continue
            if re.match(
                r"(?i)^(?:note|date|part\b|please\b|all\b|incomplete\b|any\b|"
                r"compliance\s+statement|general\s+conditions|form\s+of|"
                r"page\s+\d+|clause(?:\s+no)?\b|description\b|complied\b|"
                r"details\s+required|updated\s+details|yes\b|no\b|s/no\b|qty\b|unit\s+price|vat\b|nbt\b|other\s+taxes|"
                r"country\s+of\s+origin|fob\b|cif\b|explanatory\s+note|signature\b|tender\s*(?:no|number)?|"
                r"name\s*(?:&|and)\s*signature)\b",
                value,
            ):
                continue
            if value and not _OFFICIAL_FRANK_LABEL.fullmatch(value):
                return value
    return fallback


def extract_official_frank_ocr_name(ocr_text, fallback):
    """Use OCR only for text following an OCR-detected official frank or stamp label."""
    match = _OFFICIAL_FRANK_LABEL.search(ocr_text)
    if not match:
        return fallback
    return extract_official_frank_name(
        f"{match.group(0)}\n{ocr_text[match.end():]}",
        fallback,
    )


def has_official_frank_label(text):
    """Return whether text refers to an official frank, stamp, or seal field."""
    return bool(_OFFICIAL_FRANK_LABEL.search(text or ""))


# ── Header company-name extraction (PRIMARY method) ──────────────────────────
# Matches legal-entity suffixes that reliably identify a company name line.
_HEADER_COMPANY_SUFFIXES = re.compile(
    r"(?i)\b(?:"
    r"(?:pvt|private|public)\s*\.?\s*ltd|limited|incorporated|inc\b|"
    r"corporation|corp\b|llc|llp|lp\b|plc\b|"
    r"\(pvt\)\s*ltd|\(private\)\s*limited|\(guarantee\)\s*limited|"
    r"company|co\b|"
    r"associates?|group\b|holdings?|enterprises?|"
    r"industries|industry|traders?|trading|international|services?|"
    r"solutions?|technologies?|technology|systems?|consultants?|consulting|"
    r"engineering|engineers?|construction|builders?|suppliers?|imports?|exports?"
    r")\b"
)

_HEADER_SKIP = re.compile(
    r"(?i)^(?:"
    r"date|ref|reference|to\b|from\b|subject|attn|attention|dear|"
    r"tender|no\.?|number|page\b|clause|description|item\b|s\.?no\.?|"
    r"name\s*(?:&|and)\s*signature|signature|authorized|"
    r"compliance|general\s+conditions|special\s+conditions|"
    r"goods\b|services\b|form\b|lgls|fcc|"
    r"bid\b|bidder|this\s+document"
    r")\b"
)

# Number of non-empty lines from the top to treat as the document header
_HEADER_SCAN_LINES = 20


def extract_header_name(text, fallback):
    """
    PRIMARY company-name detection: scan the document HEADER (letterhead /
    title area, first ~20 non-empty lines) for a company name.

    Strategy:
      Pass 1 - return the first header line that contains a known legal-entity
               suffix (Ltd, Pvt Ltd, Corp, PLC, etc.).
      Pass 2 - return the longest mostly-capitalised line in the very first 5
               lines (letterhead almost always has the company name there).
    Returns *fallback* when nothing suitable is found, so the caller can try
    the Official Frank / stamp method next.
    """
    lines = [clean(ln) for ln in text.splitlines() if clean(ln)]
    header_lines = lines[:_HEADER_SCAN_LINES]

    # Pass 1: legal-entity suffix in the header
    for line in header_lines:
        if _HEADER_SKIP.match(line):
            continue
        if _OFFICIAL_FRANK_LABEL.search(line):
            continue
        if _HEADER_COMPANY_SUFFIXES.search(line):
            # Strip leading salutation noise ("To:", "From:", "Re:" …)
            value = re.sub(r"(?i)^(?:to|from|re|ref|attn)\s*:\s*", "", line)
            value = value.strip(" \t:;,-")
            if 3 < len(value) < 120:
                return value

    # Pass 2: longest mostly-capitalised line in the first 5 lines
    for line in lines[:5]:
        if _HEADER_SKIP.match(line):
            continue
        if _OFFICIAL_FRANK_LABEL.search(line):
            continue
        words = line.split()
        if len(words) < 2:
            continue
        cap_ratio = sum(1 for w in words if w and w[0].isupper()) / len(words)
        if cap_ratio >= 0.6 and 6 < len(line) < 100:
            return line

    return fallback



def ocr_embedded_images(path):
    """OCR embedded document images for image-only company stamps."""
    try:
        import pytesseract
        from PIL import Image
        from pypdf import PdfReader
    except ImportError:
        return ""
    tesseract = Path(r"C:\Program Files\Tesseract-OCR\tesseract.exe")
    if tesseract.exists():
        pytesseract.pytesseract.tesseract_cmd = str(tesseract)
    image_data = []
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        reader = PdfReader(str(path))
        image_data = [image.data for page in reader.pages for image in page.images]
    elif path.suffix.lower() in (".docx", ".xlsx", ".xlsm"):
        with zipfile.ZipFile(path) as archive:
            image_data = [
                archive.read(name)
                for name in archive.namelist()
                if name.startswith(("word/media/", "xl/media/"))
            ]
    text = []
    for data in image_data:
        try:
            image = Image.open(io.BytesIO(data)).convert("RGBA")
            background = Image.new("RGB", image.size, "white")
            background.paste(image, mask=image.getchannel("A"))
            value = pytesseract.image_to_string(background).strip()
        except (OSError, ValueError, pytesseract.TesseractNotFoundError):
            continue
        if value:
            text.append(value)
    if path.suffix.lower() == ".pdf":
        try:
            import importlib
            fitz = importlib.import_module("fitz")
            pdf = fitz.open(str(path))
            for page in pdf:
                blocks = page.get_text("blocks")
                for block in blocks:
                    if not re.search(r"(?i)official\s+frank", block[4] or ""):
                        continue
                    x0, y0, x1, y1 = block[:4]
                    clip = fitz.Rect(max(0, x0 - 120), max(0, y0 - 160),
                                     min(page.rect.width, x1 + 420),
                                     min(page.rect.height, y1 + 240))
                    pixmap = page.get_pixmap(matrix=fitz.Matrix(3, 3), clip=clip, alpha=False)
                    value = pytesseract.image_to_string(
                        Image.frombytes("RGB", (pixmap.width, pixmap.height), pixmap.samples)
                    ).strip()
                    if value:
                        text.append(value)
            pdf.close()
        except (ImportError, OSError, ValueError, pytesseract.TesseractNotFoundError):
            pass
    return "\n".join(text)


def document_has_embedded_images(path):
    """Return whether a document contains embedded image content."""
    path = Path(path)
    if path.suffix.lower() == ".pdf":
        from pypdf import PdfReader
        return any(page.images for page in PdfReader(str(path)).pages)
    if path.suffix.lower() in (".docx", ".xlsx", ".xlsm"):
        with zipfile.ZipFile(path) as archive:
            return any(
                name.startswith(("word/media/", "xl/media/"))
                for name in archive.namelist()
            )
    return False


def document_text(path):
    """Return searchable text from a normalized GCC workbook or SCC document."""
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True)
        return "\n".join(
            clean(cell.value)
            for ws in wb.worksheets
            for row in ws.iter_rows()
            for cell in row
            if clean(cell.value)
        )
    if docx is None:
        return ""
    d = docx.Document(path)
    parts = [p.text for p in d.paragraphs]
    parts.extend(cell.text for table in d.tables for row in table.rows for cell in row.cells)
    return "\n".join(clean_multiline(x) for x in parts if clean_multiline(x))

def document_summary(path, bidder, kind, display_name=None, extra_text="",
                     has_embedded_images=False, name_source=None):
    text = document_text(path)
    text = f"{text}\n{extra_text}"
    lower = text.lower()
    tender = re.search(r"(?i)\btender\s*(?:no\.?|number)\s*[:\-]?\s*([A-Za-z0-9/.\-]+)", text)
    dates = re.findall(
        r"\b(?:\d{1,2}[/-]\d{1,2}[/-]\d{2,4}|\d{4}[/-]\d{1,2}[/-]\d{1,2}"
        r"|\d{1,2}\s+(?:January|February|March|April|May|June|July|August|September|October|November|December)"
        r"\s+\d{4})\b",
        text,
        flags=re.IGNORECASE,
    )
    signature = bool(re.search(r"signature\s+of\s+(?:the\s+)?tender|tenderer.?s?\s+signature|authorized\s+sign", lower))
    signed_value = bool(
        re.search(r"signature.{0,80}(yes|signed|present|✓|✔)", lower, re.DOTALL)
        or (signature and has_embedded_images)
    )
    frank = bool(re.search(r"official\s+(?:frank|stamp|seal)|official\s+seal|company\s+stamp|rubber\s+stamp", lower))
    return {
        "bidder": bidder,
        "document": display_name or path.name,
        "type": kind,
        "points": [
            f"Document type: {kind}",
            f"Bidder name: {bidder}",
            f"Bidder name source: {name_source or 'document'}",
            f"Tender number: {tender.group(1) if tender else 'Not found'}",
            f"Signature section: {'Present' if signature else 'Not found'}",
            f"Signature completed: {'Found' if signed_value else 'Not found'}",
            f"Date: {', '.join(dict.fromkeys(dates)) if dates else 'Not found'}",
            f"Official frank / stamp / seal reference: {'Present' if frank else 'Not found'}",
        ],
    }

# ----------------------------------------------------------------- GCC reader
def parse_gcc(path):
    """Read one bidder's GCC compliance workbook."""
    wb = openpyxl.load_workbook(path, data_only=True)
    for ws in wb.worksheets:
        hdr = None
        for r in range(1, min(ws.max_row, 30) + 1):
            if clean(ws.cell(r, 1).value).lower().startswith("clause"):
                hdr = r
                break
        if hdr is None:
            continue
        code = None
        for r in range(1, hdr):
            for c in range(1, 8):
                m = re.search(r"LGLS/FCC/(\d+)", clean(ws.cell(r, c).value))
                if m:
                    code = m.group(1)
        rows = []
        for r in range(hdr + 1, ws.max_row + 1):
            a, b = clean(ws.cell(r, 1).value), clean(ws.cell(r, 2).value)
            if re.match(r"(?i)signature of", b):
                break
            if not a or not b or set(b) <= set(".…- "):
                continue
            status = norm_status(ws.cell(r, 3).value)
            rows.append(dict(clause=a, desc=b, status=status,
                             raw=clean(ws.cell(r, 3).value),
                             note=clean_note(ws.cell(r, 4).value)))
        if rows:
            cat = FORM_CODE_TO_CAT.get(code)
            return dict(cat=cat, code=code, rows=rows,
                        sig=tuple(x["clause"] for x in rows))
    raise ValueError("no 'Clause No' table found")

# ----------------------------------------------------------------- SCC reader
LEGEND = ("shall be deleted", "shall be amended", "the following shall be added")

def scc_category(title):
    t = title.lower()
    g, s, f = "goods" in t, "service" in t, "foreign" in t
    if g and s:
        return 1 if f else 3
    if g:
        return 2 if f else 4
    if s:
        return 5
    return None

def _rows(table):
    out = []
    for row in table.rows:
        cells, seen = [], set()
        for c in row.cells:
            key = id(c._tc)
            cells.append("" if key in seen else clean_note(c.text))
            seen.add(key)
        out.append(cells)
    return out

def parse_scc(path):
    if docx is None:
        raise RuntimeError("python-docx is not installed (pip install python-docx)")
    d = docx.Document(path)
    title = next((clean(p.text) for p in d.paragraphs if clean(p.text)), "")
    t1, t2 = [], []
    for t in d.tables:
        rows = _rows(t)
        if not rows:
            continue
        head = " ".join(rows[0]).lower()
        if "details required" in head:
            for r in rows[1:]:
                r = (r + ["", "", "", ""])[:4]
                if not r[0] and not r[1] and not r[2]:
                    continue
                t1.append(dict(clause=r[0], req=r[1], resp=r[2], stat=r[3]))
        elif "amendment" in head:
            for r in rows[1:]:
                r = (r + ["", "", "", ""])[:4]
                ref, clause, amend, rem = r
                if not (clause or amend or rem):
                    continue
                a = amend.lower().strip()
                if not clause and not rem and any(a.startswith(x) for x in LEGEND):
                    continue                       # printed legend row 01/02/03
                if not clause and not amend.strip() and not rem:
                    continue
                t2.append(dict(ref=ref, clause=clause, amend=amend, remarks=rem))
    if not t1 and not t2:
        raise ValueError("no SCC Table 1 / Table 2 found")
    return dict(cat=scc_category(title), title=title, t1=t1, t2=t2)

def template_scc_text():
    """Blank-template wording per category, to spot 'left unchanged' cells."""
    out = {}
    if docx is None:
        return out
    for i in CATS:
        p = HERE / "templates" / f"SCC_{i}.docx"
        if p.exists():
            try:
                out[i] = {r["clause"]: r["resp"] for r in parse_scc(p)["t1"]}
            except Exception:
                pass
    return out

def scc_status(resp, template_resp, stat=""):
    st = clean(stat)
    if st:                                   # the bidder ticked Complied / Not Complied in the form's own column
        s0 = norm_status(st)
        if s0 in ("Complied", "Not Complied", "Partially Complied"):
            return s0
    r = clean_multiline(resp)
    if not r or set(r) <= set(".-_…–\n :"):
        return "Not Filled"
    if template_resp is not None and re.sub(r"\W+", "", r.lower()) == re.sub(r"\W+", "", template_resp.lower()):
        return "Not Filled"
    s = norm_status(r.split("\n")[0][:40]) if re.match(r"(?i)\s*(not\s+|non[- ]?)?compl|\s*(yes|no)\b", r) else None
    if s in ("Complied", "Not Complied", "Partially Complied"):
        return s
    return "Provided"

# ----------------------------------------------------------------- Excel styles
NAVY, GREY = "1F3864", "F2F2F2"
HFILL = PatternFill("solid", fgColor=NAVY)
SUBFILL = PatternFill("solid", fgColor="D9E1F2")
BIDFILLS = ["1F3864", "2E75B6", "375623", "7F6000", "7030A0", "833C0B", "006666", "7B2C2C"]
GREEN, RED, AMBER, BLUEISH = "C6EFCE", "FFC7CE", "FFEB9C", "DDEBF7"
thin = Side(style="thin", color="A6A6A6")
BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
WRAP = Alignment(wrap_text=True, vertical="top")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

def hcell(ws, r, c, text, fill=HFILL, color="FFFFFF"):
    x = ws.cell(r, c, text)
    x.font = Font(bold=True, color=color)
    x.fill = fill if isinstance(fill, PatternFill) else PatternFill("solid", fgColor=fill)
    x.alignment = CENTER
    x.border = BORDER
    return x

def body(ws, r, c, v, bold=False, center=False):
    x = ws.cell(r, c, v)
    x.alignment = CENTER if center else WRAP
    x.border = BORDER
    if bold:
        x.font = Font(bold=True)
    return x

def point_note(value):
    """Keep source wording and paragraph structure for readable Excel notes."""
    text = clean_note(value)
    if not text:
        return ""
    points = []
    for paragraph in re.split(r"\n{2,}", text):
        lines = [line.strip() for line in paragraph.splitlines() if line.strip()]
        if lines:
            points.append(" ".join(lines))
    return "\n\n".join(points)

def set_note_row_height(ws, row, *values):
    lines = 1
    for value in values:
        text = str(value or "")
        lines = max(lines, sum(max(1, (len(line) + 32) // 33) for line in text.splitlines()))
    ws.row_dimensions[row].height = min(300, max(30, lines * 15))

def add_status_cf(ws, rng, first):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="Complied"'], fill=PatternFill("solid", bgColor=GREEN, fgColor=GREEN)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="Not Complied"'], fill=PatternFill("solid", bgColor=RED, fgColor=RED)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'OR({first}="Not Filled",{first}="Unclear",{first}="Partially Complied")'], fill=PatternFill("solid", bgColor=AMBER, fgColor=AMBER)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'{first}="Provided"'], fill=PatternFill("solid", bgColor=BLUEISH, fgColor=BLUEISH)))

def cf_eligible(ws, rng, first):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},3)="Yes"'], fill=PatternFill("solid", bgColor=GREEN, fgColor=GREEN)))
    ws.conditional_formatting.add(rng, FormulaRule(formula=[f'LEFT({first},2)="No"'], fill=PatternFill("solid", bgColor=RED, fgColor=RED)))

def finish(ws, freeze, landscape=True):
    ws.freeze_panes = None
    ws.sheet_view.zoomScale = 90
    ws.page_setup.orientation = "landscape" if landscape else "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)

def safe_sheet(name, used):
    name = re.sub(r"[\[\]\*\?/\\:]", "-", name)[:31]
    base, n = name, 2
    while name in used:
        name = f"{base[:28]} {n}"
        n += 1
    used.add(name)
    return name


# ----------------------------------------------------------------- clause alignment
import difflib
def _nd(t):
    return re.sub(r"[^a-z0-9]+", "", t.lower())

def align_clauses(bidders):
    """Match each bidder's rows to a master clause list by DESCRIPTION similarity
    (clause number only as a tie-break), so re-worded or re-numbered sheets still line up."""
    master = []   # dict(clause, desc, rows{idx: row})
    for i, b in enumerate(bidders):
        pos = -1
        for row in b["data"]["rows"]:
            best, bs = None, 0.0
            for k, m in enumerate(master):
                if i in m["rows"]:
                    continue
                sim = difflib.SequenceMatcher(None, _nd(row["desc"]), _nd(m["desc"])).ratio()
                if row["clause"] == m["clause"]:
                    sim += 0.15
                if k > pos:
                    sim += 0.02          # prefer keeping the original order
                if sim > bs:
                    best, bs = k, sim
            if best is not None and bs >= 0.72:
                master[best]["rows"][i] = row
                pos = best
            else:
                master.insert(pos + 1, dict(clause=row["clause"], desc=row["desc"], rows={i: row}))
                pos += 1
    return master

# ----------------------------------------------------------------- GCC workbook
def build_gcc(groups, out, document_type="GCC"):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    summary = []
    for gkey, bidders in groups.items():
        cat, code = gkey[0], gkey[1]
        label = CATS.get(cat)
        if label is None and document_type == "GCT":
            label = ", ".join(b["name"] for b in bidders)
        label = label or "Unrecognised template"
        # ordered union of clauses
        clauses = align_clauses(bidders)
        ws = wb.create_sheet(safe_sheet(f"{document_type} - {label}", used))
        n = len(bidders)
        last_col = 2 + 2 * n + 1
        ws.cell(1, 1, f"{document_type} COMPLIANCE COMPARISON — {label}" + (f"  (Form LGLS/FCC/{code})" if code else "")).font = Font(bold=True, size=14, color=NAVY)
        ws.cell(2, 1, f"{n} bidder(s) • {len(clauses)} clauses • generated {datetime.now():%Y-%m-%d %H:%M}").font = Font(italic=True, color="595959")
        if document_type == "GCT":
            ws.cell(3, 1, "Document summary: GCT compliance document analyzed").font = Font(italic=True, color="595959")
        hdr = 4
        for c, t in ((1, "Clause No"), (2, "Description")):
            hcell(ws, hdr, c, t)
            ws.merge_cells(start_row=hdr, start_column=c, end_row=hdr + 1, end_column=c)
            ws.cell(hdr + 1, c).border = BORDER
        for i, b in enumerate(bidders):
            c0 = 3 + 2 * i
            fill = BIDFILLS[i % len(BIDFILLS)]
            hcell(ws, hdr, c0, b["name"], fill)
            hcell(ws, hdr, c0 + 1, None, fill)
            ws.merge_cells(start_row=hdr, start_column=c0, end_row=hdr, end_column=c0 + 1)
            hcell(ws, hdr + 1, c0, "Complied / Not Complied", SUBFILL, "000000")
            hcell(ws, hdr + 1, c0 + 1, "Explanatory Note", SUBFILL, "000000")
        r = hdr + 2
        first = r
        for m in clauses:
            body(ws, r, 1, m["clause"], center=True)
            body(ws, r, 2, m["desc"])
            for i, b in enumerate(bidders):
                row = m["rows"].get(i)
                st = row["status"] if row else "Not Filled"
                if row and st == "Unclear":
                    st = "Unclear"
                body(ws, r, 3 + 2 * i, st, center=True)
                note = point_note(row["note"] if row else "Clause not present in this bidder's sheet")
                if row and row["raw"] and row["raw"].strip().lower() != st.lower():
                    note = "\n".join(x for x in (point_note(row["raw"]), note) if x)
                body(ws, r, 4 + 2 * i, note)
                set_note_row_height(ws, r, m["desc"], st, note)
            r += 1
        lastr = r - 1
        for i in range(n):
            col = L(3 + 2 * i)
            add_status_cf(ws, f"{col}{first}:{col}{lastr}", f"{col}{first}")
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 48
        for i in range(n):
            ws.column_dimensions[L(3 + 2 * i)].width = 15
            ws.column_dimensions[L(4 + 2 * i)].width = 52
        ws.row_dimensions[hdr].height = 24
        ws.row_dimensions[hdr + 1].height = 32
        finish(ws, f"C{hdr+2}")
        for i, b in enumerate(bidders):
            rows = b["data"]["rows"]
            cnt = defaultdict(int)
            for x in rows:
                cnt[x["status"]] += 1
            summary.append((label, b["name"], Path(b["path"]).name, len(clauses), cnt))
    wb.save(out)

# ----------------------------------------------------------------- SCC workbook
def build_scc(groups, out, tmpl):
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    for gkey, bidders in groups.items():
        cat = gkey[0]
        label = CATS.get(cat, bidders[0]["data"]["title"][:40] or "Unrecognised")
        n = len(bidders)
        tmap = tmpl.get(cat, {})
        # ---------- Table 1
        keys = OrderedDict()
        for b in bidders:
            for row in b["data"]["t1"]:
                k = row["clause"] or "other"
                if k not in keys:
                    keys[k] = row["req"]
        ws = wb.create_sheet(safe_sheet(f"SCC - {label}", used))
        ws.cell(1, 1, f"SCC COMPLIANCE COMPARISON — {label}").font = Font(bold=True, size=14, color=NAVY)
        ws.cell(2, 1, f"{n} bidder(s) • {len(keys)} condition(s) • generated {datetime.now():%Y-%m-%d %H:%M}").font = Font(italic=True, color="595959")
        hdr = 4
        for c, t in ((1, "Clause No"), (2, "Description")):
            hcell(ws, hdr, c, t)
            ws.merge_cells(start_row=hdr, start_column=c, end_row=hdr + 1, end_column=c)
            ws.cell(hdr + 1, c).border = BORDER
        for i, b in enumerate(bidders):
            c0 = 3 + 2 * i
            fill = BIDFILLS[i % len(BIDFILLS)]
            hcell(ws, hdr, c0, b["name"], fill)
            hcell(ws, hdr, c0 + 1, None, fill)
            ws.merge_cells(start_row=hdr, start_column=c0, end_row=hdr, end_column=c0 + 1)
            hcell(ws, hdr + 1, c0, "Complied / Not Complied", SUBFILL, "000000")
            hcell(ws, hdr + 1, c0 + 1, "Explanatory Note", SUBFILL, "000000")
        lk = [{(r["clause"] or "other"): r for r in b["data"]["t1"]} for b in bidders]
        r = first = hdr + 2
        for k, req in keys.items():
            body(ws, r, 1, "" if k == "other" else k, center=True)
            body(ws, r, 2, req)
            for i in range(n):
                row = lk[i].get(k)
                st = scc_status(row["resp"], tmap.get(k), row.get("stat", "")) if row else "Not Filled"
                if k == "other" and st == "Not Filled":
                    st = "-"          # optional free-form row: blank is fine
                body(ws, r, 3 + 2 * i, st, center=True)
                note = point_note(row["resp"] if row else "")
                if row and len(clean(row.get("stat", ""))) > 14:          # bidder's own remark written in the status column
                    note = "\n".join(x for x in (note, point_note(row["stat"])) if x)
                body(ws, r, 4 + 2 * i, note)
                set_note_row_height(ws, r, req, st, note)
            r += 1
        lastr = r - 1
        for i in range(n):
            col = L(3 + 2 * i)
            add_status_cf(ws, f"{col}{first}:{col}{lastr}", f"{col}{first}")
        # ---------- Table 2: amendments / new additions proposed by each bidder
        r += 2
        ws.cell(r, 1, "TABLE 2 — AMENDMENTS / NEW ADDITIONS TO THE GENERAL CONDITIONS").font = Font(bold=True, size=12, color=NAVY)
        r += 1
        for c, t in ((1, "Ref"), (2, "Clause Number under the General Conditions of Contract")):
            hcell(ws, r, c, t)
        for i, b in enumerate(bidders):
            c0 = 3 + 2 * i
            fill = BIDFILLS[i % len(BIDFILLS)]
            hcell(ws, r, c0, b["name"] + " — Status", fill)
            hcell(ws, r, c0 + 1, "Amendment / New addition", fill)
        r += 1
        t2keys = OrderedDict()
        for b in bidders:
            for row in b["data"]["t2"]:
                if clean(row["clause"]):
                    t2keys.setdefault(_nd(row["clause"]), (row["ref"], clean_multiline(row["clause"])))
        t2l = [{_nd(x["clause"]): x for x in b["data"]["t2"] if clean(x["clause"])} for b in bidders]
        t2first = r
        for k, (ref, clause) in t2keys.items():
            body(ws, r, 1, ref, center=True)
            body(ws, r, 2, clause)
            for i in range(n):
                x = t2l[i].get(k)
                if x:
                    s0 = norm_status(x["remarks"]) if clean(x["remarks"]) else "Provided"
                    body(ws, r, 3 + 2 * i, s0 if s0 != "Unclear" else "Provided", center=True)
                    note = point_note(x["amend"])
                    if len(clean(x["remarks"])) > 14:
                        note = "\n".join(y for y in (note, point_note(x["remarks"])) if y)
                    body(ws, r, 4 + 2 * i, note)
                else:
                    body(ws, r, 3 + 2 * i, "-", center=True)
                    body(ws, r, 4 + 2 * i, "")
                set_note_row_height(ws, r, clause, note if x else "")
            r += 1
        if not t2keys:
            body(ws, r, 2, "No amendments or new additions proposed by any bidder")
            r += 1
        else:
            for i in range(n):
                col = L(3 + 2 * i)
                add_status_cf(ws, f"{col}{t2first}:{col}{r-1}", f"{col}{t2first}")
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 48
        for i in range(n):
            ws.column_dimensions[L(3 + 2 * i)].width = 15
            ws.column_dimensions[L(4 + 2 * i)].width = 52
        finish(ws, f"C{hdr+2}")
    wb.save(out)

# ----------------------------------------------------------------- collection
def collect(inputs, tmp, problems):
    """-> [(bidder_name, readable_path (.xlsx/.docx), original_path)]. .doc/.xls/.pdf are converted first."""
    try:
        from Backend.Technical import convert
    except ModuleNotFoundError:
        from Technical import convert
    items = []
    for spec in inputs:
        name = None
        if "=" in spec and not Path(spec).exists():
            name, spec = spec.split("=", 1)
        p = Path(spec).expanduser()
        if not p.exists():
            sys.exit(f"Not found: {p}")
        if p.suffix.lower() == ".zip":
            d = Path(tempfile.mkdtemp(dir=tmp))
            with zipfile.ZipFile(p) as z:
                z.extractall(d)
            kids = [x for x in d.iterdir() if x.name != "__MACOSX"]
            while len(kids) == 1 and kids[0].is_dir():          # zip wrapped in one top folder: bidders are inside it
                d = kids[0]
                kids = list(d.iterdir())
            p, root = d, d
        else:
            root = p if p.is_dir() else None
        files = [p] if p.is_file() else sorted(x for x in p.rglob("*") if x.is_file())
        for f in files:
            if f.name.startswith(("~$", ".")) or "__MACOSX" in f.parts:
                continue
            if f.suffix.lower() not in convert.SUPPORTED:
                continue
            bname = name or derive_bidder_name(f, root)
            try:
                rp = convert.normalise(f, Path(tempfile.mkdtemp(dir=tmp)))
            except Exception as e:
                problems.append(f"{f.name} ({bname}): {e}")
                continue
            items.append((bname, rp, f))
    return items

def _tender_suffix(tenders):
    safe = [
        re.sub(r'[\\/:*?"<>|]+', "-", str(tender).strip())
        for tender in tenders
        if str(tender).strip()
    ]
    return f"_{'-'.join(OrderedDict.fromkeys(safe))}" if safe else ""


def run(inputs, out, mode=None):
    """Process bidder files; returns dict(gcc=Path|None, scc=Path|None, bidders=[...], problems=[...]).
    mode: None = everything, "bids" = GCC/SCC only, "eval" = technical / sole-source evaluation reports only."""
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    tmp = tempfile.mkdtemp()
    gcc_groups, gct_groups, scc_groups = OrderedDict(), OrderedDict(), OrderedDict()
    seen = defaultdict(int)
    res = dict(gcc=None, gct=None, scc=None, gcc_bidders=[], gct_bidders=[], scc_bidders=[], document_summaries=[],
               tenders=[],
               problems=[], eval=None, evals=[])
    evals = []
    for bname, f, orig in collect(inputs, tmp, res["problems"]):
        original_bname = bname
        try:
            text = document_text(f)
            # ── Step 1: Try the document HEADER (letterhead / company title) ──
            document_name = extract_header_name(text, bname)
            # ── Step 2: Fallback – Official Frank / stamp field in document text ──
            if document_name == bname:
                document_name = extract_official_frank_name(text, bname)
            # ── Step 3: Fallback – OCR the embedded stamp image ──
            ocr_text = ""
            if document_name == bname and has_official_frank_label(text):
                ocr_text = ocr_embedded_images(orig)
                if ocr_text:
                    document_name = extract_official_frank_ocr_name(ocr_text, bname)
        except (OSError, ValueError):
            ocr_text = ""
            text = ""
            document_name = bname
        kind = "GCC" if f.suffix.lower().startswith(".xls") else "SCC"
        if re.search(r"(?i)general\s+conditions\s+of\s+tender|compliance\s+statement\s+for\s+general\s+conditions\s+of\s+tender", text):
            kind = "GCT"
        if mode == "eval" and kind in ("GCC", "GCT"):
            res["problems"].append(f"{orig.name}: not an evaluation report (compliance documents belong in the GCC / SCC / GCT tab)")
            continue
        if kind == "SCC":                       # a .docx may be a TET / SSET evaluation report instead of an SCC form
            try:
                from Backend.Technical import evalreport
            except ModuleNotFoundError:
                from Technical import evalreport
            try:
                ev = evalreport.parse_eval(f)
                if mode == "bids":
                    res["problems"].append(f"{orig.name}: this is an evaluation report - upload it in the Technical Evaluation tab")
                    continue
                evals.append((document_name, ev))
                continue
            except ValueError as e:
                if mode == "eval":
                    res["problems"].append(f"{orig.name}: {e}")
                    continue
                if "not an evaluation report" not in str(e):
                    res["problems"].append(f"{orig.name}: {e}")
                    continue
        bname = document_name
        try:
            data = parse_gcc(f) if kind in ("GCC", "GCT") else parse_scc(f)
        except Exception as e:
            res["problems"].append(f"{orig.name} ({bname}): {e}")
            continue
        seen[(kind, bname)] += 1
        if seen[(kind, bname)] > 1:
            bname = f"{bname} ({seen[(kind, bname)]})"
        if kind == "GCC":
            key = (data["cat"], data["code"] or "") if data["cat"] else (4, "075")   # no form code -> assume Goods (Local)
            gcc_groups.setdefault(key, []).append(dict(name=bname, path=str(orig), data=data))
        elif kind == "GCT":
            key = (data["cat"], data["code"] or "") if data["cat"] else ("GCT", "")
            gct_groups.setdefault(key, []).append(dict(name=bname, path=str(orig), data=data))
        else:
            key = (data["cat"],) if data["cat"] else (None, data["title"])
            scc_groups.setdefault(key, []).append(dict(name=bname, path=str(orig), data=data))
        summary = document_summary(
            f, bname, kind, orig.name, ocr_text,
            document_has_embedded_images(orig),
            ("Header" if extract_header_name(text, bname) != bname
             else "Official frank / stamp / seal")
            if document_name != original_bname
            else "filename or folder fallback",
        )
        res["document_summaries"].append(summary)
        tender = next((point.split(":", 1)[1].strip() for point in summary["points"]
                       if point.lower().startswith("tender number:")), "")
        if tender and tender != "Not found" and tender not in res["tenders"]:
            res["tenders"].append(tender)
    for g in (gcc_groups, gct_groups, scc_groups):        # same bidder order in every sheet, whatever order files arrived in
        for lst in g.values():
            lst.sort(key=lambda b: b["name"].lower())
    if gcc_groups:
        res["gcc"] = out / f"GCC_Compliance_Comparison{_tender_suffix(res['tenders'])}.xlsx"
        build_gcc(gcc_groups, res["gcc"])
        res["gcc_bidders"] = [b["name"] for v in gcc_groups.values() for b in v]
    if gct_groups:
        res["gct"] = out / f"GCT_Compliance_Comparison{_tender_suffix(res['tenders'])}.xlsx"
        build_gcc(gct_groups, res["gct"], document_type="GCT")
        res["gct_bidders"] = [b["name"] for v in gct_groups.values() for b in v]
    if scc_groups:
        res["scc"] = out / f"SCC_Compliance_Comparison{_tender_suffix(res['tenders'])}.xlsx"
        build_scc(scc_groups, res["scc"], template_scc_text())
        res["scc_bidders"] = [b["name"] for v in scc_groups.values() for b in v]
    if evals:
        from Technical import evalreport
        for _, ev in evals:
            tender = ev["meta"].get("Tender No.", "")
            if tender and tender not in res["tenders"]:
                res["tenders"].append(tender)
        res["eval"] = out / f"Technical_Evaluation_Compliance{_tender_suffix(res['tenders'])}.xlsx"
        evalreport.build_eval_workbook(evals, res["eval"])
        res["evals"] = [evalreport.eval_to_json(lbl, ev) for lbl, ev in evals]
    return res

def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("inputs", nargs="+", help="bidder files, folders or .zip (optionally 'Name=path')")
    ap.add_argument("-o", "--out", default="output", help="output folder (default: ./output)")
    a = ap.parse_args()
    res = run(a.inputs, a.out)
    if res["gcc"]:
        print(f"GCC: {len(res['gcc_bidders'])} bidder(s) -> {res['gcc']}")
    if res["gct"]:
        print(f"GCT: {len(res['gct_bidders'])} bidder(s) -> {res['gct']}")
    if res["scc"]:
        print(f"SCC: {len(res['scc_bidders'])} bidder(s) -> {res['scc']}")
    if res["problems"]:
        print("\nCould not read:")
        for p in res["problems"]:
            print("  -", p)
    if res["eval"]:
        print(f"Evaluation reports: {len(res['evals'])} -> {res['eval']}")
    if not res["gcc"] and not res["gct"] and not res["scc"] and not res["eval"]:
        sys.exit("No readable GCC / GCT / SCC / evaluation files found (.xlsx .xls .docx .doc .pdf).")

if __name__ == "__main__":
    main()
