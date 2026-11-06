
import re, shutil, subprocess, tempfile
from pathlib import Path

SUPPORTED = (".xlsx", ".xlsm", ".xls", ".docx", ".doc", ".pdf")


def _soffice(src, fmt, outdir):
    exe = shutil.which("soffice") or shutil.which("libreoffice")
    if not exe:
        raise RuntimeError("LibreOffice is needed to read .doc / .xls files (not found on this computer)")
    prof = tempfile.mkdtemp()
    try:
        subprocess.run([exe, f"-env:UserInstallation=file://{prof}", "--headless", "--convert-to", fmt,
                        "--outdir", str(outdir), str(src)], check=True, capture_output=True, timeout=180)
    except subprocess.TimeoutExpired:
        raise RuntimeError("converting the file timed out")
    except subprocess.CalledProcessError as e:
        raise RuntimeError("could not convert the file: " + e.stderr.decode(errors="ignore")[:200])
    finally:
        shutil.rmtree(prof, ignore_errors=True)
    out = Path(outdir) / (Path(src).stem + "." + fmt)
    if not out.exists():
        raise RuntimeError("conversion produced no output")
    return out


def _clean(s):
    return re.sub(r"\s+", " ", (s or "").replace("\x00", "")).strip()


def _pdf_content(path):
    """-> (text_lines_outside_tables, tables) ; tables = list of rows (list of str), page order."""
    import pdfplumber
    lines, tables = [], []
    with pdfplumber.open(path) as pdf:
        for page in pdf.pages:
            found = page.find_tables()
            boxes = [t.bbox for t in found]
            def outside(o, boxes=boxes):
                if o.get("object_type") != "char":
                    return True
                x, y = (o["x0"] + o["x1"]) / 2, (o["top"] + o["bottom"]) / 2
                return not any(b[0] <= x <= b[2] and b[1] <= y <= b[3] for b in boxes)
            txt = page.filter(outside).extract_text() or ""
            lines += [l.strip() for l in txt.splitlines() if l.strip()]
            for t in found:
                rows = [[(c or "").replace("\x00", "").strip() for c in r] for r in t.extract()]
                if rows:
                    tables.append(rows)
    return lines, tables


def _pdf_to_xlsx(lines, tables, out):
    import openpyxl
    gcc, ncol = [], None
    for t in tables:                       # header table, then its continuation tables on later pages
        if not t:
            continue
        if _clean(t[0][0]).lower().startswith("clause") and "complied" in " ".join(t[0]).lower():
            gcc.append(t)
            ncol = len(t[0])
        elif ncol and len(t[0]) == ncol:
            gcc.append(t)
        else:
            ncol = None
    if not gcc:
        return None
    wb = openpyxl.Workbook()
    ws = wb.active
    r = 1
    for l in lines[:15]:
        ws.cell(r, 1, l)
        r += 1
    for t in gcc:
        for row in t:
            if not any(_clean(c) for c in row):
                continue
            if r > 1 and _clean(row[0]).lower().startswith("clause") and ws.cell(r - 1, 1).value is not None \
                    and any(_clean(ws.cell(i, 1).value).lower().startswith("clause") for i in range(1, r)):
                continue                                      # header repeated on later pages
            for c, v in enumerate(row[:4], 1):
                ws.cell(r, c, v)
            r += 1
    # signature block / trailing text so the reader stops cleanly
    wb.save(out)
    return out


def _pdf_to_docx(lines, tables, out):
    import docx
    d = docx.Document()
    # the readers identify a document by its first paragraph, but PDFs often start with a form code / header
    k = next((i for i, l in enumerate(lines) if re.search(r"(?i)evaluation report|conditions of contract", l)), 0)
    lines = lines[k:k + 1] + lines[:k] + lines[k + 1:]
    for l in lines:
        d.add_paragraph(l)
    merged = []                            # a table split over a page break arrives as separate tables: rejoin
    for rows in tables:
        prev = merged[-1] if merged else None
        same = prev and len(rows[0]) >= 3 and len(prev[0]) == len(rows[0])
        known = same and re.search(r"(?i)details required|amendment|evaluation area", " ".join(prev[0]))
        if same and (known or len(prev) == 1 or [_clean(c) for c in rows[0]] == [_clean(c) for c in prev[0]]):
            if [_clean(c) for c in rows[0]] == [_clean(c) for c in prev[0]]:
                rows = rows[1:]            # repeated header row
            prev.extend(rows)
        else:
            merged.append(list(rows))
    for rows in merged:
        n = max(len(r) for r in rows)
        t = d.add_table(rows=0, cols=n)
        for r in rows:
            cells = t.add_row().cells
            for i, v in enumerate(r):
                cells[i].text = v
        d.add_paragraph("")
    d.save(out)
    return out


def normalise(path, outdir):
    """Return a .xlsx/.docx path the readers can open (the original if it already is one)."""
    path, outdir = Path(path), Path(outdir)
    s = path.suffix.lower()
    if s in (".xlsx", ".xlsm", ".docx"):
        return path
    outdir.mkdir(parents=True, exist_ok=True)
    if s == ".doc":
        return _soffice(path, "docx", outdir)
    if s == ".xls":
        return _soffice(path, "xlsx", outdir)
    if s == ".pdf":
        lines, tables = _pdf_content(path)
        if not lines and not tables:
            raise RuntimeError("no readable text in this PDF (scanned image? export it from the source file instead)")
        out = _pdf_to_xlsx(lines, tables, outdir / (path.stem + ".xlsx"))
        return out or _pdf_to_docx(lines, tables, outdir / (path.stem + ".docx"))
    raise RuntimeError("unsupported file type")
