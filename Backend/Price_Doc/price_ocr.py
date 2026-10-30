import re
import numpy as np

OCR_DPI = 300


# ----------------------------------------------------------------------------- loading
def render_pdf_page(path, page_no, dpi=OCR_DPI):
    from pdf2image import convert_from_path
    return convert_from_path(path, dpi=dpi, first_page=page_no, last_page=page_no)[0]


def load_images(path):
    from PIL import Image, ImageSequence
    im = Image.open(path)
    return [f.convert("RGB").copy() for f in ImageSequence.Iterator(im)]


def _prep(pil_img):
    """PIL -> grayscale numpy, upscaled if small, deskewed."""
    import cv2
    g = cv2.cvtColor(np.array(pil_img.convert("RGB")), cv2.COLOR_RGB2GRAY)
    h, w = g.shape
    if w < 2000:
        k = min(3.0, 2400 / w)
        g = cv2.resize(g, None, fx=k, fy=k, interpolation=cv2.INTER_CUBIC)
    return _deskew(_orient(g))


def _orient(g):
    """Fix pages scanned sideways / upside-down using Tesseract's orientation detection."""
    import pytesseract
    try:
        osd = pytesseract.image_to_osd(g, config="--psm 0")
        rot = int(re.search(r"Rotate:\s*(\d+)", osd).group(1))
        conf = float(re.search(r"Orientation confidence:\s*([\d.]+)", osd).group(1))
    except Exception:
        return g
    if rot and conf >= 1.5:
        return np.ascontiguousarray(np.rot90(g, k=-(rot // 90)))
    return g


def _deskew(g):
    import cv2
    small = cv2.resize(g, None, fx=0.25, fy=0.25, interpolation=cv2.INTER_AREA)
    bw = (small < 160).astype(np.uint8)
    best, best_a = -1, 0.0
    for a in np.arange(-4, 4.01, 0.5):
        M = cv2.getRotationMatrix2D((small.shape[1] / 2, small.shape[0] / 2), a, 1)
        r = cv2.warpAffine(bw, M, (small.shape[1], small.shape[0]), flags=cv2.INTER_NEAREST)
        score = np.var(r.sum(axis=1))
        if score > best:
            best, best_a = score, a
    if abs(best_a) < 0.2:
        return g
    M = cv2.getRotationMatrix2D((g.shape[1] / 2, g.shape[0] / 2), best_a, 1)
    return cv2.warpAffine(g, M, (g.shape[1], g.shape[0]), flags=cv2.INTER_CUBIC, borderValue=255)


# ----------------------------------------------------------------------------- tesseract helpers
def _ocr_text(img, psm=6, whitelist=None):
    """-> (text, mean_conf) for one image crop."""
    import pytesseract
    cfg = f"--psm {psm}"
    if whitelist:
        cfg += f" -c tessedit_char_whitelist={whitelist}"
    d = pytesseract.image_to_data(img, config=cfg, output_type=pytesseract.Output.DICT)
    lines, confs = {}, []
    for i, t in enumerate(d["text"]):
        if not t.strip():
            continue
        key = (d["block_num"][i], d["par_num"][i], d["line_num"][i])
        lines.setdefault(key, []).append(t)
        try:
            c = float(d["conf"][i])
        except ValueError:
            c = -1
        if c >= 0:
            confs.append(c)
    text = "\n".join(" ".join(v) for _, v in sorted(lines.items()))
    return text, (sum(confs) / len(confs) if confs else None)


# ----------------------------------------------------------------------------- A. ruled tables
def _grid(g):
    import cv2
    h, w = g.shape
    bw = cv2.adaptiveThreshold(g, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C, cv2.THRESH_BINARY_INV, 21, 12)
    hor = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(30, w // 35), 1)))
    ver = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(55, h // 50))))
    grid = cv2.dilate(cv2.bitwise_or(hor, ver), np.ones((3, 3), np.uint8))
    return grid


def _cluster(vals, tol):
    vals = sorted(vals)
    groups = []
    for v in vals:
        if groups and v - groups[-1][-1] <= tol:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [sum(g) / len(g) for g in groups]


def ruled_tables(g):
    """-> list of {"rows":[[text]], "conf":[[float]], "bbox":(x,y,w,h)} or []"""
    import cv2
    h, w = g.shape
    grid = _grid(g)
    inv = cv2.bitwise_not(grid)
    n, labels, stats, _ = cv2.connectedComponentsWithStats(inv, connectivity=4)
    cells = []
    for i in range(1, n):
        x, y, cw, ch, area = stats[i]
        if area > 0.4 * w * h or cw > 0.97 * w or ch > 0.9 * h or cw < 25 or ch < 18:
            continue
        cells.append((int(x), int(y), int(cw), int(ch)))
    if len(cells) < 4:
        return []
    # --- group cells into tables via the outer contours of the grid
    closed = cv2.dilate(grid, np.ones((9, 9), np.uint8))
    cnts, _ = cv2.findContours(closed, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    boxes = [cv2.boundingRect(c) for c in cnts]
    boxes = [b for b in boxes if b[2] > 0.25 * w and b[3] > 40]
    boxes.sort(key=lambda b: b[1])
    out = []
    for (bx, by, bw_, bh) in boxes:
        mine = [c for c in cells if bx - 5 <= c[0] + c[2] / 2 <= bx + bw_ + 5 and by - 5 <= c[1] + c[3] / 2 <= by + bh + 5]
        if len(mine) < 4:
            continue
        tops = _cluster([c[1] for c in mine], max(8, int(np.median([c[3] for c in mine]) * 0.3)))
        lefts = _cluster([c[0] for c in mine], 14)
        if len(lefts) < 2:
            continue
        rows = [[""] * len(lefts) for _ in tops]
        conf = [[None] * len(lefts) for _ in tops]
        geom = [[None] * len(lefts) for _ in tops]
        for (x, y, cw, ch) in mine:
            r = int(np.argmin([abs(y - t) for t in tops]))
            c = int(np.argmin([abs(x - l) for l in lefts]))
            text, cf = _read_cell(g, x, y, cw, ch)
            geom[r][c] = (x, y, cw, ch)
            if text:
                rows[r][c] = (rows[r][c] + "\n" + text).strip()
                conf[r][c] = cf
        raw = [list(r) for r in rows]                       # verbatim first reading, kept for the 'Raw Tables' sheet
        _numeric_pass(g, rows, conf, geom)
        out.append({"rows": rows, "conf": conf, "bbox": (bx, by, bw_, bh), "raw_rows": raw})
    return out


def _read_cell(g, x, y, cw, ch, whitelist=None, psm=6):
    import cv2
    pad = 6                                                                  # stay clear of the ruling lines
    crop = g[y + pad:y + ch - pad, x + pad:x + cw - pad].copy()
    if crop.size == 0 or (crop < 140).mean() < 0.004:                       # empty / blank cell
        return "", None
    crop = cv2.copyMakeBorder(crop, 14, 14, 14, 14, cv2.BORDER_CONSTANT, value=255)
    return _ocr_text(crop, psm=psm if ch > 80 else 7, whitelist=whitelist)


def _tight_digits(g, x, y, cw, ch):
    """Crop a cell tightly around its ink (specks removed), scale the glyphs to ~70 px and pad - this is what lets
    Tesseract read a lone '2' that sits in the middle of a large empty cell."""
    import cv2
    pad = 6
    crop = g[y + pad:y + ch - pad, x + pad:x + cw - pad]
    if crop.size == 0:
        return None
    _, bw = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    n, lab, st, _ = cv2.connectedComponentsWithStats(bw, connectivity=8)
    keep = [i for i in range(1, n) if st[i, cv2.CC_STAT_AREA] >= 25 and st[i, cv2.CC_STAT_HEIGHT] >= 8]
    if not keep:
        return None
    x0 = min(st[i, cv2.CC_STAT_LEFT] for i in keep)
    x1 = max(st[i, cv2.CC_STAT_LEFT] + st[i, cv2.CC_STAT_WIDTH] for i in keep)
    y0 = min(st[i, cv2.CC_STAT_TOP] for i in keep)
    y1 = max(st[i, cv2.CC_STAT_TOP] + st[i, cv2.CC_STAT_HEIGHT] for i in keep)
    glyph = crop[max(0, y0 - 3):y1 + 3, max(0, x0 - 3):x1 + 3]
    k = 70.0 / max(glyph.shape[0], 1)
    glyph = cv2.resize(glyph, None, fx=k, fy=k, interpolation=cv2.INTER_CUBIC)
    return cv2.copyMakeBorder(glyph, 30, 30, 30, 30, cv2.BORDER_CONSTANT, value=255)


def _reread_number(g, geo):
    """Second opinion for a numeric cell: digits-only vocabulary on the tightly cropped ink.
    Returns (text, conf) of the first reading that parses as a number."""
    from price_extractor import parse_num
    img = _tight_digits(g, *geo)
    if img is None:
        return None, None
    for psm in (7, 6, 8, 10):
        t, c = _ocr_text(img, psm=psm, whitelist="0123456789.,%-/")
        t = (t or "").replace("\n", " ").strip()
        if t and parse_num(t, True) is not None:
            return t, c
    return None, None


_WORDY = re.compile(r"(?i)^(?:included?|inclusive|nil|n/?a|none|free(?:\s+of\s+charge)?|foc|exempt\w*|excluded?|"
                    r"as\s+(?:per|above|required|applicable)\b.*|not\s+(?:applicable|required|quoted).*|tbd|by\s+others|optional)$")


def _numeric_pass(g, rows, conf, geom):
    """Every numeric cell gets a digits-only second reading.
      - first reading was not a number (e.g. 'tae', 'RS', '=')   -> use the second reading (rescue)
      - both readings are numbers but disagree                    -> keep the first, mark the cell for review
    Words such as 'Included', 'NIL', 'Free of charge' are left untouched."""
    from price_extractor import NUM_FIELDS, best_header, map_header, parse_num
    hi, hcells, span = best_header(rows)
    if hi is None:
        return
    mapping, _ = map_header(hcells)
    numcols = [c for c, f in mapping.items() if f in NUM_FIELDS]
    for r in range(hi + span, len(rows)):
        for c in numcols:
            cell, geo = rows[r][c], geom[r][c]
            if geo is None or not cell or _WORDY.match(cell.strip()):
                continue                                  # empty, or a genuine word ("Included", "Free of charge")
            first = parse_num(cell, True)
            t2, c2 = _reread_number(g, geo)
            if t2 is None:
                continue
            v2 = parse_num(t2, True)
            if first is None:
                rows[r][c], conf[r][c] = t2, min(c2 if c2 is not None else 70, 80)   # rescued -> still shown for review
            elif abs(first - v2) > 1e-9:
                conf[r][c] = min(conf[r][c] if conf[r][c] is not None else 100, 50)  # disagreement -> flag, keep first


# ----------------------------------------------------------------------------- B. unruled tables
def positional_tables(words):
    """words: [{"text","x0","x1","top","bottom","conf"}] -> tables for tables WITHOUT ruling lines.
    Finds the header line by keywords, then derives the columns from the vertical white-space gutters."""
    from price_extractor import header_score
    if not words:
        return []
    ws = sorted(words, key=lambda w: (w["top"], w["x0"]))
    lh = float(np.median([w["bottom"] - w["top"] for w in ws])) or 10.0
    lines, cur = [], [ws[0]]
    for w in ws[1:]:
        if abs(w["top"] - cur[0]["top"]) <= lh * 0.6:
            cur.append(w)
        else:
            lines.append(sorted(cur, key=lambda a: a["x0"]))
            cur = [w]
    lines.append(sorted(cur, key=lambda a: a["x0"]))

    def phrases(line, gap):
        out = []
        for w in line:
            if out and w["x0"] - out[-1]["x1"] <= gap:
                out[-1]["text"] += " " + w["text"]
                out[-1]["x1"] = w["x1"]
            else:
                out.append(dict(w))
        return out

    hi = None
    for i, line in enumerate(lines):
        ph = phrases(line, lh * 1.0)
        if len(ph) >= 3 and header_score([p["text"] for p in ph]):
            hi = i
            break
    if hi is None:
        return []
    body = []
    for line in lines[hi:]:
        txt = " ".join(w["text"] for w in line)
        if line is not lines[hi] and re.search(r"name\s*&?\s*signature|official\s*frank|^note\b|^\W*signature", txt, re.I):
            break
        body.append(line)
    # ---- columns from gutters
    x_min = min(w["x0"] for l in body for w in l)
    x_max = max(w["x1"] for l in body for w in l)
    occ = np.zeros(int(x_max - x_min) + 3, bool)
    for l in body:
        for w in l:
            occ[int(w["x0"] - x_min):int(w["x1"] - x_min) + 1] = True
    hph = phrases(lines[hi], lh * 1.6)          # header phrases: a gutter must not split one of them
    thr, cuts, start = max(3, int(lh * 0.8)), [], None
    for i, v in enumerate(occ):
        if not v and start is None:
            start = i
        elif v and start is not None:
            if i - start >= thr and start > 0:
                c = x_min + (start + i) / 2
                if not any(p["x0"] < c < p["x1"] for p in hph):
                    cuts.append(c)
            start = None

    def to_rows(cut_list):
        rows, conf = [], []
        for l in body:
            cells = [[] for _ in range(len(cut_list) + 1)]
            for w in l:
                cells[sum((w["x0"] + w["x1"]) / 2 > c for c in cut_list)].append(w)
            rows.append([" ".join(w["text"] for w in c) for c in cells])
            conf.append([(float(np.mean([w["conf"] for w in c if w["conf"] is not None]))
                          if any(w["conf"] is not None for w in c) else None) for c in cells])
        return rows, conf
    rows, conf = to_rows(cuts)
    if not header_score(rows[0]) and not (len(rows) > 1 and header_score([(a + " " + b).strip() for a, b in zip(rows[0], rows[1])])):
        # gutters did not line up with the header -> use header phrases
        ph = phrases(body[0], lh * 1.0)
        centres = [(p["x0"] + p["x1"]) / 2 for p in ph]
        rows, conf = to_rows([(centres[i] + centres[i + 1]) / 2 for i in range(len(centres) - 1)])
    rows, conf = _group_rows(rows, conf)
    return [{"rows": rows, "conf": conf, "bbox": None}]


def _group_rows(rows, conf):
    """Without cell borders a wrapped cell becomes several text lines. If the table has an S/No column,
    every line up to the next S/No (or Total line) belongs to the same row - join them cell by cell."""
    from price_extractor import TOTAL_RX, best_header, map_header
    hi, hcells, span = best_header(rows)
    if hi is None:
        return rows, conf
    mapping, _ = map_header(hcells)
    sno = next((c for c, f in mapping.items() if f == "sno"), None)
    desc = next((c for c, f in mapping.items() if f == "desc"), None)
    if sno is None:
        return rows, conf
    starts = lambda r: bool(re.fullmatch(r"[A-Za-z]?\d{1,3}[\.\)]?", r[sno].strip())) or \
        bool(desc is not None and TOTAL_RX.search(r[desc]) and not r[sno].strip())
    body0 = hi + span
    if sum(starts(r) for r in rows[body0:]) < 1:
        return rows, conf
    out, oc = list(rows[:body0]), list(conf[:body0])
    for r, c in zip(rows[body0:], conf[body0:]):
        if starts(r) or len(out) == body0:
            out.append(list(r)); oc.append(list(c))
        else:
            for i, v in enumerate(r):
                if v:
                    out[-1][i] = (out[-1][i] + " " + v).strip()
    return out, oc


# ----------------------------------------------------------------------------- entry point
def ocr_image(pil_img, page_no):
    """Whole page -> (tables, text_lines)."""
    import pytesseract
    g = _prep(pil_img)
    tables = ruled_tables(g)
    for t in tables:
        t["page"], t["method"] = page_no, "ocr"
    d = pytesseract.image_to_data(g, config="--psm 6", output_type=pytesseract.Output.DICT)
    words, lines = [], {}
    for i, t in enumerate(d["text"]):
        if not t.strip():
            continue
        try:
            c = float(d["conf"][i])
        except ValueError:
            c = None
        words.append({"text": t, "x0": d["left"][i], "x1": d["left"][i] + d["width"][i], "top": d["top"][i],
                      "bottom": d["top"][i] + d["height"][i], "conf": c if c is not None and c >= 0 else None})
        lines.setdefault((d["block_num"][i], d["par_num"][i], d["line_num"][i]), []).append(t)
    text = [" ".join(v) for _, v in sorted(lines.items())]
    if not tables:                                    # no ruling lines found -> use word positions
        for t in positional_tables(words):
            t["page"], t["method"] = page_no, "ocr"
            tables.append(t)
    return tables, text
