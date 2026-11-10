"""Price schedule reader.

The SLT "Sample Price Schedule" has two tables:
    (A) FOR LOCAL SUPPLIES   S/No | Description | Qty | Unit Price (excl. taxes) | VAT % | NBT % | Other taxes % | Unit Price (incl. taxes)
    (B) FOR FOREIGN SUPPLIES S/No | Description | Qty | Country of Origin | FOB/FCA Unit Price | FOB/FCA Total Cost |
                             Total Insurance | Total Freight | Total CIF/CIP Cost | Local Agent's Commission
A bidder fills in one of them (or both).  Columns are found by their header words, so blank columns,
re-ordered columns and PDF / Word / Excel layouts all work.

Quoted total:   local   = sum( Qty x Unit Price incl. taxes )
                foreign = sum( Total CIF/CIP Cost )   (+ Local Agent's Commission if switched on in the Excel)
Anything the bidder left blank but that can be calculated is calculated, as an editable formula shaded amber.
"""
import re, sys, tempfile
from collections import OrderedDict, defaultdict
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

try:
    from ..goods_services import consolidate as C
except ImportError:
    try:
        from goods_services import consolidate as C
    except ImportError:                        # pragma: no cover - direct CLI use
        from Backend.goods_services import consolidate as C

try:
    import docx
except ImportError:                                   # pragma: no cover
    docx = None

STATUSES = ["To be evaluated", "Evaluated", "Not responsive"]
AMBER = PatternFill("solid", fgColor="FFEB9C")
INPUT = PatternFill("solid", fgColor="FFF2CC")

# --------------------------------------------------------------------------- reading
def num(v):
    if v is None:
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).replace("\n", " ")
    s = re.sub(r"(?i)\b(rs|lkr|usd|eur|gbp|inr|cny|jpy|aud|sgd)\b\.?|[$€£%,]|\s", "", s)
    if not s or not re.fullmatch(r"-?\d+(\.\d+)?", s):
        return None
    return float(s)


def _text(v):
    text = C.clean_multiline(v) if v is not None else ""
    return re.sub(r"\(cid:415\)", "ti", re.sub(r"\(cid:414\)", "tf", text, flags=re.IGNORECASE),
                  flags=re.IGNORECASE)


def read_grids(path):
    """-> (grids, text).  grid = list of rows (list of cell values);  text = all running text of the file."""
    path = Path(path)
    grids, text = [], []
    if path.suffix.lower() in (".xlsx", ".xlsm"):
        wb = openpyxl.load_workbook(path, data_only=True)
        for ws in wb.worksheets:
            g = [[c for c in row] for row in ws.iter_rows(values_only=True)]
            grids.append(g)
            text += [str(c) for row in g for c in row if isinstance(c, str)]
    else:
        d = docx.Document(path)
        text += [p.text for p in d.paragraphs]
        for t in d.tables:
            g = C._rows(t)
            grids.append(g)
            text += [c for row in g for c in row]
    return grids, "\n".join(text)


# --------------------------------------------------------------------------- locating the two tables
LOCAL_FIELDS = ["sno", "desc", "qty", "unit_excl", "vat", "sscl", "other", "unit_incl"]
FOREIGN_FIELDS = ["sno", "desc", "qty", "origin", "fob_unit", "fob_total", "insurance", "freight", "cif", "commission"]


def _compact(t):
    return re.sub(r"\s+", "", t.lower())            # PDF cells wrap inside words ("Insura/nce")


def _hdr_norm(row):
    return _compact(" ".join(_text(c) for c in row))


def _is_local_header(row):
    h = _hdr_norm(row)
    return "unitprice" in h and ("vat" in h or "nbt" in h)


def _is_foreign_header(row):
    h = _hdr_norm(row)
    return ("fob" in h or "cif" in h) and ("freight" in h or "insurance" in h or "origin" in h or "cif" in h)


def _local_field(t):
    t = _compact(t)
    if re.search(r"s/no|^no\.?$|serial|itemno", t):
        return "sno"
    if "description" in t or t == "item":
        return "desc"
    if t.startswith("qty") or t == "quantity":
        return "qty"
    if "vat" in t:
        return "vat"
    if "sscl" in t:
        return "sscl"
    if "other" in t:
        return "other"
    if "price" in t:
        return "unit_incl" if "including" in t or "incl" in t else "unit_excl"
    return None


def _foreign_field(t):
    t = _compact(t)
    if re.search(r"s/no|^no\.?$|serial|itemno", t):
        return "sno"
    if "description" in t or "descrip" in t:
        return "desc"
    if t.startswith("qty") or t == "quantity":
        return "qty"
    if "origin" in t or "country" in t:
        return "origin"
    if "commission" in t or "agent" in t:
        return "commission"
    if "cif" in t or "cip" in t:
        return "cif"
    if "insurance" in t:
        return "insurance"
    if "freight" in t:
        return "freight"
    if "fob" in t or "fca" in t:
        return "fob_total" if "total" in t else "fob_unit"
    return None


SECOND_LINE = re.compile(r"(?i)excluding|including|specify|taxes|\(%\)|\(rs")


def _header_map(grid, i, fieldfn, order):
    """Column -> field for the header starting at row i; returns (map, first_data_row)."""
    rows = [grid[i]]
    j = i + 1
    while j < len(grid):                                   # a header printed over two rows
        cells = [_text(c) for c in grid[j] if _text(c)]
        if cells and all(SECOND_LINE.search(c) and num(c) is None for c in cells):
            rows.append(grid[j])
            j += 1
        else:
            break
    ncol = max(len(r) for r in rows)
    mapping = {}
    for c in range(ncol):
        t = " ".join(_text(r[c]) for r in rows if c < len(r))
        f = fieldfn(t) if t.strip() else None
        if f and f not in mapping.values():
            mapping[c] = f
    if len(mapping) < 4:                                   # unreadable header: fall back to the printed order
        cols = [c for c in range(ncol) if any(_text(r[c]) for r in rows if c < len(r))] or list(range(ncol))
        mapping = {c: f for c, f in zip(cols, order)}
    return mapping, j


STOP = re.compile(r"(?i)^\s*(i\s*/\s*we\s+agree|note\b|name\s*&|signature|\(?b\)|\(?a\)|sample price|tender for)")


def _extract_part(grid, start, kind):
    fieldfn, order = (_local_field, LOCAL_FIELDS) if kind == "local" else (_foreign_field, FOREIGN_FIELDS)
    mapping, j = _header_map(grid, start, fieldfn, order)
    items, stated = [], None
    blank_run = 0
    while j < len(grid):
        row = grid[j]
        j += 1
        if _is_local_header(row) or _is_foreign_header(row):
            j -= 1
            break
        first = next((_text(c) for c in row if _text(c)), "")
        if not first:
            blank_run += 1
            if blank_run >= 4 and items:
                break
            continue
        blank_run = 0
        if STOP.match(first) and not re.fullmatch(r"\d+", first):
            break
        rec = {f: (row[c] if c < len(row) else None) for c, f in mapping.items()}
        body = [f for f, v in rec.items() if f != "sno" and _text(v)]
        if not body:
            continue                                       # numbered but empty line of the blank form
        desc = _text(rec.get("desc"))
        if re.search(r"(?i)\b(grand\s+)?total\b", desc) or (re.search(r"(?i)\btotal\b", first) and not desc):
            nums = [num(v) for f, v in rec.items() if f not in ("sno", "desc") and num(v) is not None]
            if nums:
                stated = nums[-1]
            continue
        items.append(rec)
    return items, stated, j


def parse_price(path):
    """-> dict(parts=[dict(kind, items, stated)], currency_text, tender)."""
    grids, text = read_grids(path)
    parts = []
    for g in grids:
        i = 0
        while i < len(g):
            kind = "local" if _is_local_header(g[i]) else "foreign" if _is_foreign_header(g[i]) else None
            if not kind:
                i += 1
                continue
            items, stated, nxt = _extract_part(g, i, kind)
            if items:
                parts.append(dict(kind=kind, items=items, stated=stated))
            i = max(nxt, i + 1)
    m = re.search(r"(?i)tender\s*no\.?\s*[:\-]?\s*([A-Za-z0-9/\-\.]{3,})", text)
    return dict(parts=parts, text=text, tender=m.group(1).strip(".") if m else "")


CURRENCY = re.compile(
    r"\b(USD|EUR|GBP|JPY|CNY|INR|AUD|SGD|CHF|AED|LKR)\b"
    r"|US\s*\$|[$€£₹]|\bRs\.?"
, re.IGNORECASE)


def detect_currency(text, kind):
    found = []
    for m in CURRENCY.finditer(text):
        s = m.group(0)
        found.append({
            "$": "USD",
            "€": "EUR",
            "£": "GBP",
            "₹": "INR",
        }.get(s, "LKR" if s.lower().startswith("rs") else "USD" if s.upper().startswith("US") else s.upper()))
    if kind == "local":
        return "LKR"
    for c in found:
        if c != "LKR":
            return c
    if "LKR" in found:
        return "LKR"
    return ""


# --------------------------------------------------------------------------- numbers for the UI
def part_total(kind, items, include_commission=False):
    """Compute the quoted total of one part the same way the Excel formulas do. -> (total, notes)"""
    total, notes = 0.0, []
    for k, it in enumerate(items, 1):
        qty = num(it.get("qty"))
        if kind == "local":
            incl = num(it.get("unit_incl"))
            if incl is None:
                ex = num(it.get("unit_excl"))
                if ex is None:
                    notes.append(f"line {k}: no unit price")
                    continue
                incl = ex * (1 + sum(num(it.get(x)) or 0 for x in ("vat", "sscl", "other")) / 100)
                notes.append(f"line {k}: price including taxes calculated")
            if qty is None:
                notes.append(f"line {k}: quantity missing (counted as 1)")
                qty = 1
            total += qty * incl
        else:
            cif = num(it.get("cif"))
            if cif is None:
                fob = num(it.get("fob_total"))
                if fob is None and num(it.get("fob_unit")) is not None:
                    fob = (qty if qty is not None else 1) * num(it.get("fob_unit"))
                if fob is None:
                    notes.append(f"line {k}: no cost")
                    continue
                cif = fob + (num(it.get("insurance")) or 0) + (num(it.get("freight")) or 0)
                notes.append(f"line {k}: CIF/CIP cost calculated")
            total += cif + ((num(it.get("commission")) or 0) if include_commission else 0)
    return total, notes


# --------------------------------------------------------------------------- Excel
def _hcell(ws, r, c, text, fill="1F3864", color="FFFFFF"):
    x = ws.cell(r, c, text)
    x.font = Font(bold=True, color=color)
    x.fill = PatternFill("solid", fgColor=fill)
    x.alignment = C.CENTER
    x.border = C.BORDER
    return x


def _body(ws, r, c, v, fmt=None, fill=None, bold=False, center=False):
    x = C.body(ws, r, c, v, bold=bold, center=center)
    if fmt:
        x.number_format = fmt
    if fill:
        x.fill = fill
    return x


MONEY = "#,##0.00"


def build_workbook(entries, out):
    """Create the two editable price-comparison tabs from local-supply entries."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    used = set()
    local_entries = [e for e in entries if e["kind"] == "local"]
    _comparison_sheet(wb, C.safe_sheet("Unit Price Excl Taxes", used), local_entries, "unit_excl",
                      "Unit Price (excl. taxes)")
    _comparison_sheet(wb, C.safe_sheet("Unit Price Incl Taxes", used), local_entries, "unit_incl",
                      "Unit Price (incl. taxes)")
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    wb.save(out)


def _comparison_sheet(wb, name, entries, price_field, price_title):
    """Create the bidder-column comparison layout used by the uploaded price template."""
    ws = wb.create_sheet(name)
    bidders = sorted({e["bidder"] for e in entries}, key=str.lower)
    items = OrderedDict()
    for entry in entries:
        for index, item in enumerate(entry["items"], 1):
            description = re.sub(r"\s+", " ", _text(item.get("desc")))
            key = description.casefold() if description else f"__row_{index}"
            row = items.setdefault(key, {"sno": _text(item.get("sno")) or index, "description": description,
                                         "prices": {}})
            value = num(item.get(price_field))
            if price_field == "unit_incl" and value is None:
                excluded = num(item.get("unit_excl"))
                if excluded is not None:
                    taxes = sum(num(item.get(field)) or 0 for field in ("vat", "sscl", "other"))
                    value = excluded * (1 + taxes / 100)
            if value is not None:
                row["prices"][entry["bidder"]] = value
    ws["A1"] = f"LOCAL SUPPLIES — {price_title}"
    ws["A1"].font = Font(bold=True, size=13, color=C.NAVY)
    ws["A2"] = "Comparison of bidder unit prices. Cells are editable. Final Amount is calculated automatically for each company."
    ws["A2"].font = Font(italic=True, color="595959")
    header_row = 4
    _hcell(ws, header_row, 1, "S/No")
    _hcell(ws, header_row, 2, "Description")
    for column, bidder in enumerate(bidders, 3):
        _hcell(ws, header_row, column, bidder)
    first_data_row = header_row + 1
    for row_number, item in enumerate(items.values(), first_data_row):
        _body(ws, row_number, 1, item["sno"], center=True)
        _body(ws, row_number, 2, item["description"])
        for column, bidder in enumerate(bidders, 3):
            value = item["prices"].get(bidder)
            _body(ws, row_number, column, value, fmt=MONEY, fill=INPUT, center=True)
    final_row = max(first_data_row, header_row + len(items) + 1)
    _body(ws, final_row, 1, "", fill=PatternFill("solid", fgColor="D9EAF7"))
    _body(ws, final_row, 2, "FINAL AMOUNT", fill=PatternFill("solid", fgColor="D9EAF7"), bold=True)
    for column in range(3, len(bidders) + 3):
        bidder = bidders[column - 3]
        total = sum(item["prices"].get(bidder, 0) or 0 for item in items.values())
        _body(ws, final_row, column, total,
              fmt=MONEY, fill=PatternFill("solid", fgColor="D9EAF7"), bold=True, center=True)
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 42
    for column in range(3, len(bidders) + 3):
        ws.column_dimensions[L(column)].width = 22
    C.finish(ws, "C5")
    ws.freeze_panes = "C5"


def _detail_sheet(wb, name, e):
    ws = wb.create_sheet(name)
    local = e["kind"] == "local"
    ws["A1"] = f"{e['bidder']} — PRICE SCHEDULE ({'Part A: Local supplies' if local else 'Part B: Foreign supplies'})"
    ws["A1"].font = Font(bold=True, size=13, color=C.NAVY)
    ws["A2"] = f"Source: {e['source']}" + (f"   •   Tender No: {e['tender']}" if e.get("tender") else "") + f"   •   Currency: {e['currency'] or '(not stated)'}"
    ws["A2"].font = Font(italic=True, color="595959")
    heads = (["S/No", "Description", "Qty", "Unit Price (excl. taxes)", "VAT (%)", "NBT (%)", "Other Taxes (%)", "Unit Price (incl. taxes)"] if local else
             ["S/No", "Description", "Qty", "Country of Origin", "FOB/FCA Unit Price", "FOB/FCA Total Cost", "Total Insurance",
              "Total Freight", "Total CIF/CIP Cost", "Local Agent's Commission"])
    if not local:
        ws["A3"] = "Include Local Agent's Commission in the total?"
        ws["A3"].font = Font(bold=True)
        ws.merge_cells("A3:C3")
        ws["D3"] = "No"
        ws["D3"].fill = INPUT
        ws["D3"].border = C.BORDER
        ws["D3"].alignment = C.CENTER
        dv = DataValidation(type="list", formula1='"Yes,No"', allow_blank=False)
        ws.add_data_validation(dv)
        dv.add("D3")
    hdr = 4
    for c, t in enumerate(heads, 1):
        _hcell(ws, hdr, c, t)
    r = first = hdr + 1
    for k, it in enumerate(e["items"], 1):
        v = lambda f: it.get(f)
        sno = _text(v("sno")) or k
        _body(ws, r, 1, sno, center=True)
        _body(ws, r, 2, re.sub(r"\s+", " ", _text(v("desc"))))
        q = num(v("qty"))
        _body(ws, r, 3, q if q is not None else _text(v("qty")) or None, fmt="#,##0.##", center=True)
        if local:
            for col, f in ((4, "unit_excl"), (5, "vat"), (6, "nbt"), (7, "other")):
                n = num(v(f))
                _body(ws, r, col, n if n is not None else (_text(v(f)) or None), fmt=MONEY if col == 4 else "0.##")
            n = num(v("unit_incl"))
            if n is not None:
                _body(ws, r, 8, n, fmt=MONEY)
            elif num(v("unit_excl")) is not None:
                _body(ws, r, 8, f"=D{r}*(1+(N(E{r})+N(F{r})+N(G{r}))/100)", fmt=MONEY, fill=AMBER)
            else:
                _body(ws, r, 8, None, fmt=MONEY, fill=AMBER)
        else:
            _body(ws, r, 4, _text(v("origin")))
            fu, ft = num(v("fob_unit")), num(v("fob_total"))
            _body(ws, r, 5, fu if fu is not None else (_text(v("fob_unit")) or None), fmt=MONEY)
            if ft is not None:
                _body(ws, r, 6, ft, fmt=MONEY)
            elif fu is not None:
                _body(ws, r, 6, f"=IF(ISNUMBER(C{r}),C{r},1)*E{r}", fmt=MONEY, fill=AMBER)
            else:
                _body(ws, r, 6, None, fmt=MONEY, fill=AMBER)
            for col, f in ((7, "insurance"), (8, "freight")):
                n = num(v(f))
                _body(ws, r, col, n if n is not None else (_text(v(f)) or None), fmt=MONEY)
            cif = num(v("cif"))
            if cif is not None:
                _body(ws, r, 9, cif, fmt=MONEY)
            else:
                _body(ws, r, 9, f"=N(F{r})+N(G{r})+N(H{r})", fmt=MONEY, fill=AMBER)
            cm = num(v("commission"))
            _body(ws, r, 10, cm if cm is not None else (_text(v("commission")) or None), fmt=MONEY)
        r += 1
    last = r - 1
    tcol = "H" if local else "I"
    _body(ws, r, 2, "TOTAL", bold=True)
    for c in range(1, len(heads) + 1):
        if c != 2:
            _body(ws, r, c, None)
    if local:
        total_value = sum(
            (num(it.get("qty")) if num(it.get("qty")) is not None else 1)
            * (num(it.get("unit_incl")) if num(it.get("unit_incl")) is not None else
               (num(it.get("unit_excl")) or 0) * (1 + sum(num(it.get(x)) or 0 for x in ("vat", "sscl", "other")) / 100))
            for it in e["items"]
        )
    else:
        total_value = part_total("foreign", e["items"])[0]
    _body(ws, r, 8 if local else 9, total_value, fmt=MONEY, bold=True)
    if not local:
        commission_total = sum(num(it.get("commission")) or 0 for it in e["items"])
        _body(ws, r, 10, commission_total, fmt=MONEY, bold=True)
    total_cell = f"{tcol}{r}"
    if e.get("stated") is not None:
        r += 1
        _body(ws, r, 2, "Total stated in the bidder's document", bold=False)
        _body(ws, r, 8 if local else 9, e["stated"], fmt=MONEY)
        r += 1
        _body(ws, r, 2, "Difference (calculated − stated)")
        _body(ws, r, 8 if local else 9, f"={tcol}{r-2}-{tcol}{r-1}", fmt=MONEY)
    widths = [7, 42, 8, 18, 14, 14, 14, 16, 16, 18, 16][: len(heads)]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[L(i)].width = w
    ws.row_dimensions[hdr].height = 34
    C.finish(ws, f"C{hdr+1}")
    return total_cell


# --------------------------------------------------------------------------- ranking
def _ordinal(n):
    return f"{n}{'th' if 10 <= n % 100 <= 20 else {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')}"


def rank_rows(res):
    """Financial evaluation status = place by price: the lowest quoted total is 1st.
    Totals in different currencies cannot be compared, so each currency is ranked on its own."""
    res.setdefault("errors", [])
    missing_currency = [r["bidder"] for r in res["rows"] if not r["currency"]]
    if missing_currency:
        bidders = ", ".join(dict.fromkeys(missing_currency))
        res["errors"].append(
            "Cannot compare prices because the currency is not stated in the foreign "
            f"price schedule(s) for: {bidders}. Specify the currency and try again."
        )
        return

    groups = OrderedDict()
    for r in res["rows"]:
        groups.setdefault(r["currency"], []).append(r)
    for cur, lst in groups.items():
        priced = sorted((r for r in lst if r["total"] > 0), key=lambda r: r["total"])
        place, prev = 0, None
        for k, r in enumerate(priced, 1):
            if r["total"] != prev:
                place, prev = k, r["total"]
            tie = sum(1 for x in priced if x["total"] == r["total"]) > 1
            r["rank"] = place
            r["status"] = (f"{_ordinal(place)} — Lowest evaluated price" if place == 1 else _ordinal(place)) + (" (tie)" if tie else "")
        for r in lst:
            if r["total"] <= 0:
                r["rank"] = None
                r["status"] = "No price found"
    res["rows"].sort(key=lambda r: (list(groups).index(r["currency"]), r["rank"] is None, r["rank"] or 0, r["bidder"].lower()))
    if len(groups) > 1:
        res["notes"].append("Prices are in different currencies (" + ", ".join(c or "currency not stated" for c in groups) +
                            "), so bidders are ranked within each currency only. Convert to one currency to compare them directly.")


# --------------------------------------------------------------------------- driver
def run(inputs, out):
    """-> dict(file, rows=[...], problems=[...])"""
    out = Path(out)
    out.mkdir(parents=True, exist_ok=True)
    tmp = tempfile.mkdtemp()
    problems = []
    entries = []
    for bname, f, orig in C.collect(inputs, tmp, problems):
        if f.suffix.lower().startswith(".xls") and f.suffix.lower() not in (".xlsx", ".xlsm"):
            continue
        try:
            data = parse_price(f)
        except Exception as ex:
            problems.append(f"{orig.name} ({bname}): {ex}")
            continue
        if not data["parts"]:
            problems.append(f"{orig.name} ({bname}): no filled-in price schedule found (is Part A / Part B filled in?)")
            continue
        for p in data["parts"]:
            entries.append(dict(bidder=bname, kind=p["kind"], items=p["items"], stated=p["stated"],
                                currency=detect_currency(data["text"], p["kind"]), source=orig.name, tender=data["tender"]))
    res = dict(file=None, rows=[], problems=problems, notes=[], errors=[])
    if entries:
        res["file"] = out / "Price_Schedules.xlsx"
        build_workbook(entries, res["file"])
        rows = OrderedDict()
        for e in sorted(entries, key=lambda x: (x["bidder"].lower(), x["kind"])):
            total, notes = part_total(e["kind"], e["items"])
            label = "Local" if e["kind"] == "local" else "Foreign"
            warn = [f"{label}: {n}" for n in notes]
            if e["stated"] is not None and abs(e["stated"] - total) > 0.5:
                warn.append(f"{label}: calculated total {total:,.2f} differs from the total stated in the document ({e['stated']:,.2f})")
            row = rows.setdefault((e["bidder"], e["currency"]), dict(
                bidder=e["bidder"], currency=e["currency"], total=0.0,
                unit_price_excl_taxes=None, unit_price_incl_taxes=None,
                lines=0, status="To be evaluated", source=[], warnings=[]))
            row["total"] += total
            if e["kind"] == "local":
                excl = sum(num(item.get("unit_excl")) or 0 for item in e["items"])
                incl = sum(
                    num(item.get("unit_incl")) if num(item.get("unit_incl")) is not None else
                    (num(item.get("unit_excl")) or 0) * (
                        1 + sum(num(item.get(field)) or 0 for field in ("vat", "sscl", "other")) / 100
                    )
                    for item in e["items"]
                )
                row["unit_price_excl_taxes"] = (row["unit_price_excl_taxes"] or 0) + excl
                row["unit_price_incl_taxes"] = (row["unit_price_incl_taxes"] or 0) + incl
            row["lines"] += len(e["items"])
            row["warnings"] += warn
            if e["source"] not in row["source"]:
                row["source"].append(e["source"])
        for row in rows.values():
            row["source"] = ", ".join(row["source"])
            res["rows"].append(row)
        rank_rows(res)
    return res
