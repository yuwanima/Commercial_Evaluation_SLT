
import re
from datetime import datetime
from pathlib import Path

import docx
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation

from goods_services import consolidate as C

STATUS_START = re.compile(
    r"^\s*((fully|partially|not|non)[\s-]*)?(complied|compliant|complies|compliance)\b"
    r"|^\s*(yes|no|n/?a|pass|fail|partial|qualified|disqualified|not qualified)\b|^\s*[✓✔√✗✘×]|^\s*does not comply", re.I)
OVERALL_ROW = re.compile(r"overall|final result|^result|recommend|qualif", re.I)

def status_class(t):
    t = C.clean(t).lower()
    if not t or set(t) <= set("-–— "):
        return "blank"
    if re.search(r"not\s*compl|non[- ]?compl|does not|disqualif|not\s*qualif|^no\b|^fail|[✗✘×]", t):
        return "bad"
    if "partial" in t:
        return "partial"
    if re.search(r"complied|compliant|complies|^yes|^pass|^qualified|[✓✔√]", t):
        return "ok"
    return "other"

def _is_status(t):
    return bool(STATUS_START.match(C.clean(t)))

def _share(cells):
    cells = [c for c in cells if C.clean(c)]
    return (sum(_is_status(c) for c in cells) / len(cells), len(cells)) if cells else (0, 0)

# ----------------------------------------------------------------------------- parsing
def parse_eval(path):
    d = docx.Document(path)
    heads = [C.clean(p.text) for p in d.paragraphs if C.clean(p.text)][:10]
    # the title is not always the first line (a form code such as "SCMD/RTE/001" often comes first)
    title = next((h for h in heads if "EVALUATION REPORT" in h.upper()), heads[0] if heads else "")
    if "EVALUATION REPORT" not in title.upper():
        raise ValueError("not an evaluation report")
    rtype = "SSET" if "SOLE SOURCE" in title.upper() else "TET"
    tables = [C._rows(t) for t in d.tables]

    # ---- header details
    meta = {}
    for rows in tables[:2]:
        for r in rows:
            for cell in r:
                m = re.match(r"(?is)\s*(subject|tender no|user department|forwarded by)[\s:.]*(.*)", cell)
                if m and C.clean(m.group(2)):
                    meta[m.group(1).title().replace("No", "No.")] = C.clean_multiline(m.group(2))
    for rows in tables[:3]:
        for r in rows:
            if len(r) >= 2 and re.search(r"date", r[0], re.I) and C.clean(r[1]):
                meta[C.clean(r[0])] = C.clean(r[1])

    # ---- details of bids
    bids = None
    for rows in tables:
        h = " ".join(rows[0]).lower() if rows else ""
        if "bidder" in h and ("brand" in h or "validity" in h or "authority" in h):
            body = [r for r in rows[1:] if any(C.clean(x) for x in r[1:]) or (C.clean(r[0]) and not re.fullmatch(r"\d+", C.clean(r[0])))]
            if body:
                bids = dict(headers=[C.clean(x).replace("\n", " ") for x in rows[0]], rows=body)
            break

    # ---- compliance matrices
    matrices = []
    for rows in tables:
        rows = [r for r in rows if any(C.clean(x) for x in r)]
        if len(rows) < 2 or len(rows[0]) < 2:
            continue
        head_share, _ = _share(rows[0][1:])
        has_header = head_share < 0.5
        body = rows[1:] if has_header else rows
        share, n = _share([c for r in body for c in r[1:]])
        if n < 2 or share < 0.6:
            continue
        if has_header:
            hdr = [C.clean(x) for x in rows[0]]
            if re.search(r"bidder|company|vendor|supplier|tenderer", hdr[0], re.I):      # bidders down the side -> transpose
                areas = hdr[1:]
                bidders = [C.clean(r[0]) for r in body]
                body = [[areas[j]] + [C.clean(r[j + 1]) if j + 1 < len(r) else "" for r in body] for j in range(len(areas))]
                hdr = ["Evaluation Area"] + bidders
            key = tuple(x.lower() for x in hdr[1:])
            cur = next((m for m in matrices if m["key"] == key), None)
            if cur is None:
                cur = dict(key=key, headers=hdr, rows=[], overall=None)
                matrices.append(cur)
        else:
            cur = next((m for m in reversed(matrices) if len(m["headers"]) == len(rows[0])), None)
            if cur is None:
                continue
        for r in body:
            r = (r + [""] * len(cur["headers"]))[:len(cur["headers"])]
            area = C.clean_multiline(r[0])
            if [C.clean(x).lower() for x in r] == [x.lower() for x in cur["headers"]]:
                continue                                  # repeated header row at a page break
            vals = [C.clean_multiline(x) for x in r[1:]]
            if OVERALL_ROW.search(area) and all(not v or status_class(v) in ("ok", "bad", "other") for v in vals):
                cur["overall"] = vals
            else:
                cur["rows"].append((area, vals))
    if not matrices:
        raise ValueError("no compliance matrix (Evaluation Area x bidders) found in this report")
    return dict(type=rtype, title=title, meta=meta, bids=bids, matrices=matrices)

def eval_to_json(label, ev):
    def cell(t):
        return dict(t=t, c=status_class(t))
    all_bidders, not_eligible = [], set()
    for m in ev["matrices"]:
        for bn in m["headers"][1:]:
            if bn not in all_bidders:
                all_bidders.append(bn)
        for area, vals in m["rows"]:
            for j, v in enumerate(vals):
                if status_class(v) in ("partial", "bad") and j + 1 < len(m["headers"]):
                    not_eligible.add(m["headers"][1 + j])
    shortlist = dict(eligible=[b for b in all_bidders if b not in not_eligible],
                     excluded=[b for b in all_bidders if b in not_eligible])
    return dict(label=label, type=ev["type"], title=ev["title"], meta=ev["meta"], bids=ev["bids"],
                shortlist=shortlist,
                matrices=[dict(headers=m["headers"],
                               rows=[dict(area=a, cells=[cell(v) for v in vals]) for a, vals in m["rows"]],
                               overall=[cell(v) for v in m["overall"]] if m["overall"] else None,
                               eligible=[(m["headers"][1 + j] not in not_eligible) for j in range(len(m["headers"]) - 1)])
                          for m in ev["matrices"]])

# ----------------------------------------------------------------------------- Excel
OK, BAD, PART, PEND = "C6EFCE", "FFC7CE", "FFEB9C", "FFF2CC"

def _cf(ws, rng, first, kind):
    f = lambda formula, color: ws.conditional_formatting.add(
        rng, FormulaRule(formula=[formula], fill=PatternFill("solid", bgColor=color, fgColor=color)))
    if kind == "status":
        f(f'LEFT({first},8)="Complied"', OK)
        f(f'LEFT({first},7)="Partial"', PART)
        f(f'OR(LEFT({first},3)="Not",LEFT({first},3)="Non")', BAD)
    else:
        f(f'LEFT({first},4)="Qual"', "A9D08E")
        f(f'OR(LEFT({first},4)="Disq",LEFT({first},3)="Not")', "F4B084")
        f(f'OR({first}="",{first}="Pending")', PEND)

def _cf_eligible(ws, rng, first):
    ws.conditional_formatting.add(
        rng,
        FormulaRule(
            formula=[f'LEFT({first},3)="Yes"'],
            fill=PatternFill("solid", bgColor=OK, fgColor=OK),
        ),
    )
    ws.conditional_formatting.add(
        rng,
        FormulaRule(
            formula=[f'LEFT({first},2)="No"'],
            fill=PatternFill("solid", bgColor=BAD, fgColor=BAD),
        ),
    )

def build_eval_workbook(evals, out):
    """evals: list of (label, parsed_report)."""
    wb = Workbook()
    wb.remove(wb.active)
    used = set()
    multi = len(evals) > 1
    for idx, (label, ev) in enumerate(evals, 1):
        pre = f"R{idx} " if multi else ""
        ws = wb.create_sheet(C.safe_sheet(pre + "Compliance Matrix", used))
        meta = ev["meta"]
        kind_name = "TECHNICAL EVALUATION" if ev["type"] == "TET" else "SOLE SOURCE EVALUATION"
        ws["A1"] = f"{kind_name} — COMPLIANCE MATRIX"
        ws["A1"].font = Font(bold=True, size=14, color=C.NAVY)
        info = "   |   ".join(f"{k}: {v}" for k, v in (("Tender No.", meta.get("Tender No.")), ("Subject", meta.get("Subject")),
                                                      ("User Department", meta.get("User Department"))) if v)
        ws["A2"] = info or f"Source: {label}"
        ws["A2"].font = Font(italic=True, color="595959")
        r = 4
        nc_rows = []
        not_eligible = set()      # bidders with >=1 partial/not-complied item, across all matrices in this report
        all_bidders = []
        first_header_row = r
        for mx in ev["matrices"]:
            n = len(mx["headers"]) - 1
            bidder_names = mx["headers"][1:]
            for bn in bidder_names:
                if bn not in all_bidders:
                    all_bidders.append(bn)
            C.hcell(ws, r, 1, mx["headers"][0] or "Evaluation Area")
            for j in range(n):
                C.hcell(ws, r, 2 + j, mx["headers"][1 + j])
            ws.row_dimensions[r].height = 26
            hdr = r
            r += 1
            first = r
            dv = DataValidation(type="list", formula1='"Complied,Partially complied,Not complied,N/A"',
                                allow_blank=True, showErrorMessage=False)
            ws.add_data_validation(dv)
            for area, vals in mx["rows"]:
                x = C.body(ws, r, 1, area, bold=True)
                x.fill = PatternFill("solid", fgColor="DDEBF7")
                for j in range(n):
                    v = vals[j] if j < len(vals) else ""
                    C.body(ws, r, 2 + j, v, center=True)
                    if status_class(v) in ("partial", "bad"):
                        nc_rows.append((mx["headers"][1 + j], area, v))
                        not_eligible.add(mx["headers"][1 + j])
                r += 1
            last = r - 1
            if last >= first:
                rng = f"B{first}:{L(1 + n)}{last}"
                dv.add(rng)
                _cf(ws, rng, f"B{first}", "status")
            # overall result row
            ovr = r
            x = C.body(ws, r, 1, "Overall result", bold=True)
            x.fill = PatternFill("solid", fgColor="E2EFDA")
            dvo = DataValidation(type="list", formula1='"Qualified,Disqualified,Pending"', allow_blank=True, showErrorMessage=False)
            ws.add_data_validation(dvo)
            for j in range(n):
                v = (mx["overall"][j] if mx["overall"] and j < len(mx["overall"]) else "") or "Pending"
                y = C.body(ws, r, 2 + j, v, bold=True, center=True)
            dvo.add(f"B{ovr}:{L(1 + n)}{ovr}")
            _cf(ws, f"B{ovr}:{L(1 + n)}{ovr}", f"B{ovr}", "overall")
            r += 2
            # live summary
            labels = [("Fully complied (no.)", 'COUNTIF({c}{a}:{c}{b},"complied*")'),
                      ("Partially complied (no.)", 'COUNTIF({c}{a}:{c}{b},"partial*")'),
                      ("Not complied (no.)", 'COUNTIF({c}{a}:{c}{b},"not*")+COUNTIF({c}{a}:{c}{b},"non*")')]
            srow = r
            for k, (lbl, f) in enumerate(labels):
                C.body(ws, r, 1, lbl, bold=True)
                for j in range(n):
                    col = L(2 + j)
                    C.body(ws, r, 2 + j, "=" + f.format(c=col, a=first, b=max(last, first)), bold=True, center=True)
                r += 1
            # Eligibility is derived strictly from the compliance cells, not from whatever the
            # "Overall result" row says — a single partial/not-complied item excludes the bidder
            # from commercial evaluation, so this cannot silently say "Qualified" by mistake.
            elig_row = r
            C.body(ws, r, 1, "Eligible for Commercial Evaluation", bold=True)
            for j in range(n):
                col = L(2 + j)
                x = C.body(ws, r, 2 + j,
                           f'=IF({col}{srow+1}+{col}{srow+2}=0,"Yes — proceed","No — excluded (non-compliance)")',
                           bold=True, center=True)
            _cf_eligible(ws, f"B{elig_row}:{L(1+n)}{elig_row}", f"B{elig_row}")
            r += 1
            C.body(ws, r, 1, "Flag: report's stated result conflicts with compliance", bold=False)
            for j in range(n):
                col = L(2 + j)
                x = C.body(ws, r, 2 + j,
                           f'=IF(AND(LEFT({col}{ovr},4)="Qual",LEFT({col}{elig_row},2)="No"),"Report says Qualified but has non-compliance — check",'
                           f'IF(AND(LEFT({col}{ovr},4)="Disq",LEFT({col}{elig_row},3)="Yes"),"Report says Disqualified with no non-compliance — check",""))',
                           center=True)
                x.font = Font(italic=True, size=9, color="B42318")
            r += 3
        ws.column_dimensions["A"].width = 36
        ncols = max(len(m["headers"]) for m in ev["matrices"])
        for j in range(2, ncols + 1):
            ws.column_dimensions[L(j)].width = 22
        C.finish(ws, f"B{first_header_row+1}")
        ws.cell(r, 1, f"Generated {datetime.now():%Y-%m-%d %H:%M} from {label}. Cells are editable (drop-down lists); counts and review checks update automatically.").font = Font(italic=True, size=9, color="7F7F7F")

        # ---- non-compliances (report instruction 4b): a record of why each bidder was excluded.
        # Non-compliant / partially-compliant bidders do NOT go forward to commercial evaluation,
        # so this sheet is an audit trail, not a worksheet for deciding whether to still accept them.
        ws2 = wb.create_sheet(C.safe_sheet(pre + "Non-compliances", used))
        ws2["A1"] = "NON-COMPLIANCES — bidders below are excluded from commercial evaluation (Section 4b)"
        ws2["A1"].font = Font(bold=True, size=13, color=C.NAVY)
        for c, h in enumerate(["Bidder", "Evaluation Area", "Status in report", "Impact (for the record)"], 1):
            C.hcell(ws2, 3, c, h)
        k = 4
        for b, a, v in nc_rows:
            for c, val in enumerate([b, a, v, ""], 1):
                C.body(ws2, k, c, val, bold=(c == 1))
            k += 1
        if not nc_rows:
            C.body(ws2, 4, 1, "No non-compliances reported — every bidder is eligible for commercial evaluation.")
            ws2.merge_cells("A4:D4")
        for c, w in zip("ABCD", (26, 46, 26, 60)):
            ws2.column_dimensions[c].width = w
        C.finish(ws2, "A4")

        # ---- commercial evaluation shortlist: only bidders with ZERO non-compliance, per the report's
        # own rule that a single non-compliance or partial compliance excludes a bid from commercial evaluation.
        ws4 = wb.create_sheet(C.safe_sheet(pre + "Commercial Eval Shortlist", used))
        ws4["A1"] = "SHORTLIST FOR COMMERCIAL EVALUATION — fully-complied bidders only"
        ws4["A1"].font = Font(bold=True, size=13, color=C.NAVY)
        ws4["A2"] = "Any bidder with a partial or non-compliance in the Compliance Matrix is excluded here automatically."
        ws4["A2"].font = Font(italic=True, color="595959")
        bid_hdrs = ev["bids"]["headers"][1:] if ev["bids"] else []
        bid_by_name = {}
        if ev["bids"]:
            for row in ev["bids"]["rows"]:
                if row and C.clean(row[0]):
                    bid_by_name[C.clean(row[0]).lower()] = row[1:]
        heads = ["Bidder", "Status"] + bid_hdrs
        for c, h in enumerate(heads, 1):
            C.hcell(ws4, 4, c, h)
        k = 5
        shortlisted = [b for b in all_bidders if b not in not_eligible]
        for b in shortlisted:
            vals = [b, "Eligible — proceed to commercial evaluation"] + list(bid_by_name.get(b.lower(), [""] * len(bid_hdrs)))
            for c, v in enumerate(vals, 1):
                x = C.body(ws4, k, c, v, bold=(c == 1))
                if c == 2:
                    x.fill = PatternFill("solid", fgColor=C.GREEN)
            k += 1
        excluded = [b for b in all_bidders if b in not_eligible]
        if excluded:
            k += 1
            ws4.cell(k, 1, "Excluded (non-compliance / partial compliance):").font = Font(bold=True, color="B42318")
            k += 1
            for b in excluded:
                x = C.body(ws4, k, 1, b, bold=True)
                y = C.body(ws4, k, 2, "Not eligible")
                y.fill = PatternFill("solid", fgColor=C.RED)
                k += 1
        if not shortlisted:
            ws4.cell(5, 1, "No bidder fully complied — none are eligible for commercial evaluation as things stand.").font = Font(italic=True, color="B42318")
        ws4.column_dimensions["A"].width = 24
        ws4.column_dimensions["B"].width = 36
        for c in range(3, 3 + len(bid_hdrs)):
            ws4.column_dimensions[L(c)].width = 22
        C.finish(ws4, "A5")

        # ---- bid details
        if ev["bids"]:
            ws3 = wb.create_sheet(C.safe_sheet(pre + "Bid Details", used))
            ws3["A1"] = "DETAILS OF BIDS"
            ws3["A1"].font = Font(bold=True, size=13, color=C.NAVY)
            for c, h in enumerate(ev["bids"]["headers"], 1):
                C.hcell(ws3, 3, c, h)
            for i, row in enumerate(ev["bids"]["rows"], 4):
                for c, v in enumerate(row, 1):
                    C.body(ws3, i, c, v)
            for c in range(1, len(ev["bids"]["headers"]) + 1):
                ws3.column_dimensions[L(c)].width = 24
            C.finish(ws3, "A4")
    wb.save(out)
