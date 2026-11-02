from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

from excel_writer import (BORDER, FONT, HDR_FILL, NOTE_FILL, _font, _hdr, _is_empty, _kv_sheet, _label, _put, _src, _title,
                          BLANK_FILL)

YN = ["Yes", "No"]
CN = ["Compliant", "Non-compliant"]
QD = ["Qualified", "Disqualified"]

T1 = [("tender_no", "Tender No", "", None), ("user_department", "User Department", "", None),
      ("forwarded_by", "Forwarded by", "", None), ("_budget", "Budget Confirmed", "CAPEX / OPEX amounts", None),
      ("_jtb", "Previous JTB Reference (if any)", "JTB Ref No and Date", None),
      ("published", "Tender Published Date", "", None), ("closed", "Tender Closed Date", "", None),
      ("opened", "Technical Bid Opened Date", "", None), ("submitted", "TET Report Submission Date", "", None),
      ("bond_expiry", "Bid Bond Expiry Date(s)", "", None),
      ("method", "Method of Tendering", "Open Tender / Registered Suppliers / Shortlisted / Sole Source", None),
      ("purpose", "Purpose of Tender", "", None),
      ("description", "Short Description of Goods/Services",
       "Include details on installation, commissioning, training and the time period covered", None)]
T2 = [("need", "How the need arose", "", None),
      ("approval", "Relevant Approval", "Governance Board / DPAC approval etc.", None),
      ("project", "Project No. & Name", "if applicable", None),
      ("account", "Account Code & Cost Center", "if OPEX", None),
      ("budget_res", "Budget Reserved/Total Available", "", None),
      ("process", "Procurement process up to bid opening",
       "Dates of floating, bidders' conference, extensions, closing, opening", None)]
# (key, label, hint, dropdown)
T3 = [("name", "Bidder's Name", "", None), ("agent", "Local Agent", "", None),
      ("loa", "Letter of Authority (Yes/No)", "", YN), ("brand", "Product Brand & Model", "", None),
      ("validity", "Bid Validity (as stated in bid)", "", None), ("bond", "Bid Bond Expiry Date", "", None)]
T4 = [("c_general", "Compliance with General Conditions", "Compliant / Non-compliant", CN),
      ("c_technical", "Compliance with Technical Conditions", "Compliant / Non-compliant", CN),
      ("c_special", "Compliance with Special Conditions", "Compliant / Non-compliant", CN),
      ("c_contract", "Compliance with Contract Conditions", "Compliant / Non-compliant", CN),
      ("major", "Major Non-compliances (disqualifying)", "List", None),
      ("minor", "Minor/Acceptable Non-compliances", "List", None),
      ("impact", "Impact of Non-compliances", "", None),
      ("status", "Qualified / Disqualified", "Qualified / Disqualified", QD),
      ("negotiate", "Commercial terms to be negotiated (non-compliance to financial terms)", "", None)]
T5 = [("rec_fin", "Recommended for Financial Bid Opening (Yes/No)", "", YN),
      ("bom", "Evaluated BOM included? (Yes/No)", "", YN),
      ("missing", "Missing price items needing separate sealed envelope? (Yes/No - list items)", "", None),
      ("prev_prices", "Previous purchase price (with period & quantity)", "", None),
      ("ref_price", "Reference price of similar items (market/previous)", "", None)]


S3, S4, S5, S6 = "03 Bidder Details", "04 Compliance Evaluation", "05 Commercial Evaluation", "06 Scores & Ranking"
Q4, Q6 = f"'{S4}'", f"'{S6}'"            # quoted sheet names for formulas
AUTO_FILL = PatternFill("solid", fgColor="E8F0FE")   # light blue = calculated by formula (can be overwritten)
IN_FILL = PatternFill("solid", fgColor="FFF2CC")
BLUE = "0000FF"

# ---- Procurement Manual Part II, Section 6 (Evaluating Responses) -----------------
IMPORTANCE = [("Most Important", 60), ("Important", 30), ("Least Important", 10)]     # s.5.1.4
DEGREE = [("Comply", 10), ("Partial Comply", 5), ("Non Comply", 0)]                  # s.5.1.5
PRICE_WEIGHT_DEFAULT = 70      # Part I s.4.7.3: standard equipment -> commercial 70 : technical 30
CRITERIA = [("Compliance with Technical Conditions", "Most Important", "c_technical"),
            ("Compliance with Special Conditions", "Important", "c_special"),
            ("Compliance with General Conditions", "Least Important", "c_general"),
            ("Compliance with Contract Conditions", "Least Important", "c_contract")]
EXTRA_CRITERIA_ROWS = 6


def _dv(ws, options):
    dv = DataValidation(type="list", formula1='"' + ",".join(options) + '"', allow_blank=True, showErrorMessage=False)
    ws.add_data_validation(dv)
    return dv


def _fcell(ws, r, c, formula, fmt=None, bold=False, center=False):
    """A calculated cell (formula, light-blue)."""
    cell = ws.cell(r, c, formula)
    cell.font = _font(bold=bold)
    cell.fill = AUTO_FILL
    cell.border = BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="top", horizontal="center" if center else None)
    if fmt:
        cell.number_format = fmt
    return cell


def _inp(ws, r, c, value=None, fmt=None, dv=None):
    """A number/choice the member types in (blue text, yellow)."""
    cell = ws.cell(r, c, value)
    cell.font = _font(color=BLUE)
    cell.fill = IN_FILL
    cell.border = BORDER
    cell.alignment = Alignment(wrap_text=True, vertical="top")
    if fmt:
        cell.number_format = fmt
    if dv is not None:
        dv.add(cell)
    return cell


def _ordinal(rk):
    return (f'{rk}&CHOOSE(IF(AND(MOD({rk},100)>=11,MOD({rk},100)<=13),0,MOD({rk},10))+1,'
            f'"th","st","nd","rd","th","th","th","th","th","th")')


def _matrix_sheet(wb, name, title, rows, reports, extra=None):
    """Generic per-bidder table. Returns one layout dict per report: {"rows": {key: row}, "first_col": 2}."""
    ws = wb[name]
    _title(ws, title, 2)
    r, maxcols, dvs, layouts = 4, 1, {}, []
    for rep in reports:
        bidders = rep["bidders"] or [{}]
        maxcols = max(maxcols, len(bidders))
        ws.cell(r, 1, f"Report: {_src(rep).replace(chr(10), ' ')}").font = _font(bold=True, color="1F3864", size=11)
        r += 1
        _hdr(ws.cell(r, 1), "FIELD")
        for j in range(len(bidders)):
            nm = (bidders[j].get("name") or "").strip()
            _hdr(ws.cell(r, 2 + j), f"BIDDER {j + 1}" + (f"\n{nm}" if nm else ""))
        ws.row_dimensions[r].height = 30
        r += 1
        rowmap = {}
        for key, label, hint, options in rows:
            _label(ws, r, label)
            if hint:
                ws.cell(r, 1).comment = Comment(hint, "TET tool", width=240, height=50)
            dv = None
            if options:
                dv = dvs.get(tuple(options)) or _dv(ws, options)
                dvs[tuple(options)] = dv
            for j, b in enumerate(bidders):
                _put(ws, r, 2 + j, b.get(key, ""), "", dv, bold=(key == "name"))
            rowmap[key] = r
            r += 1
        layout = {"rows": rowmap, "n": len(bidders), "title_row": r}
        if extra:
            r = extra(ws, r, rep, bidders, rowmap, layout)
        layouts.append(layout)
        r += 2
    ws.column_dimensions["A"].width = 46
    for j in range(maxcols):
        ws.column_dimensions[get_column_letter(2 + j)].width = 34
    ws.freeze_panes = "B1"
    return layouts


# ---------------------------------------------------------------- table 4 extra row
def _overall_row(ws, r, rep, bidders, rowmap, layout):
    _label(ws, r, "OVERALL: Compliant / Non-compliant (auto)")
    ws.cell(r, 1).comment = Comment(
        "Non-compliant if any of the four conditions is Non-compliant or a major non-compliance is listed; "
        "otherwise Compliant. Formula - overwrite if the team decides differently.", "TET tool", width=280, height=80)
    c1, c4, rm = rowmap["c_general"], rowmap["c_contract"], rowmap["major"]
    for j in range(len(bidders)):
        L = get_column_letter(2 + j)
        f = (f'=IF(AND(COUNTA({L}{c1}:{L}{c4})=0,{L}{rm}=""),"",'
             f'IF(OR(COUNTIF({L}{c1}:{L}{c4},"Non-compliant")>0,'
             f'AND({L}{rm}<>"",ISERROR(MATCH(LOWER(TRIM({L}{rm})),{{"none","nil","n/a","-","no"}},0)))),'
             f'"Non-compliant","Compliant"))')
        _fcell(ws, r, 2 + j, f, bold=True, center=True)
    rng = f"B{r}:{get_column_letter(1 + len(bidders))}{r}"
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Non-compliant"'], font=Font(name=FONT, bold=True, color="C00000")))
    ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Compliant"'], font=Font(name=FONT, bold=True, color="2E7D32")))
    layout["overall_row"] = r
    return r + 1


# ---------------------------------------------------------------- sheet 06: scores
def _scores_sheet(wb, reports, l4):
    ws = wb[S6]
    _title(ws, "06 Scores & Ranking (Procurement Manual Part II, Section 6)", 2)
    ws["A2"] = "Light-blue = calculated (overwrite if needed). Yellow/blue = type here. Place is calculated from the scores."
    # ---- legend / scales
    ws["A4"] = "SCORING SCALES - editable. Source: Procurement Manual Part II, Section 6, 5.1.4 (importance) and 5.1.5 (compliance scale)"
    ws["A4"].font = _font(bold=True, color="1F3864")
    _hdr(ws.cell(5, 1), "Importance"); _hdr(ws.cell(5, 2), "Weight")
    _hdr(ws.cell(5, 4), "Degree of compliance"); _hdr(ws.cell(5, 5), "Score")
    for i, ((il, iw), (dl, ds)) in enumerate(zip(IMPORTANCE, DEGREE)):
        _label(ws, 6 + i, il); _inp(ws, 6 + i, 2, iw)
        ws.cell(6 + i, 4, dl).font = _font(bold=True); ws.cell(6 + i, 4).border = BORDER
        ws.cell(6 + i, 4).fill = PatternFill("solid", fgColor="EDEDED"); _inp(ws, 6 + i, 5, ds)
    notes = [
        "Assumptions (not fixed by the Manual - set per tender before bids open; edit as approved in the tender strategy):",
        "  - The four 'Compliance with ... Conditions' rows are pre-scored from Table 4 (Compliant = Comply, Non-compliant = Non Comply). "
        "Their importance (Technical = Most, Special = Important, General/Contract = Least) is an assumption.",
        "  - Add the tender's own scored criteria in the blank rows (importance + degree of compliance per bidder).",
        "  - Technical score = raw score / maximum x 100. Commercial score = lowest eligible price / bidder price x 100 (Attachment A).",
        "  - Price weight defaults to 70 : 30 (Part I s.4.7.3, standard equipment); use the weighting approved by TSC / CEO.",
        "  - Before prices are known the place is a TECHNICAL ranking only. Per the TET ToR it may be used for commercial weighting only if CEO approval was obtained upfront.",
        "  - A bidder marked 'Disqualified' in Table 4 is not ranked (drop-dead criteria, Part I s.4.7.3).",
    ]
    for i, t in enumerate(notes):
        c = ws.cell(10 + i, 1, t)
        c.font = _font(size=9, italic=(i > 0), bold=(i == 0), color="595959")
    IMP_L, IMP_W = "$A$6:$A$8", "$B$6:$B$8"
    S_COMPLY, S_PART, S_NON = "$E$6", "$E$7", "$E$8"
    dv_imp = _dv(ws, [x for x, _ in IMPORTANCE])
    dv_deg = _dv(ws, [x for x, _ in DEGREE])

    r = 10 + len(notes) + 1
    maxcols, layouts = 1, []
    for ri, rep in enumerate(reports):
        bidders = rep["bidders"] or [{}]
        n = len(bidders)
        maxcols = max(maxcols, n)
        first, last = 4, 3 + n
        Lf, Ll = get_column_letter(first), get_column_letter(last)
        rows4 = l4[ri]["rows"]
        ws.cell(r, 1, f"Report: {_src(rep).replace(chr(10), ' ')}").font = _font(bold=True, color="1F3864", size=11)
        r += 1
        pw_r, tw_r, basis_r = r, r + 1, r + 2
        _label(ws, pw_r, "Price weighting (%)"); _inp(ws, pw_r, 2, PRICE_WEIGHT_DEFAULT)
        _label(ws, tw_r, "Technical weighting (%)"); _fcell(ws, tw_r, 2, f"=100-B{pw_r}")
        _label(ws, basis_r, "Ranking basis")
        r = basis_r + 2
        _hdr(ws.cell(r, 1), "CRITERION"); _hdr(ws.cell(r, 2), "IMPORTANCE"); _hdr(ws.cell(r, 3), "WEIGHT")
        for j, b in enumerate(bidders):
            nm = (b.get("name") or "").strip()
            _hdr(ws.cell(r, first + j), f"BIDDER {j + 1}" + (f"\n{nm}" if nm else ""))
        ws.row_dimensions[r].height = 30
        r += 1
        k0 = r
        for ci in range(len(CRITERIA) + EXTRA_CRITERIA_ROWS):
            crit = CRITERIA[ci] if ci < len(CRITERIA) else None
            _inp(ws, r, 1, crit[0] if crit else None)
            _inp(ws, r, 2, crit[1] if crit else None, dv=dv_imp)
            _fcell(ws, r, 3, f'=IF(B{r}="","",IFERROR(INDEX({IMP_W},MATCH(B{r},{IMP_L},0)),""))', center=True)
            for j in range(n):
                L4 = get_column_letter(2 + j)
                if crit:
                    src = f"{Q4}!{L4}{rows4[crit[2]]}"
                    cell = _fcell(ws, r, first + j, f'=IF({src}="Compliant","Comply",IF({src}="Non-compliant","Non Comply",""))', center=True)
                    cell.font = _font(color="008000")        # green = pulled from another sheet
                    dv_deg.add(cell)
                else:
                    _inp(ws, r, first + j, None, dv=dv_deg)
            r += 1
        k1 = r - 1
        rr = {}
        for key, label in (("raw", "Raw score (sum of weight x score)"), ("max", "Maximum possible raw score"),
                           ("tech", "TECHNICAL SCORE (/100)"), ("elig", "Eligible for ranking? (not disqualified)"),
                           ("price", "Evaluated total price (enter when CET prices are known)"),
                           ("pfx", "Price used for comparison (eligible bidders only)"),
                           ("comm", "COMMERCIAL SCORE (/100, lowest eligible price = 100)"),
                           ("final", "FINAL SCORE"), ("place", "PLACE")):
            rr[key] = r
            _label(ws, r, label)
            if key in ("tech", "comm", "final", "place"):
                ws.cell(r, 1).font = _font(bold=True, color="1F3864")
            r += 1
        rng = lambda k: f"${Lf}${rr[k]}:${Ll}${rr[k]}"
        ws.cell(basis_r, 2).value = (f'=IF(AND(COUNT({rng("pfx")})>0,COUNT({rng("pfx")})=COUNTIF({rng("elig")},"Yes")),'
                                     f'"Combined (price + technical)","Technical only (prices not entered yet)")')
        ws.merge_cells(start_row=basis_r, start_column=2, end_row=basis_r, end_column=5)
        ws.cell(basis_r, 2).fill = AUTO_FILL; ws.cell(basis_r, 2).font = _font(bold=True)
        for j in range(n):
            L = get_column_letter(first + j)
            L4 = get_column_letter(2 + j)
            deg = f"{L}{k0}:{L}{k1}"
            wts = f"$C${k0}:$C${k1}"
            _fcell(ws, rr["raw"], first + j,
                   f'=SUMPRODUCT({wts},({deg}="Comply")*{S_COMPLY}+({deg}="Partial Comply")*{S_PART}+({deg}="Non Comply")*{S_NON})', "0.0")
            _fcell(ws, rr["max"], first + j, f"=SUM({wts})*{S_COMPLY}", "0.0")
            _fcell(ws, rr["tech"], first + j,
                   f'=IF(OR(COUNTIF({deg},"?*")=0,{L}{rr["max"]}=0),"",{L}{rr["raw"]}/{L}{rr["max"]}*100)', "0.00", bold=True)
            _fcell(ws, rr["elig"], first + j, f'=IF({Q4}!{L4}{rows4["status"]}="Disqualified","No","Yes")', center=True)
            _inp(ws, rr["price"], first + j, None, "#,##0.00")
            _fcell(ws, rr["pfx"], first + j, f'=IF(AND(ISNUMBER({L}{rr["price"]}),{L}{rr["elig"]}="Yes"),{L}{rr["price"]},"")', "#,##0.00")
            _fcell(ws, rr["comm"], first + j, f'=IF(ISNUMBER({L}{rr["pfx"]}),MIN({rng("pfx")})/{L}{rr["pfx"]}*100,"")', "0.00", bold=True)
            _fcell(ws, rr["final"], first + j,
                   f'=IF(AND({L}{rr["elig"]}="Yes",ISNUMBER({L}{rr["tech"]})),'
                   f'IF(LEFT($B${basis_r},8)="Combined",$B${pw_r}/100*{L}{rr["comm"]}+$B${tw_r}/100*{L}{rr["tech"]},{L}{rr["tech"]}),"")',
                   "0.00", bold=True)
            rk = f'RANK({L}{rr["final"]},{rng("final")})'
            _fcell(ws, rr["place"], first + j,
                   f'=IF(ISNUMBER({L}{rr["final"]}),{_ordinal(rk)},IF({L}{rr["elig"]}="No","Disqualified",""))', bold=True, center=True)
        layouts.append({"rr": rr, "first": first, "n": n, "basis_r": basis_r})
        r += 2
    ws.column_dimensions["A"].width = 52
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 10
    for j in range(maxcols):
        ws.column_dimensions[get_column_letter(4 + j)].width = 26
    return layouts


# ---------------------------------------------------------------- sheet 05: commercial + scores summary
def _table5(wb, reports, l4, l6):
    ws = wb[S5]
    _title(ws, "5. Recommendations for Commercial Evaluation", 2)
    r, maxcols, dvs = 4, 1, {}
    for ri, rep in enumerate(reports):
        bidders = rep["bidders"] or [{}]
        n = len(bidders)
        maxcols = max(maxcols, n)
        ws.cell(r, 1, f"Report: {_src(rep).replace(chr(10), ' ')}").font = _font(bold=True, color="1F3864", size=11)
        r += 1
        _hdr(ws.cell(r, 1), "FIELD")
        for j in range(n):
            nm = (bidders[j].get("name") or "").strip()
            _hdr(ws.cell(r, 2 + j), f"BIDDER {j + 1}" + (f"\n{nm}" if nm else ""))
        ws.row_dimensions[r].height = 30
        r += 1
        # compliance, linked from table 4
        _label(ws, r, "Compliant / Non-compliant (from Table 4)")
        for j in range(n):
            L = get_column_letter(2 + j)
            _fcell(ws, r, 2 + j, f'={Q4}!{L}{l4[ri]["overall_row"]}', bold=True, center=True)
        rng = f"B{r}:{get_column_letter(1 + n)}{r}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Non-compliant"'], font=Font(name=FONT, bold=True, color="C00000")))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Compliant"'], font=Font(name=FONT, bold=True, color="2E7D32")))
        r += 1
        for key, label, hint, options in T5:
            _label(ws, r, label)
            dv = None
            if options:
                dv = dvs.get(tuple(options)) or _dv(ws, options)
                dvs[tuple(options)] = dv
            for j, b in enumerate(bidders):
                _put(ws, r, 2 + j, b.get(key, ""), "", dv)
            r += 1
        _label(ws, r, "More Suitable (as stated in the TET report)")
        for j, b in enumerate(bidders):
            _put(ws, r, 2 + j, b.get("suitable", ""))
        r += 2
        # ---- separate scores table under the bidder names
        c = ws.cell(r, 1, "SCORES AND PLACE (calculated on sheet 06 - Procurement Manual Part II, Section 6)")
        c.font = _font(bold=True, color="FFFFFF"); c.fill = HDR_FILL
        for j in range(n):
            ws.cell(r, 2 + j).fill = HDR_FILL
        r += 1
        rr, first = l6[ri]["rr"], l6[ri]["first"]
        for key, label, fmt in (("tech", "Technical score (/100)", "0.00"), ("comm", "Commercial score (/100)", "0.00"),
                                ("final", "Final score", "0.00"), ("place", "PLACE (1st / 2nd / 3rd ...)", None)):
            _label(ws, r, label)
            for j in range(n):
                L6 = get_column_letter(first + j)
                ref = f"{Q6}!{L6}{rr[key]}"
                _fcell(ws, r, 2 + j, f'=IF({ref}="","",{ref})', fmt, bold=(key in ("final", "place")), center=(key == "place"))
            r += 1
        _label(ws, r, "Ranking basis")
        ws.cell(r, 1).font = _font(italic=True, size=9)
        ws.cell(r, 2, f"={Q6}!B{l6[ri]['basis_r']}").font = _font(italic=True, size=9)
        r += 3
    ws.column_dimensions["A"].width = 46
    for j in range(maxcols):
        ws.column_dimensions[get_column_letter(2 + j)].width = 34
    ws.freeze_panes = "B1"
    return ws


def build_workbook(reports, out_path):
    wb = Workbook()
    wb.remove(wb.active)
    _kv_sheet(wb, "01 Tender & Timeline", "01 Tender & Timeline Header", T1, reports)
    _kv_sheet(wb, "02 Background Summary", "02 Background Summary", T2, reports)
    for nm in (S3, S4, S5, S6):
        wb.create_sheet(nm)
    _matrix_sheet(wb, S3, "3. Bidder Details Table (key comparison table)", T3, reports)
    l4 = _matrix_sheet(wb, S4, "4. Compliance Evaluation Table", T4, reports, extra=_overall_row)
    l6 = _scores_sheet(wb, reports, l4)
    _table5(wb, reports, l4, l6)
    _log_sheet(wb, reports)
    for w in wb.worksheets:
        w.sheet_properties.tabColor = "0B6E4F" if w.title[0].isdigit() else "7F7F7F"
        w.page_setup.orientation = "landscape"
        w.page_setup.fitToWidth = 1
        w.page_setup.fitToHeight = 0
        w.sheet_properties.pageSetUpPr.fitToPage = True
    wb.calculation.fullCalcOnLoad = True      # Excel recalculates every formula when the file is opened
    wb.save(out_path)
    return out_path


def _log_sheet(wb, reports):
    ws = wb.create_sheet("Extraction Log")
    ws["A1"] = "Extraction Log (for the reviewer - not part of the tables)"
    ws["A1"].font = _font(bold=True, size=13)
    ws.sheet_view.showGridLines = False
    for j, h in enumerate(("Source file", "Tender-level fields found", "Bidders found", "Empty bidder cells", "Notes")):
        _hdr(ws.cell(3, 1 + j), h)
    tender_keys = [(k, l) for k, l, *_ in T1 + T2 if not k.startswith("_")]
    bkeys = [k for k, *_ in T3 + T4 + T5]
    for i, rep in enumerate(reports):
        found = sum(1 for k, _ in tender_keys if k in rep["fields"])
        empty = sum(1 for b in rep["bidders"] for k in bkeys if not b.get(k))
        notes = list(rep["notes"])
        miss = [l for k, l in tender_keys if k not in rep["fields"]]
        if miss:
            notes.append("Not found: " + "; ".join(miss))
        if not rep["bidders"]:
            notes.append("No bidders found.")
        notes.append("Scores/place (sheet 06) are calculated from Table 4 compliance and any criteria/prices typed in - "
                     "the TET report itself is not read for scores.")
        row = [rep["source"], f"{found} of {len(tender_keys)}", len(rep["bidders"]),
               f"{empty} of {len(rep['bidders']) * len(bkeys)}", "\n".join(notes) or "-"]
        for j, v in enumerate(row):
            c = ws.cell(4 + i, 1 + j, v)
            c.font = _font()
            c.border = BORDER
            c.alignment = Alignment(wrap_text=True, vertical="top")
    for col, w in zip("ABCDE", (34, 24, 14, 20, 90)):
        ws.column_dimensions[col].width = w
