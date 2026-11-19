
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
    r"^\s*((fully|partially|not|non)[\s-]*)?(comply|complied|compliant|complies|compliance)\b"
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
    if re.search(r"\bcomply\b", t):
        return "ok"
    if re.search(r"complied|compliant|complies|^yes|^pass|^qualified|[✓✔√]", t):
        return "ok"
    return "other"

def _is_ineligible_status(t):
    return status_class(t) == "bad"

def _recommendation_value(headers, values, pattern):
    for header, value in zip(headers, values):
        if re.search(pattern, header, re.I) and C.clean(value):
            return C.clean_multiline(value)
    return "Not available in technical report"

def _recommendation_table(all_bidders, not_eligible, bids):
    bid_values = {}
    if bids:
        headers = bids["headers"][1:]
        for row in bids["rows"]:
            if row and C.clean(row[0]):
                bid_values[_bidder_match_key(C.clean(row[0]))] = row[1:]
    table = []
    for bidder in all_bidders:
        values = bid_values.get(_bidder_match_key(bidder), [])
        headers = bids["headers"][1:] if bids else []
        table.append({
            "bidder": bidder,
            "financial_bid_opening": "Yes" if bidder not in not_eligible else "No",
            "evaluated_bom_included": _recommendation_value(
                headers, values, r"\b(bom|bill of quantities|bill of materials)\b"),
            "missing_price_items": _recommendation_value(
                headers, values, r"missing.*price|price.*missing|separate.*sealed"),
            "previous_purchase_price": _recommendation_value(
                headers, values, r"previous.*purchase|purchase.*price"),
            "reference_price": _recommendation_value(
                headers, values, r"reference.*price|market.*price"),
            "more_suitable": "Not enough comparable data",
        })
    comparable = []
    for item in table:
        known = sum(
            item[key] != "Not available in technical report"
            for key in ("evaluated_bom_included", "missing_price_items",
                        "previous_purchase_price", "reference_price")
        )
        if item["financial_bid_opening"] == "Yes" and known:
            score = 0
            bom = item["evaluated_bom_included"].lower()
            missing = item["missing_price_items"].lower()
            if bom.startswith("yes") or "included" in bom:
                score += 2
            if missing.startswith("no") or "none" in missing:
                score += 2
            if item["previous_purchase_price"] != "Not available in technical report":
                score += 1
            if item["reference_price"] != "Not available in technical report":
                score += 1
            comparable.append((score, item))
    if comparable:
        comparable.sort(key=lambda pair: pair[0], reverse=True)
        rank = 0
        previous_score = None
        for position, (score, item) in enumerate(comparable, 1):
            if score != previous_score:
                rank = position
                previous_score = score
            item["more_suitable"] = f"{rank} place"
    return table


def _parse_amount_currency(value):
    """Parse a quoted amount without guessing a currency that is not stated."""
    raw = C.clean_multiline(value)
    currency_match = re.search(r"(?i)\b(LKR|Rs\.?|USD|EUR|GBP|JPY|CNY|INR|AUD|SGD|CHF|AED)\b|US\s*\$|[$€£₹]", raw)
    currency = ""
    if currency_match:
        token = currency_match.group(0).upper().replace(" ", "")
        currency = {"RS": "LKR", "RS.": "LKR", "$": "USD", "US$": "USD", "€": "EUR", "£": "GBP", "₹": "INR"}.get(token, token)
    cleaned = re.sub(r"(?i)\b(LKR|Rs\.?|USD|EUR|GBP|JPY|CNY|INR|AUD|SGD|CHF|AED)\b|US\s*\$|[$€£₹,\s]", "", raw)
    match = re.search(r"-?\d+(?:\.\d+)?", cleaned)
    amount = float(match.group(0)) if match else None
    return amount, currency


def _extract_commercial_prices(tables):
    """Extract explicit bidder/quoted-total tables from a TET report as fallback evidence."""
    found = []
    for rows in tables:
        cleaned = [[C.clean_multiline(cell) for cell in row] for row in rows]
        for header_index, header in enumerate(cleaned[:4]):
            bidder_col = next((i for i, cell in enumerate(header)
                               if re.search(r"\b(bidder|tenderer|supplier|vendor|company)\b", cell, re.I)), None)
            total_col = next((i for i, cell in enumerate(header)
                              if re.search(r"quoted\s+total|evaluated\s+total|(quoted|evaluated|tender|bid|total).*(price|amount|cost)|(price|amount|cost).*(quoted|evaluated|total)", cell, re.I)), None)
            if bidder_col is None or total_col is None:
                continue
            for row in cleaned[header_index + 1:]:
                if max(bidder_col, total_col) >= len(row):
                    continue
                bidder = C.clean(row[bidder_col])
                raw_amount = C.clean(row[total_col])
                if not bidder or not raw_amount or re.search(r"\b(total|grand total)\b", bidder, re.I):
                    continue
                amount, currency = _parse_amount_currency(raw_amount)
                if amount is not None and amount > 0:
                    found.append({"bidder": bidder, "amount": amount, "currency": currency,
                                  "raw_amount": raw_amount, "source": "TET report quoted-total table"})
            if found:
                return found
    return found


def _extract_technical_scores(tables):
    """Extract explicit numeric technical scores; compliance labels are not converted to marks."""
    scores = {}
    for rows in tables:
        cleaned = [[C.clean_multiline(cell) for cell in row] for row in rows]
        for header_index, header in enumerate(cleaned[:4]):
            bidder_col = next((i for i, cell in enumerate(header)
                               if re.search(r"\b(bidder|tenderer|supplier|vendor|company)\b", cell, re.I)), None)
            score_col = next((i for i, cell in enumerate(header)
                              if re.search(r"technical.*score|score.*technical|technical.*marks|marks.*technical", cell, re.I)), None)
            if bidder_col is None or score_col is None:
                continue
            max_col = next((i for i, cell in enumerate(header)
                            if re.search(r"(max|maximum).*(score|marks)|(score|marks).*(max|maximum)", cell, re.I)), None)
            for row in cleaned[header_index + 1:]:
                if max(bidder_col, score_col) >= len(row):
                    continue
                bidder = C.clean(row[bidder_col])
                raw_score = C.clean(row[score_col])
                if not bidder or not raw_score:
                    continue
                ratio = re.search(r"(-?\d+(?:\.\d+)?)\s*/\s*(-?\d+(?:\.\d+)?)", raw_score)
                if ratio and float(ratio.group(2)) > 0:
                    score = float(ratio.group(1)) / float(ratio.group(2)) * 100
                else:
                    numeric = re.search(r"-?\d+(?:\.\d+)?", raw_score.replace(",", ""))
                    if not numeric:
                        continue
                    score = float(numeric.group(0))
                    if max_col is not None and max_col < len(row):
                        maximum = re.search(r"\d+(?:\.\d+)?", row[max_col].replace(",", ""))
                        if maximum and float(maximum.group(0)) > 0:
                            score = score / float(maximum.group(0)) * 100
                if 0 <= score <= 100:
                    scores[bidder] = round(score, 4)
    return scores


# A technical report often contains only compliance labels, even though the
# decision makers still need a comparable technical score for every bidder.
# These are the marks prescribed by the SLT procurement manual.  The final
# comparable score is normalized to /100 below, but these marks are the actual
# values used for every compliance criterion.
MAX_COMPLIANCE_MARK = 10.0
COMPLIANCE_SCORE = {
    "ok": 10.0,
    "partial": 5.0,
    "bad": 0.0,
}


def _calculated_technical_scores(matrices):
    """Calculate a percentage score from explicit matrix compliance labels.

    Blank, N/A and unrecognised cells are omitted from the denominator.  The
    returned counts make the result auditable in the API and workbook.
    """
    stats = {}
    for matrix in matrices:
        bidders = matrix.get("headers", [])[1:]
        for bidder in bidders:
            stats.setdefault(bidder, {"earned": 0.0, "criteria": 0,
                                      "complied": 0, "partial": 0,
                                      "not_complied": 0})
        for _area, values in matrix.get("rows", []):
            for index, value in enumerate(values):
                if index >= len(bidders):
                    continue
                status = status_class(value)
                points = COMPLIANCE_SCORE.get(status)
                if points is None:
                    continue
                item = stats[bidders[index]]
                item["earned"] += points
                item["criteria"] += 1
                item[{"ok": "complied", "partial": "partial",
                      "bad": "not_complied"}[status]] += 1

    scores = {}
    for bidder, item in stats.items():
        criteria = item["criteria"]
        if criteria:
            scores[bidder] = {
                "score": round(item["earned"] / (criteria * MAX_COMPLIANCE_MARK) * 100, 2),
                "source": "calculated_from_compliance",
                "marks_earned": round(item["earned"], 2),
                "maximum_marks": round(criteria * MAX_COMPLIANCE_MARK, 2),
                "criteria_scored": criteria,
                "complied": item["complied"],
                "partially_complied": item["partial"],
                "not_complied": item["not_complied"],
            }
        else:
            scores[bidder] = {
                "score": None,
                "source": "not_available",
                "marks_earned": 0,
                "maximum_marks": 0,
                "criteria_scored": 0,
                "complied": 0,
                "partially_complied": 0,
                "not_complied": 0,
            }
    return scores


def _bidder_match_key(value):
    words = re.findall(r"[a-z0-9]+", str(value or "").casefold().replace("&", " and "))
    suffixes = {("private", "limited"), ("private", "ltd"), ("pvt", "limited"),
                ("pvt", "ltd"), ("limited",), ("ltd",), ("plc",),
                ("incorporated",), ("inc",), ("llc",), ("company",), ("co",)}
    changed = True
    while words and changed:
        changed = False
        for suffix in sorted(suffixes, key=len, reverse=True):
            if len(words) >= len(suffix) and tuple(words[-len(suffix):]) == suffix:
                words = words[:-len(suffix)]
                changed = True
                break
    return "".join(words)


def _technical_verdict(value):
    """Return 'eligible', 'excluded', or None from an explicit overall qualification statement."""
    text = C.clean(value).casefold()
    if not text:
        return None
    if re.search(r"not\s+(?:technically\s+)?qualif|disqualif|not eligible|failed|fail(?:ed)?|reject(?:ed)?", text):
        return "excluded"
    if re.search(r"technically\s+qualified|qualified|eligible|proceed|pass(?:ed)?|responsive", text):
        return "eligible"
    return None

def _is_status(t):
    return bool(STATUS_START.match(C.clean(t)))

def _share(cells):
    cells = [c for c in cells if C.clean(c)]
    return (sum(_is_status(c) for c in cells) / len(cells), len(cells)) if cells else (0, 0)

_BIDDER_WORDS = re.compile(r"\b(bidder|tenderer|supplier|vendor|company|manufacturer)\b", re.I)
_BID_DETAIL_WORDS = re.compile(
    r"\b(name|brand|model|make|address|country|origin|registration|contact|"
    r"authorization|authority|validity|quotation|offer|bid|price|agent)\b", re.I)

def _bid_details(rows):
    """Return a bidder-detail table as headers and rows, or None for other tables."""
    if len(rows) < 2:
        return None
    cleaned = [[C.clean_multiline(c) for c in row] for row in rows]
    for header_index in range(min(3, len(cleaned) - 1)):
        header = cleaned[header_index]
        header_text = " ".join(header)
        if not _BIDDER_WORDS.search(header_text) or not _BID_DETAIL_WORDS.search(header_text):
            continue
        headers = [C.clean(x).replace("\n", " ") for x in header]
        body = [
            row for row in cleaned[header_index + 1:]
            if any(C.clean(x) for x in row[1:])
            or (C.clean(row[0]) and not re.fullmatch(r"\d+", C.clean(row[0])))
        ]
        if not body:
            continue
        status_share, status_count = _share([cell for row in body for cell in row[1:]])
        if status_count >= 2 and status_share >= 0.6:
            continue
        width = len(headers)
        normalized_rows = [(row + [""] * width)[:width] for row in body]
        bidder_col = next((i for i, value in enumerate(headers)
                           if re.search(r"\b(bidder|tenderer|supplier|vendor|company)\b", value, re.I)), 0)
        # Downstream evaluation/report code expects the bidder name in column zero.
        # Some official-style bidder tables place S/No before the Bidder column.
        if bidder_col:
            headers = [headers[bidder_col]] + headers[:bidder_col] + headers[bidder_col + 1:]
            normalized_rows = [
                [row[bidder_col]] + row[:bidder_col] + row[bidder_col + 1:]
                for row in normalized_rows
            ]
        return dict(headers=headers, rows=normalized_rows)
    return None

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
        bids = _bid_details(rows)
        if bids:
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
    return dict(
        type=rtype, title=title, meta=meta, bids=bids, matrices=matrices,
        commercial_prices=_extract_commercial_prices(tables),
        technical_scores=_extract_technical_scores(tables),
    )

def eval_to_json(label, ev):
    def cell(t):
        return dict(t=t, c=status_class(t))

    all_bidders, issues, stated_results, mandatory_issues = [], {}, {}, {}
    explicit_results = {}
    criterion_results = {}

    def ensure(bidder):
        if bidder and bidder not in all_bidders:
            all_bidders.append(bidder)
        issues.setdefault(bidder, [])
        stated_results.setdefault(bidder, [])
        mandatory_issues.setdefault(bidder, [])
        explicit_results.setdefault(bidder, [])
        criterion_results.setdefault(bidder, [])

    for matrix in ev["matrices"]:
        for bidder in matrix["headers"][1:]:
            ensure(bidder)
        for area, values in matrix["rows"]:
            mandatory = bool(re.search(r"\bmandatory\b|drop[ -]?dead|\[(?:M|MANDATORY)\]|\((?:M|MANDATORY)\)", area, re.I))
            for index, value in enumerate(values):
                if index + 1 >= len(matrix["headers"]):
                    continue
                bidder = matrix["headers"][index + 1]
                status = status_class(value)
                if status in ("bad", "partial"):
                    issue = dict(area=area, status=value)
                    issues[bidder].append(issue)
                    criterion_results[bidder].append(status)
                    if mandatory and status == "bad":
                        mandatory_issues[bidder].append(issue)
                elif status == "ok":
                    criterion_results[bidder].append(status)

        if matrix["overall"]:
            for index, value in enumerate(matrix["overall"]):
                if index + 1 >= len(matrix["headers"]) or not C.clean(value):
                    continue
                bidder = matrix["headers"][index + 1]
                stated_results[bidder].append(value)
                verdict = _technical_verdict(value)
                if verdict:
                    explicit_results[bidder].append(verdict)
                if verdict == "excluded":
                    issues[bidder].append(dict(area="Overall result", status=value))

    # Retain every bidder named in the TET report's bidder-details table.
    if ev.get("bids"):
        headers = ev["bids"].get("headers", [])
        status_columns = [i for i, header in enumerate(headers)
                          if re.search(r"technical.*status|overall.*result|qualification.*status|evaluation.*result", header, re.I)]
        for row in ev["bids"].get("rows", []):
            bidder = C.clean(row[0]) if row else ""
            if not bidder:
                continue
            existing = next((name for name in all_bidders
                             if _bidder_match_key(name) == _bidder_match_key(bidder)), None)
            canonical_bidder = existing or bidder
            ensure(canonical_bidder)
            for col in status_columns:
                if col < len(row):
                    verdict = _technical_verdict(row[col])
                    if verdict:
                        explicit_results[canonical_bidder].append(verdict)
                        stated_results[canonical_bidder].append(row[col])

    calculated_scores = _calculated_technical_scores(ev["matrices"])
    reported_scores = ev.get("technical_scores") or {}

    def score_for(bidder):
        reported = next((score for name, score in reported_scores.items()
                         if _bidder_match_key(name) == _bidder_match_key(bidder)), None)
        if reported is not None:
            calculated = calculated_scores.get(bidder, {})
            return {
                "score": reported,
                "source": "reported_in_technical_report",
                "marks_earned": calculated.get("marks_earned", 0),
                "maximum_marks": calculated.get("maximum_marks", 0),
                "criteria_scored": calculated.get("criteria_scored", 0),
                "complied": calculated.get("complied", 0),
                "partially_complied": calculated.get("partially_complied", 0),
                "not_complied": calculated.get("not_complied", 0),
            }
        return calculated_scores.get(bidder, {
            "score": None, "source": "not_available", "marks_earned": 0,
            "maximum_marks": 0, "criteria_scored": 0,
            "complied": 0, "partially_complied": 0, "not_complied": 0,
        })

    recommendations = []
    eligible, excluded, review = [], [], []
    for bidder in all_bidders:
        results = set(explicit_results[bidder])
        has_conflict = len(results) > 1
        if mandatory_issues[bidder]:
            eligibility_status = "excluded"
            reason = "Failed an explicitly identified mandatory/drop-dead criterion."
        elif has_conflict:
            eligibility_status = "review_required"
            reason = "The report contains conflicting overall qualification statements."
        elif results == {"excluded"}:
            eligibility_status = "excluded"
            reason = "The TET report explicitly records the bidder as not qualified/disqualified."
        elif results == {"eligible"}:
            eligibility_status = "eligible"
            reason = "The TET report explicitly records the bidder as technically qualified."
        elif criterion_results[bidder] and all(x == "ok" for x in criterion_results[bidder]):
            eligibility_status = "eligible"
            reason = "All extracted technical compliance entries are marked compliant; no separate overall result was found."
        else:
            eligibility_status = "review_required"
            reason = "The report does not provide enough explicit information to confirm technical eligibility; review the TET decision."

        if eligibility_status == "eligible":
            eligible.append(bidder)
            recommendation = "Proceed to commercial evaluation"
        elif eligibility_status == "excluded":
            excluded.append(bidder)
            recommendation = "Do not proceed to commercial evaluation"
        else:
            review.append(bidder)
            recommendation = "Manual TET review required before commercial evaluation"

        technical_score_details = score_for(bidder)
        recommendations.append(dict(
            bidder=bidder, eligible=eligibility_status == "eligible",
            eligibility_status=eligibility_status, eligibility_reason=reason,
            technical_score=technical_score_details["score"],
            technical_score_source=technical_score_details["source"],
            technical_marks=technical_score_details["marks_earned"],
            maximum_technical_marks=technical_score_details["maximum_marks"],
            criteria_scored=technical_score_details["criteria_scored"],
            complied=technical_score_details["complied"],
            partially_complied=technical_score_details["partially_complied"],
            not_complied=technical_score_details["not_complied"],
            recommendation=recommendation, stated_results=stated_results[bidder],
            issues=issues[bidder], mandatory_issues=mandatory_issues[bidder],
        ))

    recommendation_table = _recommendation_table(all_bidders, set(excluded + review), ev.get("bids"))
    return dict(
        label=label, type=ev["type"], title=ev["title"], meta=ev["meta"], bids=ev.get("bids"),
        shortlist=dict(eligible=eligible, excluded=excluded, review_required=review),
        recommendations=recommendations, recommendation_table=recommendation_table,
        commercial_prices=ev.get("commercial_prices", []), technical_scores=ev.get("technical_scores", {}),
        technical_score_basis=(
            "Explicit numeric technical score from the report when available; "
            "otherwise calculated using SLT procurement marks: Comply=10, "
            "Partial Comply=5, Non Comply=0. Blank/N/A/unclear cells are "
            "excluded; calculated marks are normalized to a score out of 100."
        ),
        matrices=[dict(headers=m["headers"],
                       rows=[dict(area=a, cells=[cell(v) for v in vals]) for a, vals in m["rows"]],
                       overall=[cell(v) for v in m["overall"]] if m["overall"] else None,
                       eligible=[next((r["eligibility_status"] == "eligible" for r in recommendations
                                       if _bidder_match_key(r["bidder"]) == _bidder_match_key(m["headers"][1 + j])), False)
                                 for j in range(len(m["headers"]) - 1)])
                for m in ev["matrices"]],
    )

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

def _build_technical_score_sheet(wb, sheet_name, label, records, tender_no=""):
    """Add a dedicated, auditable score summary worksheet."""
    ws = wb.create_sheet(sheet_name)
    ws["A1"] = "TECHNICAL EVALUATION SCORES"
    ws["A1"].font = Font(bold=True, size=14, color=C.NAVY)
    ws["A2"] = f"Source: {label}" + (f"  |  Tender No: {tender_no}" if tender_no else "")
    ws["A2"].font = Font(italic=True, color="595959")
    headers = ["Company", "Technical score /100", "Marks earned", "Maximum marks",
               "Score source", "Criteria scored", "Comply", "Partial Comply",
               "Non Comply", "Eligibility"]
    for column, header in enumerate(headers, 1):
        C.hcell(ws, 4, column, header)
    for row_number, record in enumerate(records, 5):
        source = record.get("technical_score_source")
        source_text = ("Reported in technical report" if source == "reported_in_technical_report"
                       else "Calculated from compliance labels" if source == "calculated_from_compliance"
                       else "Not available")
        values = [
            record.get("bidder", ""), record.get("technical_score"),
            record.get("technical_marks", 0), record.get("maximum_technical_marks", 0),
            source_text, record.get("criteria_scored", 0), record.get("complied", 0),
            record.get("partially_complied", 0), record.get("not_complied", 0),
            record.get("eligibility_status", "review_required").replace("_", " "),
        ]
        for column, value in enumerate(values, 1):
            C.body(ws, row_number, column, value, bold=(column == 1), center=(column > 1))
    note_row = max(5, 5 + len(records))
    C.body(ws, note_row, 1, "Scoring basis: Comply=10, Partial Comply=5, Non Comply=0. Blank/N/A/unclear cells are excluded. Explicit numeric report scores take precedence.")
    ws.merge_cells(start_row=note_row, start_column=1, end_row=note_row, end_column=len(headers))
    ws.cell(note_row, 1).font = Font(italic=True, size=9, color="595959")
    widths = [28, 20, 16, 16, 32, 16, 12, 18, 14, 20]
    for column, width in enumerate(widths, 1):
        ws.column_dimensions[L(column)].width = width
    C.finish(ws, "A5")
    return ws


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
        eligibility_records = eval_to_json(label, ev)["recommendations"]
        _build_technical_score_sheet(
            wb,
            C.safe_sheet((f"R{idx} " if multi else "") + "Technical Scores", used),
            label,
            eligibility_records,
            tender_no=meta.get("Tender No.", ""),
        )
        eligibility_by_name = {_bidder_match_key(item["bidder"]): item
                               for item in eligibility_records}
        # A non-mandatory issue is not an automatic disqualification. Only bidders that
        # the TET report excludes, or whose status needs review, are held from price opening.
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
                    if _is_ineligible_status(v):
                        nc_rows.append((mx["headers"][1 + j], area, v))
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
                if _is_ineligible_status(v):
                    nc_rows.append((mx["headers"][1 + j], "Overall result", v))
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
            # Eligibility follows the explicit TET overall result and mandatory/drop-dead
            # criteria; other issues remain visible for the evaluator's documented review.
            elig_row = r
            C.body(ws, r, 1, "Eligible for Commercial Evaluation", bold=True)
            for j in range(n):
                bidder = mx["headers"][1 + j]
                record = eligibility_by_name.get(_bidder_match_key(bidder), {})
                status = record.get("eligibility_status", "review_required")
                value = ("Yes — proceed" if status == "eligible" else
                         "No — excluded" if status == "excluded" else
                         "Review — TET decision required")
                C.body(ws, r, 2 + j, value, bold=True, center=True)
            _cf_eligible(ws, f"B{elig_row}:{L(1+n)}{elig_row}", f"B{elig_row}")
            r += 1
            C.body(ws, r, 1, "Flag: eligibility evidence / TET review", bold=False)
            for j in range(n):
                bidder = mx["headers"][1 + j]
                record = eligibility_by_name.get(_bidder_match_key(bidder), {})
                flag = record.get("eligibility_reason", "") if record.get("eligibility_status") != "eligible" else ""
                x = C.body(ws, r, 2 + j, flag, center=True)
                if flag:
                    x.font = Font(italic=True, size=9, color="B42318")
            r += 3
        ws.column_dimensions["A"].width = 36
        ws.column_dimensions["A"].width = 58
        ncols = max(max(len(m["headers"]) for m in ev["matrices"]), len(all_bidders) + 1)
        for j in range(2, ncols + 1):
            ws.column_dimensions[L(j)].width = 22
        C.finish(ws, f"B{first_header_row+1}")
        ws.cell(r, 1, f"Generated {datetime.now():%Y-%m-%d %H:%M} from {label}. Cells are editable (drop-down lists); counts and review checks update automatically.").font = Font(italic=True, size=9, color="7F7F7F")

        not_eligible = {
            bidder for bidder in all_bidders
            if not any(item["eligibility_status"] == "eligible"
                       and _bidder_match_key(item["bidder"]) == _bidder_match_key(bidder)
                       for item in eligibility_records)
        }

        # ---- non-compliances (report instruction 4b): a record of why each bidder was excluded.
        # Non-compliant bidders do NOT go forward to commercial evaluation, so this sheet
        # is an audit trail rather than a worksheet for deciding whether to accept them.
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

        # ---- commercial evaluation shortlist: bidders with no non-compliance.
        ws4 = wb.create_sheet(C.safe_sheet(pre + "Commercial Eval Shortlist", used))
        ws4["A1"] = "SHORTLIST FOR COMMERCIAL EVALUATION — no non-compliance"
        ws4["A1"].font = Font(bold=True, size=13, color=C.NAVY)
        ws4["A2"] = "Bidders with partial compliance may proceed; bidders with non-compliance are excluded automatically."
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

        wb.remove(ws2); wb.remove(ws4)        # Excel shows the same table as the screen: no extra sheets

        # ---- bid details (not shown on screen, so not in the Excel either)
        if ev["bids"] and False:
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
