"""Post-technical commercial comparison and award recommendation support.

The module deliberately separates a price-ranked front-runner from a final award
recommendation.  A weighted recommendation is only calculated when the uploaded
records contain a TET technical score, a commercial score from price schedules or
an explicitly labelled evaluation result for each eligible bidder, and approved
technical/commercial weightings that total 100%.
The final award remains subject to the appropriate SLT Tender Board/DOA approval.
"""
from __future__ import annotations

import re
import tempfile
import zipfile
from collections import defaultdict
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    from . import evalreport
except ImportError:  # pragma: no cover - direct module use
    import evalreport

try:
    from ..score_calculator import calculate_weighted_scores
except ImportError:  # pragma: no cover - direct module use
    from score_calculator import calculate_weighted_scores


def bidder_key(value: str) -> str:
    """Normalise bidder labels, including common legal suffixes, for exact matching."""
    words = re.findall(r"[a-z0-9]+", str(value or "").casefold().replace("&", " and "))
    # Uploaded filenames frequently become the fallback bidder label.  Remove
    # document-role words before matching, otherwise "Acme" and
    # "Acme Price Document" are treated as different bidders.
    document_suffixes = {
        ("price", "document"), ("price", "documents"), ("price", "schedule"), ("price", "doc"),
        ("commercial", "document"), ("commercial", "evaluation"),
        ("gcc", "form"), ("scc", "form"), ("gct", "form"), ("form",),
    }
    changed = True
    while words and changed:
        changed = False
        for suffix in sorted(document_suffixes, key=len, reverse=True):
            if len(words) >= len(suffix) and tuple(words[-len(suffix):]) == suffix:
                words = words[:-len(suffix)]
                changed = True
                break
    suffixes = {
        ("private", "limited"), ("private", "ltd"), ("pvt", "limited"), ("pvt", "ltd"),
        ("proprietary", "limited"), ("proprietary", "ltd"), ("(pvt)", "ltd"),
        ("limited",), ("ltd",), ("plc",), ("incorporated",), ("inc",), ("llc",),
        ("company",), ("co",),
    }
    changed = True
    while words and changed:
        changed = False
        for suffix in sorted(suffixes, key=len, reverse=True):
            if len(words) >= len(suffix) and tuple(words[-len(suffix):]) == suffix:
                words = words[:-len(suffix)]
                changed = True
                break
    return "".join(words)


def _parse_number(value):
    if value is None:
        return None
    raw = str(value).replace(",", "").replace(" ", "")
    match = re.search(r"-?\d+(?:\.\d+)?", raw)
    return float(match.group(0)) if match else None


def _score_value(value, maximum=None):
    """Return a 0-100 score only when the document gives a numeric score."""
    raw = str(value or "").strip().replace(",", "")
    ratio = re.search(r"(-?\d+(?:\.\d+)?)\s*/\s*(-?\d+(?:\.\d+)?)", raw)
    if ratio and float(ratio.group(2)) > 0:
        score = float(ratio.group(1)) / float(ratio.group(2)) * 100
    else:
        score = _parse_number(raw)
        max_value = _parse_number(maximum)
        if score is None:
            return None
        if max_value is not None:
            if max_value <= 0:
                return None
            score = score / max_value * 100
        elif "%" not in raw and score > 100:
            return None
    return round(score, 4) if 0 <= score <= 100 else None


def _read_normalized_document(path: Path):
    """Read text and tables from the normalized .docx/.xlsx files produced by collect()."""
    text_parts, tables = [], []
    suffix = path.suffix.casefold()
    if suffix == ".docx":
        from docx import Document
        doc = Document(path)
        text_parts.extend(p.text for p in doc.paragraphs if p.text)
        for table in doc.tables:
            rows = [[cell.text for cell in row.cells] for row in table.rows]
            tables.append(rows)
            text_parts.extend(" | ".join(row) for row in rows)
    elif suffix in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        workbook = load_workbook(path, read_only=True, data_only=True)
        for sheet in workbook.worksheets:
            rows = [[cell for cell in row] for row in sheet.iter_rows(values_only=True)]
            tables.append([["" if cell is None else str(cell) for cell in row] for row in rows])
            text_parts.extend(" | ".join(str(cell) for cell in row if cell is not None) for row in rows)
        workbook.close()
    return "\n".join(text_parts), tables


def _extract_scores_from_tables(tables):
    """Extract explicitly named commercial/financial scores from submitted evaluation docs."""
    scores = {}
    evidence = {}
    for rows in tables:
        for header_idx, row in enumerate(rows[:4]):
            header = [str(v or "").strip() for v in row]
            bidder_col = next((i for i, value in enumerate(header)
                               if re.search(r"\b(bidder|tenderer|supplier|vendor|company)\b", value, re.I)), None)
            score_col = next((i for i, value in enumerate(header)
                               if re.search(r"commercial.*(score|mark)|financial.*(score|mark)|score.*(commercial|financial)", value, re.I)), None)
            if bidder_col is None or score_col is None:
                continue
            max_col = next((i for i, value in enumerate(header)
                            if re.search(r"(max|maximum).*(score|mark)|(score|mark).*(max|maximum)", value, re.I)), None)
            for values in rows[header_idx + 1:]:
                if max(bidder_col, score_col) >= len(values):
                    continue
                bidder = str(values[bidder_col] or "").strip()
                raw = str(values[score_col] or "").strip()
                if not bidder or not raw:
                    continue
                maximum = values[max_col] if max_col is not None and max_col < len(values) else None
                score = _score_value(raw, maximum)
                if score is not None:
                    scores[bidder_key(bidder)] = score
                    evidence[bidder_key(bidder)] = {"bidder": bidder, "raw_score": raw, "score": score}
    return scores, evidence


def _has_explicit_approval(text):
    positive = bool(re.search(
        r"\b(?:approved\s+by|approved\s+the|has\s+approved|was\s+approved|"
        r"approval\s+granted|approval\s+(?:was\s+)?received|signed\s+off\s+by|endorsed\s+by|resolved\s+to)\b",
        text, re.I,
    ))
    negative = bool(re.search(
        r"\b(?:not\s+approved|pending\s+approval|subject\s+to\s+approval|"
        r"to\s+be\s+approved|will\s+be\s+approved|for\s+approval|should\s+be\s+approved)\b",
        text, re.I,
    ))
    return positive and not negative


def _extract_weights(text, tables, source_name):
    """Extract explicitly labelled technical/commercial weightages from a tender strategy."""
    technical = commercial = None
    source_has_weight_context = bool(re.search(r"weight(?:age|ing)?|tender strategy|evaluation criteria", text, re.I))

    # Prefer table rows: these usually state the criterion and percentage separately.
    for rows in tables:
        for row in rows:
            joined = " ".join(str(cell or "") for cell in row)
            pct = re.search(r"(\d+(?:\.\d+)?)\s*%", joined)
            if not pct:
                continue
            value = float(pct.group(1))
            if re.search(r"technical", joined, re.I) and re.search(r"weight|evaluation|score|criteria", joined, re.I):
                technical = value
            if re.search(r"commercial|financial|price|cost", joined, re.I) and re.search(r"weight|evaluation|score|criteria", joined, re.I):
                commercial = value

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    for line in lines:
        if not re.search(r"weight(?:age|ing)?", line, re.I):
            continue
        value = re.search(r"(\d+(?:\.\d+)?)\s*%", line)
        if not value:
            continue
        pct = float(value.group(1))
        if re.search(r"technical", line, re.I):
            technical = pct
        if re.search(r"commercial|financial|price|cost", line, re.I):
            commercial = pct

    # Pair-style tender strategy statements, e.g. "Technical 30%; Commercial 70%".
    if source_has_weight_context and (technical is None or commercial is None):
        compact = " ".join(lines)
        pair = re.search(
            r"technical(?:\s+evaluation)?[^%]{0,45}?(\d+(?:\.\d+)?)\s*%[^%]{0,100}?"
            r"(?:commercial|financial|price|cost)(?:\s+evaluation)?[^%]{0,45}?(\d+(?:\.\d+)?)\s*%",
            compact, re.I,
        )
        if pair:
            technical = technical if technical is not None else float(pair.group(1))
            commercial = commercial if commercial is not None else float(pair.group(2))
        reverse_pair = re.search(
            r"(?:commercial|financial|price|cost)(?:\s+evaluation)?[^%]{0,45}?(\d+(?:\.\d+)?)\s*%[^%]{0,100}?"
            r"technical(?:\s+evaluation)?[^%]{0,45}?(\d+(?:\.\d+)?)\s*%",
            compact, re.I,
        )
        if reverse_pair:
            commercial = commercial if commercial is not None else float(reverse_pair.group(1))
            technical = technical if technical is not None else float(reverse_pair.group(2))

    if technical is None or commercial is None:
        return None
    approved = _has_explicit_approval(text)
    if technical < 0 or commercial < 0 or abs(technical + commercial - 100) > 0.01:
        return {"technical": technical, "commercial": commercial, "valid": False,
                "approved": approved, "source": source_name,
                "reason": "Extracted technical and commercial weightages do not total 100%."}
    return {"technical": technical, "commercial": commercial, "valid": True,
            "approved": approved, "source": source_name}


def _extract_explicit_award_confirmation(text, tables, source_name):
    """Extract only an explicit award/approval notice, never a TET recommendation."""
    source = Path(source_name).name
    board_approval = bool(
        re.search(r"(?:tender board|joint tender board|\bJTB\b|\bSTB\b|board of directors)", text, re.I)
        and _has_explicit_approval(text)
        and re.search(r"\b(award|awarded|successful bidder)\b", text, re.I)
    )
    award_notice = bool(
        re.search(r"letter of award|notice of award|award notification", source + "\n" + text, re.I)
        and re.search(r"\b(awarded|successful bidder|selected bidder|award is made)\b", text, re.I)
    )
    if not (board_approval or award_notice):
        return None

    candidates = []
    evidence_lines = []
    patterns = [
        r"\b(?:the\s+)?(?:tender|contract)\s+(?:is\s+)?(?:hereby\s+)?awarded\s+to\s+(.+?)(?=[.;\n]|$)",
        r"\b(?:awarded|successful|selected)\s+(?:bidder|tenderer|supplier)\s*[:\-]\s*(.+?)(?=[.;\n]|$)",
        r"\b(?:tender|contract)\s+award\s+(?:is\s+)?(?:approved\s+)?to\s+(.+?)(?=[.;\n]|$)",
        r"\baward(?:ed|ing)?\s+to\s+(.+?)(?=[.;\n]|$)",
        r"\baward(?:ed)?\s+in\s+favour\s+of\s+(.+?)(?=[.;\n]|$)",
    ]
    for line in text.splitlines():
        clean_line = " ".join(line.split())
        if not clean_line:
            continue
        for pattern in patterns:
            match = re.search(pattern, clean_line, re.I)
            if not match:
                continue
            candidate = match.group(1).strip(" \t:,-")
            candidate = re.sub(r"^(?:M/s\.?|Messrs\.?|The)\s+", "", candidate, flags=re.I)
            candidate = re.sub(r"\s+(?:for the sum|for the amount|at the price|with effect|subject to|in accordance with).*$", "", candidate, flags=re.I)
            if len(candidate) >= 3 and not re.search(r"recommend|propos|subject to approval|for approval", candidate, re.I):
                candidates.append(candidate)
                evidence_lines.append(clean_line[:500])
                break

    # A structured award decision table can be clearer than prose. Require an
    # explicit approval context above and an unambiguous awarded/successful column.
    for rows in tables:
        for header_index, header in enumerate(rows[:4]):
            award_col = next((i for i, cell in enumerate(header)
                              if re.search(r"\b(awarded bidder|successful bidder|awarded supplier|selected bidder)\b", str(cell or ""), re.I)), None)
            if award_col is None:
                continue
            for row in rows[header_index + 1:]:
                if award_col < len(row):
                    candidate = str(row[award_col] or "").strip()
                    if candidate and len(candidate) >= 3:
                        candidates.append(candidate)
                        evidence_lines.append("Award decision table: " + " | ".join(str(v or "") for v in row)[:500])
            if candidates:
                break

    unique = {}
    for candidate in candidates:
        key = bidder_key(candidate)
        if key:
            unique.setdefault(key, candidate)
    if len(unique) != 1:
        if unique:
            return {"status": "conflicting", "bidders": list(unique.values()), "source": source,
                    "evidence": evidence_lines[:5]}
        return None
    return {"status": "explicit", "bidder": next(iter(unique.values())), "source": source,
            "evidence": evidence_lines[:3], "approval_basis": "Tender Board/authority approval wording or formal award notice"}


def inspect_input_documents(inputs, work_dir=None):
    """Scan uploaded documents (and ZIP contents) for approved weightages/CET scores.

    Uses the same normalisation/collection helper as the existing application so PDF,
    DOC, and XLS input formats can be inspected where the host has the needed tools.
    """
    try:
        from ..goods_services import consolidate
    except ImportError:  # pragma: no cover
        from goods_services import consolidate

    temp_root = Path(work_dir or tempfile.mkdtemp(prefix="award-scan-"))
    temp_root.mkdir(parents=True, exist_ok=True)
    problems = []
    collected = consolidate.collect(inputs, str(temp_root), problems)
    weights_found, commercial_scores, score_evidence, award_confirmations = [], {}, {}, []
    for _bidder, normalized_path, original_path in collected:
        try:
            text, tables = _read_normalized_document(Path(normalized_path))
        except Exception as exc:
            problems.append(f"{Path(original_path).name}: award metadata scan failed: {exc}")
            continue
        # Only recognize an explicit strategy/evaluation weighting record.
        weights = _extract_weights(text, tables, Path(original_path).name)
        if weights:
            weights_found.append(weights)
        award = _extract_explicit_award_confirmation(text, tables, Path(original_path).name)
        if award:
            award_confirmations.append(award)
        scores, evidence = _extract_scores_from_tables(tables)
        commercial_scores.update(scores)
        score_evidence.update(evidence)

    valid_weights = [item for item in weights_found if item.get("valid")]
    invalid_weights = [item for item in weights_found if not item.get("valid")]
    weights = None
    if invalid_weights:
        weights = {"valid": False, "approved": False,
                   "source": ", ".join(x["source"] for x in weights_found),
                   "reason": "An incomplete or invalid technical/commercial weighting record was found; resolve it before scoring."}
    elif valid_weights:
        unique = {(round(x["technical"], 4), round(x["commercial"], 4)) for x in valid_weights}
        approval_states = {bool(x.get("approved")) for x in valid_weights}
        if len(unique) == 1 and len(approval_states) == 1:
            weights = valid_weights[0]
        else:
            weights = {"valid": False, "approved": False,
                       "source": ", ".join(x["source"] for x in valid_weights),
                       "reason": "Conflicting technical/commercial weightages or approval evidence were found in the uploaded documents."}
    return {"weights": weights, "commercial_scores": commercial_scores,
            "commercial_score_evidence": score_evidence,
            "award_confirmations": award_confirmations, "problems": problems}


def _technical_results(evaluations):
    result = {}
    sources = defaultdict(list)
    for evaluation in evaluations or []:
        for rec in evaluation.get("recommendations", []):
            bidder = rec.get("bidder", "")
            key = bidder_key(bidder)
            if not key:
                continue
            sources[key].append(evaluation.get("label", "TET report"))
            current = result.setdefault(key, {
                "bidder": bidder, "eligibility_status": rec.get("eligibility_status", "review_required"),
                "eligibility_reason": rec.get("eligibility_reason", "No eligibility conclusion was extracted."),
                "technical_score": rec.get("technical_score"), "stated_results": [], "issues": [],
            })
            incoming_status = rec.get("eligibility_status", "review_required")
            if current["eligibility_status"] != incoming_status:
                current["eligibility_status"] = "review_required"
                current["eligibility_reason"] = "Conflicting technical eligibility results were found across uploaded TET reports."
            if current["technical_score"] is None and rec.get("technical_score") is not None:
                current["technical_score"] = rec["technical_score"]
            current["stated_results"].extend(rec.get("stated_results", []))
            current["issues"].extend(rec.get("issues", []))
    for key, item in result.items():
        item["sources"] = list(dict.fromkeys(sources[key]))
    return result


def build_award_assessment(evaluations, price_rows, source_metadata=None):
    """Build an evidence-based commercial ranking / award recommendation payload."""
    metadata = source_metadata or {}
    technical = _technical_results(evaluations)
    eligible_keys = {key for key, item in technical.items() if item["eligibility_status"] == "eligible"}
    excluded_keys = {key for key, item in technical.items() if item["eligibility_status"] == "excluded"}
    review_keys = set(technical) - eligible_keys - excluded_keys

    # Prefer extracted completed price schedules. Use an explicitly labelled quoted-total
    # table in the TET report only when no price schedule was extracted for that bidder.
    price_map = defaultdict(list)
    for row in price_rows or []:
        key = bidder_key(row.get("bidder"))
        if not key:
            continue
        total = row.get("evaluated_total", row.get("total"))
        try:
            total = float(total) if total is not None else None
        except (TypeError, ValueError):
            total = None
        price_map[key].append({
            "bidder": row.get("bidder"), "currency": str(row.get("currency") or "").upper(),
            "amount": total, "source": row.get("source") or "Price schedule",
            "warnings": list(row.get("warnings") or []),
            "commission_review_required": bool(row.get("commission_review_required")),
            "price_basis": "Parsed price schedule",
        })

    report_prices = defaultdict(list)
    for evaluation in evaluations or []:
        for price in evaluation.get("commercial_prices", []):
            key = bidder_key(price.get("bidder"))
            if key:
                report_prices[key].append({
                    "bidder": price.get("bidder"), "currency": str(price.get("currency") or "").upper(),
                    "amount": price.get("amount"), "source": evaluation.get("label", "TET report"),
                    "warnings": [], "commission_review_required": False,
                    "price_basis": "Quoted total in TET report (fallback)",
                })

    bidders = []
    for key, item in technical.items():
        price_options = price_map.get(key) or report_prices.get(key) or []
        currencies = {p["currency"] for p in price_options if p.get("currency")}
        amounts = [p.get("amount") for p in price_options if p.get("amount") is not None and p.get("amount") > 0]
        price = None
        if len(price_options) == 1 and len(amounts) == 1:
            price = dict(price_options[0])
        elif len(price_options) > 1 and len(currencies) == 1 and len(amounts) == len(price_options):
            # More than one schedule in a single currency is not automatically summed:
            # it may be duplicated or represent different lots. Require manual review.
            price = {"bidder": item["bidder"], "currency": next(iter(currencies)), "amount": None,
                     "source": ", ".join(dict.fromkeys(p["source"] for p in price_options)),
                     "warnings": ["Multiple price records were found for this bidder; confirm whether they are separate lots or duplicate submissions."],
                     "commission_review_required": any(p["commission_review_required"] for p in price_options),
                     "price_basis": "Multiple price records — review required"}
        elif not price_options:
            price = None
        else:
            price = {"bidder": item["bidder"], "currency": ", ".join(sorted(currencies)), "amount": None,
                     "source": ", ".join(dict.fromkeys(p["source"] for p in price_options)),
                     "warnings": ["More than one currency or incomplete price records were found; do not compare until reconciled."],
                     "commission_review_required": any(p["commission_review_required"] for p in price_options),
                     "price_basis": "Currency/record reconciliation required"}

        bidders.append({
            "bidder": item["bidder"], "eligibility_status": item["eligibility_status"],
            "eligibility_reason": item["eligibility_reason"], "technical_score": item["technical_score"],
            "commercial_score": metadata.get("commercial_scores", {}).get(key),
            "price": price, "price_rank": None, "normalized_price_score": None,
            "overall_score": None, "overall_rank": None,
            "sources": item["sources"], "technical_issues": item["issues"],
        })

    # Add a report-only commercial price entry for bidders named in the price table but not
    # present in the TET bidder list; these are not considered eligible until reconciled.
    for key, options in report_prices.items():
        if key not in technical and key not in price_map:
            p = options[0] if len(options) == 1 else None
            bidders.append({"bidder": (p or {}).get("bidder", "Unmatched bidder"),
                            "eligibility_status": "review_required",
                            "eligibility_reason": "The bidder appears in a quoted-price table but was not matched to a TET bidder.",
                            "technical_score": None, "commercial_score": None,
                            "price": p, "price_rank": None, "normalized_price_score": None,
                            "overall_score": None, "overall_rank": None,
                            "sources": [(p or {}).get("source", "TET report")], "technical_issues": []})

    comparable = [b for b in bidders if b["eligibility_status"] == "eligible" and b.get("price")
                  and b["price"].get("amount") is not None and b["price"].get("currency")]
    currencies = {b["price"]["currency"] for b in comparable}
    one_currency = len(currencies) == 1 and len(comparable) >= 1
    price_commercial_scores = {}
    if one_currency:
        for row in price_rows or []:
            score = row.get("price_score")
            key = bidder_key(row.get("bidder"))
            if key and score is not None:
                price_commercial_scores[key] = score
    all_eligible_priced = bool(eligible_keys) and all(
        next((b for b in bidders if bidder_key(b["bidder"]) == key and b.get("price")
              and b["price"].get("amount") is not None and b["price"].get("currency")), None)
        for key in eligible_keys
    )
    commission_review = any(b.get("price", {}).get("commission_review_required")
                            for b in bidders if b.get("price"))

    # Price ranks remain useful within each currency even when a direct
    # cross-currency comparison is not possible.  Previously the award table
    # left every rank blank as soon as one bidder quoted in a different
    # currency, although the price schedule correctly showed per-currency
    # ranks.  Keep the rank scope explicit so it cannot be mistaken for a
    # cross-currency ranking.
    if all_eligible_priced:
        for currency in sorted(currencies):
            currency_bidders = sorted(
                (b for b in comparable if b["price"]["currency"] == currency),
                key=lambda b: (b["price"]["amount"], bidder_key(b["bidder"])),
            )
            last_amount, rank = None, 0
            for position, bidder in enumerate(currency_bidders, 1):
                amount = bidder["price"]["amount"]
                if amount != last_amount:
                    rank = position
                    last_amount = amount
                bidder["price_rank"] = rank
                bidder["price_rank_scope"] = currency
        if one_currency:
            lowest = min(b["price"]["amount"] for b in comparable)
            for bidder in comparable:
                amount = bidder["price"]["amount"]
                bidder["normalized_price_score"] = round((lowest / amount) * 100, 4) if amount else None

    for bidder in bidders:
        price_score = price_commercial_scores.get(bidder_key(bidder["bidder"]))
        if price_score is not None:
            bidder["commercial_score"] = price_score

    weights = metadata.get("weights")
    weights_valid = bool(weights and weights.get("valid") and weights.get("approved"))
    has_technical_scores = bool(eligible_keys) and all(
        next((b for b in bidders if bidder_key(b["bidder"]) == key and b.get("technical_score") is not None), None)
        for key in eligible_keys
    )
    has_commercial_scores = bool(eligible_keys) and all(
        next((b for b in bidders if bidder_key(b["bidder"]) == key and b.get("commercial_score") is not None), None)
        for key in eligible_keys
    )
    # Calculate each bidder independently. A missing score for one bidder must
    # not hide the weighted components that can already be calculated for the
    # others. calculate_weighted_scores leaves the unavailable component and
    # overall_score as None until both source scores are present.
    if weights_valid:
        for bidder in bidders:
            if bidder["eligibility_status"] != "eligible":
                continue
            weighted = calculate_weighted_scores(
                bidder["technical_score"], bidder["commercial_score"],
                weights["technical"], weights["commercial"],
            )
            bidder.update({key: float(value) if value is not None else None
                           for key, value in weighted.items()})
        overall_ranked = sorted((b for b in bidders if b["overall_score"] is not None),
                                key=lambda b: (-b["overall_score"], bidder_key(b["bidder"])))
        rank, previous = 0, None
        for position, bidder in enumerate(overall_ranked, 1):
            if bidder["overall_score"] != previous:
                rank, previous = position, bidder["overall_score"]
            bidder["overall_rank"] = rank

    final_ready = (len(eligible_keys) >= 1 and weights_valid and has_technical_scores and has_commercial_scores
                   and one_currency and all_eligible_priced and not review_keys and not commission_review)
    recommended = None
    if final_ready:
        recommended = next((b for b in bidders if b.get("overall_rank") == 1), None)
        decision = "recommendation_ready" if recommended else "review_required"
    elif one_currency and all_eligible_priced:
        recommended = next((b for b in bidders if b.get("price_rank") == 1), None)
        decision = "provisional_price_leader" if recommended else "review_required"
    else:
        decision = "review_required"

    # A bidder is described as actually awarded only when an uploaded approval/award
    # record explicitly says so. A TET/CET recommendation alone is not an award.
    award_records = metadata.get("award_confirmations", []) or []
    award_candidates = {}
    award_conflict = any(record.get("status") == "conflicting" for record in award_records)
    for record in award_records:
        if record.get("status") == "explicit" and record.get("bidder"):
            award_candidates.setdefault(bidder_key(record["bidder"]), record)
    awarded_bidder = None
    award_confirmation_status = "not_found"
    award_confirmation_sources = []
    if award_conflict or len(award_candidates) > 1:
        award_confirmation_status = "review_required"
    elif len(award_candidates) == 1:
        key, record = next(iter(award_candidates.items()))
        award_confirmation_sources.append(record.get("source", "Award document"))
        matched = next((item for bidder_key_value, item in technical.items() if bidder_key_value == key), None)
        if matched and matched.get("eligibility_status") == "eligible":
            awarded_bidder = matched["bidder"]
            award_confirmation_status = "confirmed_from_document"
        else:
            award_confirmation_status = "review_required"

    blockers = []
    if award_confirmation_status == "review_required":
        blockers.append("An uploaded award/approval record was found but its bidder could not be matched unambiguously to one technically eligible bidder, or uploaded records conflict.")
    if not eligible_keys:
        blockers.append("No bidder has a confirmed technically eligible status in the TET report.")
    if review_keys:
        blockers.append("One or more bidders have unresolved/conflicting technical eligibility; resolve this with the TET before commercial ranking.")
    if not one_currency:
        blockers.append("Eligible bidders' prices are missing, incomplete, or expressed in different currencies; a common-currency comparison is not available.")
    if not all_eligible_priced:
        blockers.append("A parsed evaluated price is not available for every technically eligible bidder.")
    if commission_review:
        blockers.append("At least one price schedule has a separately stated local-agent commission that requires currency/amount confirmation.")
    if not weights_valid:
        if weights and weights.get("valid") and not weights.get("approved"):
            blockers.append("A technical/commercial weighting was found, but the uploaded evidence does not show that it was approved before bid opening.")
        else:
            blockers.append("The uploaded documents do not establish one valid, approved technical/commercial weighting totalling 100%.")
    if not has_technical_scores:
        blockers.append("Numeric technical evaluation scores were not extracted for every technically eligible bidder; compliance labels are not converted into marks.")
    if not has_commercial_scores:
        blockers.append("Numeric commercial scores were not available for every technically eligible bidder from the price schedules or a commercial evaluation report.")

    if decision == "recommendation_ready":
        status_text = "Preferred bidder identified — submit recommendation for Tender Board approval"
        basis = "Highest overall technical and commercial weighted score extracted from the submitted evaluation records."
    elif decision == "provisional_price_leader":
        status_text = "Provisional lowest evaluated-price bidder — NOT an award decision"
        basis = "Lowest extracted price among bidders whose TET status is confirmed eligible. This is a price-only comparison, not a final award recommendation."
    else:
        status_text = "Award recommendation cannot be determined from the uploaded evidence"
        basis = "Required technical/commercial evaluation evidence is missing, inconsistent, or requires confirmation."

    return {
        "decision": decision, "status_text": status_text, "basis": basis,
        "recommended_bidder": recommended["bidder"] if recommended else None,
        "recommended_score": recommended.get("overall_score") if recommended else None,
        "recommended_price": recommended.get("price") if recommended else None,
        "awarded_bidder": awarded_bidder,
        "award_confirmation_status": award_confirmation_status,
        "award_confirmation_sources": award_confirmation_sources,
        "weights": weights, "weights_valid": weights_valid,
        "technical_scores_complete": has_technical_scores,
        "commercial_scores_complete": has_commercial_scores,
        "prices_comparable": one_currency and all_eligible_priced,
        "currency": next(iter(currencies)) if one_currency else None,
        "eligible_count": len(eligible_keys), "excluded_count": len(excluded_keys),
        "review_count": len(review_keys), "bidders": sorted(bidders, key=lambda b: (
            b["eligibility_status"] != "eligible", b["overall_rank"] or 999,
            b["price_rank"] or 999, bidder_key(b["bidder"]))),
        "blockers": list(dict.fromkeys(blockers)),
        "manual_notes": [
            "This report is an evaluation aid. It does not itself award the tender; an awarded bidder is shown only when an explicit approval/award record is found in the uploaded documents.",
            "The preferred bidder remains subject to CET/TET review, negotiation where permitted, Tender Board/DOA approval, and contract execution.",
            "Commercial comparison must include whole-of-life costs, cost loaders, relevant taxes/duties, freight, installation, training, maintenance, licences, spares and other tender-specific costs where applicable.",
            "Qualitative factors must not outweigh the quantitative technical and commercial factors; use them only as tie-breakers where applicable.",
        ],
    }


def build_award_workbook(assessment, output_path, tender=""):
    """Create an auditable workbook showing eligibility, prices, scores and blockers."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Award Assessment"
    navy = "1F3864"
    pale = "D9EAF7"
    green = "E2F0D9"
    amber = "FFF2CC"
    red = "FCE4D6"

    ws.merge_cells("A1:M1")
    ws["A1"] = "COMMERCIAL EVALUATION — AWARD ASSESSMENT"
    ws["A1"].font = Font(bold=True, color=navy, size=16)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[1].height = 30
    ws["A2"] = "Tender No."
    ws["B2"] = tender or "Not extracted"
    ws["A3"] = "Assessment status"
    ws["B3"] = assessment.get("status_text")
    ws["A4"] = "Preferred bidder (subject to approval)"
    ws["B4"] = assessment.get("recommended_bidder") or "Not determined"
    ws["A5"] = "Confirmed awarded bidder (document evidence)"
    ws["B5"] = assessment.get("awarded_bidder") or "No explicit award approval/notice detected"
    ws["A6"] = "Evaluation basis"
    for row in range(2, 7):
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=13)
        ws.cell(row, 2).alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
        ws.row_dimensions[row].height = 30 if row < 6 else 46
    ws["B6"] = assessment.get("basis", "")
    for cell in ("A2", "A3", "A4", "A5", "A6"):
        ws[cell].font = Font(bold=True, color=navy)
        ws[cell].alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    ws["B3"].fill = PatternFill("solid", fgColor=green if assessment.get("decision") == "recommendation_ready" else amber)
    ws["B4"].fill = PatternFill("solid", fgColor=green if assessment.get("recommended_bidder") else amber)

    headers = ["Bidder", "TET eligibility", "Technical score / 100", "Technical weighting %",
               "Weighted technical score", "Currency", "Evaluated price", "Price rank",
               "Commercial score / 100", "Commercial weighting %", "Weighted commercial score",
               "Overall weighted score", "Source / notes"]
    header_row = 7
    for col, header in enumerate(headers, 1):
        cell = ws.cell(header_row, col, header)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor=navy)
        cell.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[header_row].height = 32

    for row_idx, bidder in enumerate(assessment.get("bidders", []), header_row + 1):
        price = bidder.get("price") or {}
        notes = [bidder.get("eligibility_reason", "")]
        if price.get("source"):
            notes.append(f"Price source: {price['source']}")
        notes.extend(price.get("warnings", []))
        weights = assessment.get("weights") or {}
        values = [
            bidder.get("bidder"), bidder.get("eligibility_status"), bidder.get("technical_score"),
            weights.get("technical"), bidder.get("weighted_technical_score"),
            price.get("currency"), price.get("amount"), bidder.get("price_rank"),
            bidder.get("commercial_score"), weights.get("commercial"),
            bidder.get("weighted_commercial_score"), bidder.get("overall_score"),
            " | ".join(x for x in notes if x),
        ]
        for col, value in enumerate(values, 1):
            cell = ws.cell(row_idx, col, value)
            cell.alignment = Alignment(vertical="top", wrap_text=(col in (1, 2, 13)))
        if bidder.get("eligibility_status") == "eligible":
            ws.cell(row_idx, 2).fill = PatternFill("solid", fgColor=green)
        elif bidder.get("eligibility_status") == "excluded":
            ws.cell(row_idx, 2).fill = PatternFill("solid", fgColor=red)
        else:
            ws.cell(row_idx, 2).fill = PatternFill("solid", fgColor=amber)
        if bidder.get("price") and bidder["price"].get("amount") is not None:
            ws.cell(row_idx, 7).number_format = '#,##0.00'
        for col in (3, 4, 5, 9, 10, 11, 12):
            ws.cell(row_idx, col).number_format = '0.00'

    notes_row = header_row + len(assessment.get("bidders", [])) + 3
    ws.cell(notes_row, 1, "Outstanding checks / blockers")
    ws.cell(notes_row, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(notes_row, 1).fill = PatternFill("solid", fgColor=navy)
    ws.merge_cells(start_row=notes_row, start_column=1, end_row=notes_row, end_column=13)
    for offset, note in enumerate(assessment.get("blockers", []) or ["No outstanding automated checks were identified."], 1):
        r = notes_row + offset
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.cell(r, 1, f"• {note}")
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30

    manual_row = notes_row + len(assessment.get("blockers", [])) + 2
    ws.merge_cells(start_row=manual_row, start_column=1, end_row=manual_row, end_column=13)
    ws.cell(manual_row, 1, "Manual / governance notes")
    ws.cell(manual_row, 1).font = Font(bold=True, color="FFFFFF")
    ws.cell(manual_row, 1).fill = PatternFill("solid", fgColor=navy)
    for offset, note in enumerate(assessment.get("manual_notes", []), 1):
        r = manual_row + offset
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=13)
        ws.cell(r, 1, f"• {note}")
        ws.cell(r, 1).alignment = Alignment(wrap_text=True, vertical="top")
        ws.row_dimensions[r].height = 30

    widths = [27, 20, 18, 17, 22, 13, 20, 12, 20, 17, 22, 20, 65]
    for col, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col)].width = width
    ws.freeze_panes = "A8"
    ws.auto_filter.ref = f"A{header_row}:M{header_row + len(assessment.get('bidders', []))}"
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and cell.row > 6:
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)
    return output_path
