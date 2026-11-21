"""Read bidders' price documents (Word / PDF / Excel, folders, zips) and store every price line in the database."""
import os, sys, tempfile, re
# folder that holds consolidate.py / pricesched.py (the bid_compliance project); set BID_COMPLIANCE_DIR if it is elsewhere
sys.path.insert(0, os.environ.get("BID_COMPLIANCE_DIR", "../bid_compliance"))
import consolidate as C, pricesched as P, tender_data as T


def _lines(kind, items):
    out = []
    for it in items:
        qty = P.num(it.get("qty"))
        desc = re.sub(r"\s+", " ", P._text(it.get("desc"))).strip()
        if kind == "local":
            base_unit = P.num(it.get("unit_excl"))
            incl = P.num(it.get("unit_incl"))
            if incl is None and base_unit is not None:
                incl = base_unit * (1 + sum(P.num(it.get(x)) or 0 for x in ("vat", "nbt", "other")) / 100)
            total = incl * (qty if qty is not None else 1) if incl is not None else None
            unit = base_unit
        else:
            total, notes = P.part_total("foreign", [it])           # CIF for the line (same rule as the Excel sheet)
            total = total or None
            unit = total / qty if total and qty else total
            base_unit = P._foreign_unit_price(it)
        out.append(dict(description=desc, qty=qty, unit_price=unit, unit_price_excl_tax=base_unit,
                        line_total=total,
                        origin=P._text(it.get("origin")) or None))
    return out


def ingest(inputs, tender_no=None, user=None):
    """inputs: list of files/folders. tender_no: if None the number printed in the document is used.
    -> dict(stored=[...], problems=[...])"""
    tmp, problems, stored = tempfile.mkdtemp(), [], []
    for company, f, orig in C.collect(inputs, tmp, problems):
        try:
            data = P.parse_price(f)
        except Exception as ex:
            problems.append(f"{orig.name}: {ex}"); continue
        tn = tender_no or data["tender"]
        if not tn:
            problems.append(f"{orig.name}: no tender number given and none found in the document"); continue
        if not data["parts"]:
            problems.append(f"{orig.name}: no filled-in price schedule found"); continue
        for part in data["parts"]:
            cur = P.detect_currency(data["text"], part["kind"])
            n = T.store_bid_lines(tn, company, part["kind"], cur, _lines(part["kind"], part["items"]), orig.name, user)
            stored.append(dict(tender=tn, company=company, part=part["kind"], currency=cur, lines=n))
            if part["kind"] == "foreign" and not cur:
                problems.append(f"{orig.name}: currency not found in the document - check before using prices")
    return dict(stored=stored, problems=problems)
