"""Shared data layer for the commercial evaluation: USD rates + previous purchases.
Works on PostgreSQL (production) and SQLite (testing). Choose with the DATABASE_URL setting:
    postgresql://user:password@server:5432/tenderdb      (all computers share this one database)
    sqlite:///tender.db                                   (single computer / testing only)
"""
import os, datetime as dt
from pathlib import Path

# DATABASE_URL comes from the environment; when it is not set, the .env sitting beside
# this file is used (…/tenderdb), so uploaded documents land in the shared database.
ENV_FILE = Path(__file__).resolve().parent / ".env"
if not os.environ.get("DATABASE_URL") and ENV_FILE.is_file():
    for _line in ENV_FILE.read_text(encoding="utf-8").splitlines():
        _key, _sep, _value = _line.strip().partition("=")
        if _sep and _key.strip() == "DATABASE_URL":
            os.environ["DATABASE_URL"] = _value.strip().strip('"').strip("'")

URL = os.environ.get("DATABASE_URL", "sqlite:///tender_test.db")
PG = URL.startswith("postgres")

if PG:
    import psycopg2
    PH = "%s"
    def connect():
        # Fail in seconds rather than minutes when the server is switched off, so a
        # dead database cannot stall an upload.
        return psycopg2.connect(URL, connect_timeout=5)
else:
    import sqlite3
    PH = "?"
    def connect():
        return sqlite3.connect(URL.replace("sqlite:///", ""))

SCHEMA = """
CREATE TABLE IF NOT EXISTS usd_rate(
    day        DATE PRIMARY KEY,
    buy        NUMERIC(12,4),
    sell       NUMERIC(12,4),
    indicative NUMERIC(12,4),
    source     TEXT NOT NULL,
    fetched_at TEXT NOT NULL
);
CREATE TABLE IF NOT EXISTS past_award(
    id            %s,
    tender_no     TEXT NOT NULL,
    item          TEXT NOT NULL,
    supplier      TEXT,
    cif_usd       NUMERIC(14,4) NOT NULL CHECK (cif_usd > 0),
    closing_date  DATE,
    po_date       DATE,
    entered_by    TEXT,
    entered_at    TEXT NOT NULL,
    UNIQUE (tender_no, item)
);
CREATE TABLE IF NOT EXISTS tender(
    tender_no       TEXT PRIMARY KEY,
    closing_date    DATE,
    po_date         DATE,
    awarded_company TEXT
);
CREATE TABLE IF NOT EXISTS bid_line(
    id           %s,
    tender_no    TEXT NOT NULL,
    company      TEXT NOT NULL,
    supply_type  TEXT NOT NULL CHECK (supply_type IN ('local','foreign')),
    line_no      INTEGER NOT NULL,
    description  TEXT,
    qty          NUMERIC(14,4),
    currency     TEXT,
    unit_price   NUMERIC(16,4),
    line_total   NUMERIC(18,4),
    origin       TEXT,
    source_file  TEXT,
    uploaded_by  TEXT,
    uploaded_at  TIMESTAMP NOT NULL,
    UNIQUE (tender_no, company, supply_type, line_no)
);
CREATE TABLE IF NOT EXISTS uploaded_document(
    id           %s,
    job          TEXT NOT NULL,
    mode         TEXT,
    bidder       TEXT,
    filename     TEXT NOT NULL,
    content_type TEXT,
    size_bytes   INTEGER NOT NULL,
    content_hash TEXT NOT NULL,
    content      %s NOT NULL,
    uploaded_at  TIMESTAMP NOT NULL,
    UNIQUE (filename, content_hash)
);
""" % ((("SERIAL PRIMARY KEY" if PG else "INTEGER PRIMARY KEY AUTOINCREMENT"),) * 3
       + (("BYTEA" if PG else "BLOB"),))


from contextlib import contextmanager


@contextmanager
def session():
    """Opens a connection, commits on success, rolls back on any error, always closes."""
    c = connect()
    try:
        yield c.cursor()
        c.commit()
    except Exception:
        c.rollback()
        raise
    finally:
        c.close()


def init():
    with session() as cur:
        for stmt in SCHEMA.split(";"):
            if stmt.strip():
                cur.execute(stmt)
        # Add content_hash to older uploaded_document tables (created before dedup).
        try:
            cur.execute("ALTER TABLE uploaded_document ADD COLUMN content_hash TEXT")
        except Exception:
            pass
        try:
            cur.execute("CREATE UNIQUE INDEX IF NOT EXISTS ux_uploaded_document_file_hash "
                        "ON uploaded_document(filename, content_hash)")
        except Exception:
            pass


def _now():
    return dt.datetime.now().isoformat(timespec="seconds")


def save_rates(rows, source):
    """rows: iterable of (date 'YYYY-MM-DD', buy, sell, indicative). Re-running is safe (upsert)."""
    sql = (f"INSERT INTO usd_rate(day,buy,sell,indicative,source,fetched_at) VALUES({','.join([PH]*6)}) "
           "ON CONFLICT(day) DO UPDATE SET buy=excluded.buy, sell=excluded.sell, "
           "indicative=excluded.indicative, source=excluded.source, fetched_at=excluded.fetched_at")
    with session() as cur:
        for d, b, s, i in rows:
            cur.execute(sql, (str(d), b, s, i, source, _now()))


def rate_on(day, kind="buy"):
    """Rate on `day`. If none was published that day (weekend/holiday) the last earlier one is used and
    exact=False is returned, so the report can say so."""
    assert kind in ("buy", "sell", "indicative")
    with session() as cur:
        cur.execute(f"SELECT day,{kind} FROM usd_rate WHERE day<={PH} AND {kind} IS NOT NULL "
                    "ORDER BY day DESC LIMIT 1", (str(day),))
        r = cur.fetchone()
    if not r:
        raise LookupError(f"no USD {kind} rate stored on or before {day}")
    return {"requested": str(day), "rate_date": str(r[0]), "rate": float(r[1]), "exact": str(r[0]) == str(day)}


def add_past_award(tender_no, item, cif_usd, supplier=None, closing_date=None, po_date=None, user=None):
    with session() as cur:
        cur.execute(f"INSERT INTO past_award(tender_no,item,supplier,cif_usd,closing_date,po_date,entered_by,entered_at) "
                    f"VALUES({','.join([PH]*8)})",
                    (tender_no, item, supplier, cif_usd, closing_date, po_date, user, _now()))


def _norm(t):
    import re
    return re.sub(r"[^a-z0-9]+", " ", (t or "").lower()).strip()


def set_tender(tender_no, closing_date=None, po_date=None, awarded_company=None):
    """Record the dates of a tender and (after the decision) which company won it."""
    with session() as cur:
        cur.execute(f"INSERT INTO tender(tender_no,closing_date,po_date,awarded_company) VALUES({PH},{PH},{PH},{PH}) "
                    "ON CONFLICT(tender_no) DO UPDATE SET "
                    "closing_date=COALESCE(excluded.closing_date,tender.closing_date), "
                    "po_date=COALESCE(excluded.po_date,tender.po_date), "
                    "awarded_company=COALESCE(excluded.awarded_company,tender.awarded_company)",
                    (tender_no, closing_date, po_date, awarded_company))


def store_bid_lines(tender_no, company, supply_type, currency, lines, source_file=None, user=None):
    """lines: dicts with description, qty, unit_price, line_total, origin. Re-uploading the same
    tender/company/part replaces its lines (no duplicates, no leftover lines from a previous upload)."""
    now = _now()
    sql = ("INSERT INTO bid_line(tender_no,company,supply_type,line_no,description,qty,currency,unit_price,"
           "line_total,origin,source_file,uploaded_by,uploaded_at) VALUES(" + ",".join([PH] * 13) + ") "
           "ON CONFLICT(tender_no,company,supply_type,line_no) DO UPDATE SET description=excluded.description,"
           "qty=excluded.qty,currency=excluded.currency,unit_price=excluded.unit_price,line_total=excluded.line_total,"
           "origin=excluded.origin,source_file=excluded.source_file,uploaded_by=excluded.uploaded_by,"
           "uploaded_at=excluded.uploaded_at")
    with session() as cur:
        # True replace: drop any lines left from an earlier (possibly larger) upload of this
        # tender/company/part, then insert the freshly parsed set. Rolled back on any failure.
        cur.execute(f"DELETE FROM bid_line WHERE tender_no={PH} AND company={PH} AND supply_type={PH}",
                    (tender_no, company, supply_type))
        for n, l in enumerate(lines, 1):
            cur.execute(sql, (tender_no, company, supply_type, n, l.get("description"), l.get("qty"), currency,
                              l.get("unit_price"), l.get("line_total"), l.get("origin"), source_file, user, now))
    return len(lines)


def store_document(job, filename, content, mode=None, bidder=None, content_type=None, user=None):
    """Keep an uploaded document (bytes and all) so it survives a restart.

    A filename identifies one stored document.  Re-uploading that filename (case
    insensitive, matching normal Windows filename behavior) always
    removes any older rows first, then stores the new content and metadata.  The
    delete also cleans up duplicate rows created by older versions that keyed
    documents by ``(filename, content_hash)``.
    """
    import hashlib
    content_hash = hashlib.sha256(content).hexdigest()
    with session() as cur:
        # Replace by filename, regardless of whether the bytes changed.  Delete
        # all matching rows so databases created before this behavior was added
        # are repaired on the next upload as well.
        cur.execute(f"DELETE FROM uploaded_document WHERE LOWER(filename)=LOWER({PH})", (filename,))
        cur.execute(
            f"INSERT INTO uploaded_document(job,mode,bidder,filename,content_type,size_bytes,"
            f"content_hash,content,uploaded_at) VALUES({','.join([PH] * 9)}) ",
            (job, mode, bidder, filename, content_type, len(content), content_hash, content, _now()))
    return len(content)


def documents_for(job):
    """Stored documents of a job, oldest first. Content is left out; use document_content() for it."""
    with session() as cur:
        cur.execute(f"SELECT id,filename,bidder,mode,size_bytes,uploaded_at FROM uploaded_document "
                    f"WHERE job={PH} ORDER BY id", (job,))
        return [dict(id=r[0], filename=r[1], bidder=r[2], mode=r[3], size_bytes=r[4],
                     uploaded_at=str(r[5])) for r in cur.fetchall()]


def document_content(document_id):
    """(filename, content_type, bytes) of one stored document, or None."""
    with session() as cur:
        cur.execute(f"SELECT filename,content_type,content FROM uploaded_document WHERE id={PH}",
                    (document_id,))
        row = cur.fetchone()
    if not row:
        return None
    return row[0], row[1], bytes(row[2])


def previous_purchased(item, current_cif_usd, limit=3, exclude_tender=None):
    """Rows for the 'Previous Purchased' block: Tender No | Price (CIF)(USD) | % of Variation (CIF).
    Sources: (1) awarded foreign bid lines stored from uploaded price documents, (2) manually entered
    past awards. Items are matched on the description (ignoring case/punctuation) - exact match only."""
    want = _norm(item)
    found = []
    with session() as cur:
        cur.execute("SELECT b.tender_no, b.description, b.unit_price, t.po_date FROM bid_line b "
                    "JOIN tender t ON t.tender_no=b.tender_no AND t.awarded_company=b.company "
                    "WHERE b.supply_type='foreign' AND b.currency='USD' AND b.unit_price IS NOT NULL")
        found += [(a, float(p), d, "uploaded bid") for a, desc, p, d in cur.fetchall() if _norm(desc) == want]
        cur.execute("SELECT tender_no, item, cif_usd, po_date FROM past_award")
        found += [(a, float(p), d, "manual entry") for a, it, p, d in cur.fetchall() if _norm(it) == want]
    found = [f for f in found if f[0] != exclude_tender]
    found.sort(key=lambda f: str(f[2] or ""), reverse=True)
    return [{"tender_no": t, "cif_usd": p, "po_date": str(d) if d else None, "source": src,
             "variation_pct": round((current_cif_usd - p) / p * 100, 2)} for t, p, d, src in found[:limit]]


def evaluation_block(item, current_cif_usd, closing_date, po_date):
    """Everything the commercial-evaluation header needs, in one call."""
    return {"rate_closing": rate_on(closing_date), "rate_po": rate_on(po_date),
            "rate_today": rate_on(dt.date.today().isoformat()),
            "previous": previous_purchased(item, current_cif_usd)}


# ---------------------------------------------------------------- CBSL history loader
def load_cbsl_xlsx(path):
    """Load USD buying/selling rates from a CBSL workbook.

    Supports both the original multi-currency CBSL layout (with a ``United
    States`` column) and the compact layout containing ``Date``, ``Buy Rate
    (LKR)``, and ``Sell Rate (LKR)`` columns.  Rows with missing or invalid
    dates/rates are skipped and reported.
    """
    import openpyxl
    ws = openpyxl.load_workbook(path, data_only=True).active

    # Some historical exports contain only the USD series and label the three
    # relevant columns explicitly instead of using the multi-currency header.
    direct_headers = {}
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 12)):
        for cell in row:
            value = str(cell.value or "").strip().casefold()
            if value == "date":
                direct_headers["date"] = cell.column
            elif value.startswith("buy rate"):
                direct_headers["buy"] = cell.column
            elif value.startswith("sell rate"):
                direct_headers["sell"] = cell.column

    # Find the direct header row without relying on a fixed number of preamble rows.
    direct_header_row = None
    if len(direct_headers) == 3:
        for row_number in range(1, min(ws.max_row, 12) + 1):
            values = {str(ws.cell(row_number, col).value or "").strip().casefold()
                      for col in direct_headers.values()}
            if "date" in values and any(v.startswith("buy rate") for v in values) \
                    and any(v.startswith("sell rate") for v in values):
                direct_header_row = row_number
                break
    if direct_header_row is not None:
        date_col = direct_headers["date"]
        buy_col = direct_headers["buy"]
        sell_col = direct_headers["sell"]
        good, skipped = [], []
        for row in ws.iter_rows(min_row=direct_header_row + 1, values_only=True):
            d = row[date_col - 1] if len(row) >= date_col else None
            buy = row[buy_col - 1] if len(row) >= buy_col else None
            sell = row[sell_col - 1] if len(row) >= sell_col else None
            if not isinstance(d, (dt.datetime, dt.date)):
                continue
            if not isinstance(buy, (int, float)) or not isinstance(sell, (int, float)):
                skipped.append((d.date().isoformat() if isinstance(d, dt.datetime) else d.isoformat(),
                                "USD rate missing"))
                continue
            good.append((d.date().isoformat() if isinstance(d, dt.datetime) else d.isoformat(), buy, sell, None))
        good.sort(key=lambda row: row[0])
        save_rates(good, f"{Path(path).name} (USD historical rates)")
        return {"loaded": len(good), "first": good[0][0], "last": good[-1][0], "skipped": skipped}

    col = None
    for row in ws.iter_rows(min_row=1, max_row=8):
        for cell in row:
            if isinstance(cell.value, str) and cell.value.strip() == "United States":
                col = cell.column
    if col is None:
        raise ValueError("could not find the 'United States' column - CBSL may have changed the layout")
    good, skipped, prev = [], [], None
    for r in ws.iter_rows(min_row=8, values_only=True):
        d, buy, sell = r[1], r[col - 1], r[col]
        if not isinstance(d, dt.datetime):
            continue                                  # 'Average October' rows, notes, blanks
        if not isinstance(buy, (int, float)) or not isinstance(sell, (int, float)):
            skipped.append((d.date().isoformat(), "USD rate missing")); continue
        if prev is not None and (d.date() <= prev or (d.date() - prev).days > 40):
            skipped.append((d.date().isoformat(), f"date out of sequence after {prev} (typo in CBSL file?)")); continue
        prev = d.date()
        good.append((d.date().isoformat(), buy, sell, None))
    save_rates(good, "CBSL IF_Buying_Selling_Exchange_Rates.xlsx (commercial bank TT averages, 9.30 AM)")
    return {"loaded": len(good), "first": good[0][0], "last": good[-1][0], "skipped": skipped}
