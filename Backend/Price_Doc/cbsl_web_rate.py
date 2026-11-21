"""Read historical USD/LKR TT rates directly from the CBSL workbook."""

import datetime as dt
import io
import time
import urllib.request
from functools import lru_cache

import openpyxl


SOURCE = (
    "https://www.cbsl.gov.lk/sites/default/files/cbslweb_documents/statistics/sheets/"
    "IF_Buying_Selling_Exchange_Rates.xlsx"
)
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
)


def _download(url=SOURCE, timeout=120):
    request = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    with urllib.request.urlopen(request, timeout=timeout) as response:
        return io.BytesIO(response.read())


def _as_date(value):
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, str):
        for fmt in ("%Y-%m-%d", "%d-%m-%Y", "%d/%m/%Y", "%d-%m-%Y"):
            try:
                return dt.datetime.strptime(value.strip(), fmt).date()
            except ValueError:
                continue
    return None


def parse_usd(fileobj):
    """Return ``{date: (buying, selling)}`` from a CBSL workbook."""
    workbook = openpyxl.load_workbook(fileobj, read_only=True, data_only=True)
    try:
        worksheet = workbook.worksheets[0]
        rows = list(worksheet.iter_rows(values_only=True))
    finally:
        workbook.close()

    usd_column = None
    for row in rows[:12]:
        for index, value in enumerate(row):
            if isinstance(value, str) and value.strip().casefold() == "united states":
                usd_column = index
                break
        if usd_column is not None:
            break
    if usd_column is None:
        raise ValueError(
            "CBSL workbook layout changed: 'United States' column was not found"
        )

    data = {}
    for row in rows:
        if len(row) <= usd_column + 1:
            continue
        day = _as_date(row[1] if len(row) > 1 else None)
        buying, selling = row[usd_column], row[usd_column + 1]
        if (
            day is not None
            and isinstance(buying, (int, float))
            and isinstance(selling, (int, float))
        ):
            data[day] = (float(buying), float(selling))
    return data


@lru_cache(maxsize=2)
def _rates(refresh_hour=None):
    """Download CBSL data once per hour so later reports see updated rates."""
    return parse_usd(_download())


def clear_cache():
    """Clear the in-memory rate cache, primarily for tests or a manual refresh."""
    _rates.cache_clear()


def web_rate(day, kind="buy"):
    """Return the CBSL rate on ``day`` or the latest earlier business-day rate."""
    if kind not in ("buy", "sell", "average"):
        raise ValueError("kind must be 'buy', 'sell' or 'average'")
    requested = _as_date(day)
    if requested is None:
        raise ValueError(f"invalid exchange-rate date: {day!r}")

    # Rate workbooks are updated by CBSL periodically. A process-lifetime cache
    # can silently leave every later report on an old workbook; refresh hourly.
    rates = _rates(int(time.time() // 3600))
    earlier = [available for available in rates if available <= requested]
    if not earlier:
        raise LookupError(f"no CBSL USD rate on or before {requested}")
    rate_date = max(earlier)
    buying, selling = rates[rate_date]
    value = {
        "buy": buying,
        "sell": selling,
        "average": round((buying + selling) / 2, 4),
    }[kind]
    return {
        "requested": requested.isoformat(),
        "rate_date": rate_date.isoformat(),
        "rate": value,
        "exact": rate_date == requested,
    }
