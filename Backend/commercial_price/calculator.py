"""Calculate SLT commercial price scores.

Price Score = (Lowest Evaluated Price / Bidder's Evaluated Price) * 100
Weighted Score = Price Score * Commercial Weight / 100
"""

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


_TWO_PLACES = Decimal("0.01")


def _decimal(value, label):
    try:
        result = value if isinstance(value, Decimal) else Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as error:
        raise ValueError(f"{label} must be a valid number.") from error
    if not result.is_finite():
        raise ValueError(f"{label} must be a finite number.")
    return result


def calculate_commercial_price_scores(bidders, commercial_weight=100):
    """Return commercial price and weighted scores for bidder evaluated prices.

    ``bidders`` is a mapping of company name to a positive evaluated price.
    Values are rounded half-up to two decimal places for reporting.  The
    calculation uses Decimal throughout to avoid binary floating-point noise.
    """
    if not bidders:
        raise ValueError("No bidders were provided.")

    weight = _decimal(commercial_weight, "Commercial weighting")
    if not Decimal("0") <= weight <= Decimal("100"):
        raise ValueError("Commercial weighting must be between 0 and 100.")

    prices = {}
    for company, value in bidders.items():
        price = _decimal(value, f"Evaluated price for {company}")
        if price <= 0:
            raise ValueError(
                f"Invalid evaluated price for {company}. "
                "The price must be greater than zero."
            )
        prices[company] = price

    lowest_price = min(prices.values())
    results = {}
    for company, price in prices.items():
        price_score = (lowest_price / price) * Decimal("100")
        weighted_score = price_score * weight / Decimal("100")
        results[company] = {
            "evaluated_price": price,
            "price_score": price_score.quantize(_TWO_PLACES, rounding=ROUND_HALF_UP),
            "weighted_score": weighted_score.quantize(_TWO_PLACES, rounding=ROUND_HALF_UP),
        }
    return results
