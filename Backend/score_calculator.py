"""Approved composite-score formulas used for each tender.

Weighted Technical Score = Technical Score * Technical Weighting / 100
Weighted Commercial Score = Commercial Score * Commercial Weighting / 100
Overall Score = Weighted Technical Score + Weighted Commercial Score
"""

from decimal import Decimal, InvalidOperation, ROUND_HALF_UP


_TWO_PLACES = Decimal("0.01")


def _number(value, label):
    try:
        number = value if isinstance(value, Decimal) else Decimal(str(value))
    except (InvalidOperation, TypeError, ValueError) as error:
        raise ValueError(f"{label} must be a valid number.") from error
    if not number.is_finite():
        raise ValueError(f"{label} must be a finite number.")
    return number


def calculate_weighted_scores(technical_score, commercial_score,
                              technical_weighting, commercial_weighting):
    """Apply the approved PDF's weighted-score formula, returning None if inputs are missing."""
    weighted_technical = None
    if technical_score is not None and technical_weighting is not None:
        technical = _number(technical_score, "Technical score")
        technical_weight = _number(technical_weighting, "Technical weighting")
        if not (Decimal(0) <= technical <= 100 and Decimal(0) <= technical_weight <= 100):
            raise ValueError("Technical scores and weightings must be between 0 and 100.")
        weighted_technical = (technical * technical_weight / 100).quantize(
            _TWO_PLACES, rounding=ROUND_HALF_UP)
    weighted_commercial = None
    if commercial_score is not None and commercial_weighting is not None:
        commercial = _number(commercial_score, "Commercial score")
        commercial_weight = _number(commercial_weighting, "Commercial weighting")
        if not (Decimal(0) <= commercial <= 100 and Decimal(0) <= commercial_weight <= 100):
            raise ValueError("Commercial scores and weightings must be between 0 and 100.")
        weighted_commercial = (commercial * commercial_weight / 100).quantize(
            _TWO_PLACES, rounding=ROUND_HALF_UP)
    if weighted_technical is None or weighted_commercial is None:
        overall = None
    else:
        overall = (weighted_technical + weighted_commercial).quantize(
            _TWO_PLACES, rounding=ROUND_HALF_UP)
    return {"weighted_technical_score": weighted_technical,
            "weighted_commercial_score": weighted_commercial,
            "overall_score": overall}
