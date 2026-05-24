import re
from dataclasses import dataclass
from datetime import datetime, timedelta

from dateutil.relativedelta import relativedelta


MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}


@dataclass
class ParseResult:
    month_labels: list[str]
    reason: str


def month_label(date_value: datetime) -> str:
    return date_value.strftime("%b-%Y")


def parse_term(term_text, start_date: datetime) -> ParseResult | None:
    if not term_text or not start_date:
        return None

    text = str(term_text).replace("\xa0", " ").strip()
    if not text:
        return None

    normalized = re.sub(r"\s+", " ", text.lower())

    if re.search(r"\b(ex[- ]?stock|current date cheque|cash|cdc in advance)\b", normalized):
        return ParseResult([month_label(start_date)], "Immediate/current date term")

    explicit = _parse_explicit_months(normalized, start_date)
    if explicit:
        return ParseResult(explicit, "Explicit month in term")

    duration = _parse_duration(normalized, start_date)
    if duration:
        return ParseResult([month_label(duration)], "Duration from LPO date")

    invoice_days = re.search(r"within\s+(\d+)\s+days", normalized)
    if invoice_days:
        end_date = start_date + timedelta(days=int(invoice_days.group(1)))
        return ParseResult([month_label(end_date)], "Within N days")

    return None


def _parse_duration(text: str, start_date: datetime) -> datetime | None:
    matches = list(
        re.finditer(
            r"(\d+)(?:\s*(?:-|to)\s*(\d+))?\s*(working\s+)?(day|days|week|weeks|month|months)",
            text,
        )
    )

    if not matches:
        return None

    best_end = None

    for match in matches:
        first = int(match.group(1))
        second = int(match.group(2)) if match.group(2) else first
        amount = max(first, second)
        unit = match.group(4)

        if unit.startswith("day"):
            candidate = start_date + timedelta(days=amount)
        elif unit.startswith("week"):
            candidate = start_date + timedelta(weeks=amount)
        else:
            candidate = start_date + relativedelta(months=amount)

        if best_end is None or candidate > best_end:
            best_end = candidate

    return best_end


def _parse_explicit_months(text: str, start_date: datetime) -> list[str] | None:
    year_match = re.search(r"\b(20\d{2})\b", text)
    year = int(year_match.group(1)) if year_match else start_date.year

    found: list[int] = []
    month_names = "|".join(sorted(MONTHS.keys(), key=len, reverse=True))

    for match in re.finditer(rf"\b({month_names})(?:\s*/\s*({month_names}))?\b", text):
        first = MONTHS[match.group(1)]
        second = MONTHS[match.group(2)] if match.group(2) else None

        for month in [first, second]:
            if month and month not in found:
                found.append(month)

    if not found:
        return None

    return [datetime(year, month, 1).strftime("%b-%Y") for month in found]