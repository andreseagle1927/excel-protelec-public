import re
from datetime import datetime, timedelta
from typing import Any, Optional


def normalize_spaces(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").strip()
    return re.sub(r"\s+", " ", text)


def to_str_id(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if value.is_integer():
            return str(int(value))
        return normalize_spaces(value)
    if isinstance(value, int):
        return str(value)
    text = normalize_spaces(value)
    if text.endswith(".0") and text[:-2].isdigit():
        return text[:-2]
    return text


def to_number(value: Any) -> Optional[float]:
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return float(value)

    text = normalize_spaces(value)
    if not text:
        return None

    text = text.replace("$", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    elif "," in text:
        text = text.replace(",", ".")

    try:
        return float(text)
    except ValueError:
        return None


def excel_serial_to_datetime(value: Any) -> Optional[datetime]:
    if isinstance(value, datetime):
        return value
    if isinstance(value, (int, float)):
        base = datetime(1899, 12, 30)
        return base + timedelta(days=float(value))
    return None
