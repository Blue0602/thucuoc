from __future__ import annotations
from datetime import datetime
from decimal import Decimal, InvalidOperation
import re
import unicodedata


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    return re.sub(r"\s+", " ", text)


def normalize_key(value) -> str:
    text = normalize_text(value).lower()
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    return re.sub(r"[^a-z0-9]+", "_", text).strip("_")


def normalize_ma_tt(value) -> str:
    text = normalize_text(value)
    if text.endswith(".0"):
        text = text[:-2]
    return text.upper()


def normalize_phone(value) -> str:
    if value is None:
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    if not digits:
        return ""
    if len(digits) == 9:
        digits = "0" + digits
    return digits


def parse_money(value) -> float:
    if value is None or value == "":
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip().replace("đ", "").replace("VND", "").replace(" ", "")
    if "," in text and "." in text:
        text = text.replace(".", "").replace(",", ".")
    else:
        text = text.replace(",", "")
    try:
        return float(Decimal(text))
    except (InvalidOperation, ValueError):
        return 0.0


def current_timestamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def format_money(value) -> str:
    try:
        return f"{float(value):,.0f} đồng".replace(",", ".")
    except Exception:
        return "0 đồng"
