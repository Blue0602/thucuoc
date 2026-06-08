from __future__ import annotations
from typing import Any, Dict, Tuple, List
import pandas as pd
from .utils import normalize_key, normalize_text, normalize_ma_tt, normalize_phone, parse_money


class ImportValidationError(Exception):
    pass


def _find_column(df: pd.DataFrame, candidates: List[str]) -> str | None:
    normalized_columns = {normalize_key(col): col for col in df.columns}
    for candidate in candidates:
        key = normalize_key(candidate)
        if key in normalized_columns:
            return normalized_columns[key]
    return None


def _build_column_mapping(df: pd.DataFrame, file_config: Dict[str, Any]) -> Dict[str, str]:
    mapping = {}
    for field, candidates in file_config.get("required_fields", {}).items():
        col = _find_column(df, candidates)
        if not col:
            raise ImportValidationError(f"Thiếu cột bắt buộc cho '{field}'. Tên chấp nhận: {candidates}")
        mapping[field] = col
    for field, candidates in file_config.get("optional_fields", {}).items():
        col = _find_column(df, candidates)
        if col:
            mapping[field] = col
    return mapping


def read_excel_flexible(uploaded_file, file_config: Dict[str, Any]) -> Tuple[pd.DataFrame, str | None]:
    sheet_name = file_config.get("sheet_name", None)
    try:
        df = pd.read_excel(uploaded_file, sheet_name=sheet_name)
    except Exception as exc:
        raise ImportValidationError(f"Không đọc được file Excel: {exc}") from exc
    if isinstance(df, dict):
        selected_sheet = next(iter(df.keys()))
        df = df[selected_sheet]
    else:
        selected_sheet = sheet_name
    df = df.dropna(how="all")
    df.columns = [normalize_text(c) for c in df.columns]
    return df, selected_sheet


def standardize_dataframe(df: pd.DataFrame, file_config: Dict[str, Any], import_type: str) -> pd.DataFrame:
    mapping = _build_column_mapping(df, file_config)
    output = pd.DataFrame()
    for canonical, source_col in mapping.items():
        output[canonical] = df[source_col]
    output["ma_tt"] = output["ma_tt"].apply(normalize_ma_tt)
    for col in ["customer_name", "representative_code", "address", "staff_name", "paid_date"]:
        if col in output.columns:
            output[col] = output[col].apply(normalize_text)
    if "phone" in output.columns:
        output["phone"] = output["phone"].apply(normalize_phone)
    for money_col in ["debt_amount", "paid_amount", "assigned_amount"]:
        if money_col in output.columns:
            output[money_col] = output[money_col].apply(parse_money)
    defaults = {"customer_name":"", "representative_code":"", "address":"", "phone":"", "staff_name":"", "debt_amount":0.0, "paid_amount":0.0, "paid_date":"", "assigned_amount":0.0}
    for col, default in defaults.items():
        if col not in output.columns:
            output[col] = default
    output = output[output["ma_tt"].astype(str).str.len() > 0].copy()
    output["source_import_type"] = import_type
    return output


def validate_standardized(df: pd.DataFrame, import_type: str) -> Dict[str, Any]:
    warnings = []
    duplicate_count = int(df["ma_tt"].duplicated().sum()) if "ma_tt" in df.columns else 0
    if duplicate_count:
        warnings.append(f"Có {duplicate_count} dòng trùng MA_TT. Nên kiểm tra trước khi báo cáo.")
    empty_phone_count = int((df.get("phone", "") == "").sum()) if "phone" in df.columns else 0
    if import_type in {"tn08", "assignment"} and empty_phone_count:
        warnings.append(f"Có {empty_phone_count} dòng chưa có số điện thoại.")
    total_amount = 0.0
    if "debt_amount" in df.columns:
        total_amount = float(df["debt_amount"].sum())
    elif "paid_amount" in df.columns:
        total_amount = float(df["paid_amount"].sum())
    return {"row_count": int(len(df)), "duplicate_ma_tt_count": duplicate_count, "total_amount": total_amount, "warnings": warnings}
