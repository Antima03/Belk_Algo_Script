import json
from pathlib import Path
from typing import Any, Dict, List
from unittest import loader

import pandas as pd

Belk_json_file = "belk_category_forecasts.json"
Belk_excel_file = "belk_category_forecasts.xlsx"

FORECAST_COLUMNS = [
    "category",
    "pid",
    "month",
    "forecasting_method",
    "fc_by_index",
    "fc_by_trend",
    "recommended_fc",
    "ty_lw_sls_u",
    "ly_lw_sls_u"]

MONTH_ORDER = {
    "JAN": 1,
    "FEB": 2,
    "MAR": 3,
    "APR": 4,
    "MAY": 5,
    "JUN": 6,
    "JUL": 7,
    "AUG": 8,
    "SEP": 9,
    "OCT": 10,
    "NOV": 11,
    "DEC": 12,
}


def _safe_number(value: Any) -> float:
    """Convert strings/numbers to float; fallback to 0."""
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        stripped = value.replace(",", "").strip()
        if stripped:
            try:
                return float(stripped)
            except ValueError:
                return 0.0
    return 0.0


def build_rows_from_json(json_payload: Dict[str, Any]) -> List[Dict[str, Any]]:
    rows: List[Dict[str, Any]] = []

    for category, pid_map in json_payload.items():
        for pid, payload in pid_map.items():
            forecast = payload.get("forecast", {})
            loader = payload.get("loader", {})
            method = forecast.get("forecasting_method", "")
            fc_by_index = forecast.get("fc_by_index", {})
            fc_by_trend = forecast.get("fc_by_trend", {})
            recommended_fc = forecast.get("recommended_fc", {})
            ty_lw_sls_u = loader.get("ty_lw_sls_u", {})
            ly_lw_sls_u = loader.get("ly_lw_sls_u", {})

            all_months = set(fc_by_index.keys()) | set(fc_by_trend.keys()) | set(recommended_fc.keys())
            for month in sorted(all_months, key=lambda m: MONTH_ORDER.get(m.upper(), 99)):
                month_key = month.upper()
                rows.append(
                    {
                        "category": category,
                        "pid": pid,
                        "month": month_key,
                        "forecasting_method": method,
                        "fc_by_index": _safe_number(fc_by_index.get(month, 0)),
                        "fc_by_trend": _safe_number(fc_by_trend.get(month, 0)),
                        "recommended_fc": _safe_number(recommended_fc.get(month, 0)),
                        "ty_lw_sls_u": _safe_number(ty_lw_sls_u.get(month, 0)),
                        "ly_lw_sls_u":_safe_number(ly_lw_sls_u.get(month, 0)),
                    }
                )

    return rows


def export_forecast_excel(json_path: Path, output_path: Path) -> None:
    with json_path.open("r", encoding="utf-8") as fh:
        payload = json.load(fh)

    rows = build_rows_from_json(payload)
    if not rows:
        raise ValueError("No forecast data found in JSON payload.")

    df = pd.DataFrame(rows, columns=FORECAST_COLUMNS)
    df.sort_values(["category", "pid", "month"], inplace=True)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    df.to_excel(output_path, index=False)


if __name__ == "__main__":
    json_file = Path(Belk_json_file)
    output_file = Path(Belk_excel_file)
    export_forecast_excel(json_file, output_file)
    print(f"Forecast Excel written to {output_file.resolve()}")
