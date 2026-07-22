from __future__ import annotations

import json
from copy import deepcopy
from pathlib import Path


MONTHLY_DEFAULTS = [
    {"band": 1, "lower": 626_053_153, "upper": 644_466_481, "rate": 0.005},
    {"band": 2, "lower": 644_466_481, "upper": 656_742_033, "rate": 0.010},
    {"band": 3, "lower": 656_742_033, "upper": 675_155_361, "rate": 0.015},
    {"band": 4, "lower": 675_155_361, "upper": 693_568_689, "rate": 0.020},
    {"band": 5, "lower": 693_568_689, "upper": 718_119_793, "rate": 0.023},
    {"band": 6, "lower": 718_119_793, "upper": 736_533_121, "rate": 0.026},
    {"band": 7, "lower": 736_533_121, "upper": 797_910_881, "rate": 0.028},
    {"band": 8, "lower": 797_910_881, "upper": 859_288_641, "rate": 0.030},
    {"band": 9, "lower": 859_288_641, "upper": None, "rate": 0.032},
]

QUARTERLY_DEFAULTS = [
    {"band": 1, "lower": 1_878_159_459, "upper": 1_933_399_443, "rate": 0.005},
    {"band": 2, "lower": 1_933_399_443, "upper": 1_970_226_099, "rate": 0.010},
    {"band": 3, "lower": 1_970_226_099, "upper": 2_025_466_083, "rate": 0.015},
    {"band": 4, "lower": 2_025_466_083, "upper": 2_080_706_067, "rate": 0.020},
    {"band": 5, "lower": 2_080_706_067, "upper": 2_154_359_379, "rate": 0.023},
    {"band": 6, "lower": 2_154_359_379, "upper": 2_209_599_363, "rate": 0.026},
    {"band": 7, "lower": 2_209_599_363, "upper": 2_393_732_644, "rate": 0.028},
    {"band": 8, "lower": 2_393_732_644, "upper": 2_577_865_924, "rate": 0.030},
    {"band": 9, "lower": 2_577_865_924, "upper": None, "rate": 0.032},
]


def default_config() -> dict:
    monthly_ends = [31, 28, 31, 30, 31, 30, 31, 31, 30, 31, 30, 31]
    monthly = [
        {
            "period": month,
            "label": f"{month}월",
            "start": f"2026-{month:02d}-01",
            "end": f"2026-{month:02d}-{monthly_ends[month - 1]:02d}",
            "bands": deepcopy(MONTHLY_DEFAULTS),
        }
        for month in range(1, 13)
    ]
    month_ends = ["03-31", "06-30", "09-30", "12-31"]
    quarterly = [
        {
            "period": quarter,
            "label": f"{quarter}분기",
            "start": f"2026-{(quarter - 1) * 3 + 1:02d}-01",
            "end": f"2026-{month_ends[quarter - 1]}",
            "bands": deepcopy(QUARTERLY_DEFAULTS),
        }
        for quarter in range(1, 5)
    ]
    return {"monthly": monthly, "quarterly": quarterly}


def load_config(path: Path) -> dict:
    if not path.exists():
        return default_config()
    try:
        value = json.loads(path.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return default_config()
    if not isinstance(value, dict) or not value.get("monthly") or not value.get("quarterly"):
        return default_config()
    if value["monthly"] and "band" in value["monthly"][0]:
        migrated = default_config()
        for period in migrated["monthly"]:
            period["bands"] = deepcopy(value["monthly"])
        for period in migrated["quarterly"]:
            period["bands"] = deepcopy(value["quarterly"])
        return migrated
    return value


def save_config(path: Path, monthly: list[dict], quarterly: list[dict]) -> None:
    if len(monthly) != 12 or len(quarterly) != 4:
        raise ValueError("표준은 12개월, 타입B는 4분기 자료가 필요합니다.")
    for period in monthly:
        _validate_bands(period["bands"], f'{period["label"]} 기본계약')
    for period in quarterly:
        _validate_bands(period["bands"], f'{period["label"]} 타입B')
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(
        json.dumps({"monthly": monthly, "quarterly": quarterly}, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def _validate_bands(bands: list[dict], label: str) -> None:
    if len(bands) != 9:
        raise ValueError(f"{label} 구간은 9개여야 합니다.")
    previous = -1
    for index, band in enumerate(bands, 1):
        lower = int(band["lower"])
        upper = band.get("upper")
        rate = float(band["rate"])
        if lower < 0 or lower < previous:
            raise ValueError(f"{label} {index}구간의 시작 금액을 확인해주세요.")
        if upper is not None and int(upper) <= lower:
            raise ValueError(f"{label} {index}구간의 종료 금액은 시작 금액보다 커야 합니다.")
        if not 0 <= rate <= 1:
            raise ValueError(f"{label} {index}구간 수취율을 확인해주세요.")
        previous = lower


def incentive_for(amount: int, bands: list[dict]) -> tuple[int, float, int]:
    selected = None
    for band in bands:
        if amount >= int(band["lower"]):
            selected = band
        else:
            break
    if selected is None:
        return 0, 0.0, 0
    rate = float(selected["rate"])
    return int(selected["band"]), rate, round(amount * rate)


def calculate_year(monthly_amounts: list[int], config: dict) -> list[dict]:
    if len(monthly_amounts) != 12:
        raise ValueError("월별 매입액은 12개월 값이 필요합니다.")
    results = []
    for amount in monthly_amounts:
        band, rate, base = incentive_for(max(0, int(amount)), config["monthly"][len(results)]["bands"])
        results.append({"amount": int(amount), "band": band, "rate": rate, "base": base, "quarter_extra": 0, "total": base})
    for quarter in range(4):
        start = quarter * 3
        end = start + 3
        quarter_amount = sum(item["amount"] for item in results[start:end])
        q_band, q_rate, quarter_total = incentive_for(quarter_amount, config["quarterly"][quarter]["bands"])
        paid_base = sum(item["base"] for item in results[start:end])
        extra = max(quarter_total - paid_base, 0)
        results[end - 1].update({
            "quarter_amount": quarter_amount,
            "quarter_band": q_band,
            "quarter_rate": q_rate,
            "quarter_extra": extra,
            "total": results[end - 1]["base"] + extra,
        })
    return results
