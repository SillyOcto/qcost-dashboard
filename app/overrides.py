"""제조원가/중량표 보정 데이터 영속 저장소.
사용자가 (1) 사이드바에서 중량표를 새로 교체 업로드하거나 (2) 누락값 팝업에서 코드별로 직접 값을
입력하면, 원본 파일 전체가 아니라 **코드&값만** 이 폴더에 저장해 용량을 최소화하고, 다음 실행 때도
(다시 업로드하지 않아도) 자동으로 불러와 기본 단가표/중량표 위에 덮어씌운다(override).
같은 코드가 여러 번 갱신되면 가장 최근 값이 남는다(병합/upsert, 덮어쓰기 삭제 아님)."""
import os
import pandas as pd
from config import (
    OVERRIDE_DIR, COST_OVERRIDE_PARQUET, PLATE_WEIGHT_OVERRIDE_PARQUET, PRODUCT_WEIGHT_OVERRIDE_PARQUET,
)


def _load_flat(path: str, key_col: str, value_col: str) -> dict:
    if not os.path.exists(path):
        return {}
    df = pd.read_parquet(path)
    return dict(zip(df[key_col], df[value_col]))


def _save_flat(path: str, key_col: str, value_col: str, values: dict) -> None:
    os.makedirs(os.path.dirname(path), exist_ok=True)
    pd.DataFrame({key_col: list(values.keys()), value_col: list(values.values())}).to_parquet(path, index=False)


def load_cost_overrides() -> dict:
    """{매칭키: 표준원가} — 지금까지 저장된 제조원가 보정값 전부."""
    return _load_flat(COST_OVERRIDE_PARQUET, "매칭키", "표준원가")


def merge_cost_overrides(new_values: dict) -> dict:
    """new_values(매칭키→표준원가)를 기존 보정 데이터에 병합(upsert)해 저장하고, 합쳐진 결과를 반환한다."""
    current = load_cost_overrides()
    current.update({k: v for k, v in new_values.items() if v is not None})
    _save_flat(COST_OVERRIDE_PARQUET, "매칭키", "표준원가", current)
    return current


def load_plate_weight_overrides() -> dict:
    """{재공품코드: {"wet": g, "dry": g}} — 재공품(Wet)/조립(Dry) 중량 보정값."""
    if not os.path.exists(PLATE_WEIGHT_OVERRIDE_PARQUET):
        return {}
    df = pd.read_parquet(PLATE_WEIGHT_OVERRIDE_PARQUET)
    return {
        row["재공품코드"]: {"wet": row["Wet중량(g)"], "dry": row["Dry중량(g)"]}
        for _, row in df.iterrows()
    }


def merge_plate_weight_overrides(new_values: dict) -> dict:
    """new_values: {코드: {"wet": x?, "dry": y?}} — 필드 하나만 와도 기존 값(다른 필드)은 유지한 채 병합."""
    current = load_plate_weight_overrides()
    for code, vals in new_values.items():
        entry = current.setdefault(code, {"wet": None, "dry": None})
        if vals.get("wet") is not None:
            entry["wet"] = vals["wet"]
        if vals.get("dry") is not None:
            entry["dry"] = vals["dry"]
    os.makedirs(OVERRIDE_DIR, exist_ok=True)
    pd.DataFrame([
        {"재공품코드": code, "Wet중량(g)": v.get("wet"), "Dry중량(g)": v.get("dry")}
        for code, v in current.items()
    ]).to_parquet(PLATE_WEIGHT_OVERRIDE_PARQUET, index=False)
    return current


def merge_plate_lookup(base: dict, override: dict) -> dict:
    """기본 중량표(base)와 보정값(override)을 필드 단위로 합친다 — override에 wet만 있고
    dry가 없으면 base의 dry를 그대로 유지한다(override 쪽 dict를 통째로 덮어쓰면 base에만 있던
    다른 필드 값이 사라지는 문제를 방지)."""
    result = {code: dict(vals) for code, vals in base.items()}
    for code, ov in override.items():
        entry = result.setdefault(code, {"wet": None, "dry": None})
        if ov.get("wet") is not None:
            entry["wet"] = ov["wet"]
        if ov.get("dry") is not None:
            entry["dry"] = ov["dry"]
    return result


def load_product_weight_overrides() -> dict:
    """{자재코드: 총중량(g)} — 완성 제품중량 보정값."""
    return _load_flat(PRODUCT_WEIGHT_OVERRIDE_PARQUET, "자재코드", "총중량(g)")


def merge_product_weight_overrides(new_values: dict) -> dict:
    current = load_product_weight_overrides()
    current.update({k: v for k, v in new_values.items() if v is not None})
    _save_flat(PRODUCT_WEIGHT_OVERRIDE_PARQUET, "자재코드", "총중량(g)", current)
    return current
