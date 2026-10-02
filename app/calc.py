"""품질비용(Q-cost) 계산 (04_prd_appendix.md 2장 공식 확정 기준).
제조원가·부적합중량 모두 오직 실제 단가표/중량표에서만 가져온다 — 임의 대표값으로 대체하지 않고,
둘 중 하나라도 못 찾으면 그 행은 계산하지 않고(Q_cost=NaN) 화면에서 에러로 보여준다
(사용자 확정: "언제나 정확한 값이 필요하므로 모르면 계산하지 말고 에러를 띄워라")."""
import pandas as pd
from config import RECOVERY_RATE
from weight import compute_defect_weight_g


def lead_price_per_g(lme: float, fx: float) -> float:
    """g당 납가격 = LME(US$/톤) × 환율 ÷ 1,000,000 (톤→g 환산, 사용자 검산 확인됨: 2000×1380÷1e6=2.76원/g)"""
    return lme * fx / 1_000_000


def qcost_row(process: str, defect_qty: float, weight_g: float, unit_cost: float,
              lead_price: float) -> float:
    """
    재공품·조립: (극판제조원가 × 부적합수량) − (극판중량g × 회수율 × g당납가격)
    완성      : (대당제조원가 × 부적합수량) − (제품중량g × 회수율 × g당납가격)
    극판중량은 재공품=Wet중량, 조립=Dry중량 기준(확정). 이 함수는 제조원가·중량이 모두
    확보된 행에서만 호출된다 — 못 구한 경우는 add_qcost_columns에서 미리 NaN 처리한다.
    """
    recovery = RECOVERY_RATE.get(process, 0.0)
    return unit_cost * defect_qty - weight_g * recovery * lead_price


def add_qcost_columns(df: pd.DataFrame, lme: float, fx: float, cost_lookup: dict = None,
                       plate_weight_lookup: dict = None, product_weight_lookup: dict = None) -> pd.DataFrame:
    """cost_lookup(자재코드→표준원가) 또는 중량표(plate_weight_lookup/product_weight_lookup)에
    매칭키가 없으면 그 행은 Q_cost=NaN으로 남긴다 — 임의값(0 포함)으로 채우지 않는다.
    호출부(app.py)가 이 누락 건들을 에러로 표시해야 한다."""
    df = df.copy()
    cost_lookup = cost_lookup or {}
    plate_weight_lookup = plate_weight_lookup or {}
    product_weight_lookup = product_weight_lookup or {}
    lead = lead_price_per_g(lme, fx)

    df["제조원가_단가"] = df["매칭키"].map(cost_lookup)
    df["제조원가_실측여부"] = df["제조원가_단가"].notna()

    weight_g = df.apply(
        lambda r: compute_defect_weight_g(
            r["공정"], r.get("매칭키"), r.get("불량수량", 0),
            plate_weight_lookup, product_weight_lookup,
        ),
        axis=1,
    )
    df["부적합중량_실측여부"] = weight_g.notna()
    df["부적합중량g"] = weight_g  # 못 찾은 행은 NaN 그대로 둔다(0으로 대체하지 않음)

    def _cost(r):
        if pd.isna(r["제조원가_단가"]) or not r["부적합중량_실측여부"]:
            return float("nan")
        return qcost_row(r["공정"], r["불량수량"], r["부적합중량g"], r["제조원가_단가"], lead)

    df["Q_cost_원"] = df.apply(_cost, axis=1)
    df["Q_cost_백만원"] = df["Q_cost_원"] / 1_000_000
    return df
