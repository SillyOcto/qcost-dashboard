"""월별 Q-cost 시계열 비교 계산 — 전월대비(MoM)/이동평균/이상탐지/증가 기여도/YTD/재발빈도.

(v21.0 신규) 사용자가 제안한 시계열 비교 로드맵 중 "여러 달 실데이터가 있어야 계산 가능한" 항목들.
`필요자료/부적합 데이터 all`(1~8월 실데이터)이 반영되면서 처음으로 구현·검증 가능해졌다 — 이전엔
2026-08 한 달치뿐이라 이 계산들이 전부 N/A이거나 무의미한 단일값만 나왔다."""
import pandas as pd


def monthly_totals(df: pd.DataFrame) -> pd.Series:
    """연-월(Period) 인덱스의 월별 Q-cost 합계 — 여러 해가 섞여도 "8월"끼리 안 합쳐지도록 월(1~12)
    대신 실제 연-월로 묶는다(charts.render_monthly_trend()와 동일 규칙)."""
    if df.empty:
        return pd.Series(dtype=float)
    return df.groupby(df["일자_dt"].dt.to_period("M"))["Q_cost_백만원"].sum().sort_index()


def compute_mom(monthly: pd.Series) -> dict:
    """최신월 vs 그 직전월 전월대비(MoM). 전월=0이면 증감률 대신 "신규 발생"/"변화 없음"으로 표기
    (사용자 제안 예외 처리 그대로). 데이터가 2개월 미만이면 None(계산 불가)."""
    if len(monthly) < 2:
        return None
    cur_period, prev_period = monthly.index[-1], monthly.index[-2]
    cur, prev = float(monthly.iloc[-1]), float(monthly.iloc[-2])
    if prev == 0:
        status = "신규 발생" if cur > 0 else "변화 없음"
        return {
            "cur_period": str(cur_period), "prev_period": str(prev_period),
            "cur": cur, "prev": prev, "delta": cur, "pct": None, "status": status,
        }
    delta = cur - prev
    pct = delta / prev * 100
    status = "해소" if cur == 0 else ("증가" if delta > 0 else "감소")
    return {
        "cur_period": str(cur_period), "prev_period": str(prev_period),
        "cur": cur, "prev": prev, "delta": delta, "pct": pct, "status": status,
    }


def ytd_total(monthly: pd.Series) -> dict:
    """연간 누적(YTD) — 최신월이 속한 해의 1월부터 최신월까지 누적 합계와 달 수."""
    if len(monthly) == 0:
        return {"year": None, "months": 0, "total": 0.0}
    latest_year = monthly.index[-1].year
    this_year = monthly[[p.year == latest_year for p in monthly.index]]
    return {"year": latest_year, "months": len(this_year), "total": float(this_year.sum())}


def detect_anomaly(monthly: pd.Series, z_threshold: float = 2.0) -> dict:
    """최신월이 그 이전 달들 평균 대비 이상치인지 Z-score로 판정.
    이전 달이 2개월 미만이면(비교 기준선 자체가 불안정) 판정 불가로 표시한다."""
    if len(monthly) < 3:
        return {"available": False}
    history = monthly.iloc[:-1]
    cur = float(monthly.iloc[-1])
    mean, std = float(history.mean()), float(history.std(ddof=0))
    if std == 0:
        return {"available": True, "z": None, "is_anomaly": False, "mean": mean, "std": std, "cur": cur}
    z = (cur - mean) / std
    return {"available": True, "z": z, "is_anomaly": abs(z) >= z_threshold, "mean": mean, "std": std, "cur": cur}


def contribution_top(df: pd.DataFrame, group_cols: list, top_n: int = 5) -> pd.DataFrame:
    """최신월 vs 직전월, group_cols(예: ["공정","제품군","소분류"]) 기준 그룹별 증감액을 계산해
    "이번 달 증가는 어디서 왔나"를 기여도(증감액) 큰 순으로 정렬한 표를 반환한다."""
    if df.empty:
        return pd.DataFrame()
    month_key = df["일자_dt"].dt.to_period("M")
    months = sorted(month_key.dropna().unique())
    if len(months) < 2:
        return pd.DataFrame()
    cur_m, prev_m = months[-1], months[-2]
    cur = df[month_key == cur_m].groupby(group_cols, observed=True)["Q_cost_백만원"].sum()
    prev = df[month_key == prev_m].groupby(group_cols, observed=True)["Q_cost_백만원"].sum()
    combined = pd.DataFrame({"당월": cur, "전월": prev}).fillna(0)
    combined["증감액"] = combined["당월"] - combined["전월"]
    combined = combined.sort_values("증감액", ascending=False)
    return combined.reset_index().head(top_n)


def recurrence_count(df: pd.DataFrame, group_col: str = "세부내역", window: int = 3) -> pd.DataFrame:
    """최근 window개월 중 몇 개월에 실제로 발생(Q-cost>0)했는지 그룹별로 센다("재발 빈도")."""
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    d["_월"] = d["일자_dt"].dt.to_period("M")
    months = sorted(d["_월"].dropna().unique())
    if not months:
        return pd.DataFrame()
    recent_months = months[-window:]
    recent = d[d["_월"].isin(recent_months)]
    monthly_sum = recent.groupby([group_col, "_월"], observed=True)["Q_cost_백만원"].sum().reset_index()
    monthly_sum = monthly_sum[monthly_sum["Q_cost_백만원"] > 0]
    counts = monthly_sum.groupby(group_col)["_월"].nunique().reset_index(name="발생월수")
    counts["대상월수"] = len(recent_months)
    return counts.sort_values("발생월수", ascending=False)


def consecutive_increase(df: pd.DataFrame, group_cols: list, window: int = 3, top_n: int = 10) -> pd.DataFrame:
    """(v24.2) "최근 window개월 연속으로 늘어난 항목" — 예전 재발 빈도 표(최근 3개월 중 발생 개월 수)는
    여러 달 데이터가 쌓이자 거의 모든 항목이 3/3으로 나와 아무 정보도 주지 못했다(사용자: "어떤 걸
    보여주려는지 모르겠음"). 대신 group_cols(예: 공정·제품군·세부내역)별 월 합계가 최근 window개월 동안
    매달 전월보다 커진 항목만 골라 경보성으로 보여준다. 최신월 값이 큰 순으로 top_n개."""
    if df.empty:
        return pd.DataFrame()
    d = df.copy()
    d["_월"] = d["일자_dt"].dt.to_period("M")
    months = sorted(d["_월"].dropna().unique())
    if len(months) < window:
        return pd.DataFrame()
    recent = months[-window:]
    piv = (
        d[d["_월"].isin(recent)]
        .groupby(group_cols + ["_월"], observed=True)["Q_cost_백만원"].sum()
        .unstack("_월").reindex(columns=recent).fillna(0.0)
    )
    rising = piv[(piv.diff(axis=1).iloc[:, 1:] > 0).all(axis=1) & (piv[recent[-1]] > 0)]
    if rising.empty:
        return pd.DataFrame()
    out = rising.copy()
    out[f"증가폭({window}개월)"] = out[recent[-1]] - out[recent[0]]
    out = out.sort_values(recent[-1], ascending=False).head(top_n).reset_index()
    out.columns = [str(c) for c in out.columns]
    return out
