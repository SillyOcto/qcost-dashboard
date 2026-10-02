"""Q-cost 분석 차트 3종(공정별 파이 / 소분류·제품군 막대 / 세부내역 파레토), Plotly 기반.
막대 클릭 → 파레토 드릴다운은 st.plotly_chart(..., on_select='rerun')의 클릭 이벤트로 구현한다."""
import pandas as pd
import plotly.graph_objects as go

# (v15.0) 다크 테마 전환 — 사용자가 마음에 들어한 발표 자료(다크 배경 + 오렌지/청록 포인트)를 참고해
# design/design.md 2.5절 "다크 모드는 배경/텍스트를 반전하고 포인트 컬러는 400 단계(더 밝은 톤)를
# 사용해 대비를 확보"를 그대로 따른 색상이다(임의 배색 아님). WCAG 대비도 전부 계산해서 확인함 —
# 카테고리 채움색은 배경(#14191D) 대비 5:1 이상, 안쪽 레이블 텍스트(다크)는 채움색 대비 5:1 이상,
# 바깥 텍스트(밝은색)는 배경 대비 17:1로 전부 통과.
CHART_BG = "#14191D"          # 앱 페이지 배경과 동일 — 밝은 조각/막대 위 "안쪽" 텍스트 색으로도 재사용
CHART_TEXT_LIGHT = "#FBFBFB"  # 다크 배경 위 일반 텍스트(막대 바깥 레이블, 누적% 텍스트 등)
CHART_TEXT_MUTED = "#C7CCCE"  # 축 눈금 등 부차적 텍스트
CHART_GRID = "#2B363D"        # 격자·축선(은은하게)

RECON_COLOR = "#ADB2B6"     # 재공품 — dark-gray-200 (파이 + 파레토 상위3 강조색)
ASSEMBLY_COLOR = "#C7CCCE"  # 조립 — dark-gray-300
FINISHED_COLOR = "#33ACBA"  # 완성 — green-400(청록, 포인트 컬러 다크모드 톤)
HIGHLIGHT_COLOR = "#EF5C33"  # 선택된 항목 강조(막대 클릭 시) — orange-400(포인트 컬러 다크모드 톤)
MUTED_BAR_COLOR = "#858C91"  # 파레토 상위 3 제외 막대 — dark-gray-300 계열(살짝 더 어둡게)
ETC_BAR_COLOR = "#D0D0CE"   # 파레토 "기타" 묶음 막대 — light-gray(밝아서 다크 배경에서도 잘 보임)

PIE_COLORS = {"재공품": RECON_COLOR, "조립": ASSEMBLY_COLOR, "완성": FINISHED_COLOR}

# 기존엔 폰트 크기를 따로 안 정해둬서 Plotly 기본값(12px)으로 나왔는데, 읽기 힘들다는 피드백에 따라
# 전부 20% 키운다(12 * 1.2 = 14.4) — 사용자 요청 "지금 포인트보다 20% 크게" 그대로 반영.
BASE_FONT_SIZE = 14.4

# 차트 배경을 투명하게 두면 Streamlit 다크 페이지 배경이 그대로 비쳐 보여 카드처럼 붕 뜨지 않는다.
_TRANSPARENT_LAYOUT = dict(
    paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    font=dict(color=CHART_TEXT_MUTED, size=BASE_FONT_SIZE),
)


def render_pie(df):
    """공정별 Q-cost 비중 파이차트.

    (v15.1) 조각 위에 투명 Scatter 마커를 겹쳐 클릭을 받는 실험을 해봤지만(v14.15) 실제 브라우저에서도
    여전히 클릭이 안 된다는 재제보로 확인 실패 — Pie는 도넛 모양이라 각도/반지름을 좌표축에 맞춰
    계산해도 Plotly가 내부적으로 그리는 실제 반지름·중심과 100% 일치한다는 보장이 없어(Pie는 라벨
    공간 확보 등으로 반지름을 자체적으로 줄이기도 함) 어긋났을 가능성이 높음 — 화면을 직접 볼 수
    없는 환경이라 좌표를 눈으로 맞춰볼 수도 없어, 대신 100% 확실하게 동작하는 st.button 3개를
    파이 아래에 두는 방식으로 교체했다(app.py 참고). 이 함수는 이제 순수하게 시각적 표시만 한다."""
    totals = df.groupby("공정")["Q_cost_백만원"].sum()
    totals = totals.reindex(["재공품", "조립", "완성"]).fillna(0)
    fig = go.Figure()
    fig.add_pie(
        labels=totals.index, values=totals.values,
        marker=dict(colors=[PIE_COLORS.get(k, ASSEMBLY_COLOR) for k in totals.index]),
        textinfo="label+percent",
        # 조각 채움색이 전부 밝은/중간톤이라(재공품·조립·완성) 다크 텍스트가 대비가 더 좋다(계산 확인:
        # 5.2~10.9:1). 흰 텍스트는 청록(완성) 조각에서 2.7:1로 너무 낮아서 다크 텍스트로 통일.
        textfont=dict(size=BASE_FONT_SIZE, color=CHART_BG),
        hovertemplate="%{label}: %{value:,.1f} 백만원 (%{percent})<extra></extra>",
        sort=False, rotation=0, direction="clockwise",
    )

    fig.update_layout(
        height=280, margin=dict(l=20, r=20, t=20, b=20), showlegend=False,
        **_TRANSPARENT_LAYOUT,
    )
    return fig


def render_bar(df, category_col: str, selected: str = None, top_n: int = None):
    """category_col(소구분 또는 제품군) 기준 Q-cost 비중 가로 막대. 클릭 드릴다운 대상이면 selected로 강조.
    top_n을 주면 값이 큰 상위 top_n개만 그린다(귀책부서·라인처럼 카테고리가 아주 많을 때 사용,
    v20.0) — 이때도 %는 "보여주는 항목들 중 비중"이 아니라 항상 "전체 대비 비중"으로 계산한다.

    (v15.3) v14.9~v15.1에서 "짧은 막대는 클릭 영역이 좁다"는 문제를 투명 오버레이 트레이스로
    풀어봤는데, 긴 막대(예: 97%짜리)조차 두 번 클릭해야 반영된다는 재제보로 확인됨 — 즉 막대
    길이와 무관하게 트레이스가 2개(오버레이) 겹쳐 있는 것 자체가 원인으로 보임(하나의 클릭이
    어느 트레이스 것인지 Plotly가 애매하게 처리, hovermode/clickmode 조정으로도 해결 안 됨).
    트레이스를 다시 1개로 되돌리고, 대신 "짧은 막대 클릭 어려움" 문제는 막대의 실제 렌더링
    길이에 최솟값(전체 최댓값의 8%)을 둬서 완화한다 — 실제 값은 텍스트 라벨과 툴팁에 항상
    정확히 표시하므로(그리는 길이만 조정, 값 자체를 속이지 않음) 데이터 정확성 원칙에 위배되지
    않는다."""
    grp = df.groupby(category_col)["Q_cost_백만원"].sum()
    grp = grp[grp > 0].sort_values(ascending=True)
    total = grp.sum()
    if top_n is not None and len(grp) > top_n:
        grp = grp.iloc[-top_n:]  # 이미 오름차순 정렬이라 뒤쪽 top_n개가 가장 큰 값들
    colors = [HIGHLIGHT_COLOR if selected and idx == selected else FINISHED_COLOR for idx in grp.index]
    text = [f"{v/total*100:.0f}% ({v:,.1f}백만원)" if total else "" for v in grp.values]
    max_val = grp.values.max() if len(grp) else 1
    xrange_max = max_val * 1.35
    floor = max_val * 0.08
    display_values = [max(v, floor) for v in grp.values]

    fig = go.Figure()
    fig.add_bar(
        x=display_values, y=grp.index, orientation="h",
        marker_color=colors, text=text,
        # "outside"로 고정하면 막대가 길어서(거의 꽉 차서) 바깥쪽에 자리가 없는 경우 컨테이너 폭을
        # 넘어가 잘렸다(cliponaxis만으론 못 막음 — 실제 원인은 플롯 영역 자체의 폭 부족).
        # "auto"로 두면 자리가 없을 때 플롯이 알아서 막대 안쪽으로 넣어준다.
        textposition="auto",
        textfont=dict(size=BASE_FONT_SIZE),
        # 안쪽(막대 위) 텍스트는 채움색(청록/오렌지 모두 밝은 톤)과 대비가 좋은 다크 텍스트,
        # 바깥쪽(다크 페이지 배경 위) 텍스트는 밝은 텍스트로 — 각각 반대라 두 값 다 명시.
        insidetextfont=dict(size=BASE_FONT_SIZE, color=CHART_BG),
        outsidetextfont=dict(size=BASE_FONT_SIZE, color=CHART_TEXT_LIGHT),
        cliponaxis=False,
        customdata=grp.values,  # 툴팁에 표시할 "진짜" 값(막대 길이는 최솟값 보정이 들어가 다를 수 있음)
        hovertemplate="%{y}: %{customdata:,.1f} 백만원<extra></extra>",
    )
    fig.update_layout(
        height=max(180, 46 * len(grp) + 40),
        margin=dict(l=10, r=170, t=10, b=10),
        xaxis=dict(showticklabels=False, range=[0, xrange_max]),
        **_TRANSPARENT_LAYOUT,
        showlegend=False,
    )
    return fig, grp


def render_pareto(df, top_n: int = 8):
    """세부내역 파레토도: 막대(개별 비중) + 누적 % 꺾은선(보조축).
    상위 3개 항목은 색을 다르게 강조하고 금액 레이블을 붙인다(사용자 요청).

    (v15.1) 상위 3개 색을 재공품 색(회색 계열, RECON_COLOR)으로 뒀더니 주변 회색 막대들과 구분이
    잘 안 된다는 재제보 — "더 눈에 띄었으면 좋겠다, 빨간색 계열은 말고"라 청록(FINISHED_COLOR,
    완성 파이 조각·누적%선과 같은 색— 이 앱에서 유일하게 쓰는 강조색이라 일관성도 있음)으로 교체."""
    grp = df.groupby("세부내역")["Q_cost_백만원"].sum()
    grp = grp[grp > 0].sort_values(ascending=False)

    shown = grp.iloc[:top_n]
    rest = grp.iloc[top_n:].sum()
    if rest > 0.0001:
        shown = pd.concat([shown, pd.Series({"기타(그 외 전체)": rest})])

    total = shown.sum()
    cum_pct = (shown.cumsum() / total * 100) if total else shown.cumsum() * 0

    # shown은 이미 내림차순 정렬 뒤 "기타"만 맨 끝에 덧붙인 상태라, 앞 3개가 항상 진짜 상위 3개다.
    bar_colors = [
        FINISHED_COLOR if i < 3 else (MUTED_BAR_COLOR if not str(idx).startswith("기타") else ETC_BAR_COLOR)
        for i, idx in enumerate(shown.index)
    ]
    # (v24.3) 상위 3개 금액 레이블이 너무 작아 안 읽힌다는 제보 — Plotly는 기본(constraintext="both")으로
    # 막대 폭에 맞춰 글자를 자동 축소하는데 파레토 막대가 좁아서 8px 수준까지 줄어들고 있었음.
    # 축소를 끄고(constraintext="none") 글자를 키우며, 단위는 y축 제목(백만원)에 이미 있으니 정수만 쓴다.
    bar_text = [f"{v:,.0f}" if i < 3 else "" for i, v in enumerate(shown.values)]

    fig = go.Figure()
    fig.add_bar(
        x=shown.index, y=shown.values, name="Q-cost",
        marker_color=bar_colors,
        text=bar_text, textposition="outside", constraintext="none",
        # 다크 페이지 배경 위(막대 바깥쪽)라 밝은 텍스트 — RECON_COLOR 채움 자체는 상위3 막대에만
        # 있고 레이블은 막대 "바깥"에 붙으므로(outside) 배경 기준으로 밝게.
        textfont=dict(size=BASE_FONT_SIZE + 4, color=CHART_TEXT_LIGHT, weight=700),
        yaxis="y1", cliponaxis=False,
        hovertemplate="%{x}: %{y:,.1f} 백만원<extra></extra>",
    )
    fig.add_scatter(
        x=shown.index, y=cum_pct.values, name="누적 %", mode="lines+markers+text",
        text=[f"{v:.0f}%" for v in cum_pct.values], textposition="top center",
        textfont=dict(size=BASE_FONT_SIZE, color=CHART_TEXT_LIGHT),
        # (v15.2) 누적%선이 상위3 막대와 같은 청록이라 겹치는 지점에서 선이 안 보인다는 제보 —
        # 막대(상위3=청록)와 확실히 구분되도록 선은 오렌지(HIGHLIGHT_COLOR)로 변경.
        line=dict(color=HIGHLIGHT_COLOR, width=2),
        marker=dict(color=HIGHLIGHT_COLOR),
        yaxis="y2", cliponaxis=False,
        hovertemplate="누적 %{y:.0f}%<extra></extra>",
    )
    # (v14.10) 80% 기준선은 의미가 헷갈린다는 피드백에 따라 삭제 — 사용자 요청.
    fig.update_layout(
        # 맨 위 점의 "100%" 라벨이 그래프 상단 경계에 잘리던 문제 — 위쪽 여백을 넉넉히 두고
        # y2(누적 %) 축 범위도 105→120으로 늘려 라벨이 들어갈 자리를 미리 확보.
        # 1차 y축(막대)도 상위 3개 금액 레이블이 위쪽에 붙으므로 범위에 여유를 둔다.
        height=360, margin=dict(l=10, r=60, t=40, b=90),
        yaxis=dict(
            title="백만원", range=[0, shown.values.max() * 1.25] if len(shown) else None,
            gridcolor=CHART_GRID, zerolinecolor=CHART_GRID,
        ),
        yaxis2=dict(overlaying="y", side="right", range=[0, 120], showgrid=False, title="누적 %"),
        xaxis=dict(tickangle=-35, gridcolor=CHART_GRID, zerolinecolor=CHART_GRID),
        showlegend=False,
        **_TRANSPARENT_LAYOUT,
    )
    return fig


def pareto_concentration_stats(df) -> dict:
    """세부내역 파레토도 옆에 붙일 "비용 집중도" 요약 숫자 — render_pareto()와 동일하게 "세부내역"
    기준으로 묶어서 계산한다(v20.0, 사용자 요청 "비용 집중도(파레토): 상위 20% 현상이 전체의 몇 %
    차지하는지"). {건수, 총합, top3_금액, top3_비율, n_to_80, pct_of_top20pct} 반환 — 데이터가
    없으면 전부 0/None."""
    grp = df.groupby("세부내역")["Q_cost_백만원"].sum()
    grp = grp[grp > 0].sort_values(ascending=False)
    n = len(grp)
    total = grp.sum()
    if n == 0 or total <= 0:
        return {"n": 0, "total": 0.0, "top3_amount": 0.0, "top3_pct": 0.0, "n_to_80": 0, "top20pct_pct": 0.0}

    cum = grp.cumsum() / total * 100
    n_to_80 = int((cum < 80).sum()) + 1  # 누적 80%를 처음 넘는 지점까지 항목 수
    n_to_80 = min(n_to_80, n)

    top3_amount = float(grp.iloc[:3].sum())
    top3_pct = top3_amount / total * 100

    n_20pct = max(1, round(n * 0.2))
    top20pct_amount = float(grp.iloc[:n_20pct].sum())
    top20pct_pct = top20pct_amount / total * 100

    return {
        "n": n, "total": float(total), "top3_amount": top3_amount, "top3_pct": top3_pct,
        "n_to_80": n_to_80, "n_20pct": n_20pct, "top20pct_pct": top20pct_pct,
    }


def render_monthly_trend(df, show_ma=False):
    """월별 Q-cost 추이 꺾은선(단일 계열) — 연-월(예: "2026-08") 단위로 묶어서 그린다.
    여러 해에 걸친 데이터가 섞여도 "8월"끼리 잘못 합쳐지지 않도록 월(1~12) 대신 실제 연-월로 묶는다.
    (v16.1) "전체 Q-cost 발생 현황" 개요 차트와 "Q-cost 세부분석"(특정 공정·소분류·세부내역 조합)
    차트 둘 다 이 함수 하나로 그린다 — 그룹 기준(전체 vs 특정 조합)만 다르고 그리는 방식은 같다.

    (v21.0) show_ma=True면 3개월 이동평균을 보조선으로 함께 그린다(데이터가 여러 달 쌓여야 의미가
    있어 "전체 개요" 차트에서만 켜고, "세부분석"의 좁은 조합 차트에는 안 켠다)."""
    if df.empty:
        monthly = pd.Series(dtype=float)
    else:
        month_key = df["일자_dt"].dt.to_period("M")
        monthly = df.groupby(month_key)["Q_cost_백만원"].sum().sort_index()
    labels = [str(p) for p in monthly.index]
    values = list(monthly.values)

    fig = go.Figure()
    fig.add_scatter(
        x=labels, y=values, mode="lines+markers+text", name="Q-cost",
        text=[f"{v:,.1f}" for v in values], textposition="top center",
        textfont=dict(size=BASE_FONT_SIZE, color=CHART_TEXT_LIGHT),
        line=dict(color=FINISHED_COLOR, width=2.5),
        marker=dict(color=FINISHED_COLOR, size=8),
        cliponaxis=False,
        hovertemplate="%{x}: %{y:,.1f} 백만원<extra></extra>",
    )
    if show_ma and len(monthly) >= 2:
        ma3 = monthly.rolling(window=3, min_periods=1).mean()
        fig.add_scatter(
            x=labels, y=list(ma3.values), mode="lines", name="3개월 이동평균",
            line=dict(color=HIGHLIGHT_COLOR, width=2, dash="dash"),
            hovertemplate="%{x} 3개월 이동평균: %{y:,.1f} 백만원<extra></extra>",
        )
    fig.update_layout(
        height=280, margin=dict(l=10, r=20, t=30, b=40),
        xaxis=dict(type="category", gridcolor=CHART_GRID, zerolinecolor=CHART_GRID),
        yaxis=dict(
            title="백만원", gridcolor=CHART_GRID, zerolinecolor=CHART_GRID,
            rangemode="tozero", range=[0, max(values) * 1.25] if values else None,
        ),
        showlegend=show_ma,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=BASE_FONT_SIZE)),
        **_TRANSPARENT_LAYOUT,
    )
    return fig


def render_period_compare(scope_a, scope_b, label_a="기간 A", label_b="기간 B"):
    """(v24.0) 두 기간의 Q-cost를 하나의 시간축에 막대로 나란히 그린다 — 세부분석 "기간 비교"용.
    두 기간을 합친 전체 폭에 따라 일(≤45일)/주(≤200일)/월 단위로 자동 묶는다(짧은 기간을 월로 묶으면
    막대가 한두 개뿐이라 비교가 안 되고, 긴 기간을 일로 그리면 너무 촘촘해지기 때문)."""
    all_dates = pd.concat([scope_a["일자_dt"], scope_b["일자_dt"]]).dropna()
    span = (all_dates.max() - all_dates.min()).days if len(all_dates) else 0
    if span <= 45:
        freq, unit = "D", "일"
    elif span <= 200:
        freq, unit = "W-SUN", "주"
    else:
        freq, unit = "M", "월"

    def _bucket(df):
        if df.empty:
            return pd.Series(dtype=float)
        key = df["일자_dt"].dt.to_period(freq)
        return df.groupby(key)["Q_cost_백만원"].sum().sort_index()

    def _label(p):
        if freq == "D":
            return p.start_time.strftime("%m/%d")
        if freq == "M":
            return str(p)
        return f"{p.start_time:%m/%d}~"

    fig = go.Figure()
    max_val = 0.0
    for series, name, color in ((_bucket(scope_a), label_a, MUTED_BAR_COLOR),
                                (_bucket(scope_b), label_b, FINISHED_COLOR)):
        if not len(series):
            continue
        max_val = max(max_val, float(series.max()))
        fig.add_bar(
            x=[_label(p) for p in series.index], y=list(series.values), name=name,
            marker_color=color, text=[f"{v:,.1f}" for v in series.values], textposition="outside",
            textfont=dict(size=BASE_FONT_SIZE - 2, color=CHART_TEXT_LIGHT), cliponaxis=False,
            hovertemplate=f"{name} %{{x}}: %{{y:,.1f}} 백만원<extra></extra>",
        )
    fig.update_layout(
        barmode="group", height=280, margin=dict(l=10, r=20, t=30, b=40),
        xaxis=dict(type="category", title=f"{unit} 단위", gridcolor=CHART_GRID, zerolinecolor=CHART_GRID),
        yaxis=dict(title="백만원", gridcolor=CHART_GRID, zerolinecolor=CHART_GRID, rangemode="tozero",
                   range=[0, max_val * 1.3] if max_val else None),
        showlegend=True,
        legend=dict(orientation="h", yanchor="bottom", y=1.02, font=dict(size=BASE_FONT_SIZE)),
        **_TRANSPARENT_LAYOUT,
    )
    return fig
