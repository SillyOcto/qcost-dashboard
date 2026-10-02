"""Q-cost 목표(경영계획) 로딩 — `총수율 표 양식.xlsx`의 "25년 실적 및 목표" 시트.

(v24.0 신규) 사용자 확인: "월별 목표표 있음(엑셀)" → 확인 결과 총수율 표 양식 파일 안에 26년 목표가
공정/제품군/소분류/상세내역별로 년간·반기·분기·월별 컬럼으로 들어있다(월별 = 년간÷12 균등 배분).
단위는 같은 파일의 양식 시트("※ 단위 : 백만원")와 동일한 백만원으로 확인됨(값 크기도 26년 실적과
같은 자릿수).

시트 구조(1행 비어있음):
  2행: 대분류 | 중분류 | 공정 | 소분류 | 상세 내역 | 25년 기준(P-COST) ×4 | 26년 목표 ×4
  3행:                                         년간 | 반기 | 분기 | 월별 | 년간 | 반기 | 분기 | 월별
  4행~: 리프 행(병합 셀이라 중분류/공정/소분류는 위 값이 이어짐) + "소계" 행 + "○○ 합계" 행
목표 시트의 리프 이름이 우리 총수율 종합표 양식보다 더 세분화돼 있어(예: "휴식시간 초기가동"),
KPI 카드용 합계는 리프 행을 전부 더해서 쓰고 리프 단위 매핑은 하지 않는다(사용자: "동일한 항목
빼고는 기타에 넣으면 될 것" — 표에 목표 열을 붙이게 되면 그때 적용)."""
import os
import re

import pandas as pd

_PROCESS_BY_PREFIX = (("재공품", "재공품"), ("조립", "조립"), ("완성", "완성"))


def _norm(s) -> str:
    return re.sub(r"\s+", "", str(s)) if s is not None else ""


def _process_of(mid_label: str):
    label = _norm(mid_label)
    for prefix, process in _PROCESS_BY_PREFIX:
        if label.startswith(prefix):
            return process
    return None


def _family_of(raw) -> str:
    fam = _norm(raw)
    return {"EB/GC": "EBGC", "V-type": "V-type"}.get(fam, fam)


def _find_sheet(path, keyword: str):
    import openpyxl
    wb = openpyxl.load_workbook(path, read_only=True)
    for name in wb.sheetnames:
        if keyword in name:
            return name
    return None


def load_targets(path, sheet_keyword: str = "목표") -> dict:
    """{"year": 2026, "leaf": DataFrame[공정,제품군,소분류,세부내역,연간목표,월목표],
        "annual_total", "monthly_total", "by_process": {공정: {"annual","monthly"}}, "source": 시트명}
    파일/시트가 없거나 구조를 못 읽으면 {}(목표 미등록) — 임의값으로 채우지 않는다."""
    if not path or not os.path.exists(path):
        return {}
    sheet = _find_sheet(path, sheet_keyword)
    if sheet is None:
        return {}
    try:
        raw = pd.read_excel(path, sheet_name=sheet, header=None)
    except Exception:
        return {}

    # 헤더 행: "상세 내역"이 들어있는 행. 그 다음 행의 마지막 "월별" 컬럼이 목표 월별, 그 3칸 앞이 목표 년간.
    header_idx = None
    for i in range(min(10, len(raw))):
        if any(_norm(v) == "상세내역" for v in raw.iloc[i].tolist()):
            header_idx = i
            break
    if header_idx is None or header_idx + 1 >= len(raw):
        return {}
    header = raw.iloc[header_idx].tolist()
    sub_header = raw.iloc[header_idx + 1].tolist()

    col = {name: None for name in ("중분류", "공정", "소분류", "상세내역")}
    for j, v in enumerate(header):
        key = _norm(v)
        if key in col and col[key] is None:
            col[key] = j
    monthly_cols = [j for j, v in enumerate(sub_header) if _norm(v) == "월별"]
    target_hdr = [j for j, v in enumerate(header) if "목표" in _norm(v)]
    if any(c is None for c in col.values()) or not monthly_cols or not target_hdr:
        return {}
    # "26년 목표" 헤더 위치 이후에 있는 "월별" 컬럼을 목표 월별로 본다(그 앞쪽 "월별"은 25년 기준).
    month_col = [j for j in monthly_cols if j >= target_hdr[0]]
    if not month_col:
        return {}
    month_col = month_col[0]
    annual_col = month_col - 3
    year_match = re.search(r"(\d{2,4})\s*년", str(header[target_hdr[0]]))
    year = None
    if year_match:
        y = int(year_match.group(1))
        year = y + 2000 if y < 100 else y

    body = raw.iloc[header_idx + 2:].copy()
    body[[col["중분류"], col["공정"], col["소분류"]]] = body[[col["중분류"], col["공정"], col["소분류"]]].ffill()

    rows = []
    for _, r in body.iterrows():
        mid, fam, sub, leaf = r[col["중분류"]], r[col["공정"]], r[col["소분류"]], r[col["상세내역"]]
        if _norm(sub) == "소계" or "합계" in _norm(mid):
            continue
        process = _process_of(mid)
        if process is None:
            continue
        annual, monthly = pd.to_numeric(r[annual_col], errors="coerce"), pd.to_numeric(r[month_col], errors="coerce")
        if pd.isna(annual) and pd.isna(monthly):
            continue
        rows.append({
            "공정": process, "제품군": _family_of(fam), "소분류": _norm(sub),
            "세부내역": str(leaf).strip() if leaf is not None and not pd.isna(leaf) else "",
            "연간목표": float(annual) if not pd.isna(annual) else 0.0,
            "월목표": float(monthly) if not pd.isna(monthly) else 0.0,
        })
    if not rows:
        return {}
    leaf = pd.DataFrame(rows)
    by_process = {
        p: {"annual": float(g["연간목표"].sum()), "monthly": float(g["월목표"].sum())}
        for p, g in leaf.groupby("공정")
    }
    return {
        "year": year, "leaf": leaf, "source": sheet,
        "annual_total": float(leaf["연간목표"].sum()), "monthly_total": float(leaf["월목표"].sum()),
        "by_process": by_process,
    }
