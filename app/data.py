"""엑셀/CSV 원본 로딩 (Schema/ 실물 스키마 기준). Q-cost는 부적합수량만으로 계산하므로
생산량 파일은 더 이상 읽지 않는다(사용자 확인)."""
import os
import pandas as pd


def _table_ext(file_or_path) -> str:
    name = getattr(file_or_path, "name", None) or (file_or_path if isinstance(file_or_path, str) else "")
    return os.path.splitext(str(name))[1].lower()


def read_table(file_or_path):
    """실제 원본은 엑셀 파일(.xlsx/.xls)이고 더미데이터는 .csv — 둘 다 받아서 같은 방식으로 처리한다.
    classify.detect_process()가 공정을 모르는 상태에서 컬럼/라인명만 먼저 들여다볼 때도 재사용한다
    (스트림릿 업로드 파일은 seekable이라 이후 load_defect()에서 다시 읽기 전에 seek(0) 필요)."""
    ext = _table_ext(file_or_path)
    if ext in (".xlsx", ".xls"):
        return pd.read_excel(file_or_path)
    return pd.read_csv(file_or_path, encoding="utf-8-sig")


def _parse_date(series, fmt=None):
    """엑셀에서 읽으면 이미 datetime인 경우가 많고, CSV 더미데이터는 문자열이라 포맷 지정이 필요함."""
    if pd.api.types.is_datetime64_any_dtype(series):
        return series
    if fmt:
        parsed = pd.to_datetime(series, format=fmt, errors="coerce")
        if parsed.isna().mean() > 0.5:  # 포맷이 안 맞으면(엑셀 원본 등) 자유 파싱으로 재시도
            parsed = pd.to_datetime(series, errors="coerce")
        return parsed
    return pd.to_datetime(series, errors="coerce")


def load_defect(file_or_path, process: str) -> pd.DataFrame:
    """부적합 원본 로딩. A타입/B타입을 공정명이 아니라 **실제 컬럼 구성**으로 판별한다.
    (실물 파일 확인 결과 재공품도 조립과 같은 B타입 구조를 쓰는 경우가 있었음 — 04_prd_appendix.md 참고)
    - B타입 (재공품코드 컬럼 있음): 매칭키=재공품코드, 소구분=구분
    - A타입 (품명 컬럼 있음): 매칭키=품명, 소구분=부적합내용명
      날짜 컬럼명은 "일자" 또는 "년"으로 오는 경우가 있어 둘 다 확인한다.
    부적합중량g은 여기서 계산하지 않는다 — calc.add_qcost_columns가 재공품 중량정보.xls/제품 중량정보.xlsx의
    확정 기준(재공품=Wet, 조립=Dry, 완성=제품 총중량)으로 매칭키 기준 재계산한다.
    """
    df = read_table(file_or_path)
    df["공정"] = process

    if "재공품코드" in df.columns:
        df["매칭키"] = df["재공품코드"]
        df["소구분"] = df["구분"]
        date_col = "일자" if "일자" in df.columns else next(
            (c for c in ("년", "년도") if c in df.columns), None
        )
    elif "품명" in df.columns:
        df["매칭키"] = df["품명"]
        # 소구분은 텍스트 파싱이 아니라 전용 컬럼(부적합내용명)을 그대로 신뢰한다 (04_prd_appendix.md 확인)
        df["소구분"] = df["부적합내용명"]
        date_col = "일자" if "일자" in df.columns else next(
            (c for c in ("년", "년도") if c in df.columns), None
        )
    else:
        raise ValueError(
            f"{process} 부적합 파일에서 '재공품코드' 또는 '품명' 컬럼을 찾지 못했습니다. "
            f"실제 컬럼: {list(df.columns)}"
        )

    if date_col is None:
        raise ValueError(
            f"{process} 부적합 파일에서 날짜 컬럼('일자'/'년')을 찾지 못했습니다. "
            f"실제 컬럼: {list(df.columns)}"
        )

    df["일자_dt"] = _parse_date(df[date_col], fmt="%y/%m/%d")
    # 날짜를 못 읽은 행(NaT)이 하나라도 섞이면 .dt.month가 NaN을 만들어 컬럼 전체가 float64로
    # 승격되고, 필터·상세표에 "8"이 아니라 "8.0"으로 보이는 원인이 된다. 결측을 허용하는
    # nullable Int64로 캐스팅하면 값은 정수로 보이면서 NaT행은 <NA>로 남는다(둘 다 만족).
    df["월"] = df["일자_dt"].dt.month.astype("Int64")
    return df
