"""제품군/세부공정 매핑(필요자료/생산량 부적합량 라인명.xlsx 기준) + 세부내역 분류
(AGM 조립 확정 규칙 + 그 외 공정/제품군용 '부적합 유형 기준 정보.xlsx' 공통 매핑표)."""
import pandas as pd


def load_line_family_rules(xlsx_path: str) -> dict:
    """`생산량 부적합량 라인명.xlsx`(Sheet3)를 파싱해 {공정구분: [(구분명, [키워드,...]), ...]} 반환.
    - 재공품: 1연도/2연도(W/F)/2연도(건조,절단,화성) — `총수율 표 양식.xlsx`(부적합률(창원) 시트)의
      공식 리포트 대분류 3그룹 기준으로 병합/개명함(fallback 없음 — 미매칭은 오류)
    - 조립·완성: AGM/EBGC/고정형/지게차 (고정형은 키워드가 없는 fallback 카테고리)
    """
    import openpyxl
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Sheet3"]

    rows = list(ws.iter_rows(min_row=9, max_row=ws.max_row, values_only=True))
    header, data_rows = rows[0], rows[1:]
    # 열 순서(F~M, 시트 구조 확인됨): 공정구분,제품구분,세부공정,수량집계Line명(생산),키워드,비고,...
    idx_process, idx_family, idx_detail, idx_keyword = 5, 6, 7, 9

    rules: dict[str, list] = {}
    cur_process = None
    for r in data_rows:
        process = r[idx_process] or cur_process
        cur_process = process
        # 건조절단/화성처럼 제품구분(G)이 비어있고 세부공정(H)만 채워진 행은 세부공정명을 구분명으로 쓴다
        family = r[idx_family] or r[idx_detail]
        keyword_raw = r[idx_keyword]
        if not process or not family:
            continue
        keywords = []
        if keyword_raw:
            # 줄바꿈은 별개 키워드 구분자. ", "(콤마+공백)는 "paste, 도장"처럼 나열 구분자지만,
            # "AGM완성_1,2차"처럼 공백 없는 콤마는 키워드 자체의 일부이므로 쪼개지 않는다.
            for line in str(keyword_raw).split("\n"):
                for part in line.split(", "):
                    part = part.strip()
                    if part and part != "-":
                        keywords.append(part)
        rules.setdefault(process, []).append((family, keywords))
    return _merge_recon_groups(_apply_known_supplements(rules))


def _apply_known_supplements(rules: dict) -> dict:
    """생산량 부적합량 라인명.xlsx 매핑표에는 없지만, 사용자가 실제 라인명 목록(스크린샷)으로
    확인해준 추가 키워드. 원본 매핑표 파일을 직접 고치는 대신 여기서 보강한다."""
    for family, keywords in rules.get("재공품", []):
        if family == "1연도" and "1P" not in keywords:
            # 1P-1~8(터널/고온/흐름숙성) 라인은 1연도 부적합 파일에만 나타남 — 사용자 확인 완료
            keywords.append("1P")
    # ES조립/VGS조립("VGS조립_CGS_1라인" 포함, "VGS조립" 접두어만으로 CGS도 함께 매치됨)·ES완성 라인은
    # "고정형 라인에서 나온 것"으로 사용자 확인(2026-09-11) — 원래 "고정형"은 키워드 없는 fallback
    # 카테고리라 이 라인들이 매핑표 키워드(AGM COS/GC COS/지게차 자동,수동)에 안 걸려 detect_process()가
    # "미상"으로 판정하던 문제가 있었음. 기존 fallback 항목은 그대로 두고, 키워드 있는 "고정형" 항목을
    # 별도로 추가해 이 라인들만 명시적으로 매치시킨다(다른 미확인 라인은 여전히 기존 fallback으로 감).
    rules.setdefault("조립", []).append(("고정형", ["ES조립", "VGS조립"]))
    rules.setdefault("완성", []).append(("고정형", ["ES완성"]))
    # (v24.4) 매핑표의 완성 AGM 키워드는 "AGM완성_2차"/"AGM완성_1,2차"뿐이라 "AGM완성_1차 A/C"와
    # "AGM충전1W->3L" 같은 충전 라인이 전부 고정형(fallback)으로 가고 있었음 — 사용자 엑셀 계산본과
    # 대조한 결과 이 라인들은 AGM으로 집계돼 있음(AGM 검사시료·외관 파손·저전압 합계가 이 라인들을
    # 합쳐야 엑셀과 일치, 고정형은 ES완성 라인만). 원본 파일 대신 여기서 보강한다.
    for i, (family, keywords) in enumerate(rules.get("완성", [])):
        if family == "AGM":
            for kw in ("AGM완성_1차", "AGM충전"):
                if kw not in keywords:
                    keywords.append(kw)
            break
    return rules


# 원본 매핑표는 재공품을 1연도/2연도/건조절단/화성 4그룹으로 나누지만, 공식 리포트
# (총수율 표 양식.xlsx의 부적합률(창원) 시트)는 2연도(W/F)와 2연도(건조,절단,화성)를 묶어 3그룹으로 본다.
_RECON_GROUP_RENAME = {"1연도": "1연도", "2연도": "2연도(W/F)"}
_RECON_MERGE_INTO = "2연도(건조,절단,화성)"
_RECON_MERGE_SOURCES = {"건조절단", "화성"}


def _merge_recon_groups(rules: dict) -> dict:
    entries = rules.get("재공품")
    if not entries:
        return rules
    merged: list[tuple[str, list]] = []
    merged_keywords: list = []
    for family, keywords in entries:
        if family in _RECON_MERGE_SOURCES:
            merged_keywords.extend(keywords)
            continue
        merged.append((_RECON_GROUP_RENAME.get(family, family), keywords))
    if merged_keywords:
        merged.append((_RECON_MERGE_INTO, merged_keywords))
    rules["재공품"] = merged
    return rules


def classify_line(process: str, line_name: str, rules: dict) -> tuple[str, bool]:
    """(구분명, 매칭성공여부) 반환. 매칭성공여부=False면 매핑표 어디에도 속하지 않는 미분류 데이터.
    실 데이터 확인 결과 "paste"(키워드) vs "Punched_Paste"(실제 라인명)처럼 대소문자가 달라
    매칭이 실패하는 경우가 있어 대소문자 구분 없이 비교한다."""
    line_name_lower = (line_name or "").lower()
    entries = rules.get(process, [])
    for family, keywords in entries:
        if keywords and any(kw.lower() in line_name_lower for kw in keywords):
            return family, True
    # 키워드가 비어있는 항목(예: 조립/완성의 "고정형")은 명시적 fallback 카테고리
    fallback = [family for family, keywords in entries if not keywords]
    if fallback:
        return fallback[0], True
    return "미분류", False


def _strict_keyword_match(process: str, line_name: str, rules: dict) -> bool:
    """classify_line()과 달리 키워드 없는 fallback 카테고리(예: 조립의 '고정형')는 매치로 치지 않는다.
    공정 자동판별(detect_process)에서 fallback까지 매치로 인정하면 재공품 라인명도 전부 '조립의
    고정형'으로 잘못 잡히는 오탐이 생기기 때문 — 실제 키워드가 걸린 경우만 매치로 본다."""
    line_name_lower = (line_name or "").lower()
    entries = rules.get(process, [])
    return any(keywords and any(kw.lower() in line_name_lower for kw in keywords) for _, keywords in entries)


def detect_process(df, line_family_rules: dict) -> str:
    """공정 구분 없이 올라온 부적합 파일이 재공품/조립/완성 중 무엇인지 컬럼 구성·라인명으로 추정한다.

    (v14.2 정정) 애초엔 "'재공품코드' 컬럼이 없으면 무조건 완성"으로 가정했으나, 실제 운영 데이터
    (필요자료/부적합데이터/1연도·2연도·건조절단·재공품 화성 부적합 파일)로 확인해보니 **재공품 데이터도
    완성과 똑같은 A타입(품명 컬럼, 재공품코드 없음) 스키마로 오는 경우가 있음** — 이 가정 때문에
    재공품 파일 전체가 "완성"으로 잘못 라벨링되어 완성/고정형에 뒤섞이고 재공품 데이터 자체가
    화면(파이차트 등)에서 통째로 사라지는 버그가 있었다. 이제 A타입이어도 라인명을 재공품
    키워드(paste/도장/W-F/절단/화성/1P 등)에 먼저 대조해보고, 걸리면 재공품으로 판정한다.

    - B타입('재공품코드' 컬럼 있음): 재공품/조립 각각의 키워드 매핑표(생산량 부적합량 라인명.xlsx)에
      대조해(fallback 카테고리는 제외) 매치 건수가 더 많은 쪽을 채택
    - A타입('품명' 컬럼, 재공품코드 없음): 재공품 키워드에 걸리면 재공품, 아니면 완성(기본값)
    라인명 컬럼 자체가 없거나 값이 비어 B타입 판별이 안 되면 "미상"을 반환한다(자동판별 실패 —
    화면에서 사용자 확인 필요). 매핑표에 아예 없는 라인명(예: "ES조립_...", "VGS조립_...")도
    "미상"으로 남는다 — 이 경우는 코드 버그가 아니라 매핑표 자체에 그 라인 키워드가 없어서이므로
    매핑표 보강이 필요하다."""
    line_col = "라인명" if "라인명" in df.columns else ("Line명" if "Line명" in df.columns else None)
    lines = df[line_col].dropna().astype(str) if line_col else pd.Series(dtype=str)

    if "재공품코드" not in df.columns:
        if lines.empty:
            return "완성"
        if any(_strict_keyword_match("재공품", ln, line_family_rules) for ln in lines):
            return "재공품"
        return "완성"

    if lines.empty:
        return "미상"
    score = {
        proc: sum(_strict_keyword_match(proc, ln, line_family_rules) for ln in lines)
        for proc in ("재공품", "조립")
    }
    if score["재공품"] == 0 and score["조립"] == 0:
        return "미상"
    return max(score, key=score.get)


def load_plate_generation_lookup(xlsx_path: str) -> dict:
    """AGM조립 극판기준정보.xlsx 기준정보 시트 → {재공품코드: 세대} 매핑."""
    try:
        df = pd.read_excel(xlsx_path, sheet_name="기준정보", header=2)
    except Exception:
        return {}
    df = df.rename(columns=lambda c: str(c).strip())
    if "재공품 코드" not in df.columns or "세대" not in df.columns:
        return {}
    df = df.dropna(subset=["재공품 코드"])
    return dict(zip(df["재공품 코드"], df["세대"]))


def load_detail_type_rules(xlsx_path: str) -> dict:
    """`부적합 유형 기준 정보.xlsx` → {부적합유형: 총수율 유형} 매핑.
    담당자 7명이 각자 관리하는 개별 엑셀(EBGC/V-Type/고정형/AGM완성)에 공통으로 들어있던
    마스터 분류표를 추출한 것 — **완성(AGM/EBGC/고정형)**에 적용된다.
    "총수율 유형" 컬럼은 `총수율 표 양식.xlsx`(부적합률(창원) 시트)의 공식 리포트 어휘와
    정확히 일치하도록 만들어진 컬럼이라 이걸 우선 쓰고("2차 유형"은 더 세분화된 참고용이라 fallback).
    조립(AGM/EBGC/고정형/지게차)은 이 표와 무관하게 별도 확정 규칙(생파/스태커·COS·콤비/시료)을 쓴다.

    (v19.0 정정) 예전엔 "검사 시료"/"개발 시료"처럼 이 표가 시료를 세분화해둔 값도 전부 "시료" 하나로
    뭉갰었는데(총수율 종합표의 "시료" 소분류가 검사시료/개발시료로 나뉘어 있어도 항상 0으로 보이던
    원인), "불량코드내역에 검사시료/개발시료가 나뉘어 있다"는 사용자 확인에 따라 원본 표의 "총수율
    유형" 값을 그대로 살린다 — 이 표 자체가 이미 "품질 시료"→"검사 시료", "기술 시료"→"개발 시료"라는
    회사 공식 대응관계를 담고 있어(사용자가 준 원본 파일 그대로), 추가로 추측할 필요가 없다."""
    try:
        df = pd.read_excel(xlsx_path, sheet_name=0)
    except Exception:
        return {}
    df = df.rename(columns=lambda c: str(c).strip())
    if "부적합 유형" not in df.columns:
        return {}
    value_col = "총수율 유형" if "총수율 유형" in df.columns else "2차 유형"
    if value_col not in df.columns:
        return {}
    df = df.dropna(subset=["부적합 유형"])
    return {str(k).strip(): str(v).strip() for k, v in zip(df["부적합 유형"], df[value_col])}


def extract_raw_defect_text(row) -> str:
    """원본 불량 설명 텍스트를 최대한 뽑아낸다.
    - 조립(B타입)은 `불량코드내역`에 항상 채워져 있음(예: "[콤비]_융착 강도 부적합")
    - 재공품(B타입)은 `불량코드내역`이 거의 비어있고, 대신 `불량명`이
      "코드,설명,수량" 형태라 가운데 설명 부분을 파싱해야 함(실 데이터로 확인됨)
    - 완성(A타입)은 `부적합코드내용명`에 채워져 있음(예: "[고율방전]_저전류")
    """
    code_desc = row.get("불량코드내역")
    if isinstance(code_desc, str) and code_desc.strip():
        return code_desc.strip()

    # (v24.4) 조립 지게차 파일은 `부적합명`, EBGC 조립 파일은 `부적합 명`(공백 포함) 컬럼에 같은 내용
    # ("[스태커]_작업중 극판 파손" 등)이 들어있는데 여기서 안 읽어 두 제품군의 취파가 전부 "기타"로,
    # 시료가 전부 "시료"로만 떨어지던 버그(사용자가 엑셀 계산본과 대조해 발견 — 지게차/EBGC 스태커·COS·
    # 콤비·검사시료 칸이 MVP에서만 0).
    for col in ("부적합명", "부적합 명"):
        v = row.get(col)
        if isinstance(v, str) and v.strip():
            return v.strip()

    # "코드,설명,수량" 형태(재공품 `불량명`, 일부 조립 파일 `종합`)는 가운데 설명 부분만 쓴다.
    for col in ("불량명", "종합"):
        v = row.get(col)
        if isinstance(v, str) and "," in v:
            parts = v.split(",")
            if len(parts) >= 2 and parts[1].strip():
                return parts[1].strip()

    a_type_desc = row.get("부적합코드내용명")
    if isinstance(a_type_desc, str) and a_type_desc.strip():
        return a_type_desc.strip()

    return ""


def classify_detail_assembly(sub_category: str, defect_desc: str, product_code: str,
                              plate_gen_lookup: dict) -> str:
    """조립(AGM/EBGC/고정형/지게차 공통) 세부내역 분류 규칙 (04_prd_appendix.md 3장 확정,
    `총수율 표 양식.xlsx` 부적합률(창원) 시트로 재확인 — 조립극판 4개 제품군이 모두 동일하게
    생파(공정 생파/1세대 극판) + 공정부적합(스태커/COS/콤비/기타) + 조건설정(시료) 구조를 씀).
    소구분(생파/취파/시료)은 별도 컬럼값을 그대로 쓰고, 여기서는 세부내역만 정한다.
    "1세대 극판" 판별은 AGM조립 극판기준정보.xlsx 조회 기준이라 실질적으로 AGM에서만 나타나고,
    다른 제품군 재공품코드는 조회에 없어 자동으로 "공정 생파"로 떨어진다(안전한 공용 규칙)."""
    desc = defect_desc or ""
    if sub_category == "생파":
        if desc.startswith("[1연도 극판]") or desc.startswith("[1연도극판]"):
            gen = plate_gen_lookup.get(product_code, "")
            if gen == "1세대":
                return "1세대 극판"
        return "공정 생파"
    if sub_category == "취파":
        for tag in ["스태커", "COS", "콤비"]:
            if f"[{tag}]" in desc:
                return f"[{tag}]"
        return "기타"
    if sub_category == "시료":
        # (v19.0) "검사 시료"/"조건 설정"처럼 원본 `불량코드내역`이 시료를 더 세분화해둔 경우가
        # 있어(사용자 확인) 예전처럼 뭉뚱그려 "시료"로 고정하지 않고 그 값을 그대로 살린다 —
        # 총수율 종합표의 "검사시료"/"개발시료" 세분화에 이 값이 필요함(fixed_report.py 참고).
        return desc.strip() if desc.strip() else "시료"
    return sub_category if isinstance(sub_category, str) and sub_category else "기타"


def classify_detail(process: str, family: str, sub_category: str, raw_text: str,
                     product_code: str, plate_gen_lookup: dict, detail_type_rules: dict) -> str:
    """세부내역 분류.
    - 조립(AGM/EBGC/고정형/지게차 전부): 확정 규칙(1세대 극판, 스태커/COS/콤비, 시료) 공통 적용
      (예전엔 AGM만 이 규칙을 쓰고 나머지는 완성용 매핑표를 잘못 조회했음 — 총수율 표 양식.xlsx로 버그 확인 후 수정)
    - 완성(AGM/EBGC/고정형): '부적합 유형 기준 정보'의 "총수율 유형" 컬럼(공식 리포트 어휘) 조회
    - 재공품/그 외: 전용 매핑표가 없어 원본 설명 텍스트를 세부내역으로 그대로 사용

    (v19.0 정정) `detail_type_rules`(부적합 유형 기준 정보.xlsx)는 완성 전용 마스터 표인데, 예전엔
    process 구분 없이 조립만 빼고 다 조회해서 재공품도 여기 걸리면(예: 원본 텍스트가 우연히 "품질
    시료" 등과 똑같으면) 완성 기준 어휘로 바뀌어버렸다 — 재공품 1연도는 원본 텍스트("품질 시료"/
    "기술 시료")가 이미 자기 양식의 정식 세부내역 이름과 같아서 바꿀 필요가 없었는데 잘못 걸리던 것.
    이제 이 표는 완성에만 적용한다."""
    if process == "조립":
        return classify_detail_assembly(sub_category, raw_text, product_code, plate_gen_lookup)

    if process == "완성":
        mapped = detail_type_rules.get(raw_text)
        if mapped:
            return mapped

    if raw_text:
        return raw_text
    if isinstance(sub_category, str) and sub_category:
        return sub_category
    return "기타"


_RECON_CONDITION_KEYWORDS = ("조건 설정", "조건설정", "기타(보정)")


def classify_subtype(process: str, sub_category: str, detail: str) -> str:
    """대분류 하위 소분류(생파/공정부적합/조건설정) — `총수율 표 양식.xlsx`(부적합률(창원) 시트)의
    실제 리포트 구조 기준. 기존 '소구분'(생파/취파/시료 원본 컬럼값)은 리포트의 소분류 개념과
    다르다(예: 재공품 "조건 설정_초기 시운전"은 구분="취파"로 들어오지만 리포트상 조건설정 항목,
    "시료"는 항상 조건설정 하위 항목) — 이 함수가 그 간극을 메운다.

    (v18.0 정정) 원본 '생파' 플래그는 조립/완성에서는 실제로 별도 대분류(공식 양식에도 "생파" 행이
    있음)이지만, 재공품은 사용자 확인 결과 다르다 — "1연도 재공품에는 생파가 없고 시료/취파(=생산중
    순수불량)로만 나뉜다". 원본 파일에 '생파'로 표시된 재공품 행이 소량 있었던 건 이 분류 체계상
    실수/예외로 보이고, 실제 내용(세부내역 텍스트)을 보면 조건설정/공정부적합 어느 한쪽에 명백히
    속한다 — 그래서 재공품은 '생파' 플래그를 무시하고 항상 텍스트 내용으로 판정한다(총수율 종합표에
    재공품용 생파 칸이 원본에 없는 것과도 일치). 조립/완성은 기존 동작 그대로 유지."""
    detail_str = detail if isinstance(detail, str) else ""
    if process == "재공품":
        if sub_category == "시료" or detail_str == "시료":
            return "조건설정"
        if any(kw in detail_str for kw in _RECON_CONDITION_KEYWORDS):
            return "조건설정"
        return "공정부적합"
    if sub_category == "생파":
        return "생파"
    if sub_category == "시료" or detail == "시료":
        return "조건설정"
    return "공정부적합"


def assign_family_and_detail(df: pd.DataFrame, process_col="공정",
                              plate_gen_lookup: dict = None,
                              line_family_rules: dict = None,
                              detail_type_rules: dict = None) -> pd.DataFrame:
    """공정별로 컬럼 구성이 다르므로(재공품·완성=Line명, 조립=라인명) 행 단위로 값을 고른다.
    line_family_rules가 없으면 전부 '미분류'로 표시된다(매핑표 필수)."""
    plate_gen_lookup = plate_gen_lookup or {}
    line_family_rules = line_family_rules or {}
    detail_type_rules = detail_type_rules or {}
    df = df.copy()

    def _line_name(r):
        for col in ("Line명", "라인명"):
            v = r.get(col)
            if isinstance(v, str) and v:
                return v
        return ""

    classified = df.apply(
        lambda r: classify_line(r[process_col], _line_name(r), line_family_rules), axis=1
    )
    df["제품군"] = classified.apply(lambda t: t[0])
    df["제품군_매칭됨"] = classified.apply(lambda t: t[1])
    # (v20.0) 재공품·완성은 "Line명", 조립은 "라인명" 컬럼을 쓰는데, 라인별 비교 화면에서 공정과
    # 무관하게 하나의 컬럼으로 묶어 그룹화할 수 있도록 통합 컬럼을 만들어둔다.
    df["라인명_통합"] = df.apply(_line_name, axis=1).replace("", pd.NA)

    def _detail(r):
        raw_text = extract_raw_defect_text(r)
        code = r.get("재공품코드", r.get("매칭키", ""))
        return classify_detail(
            r[process_col], r["제품군"], r["소구분"], raw_text, code,
            plate_gen_lookup, detail_type_rules,
        )

    df["세부내역"] = df.apply(_detail, axis=1)
    df["소분류"] = df.apply(
        lambda r: classify_subtype(r[process_col], r["소구분"], r["세부내역"]), axis=1
    )
    # (v24.0) 세부내역보다 한 단계 아래의 원본 부적합 텍스트(예: 세부내역 "스태커" 안의
    # "[스태커]_작업중 극판 파손") — 기간 비교 세부분석에서 가장 깊은 선택 단계로 쓴다.
    raw_text = df.apply(extract_raw_defect_text, axis=1)
    df["상세내역_원본"] = raw_text.where(raw_text != "", "(미기재)")
    # (v24.5) 사용자 확인(2026-09-29): "2연도 W/F 공정에서 [주조] 태그가 달린 부적합은 녹여 재용해하므로
    # 총수율 Q-cost에는 들어가지 않음" — 해당 행은 집계에서 제외한다(사유를 남겨 화면에서 건수·금액 안내).
    is_wf_casting = (df[process_col] == "재공품") & (df["제품군"] == "2연도(W/F)") & raw_text.str.startswith("[주조]")
    df["집계제외사유"] = pd.NA
    df.loc[is_wf_casting, "집계제외사유"] = "2연도(W/F) [주조] 태그 — 재용해(총수율 Q-cost 제외)"
    return df
