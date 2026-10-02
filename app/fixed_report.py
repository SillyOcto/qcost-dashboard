"""고정 양식 종합표 — `필요자료/총수율 표 양식.xlsx`의 "폐기비용(창원) 보정전" 시트 구조를 그대로 재현한다.
사용자 요청(2026-09-11): "이 양식은 필터를 걸어도 영향을 받지 않고 그대로 보여줬으면 해. 그리고 8월만
데이터가 들어가서 8월 열만 보여주는데, 틀은 고정으로 두고 값만 채우는 형식으로 가자."

- 행(공정→제품군→소분류→세부내역)과 열(1~12월+합계)이 실제 업로드된 데이터와 무관하게 항상 그대로
  나온다 — 데이터가 없는 달/항목은 0으로 비워둘 뿐, 행/열 자체가 사라지지 않는다.
- 화면 상단의 공정/제품군/월 필터 영향을 받지 않는다(항상 전체 데이터 기준).

한계(정확히 매칭 안 되는 부분, 근사치로 처리 — app.py 화면에도 안내문 표시):
- 재공품 세부내역은 전용 매핑표가 없어 원본 텍스트를 그대로 쓰는데(classify.py 3장 참고), 원본 텍스트가
  이 양식의 깔끔한 라벨(예: "수집부")과 정확히 안 겹치는 경우가 있어 키워드 포함 여부로 근사 매칭한다.
  정확한 1:1 매핑은 담당자 확인이 필요해 보류 중(사용자 제공 매핑 검토 파일로 순차 반영 중).
- (v18.0 이전) 재공품 구분="생파"인 행은 이 양식에 재공품용 "생파" 버킷이 없어 표에서 제외했었으나,
  "1연도 재공품에는 생파가 없고 시료/취파로만 나뉜다"는 사용자 확인에 따라 classify.classify_subtype()이
  재공품의 원본 '생파' 플래그 자체를 더 이상 쓰지 않고 텍스트 내용(조건설정 키워드/시료 여부)으로만
  판정하도록 바뀌었다 — 그 결과 재공품 행은 항상 공정부적합/조건설정 둘 중 하나로만 분류되어 이 표에
  전부 반영된다(제외되는 행 없음).
- 이 양식은 "조건설정"(시료 1건)과 "시료"(검사시료/개발시료로 세분화) 소분류를 따로 두지만, 우리 쪽은
  시료 계열을 전부 하나로 합쳐서만 관리한다 — 시료 합계는 "조건설정→시료"(또는 재공품은 family별 해당
  칸)에 몰아넣고, "시료→검사시료/개발시료" 행은 항상 0으로 둔다.
"""
import re

import pandas as pd

MONTHS = list(range(1, 13))

_TAG_PREFIX_RE = re.compile(r"^\[[^\]]*\]_?")


def _strip_tag_prefix(text: str) -> str:
    """원본 텍스트 앞의 "[도장]_"/"[충진]_" 같은 생산 공정 태그를 뗀 내용만 남긴다. 2연도(건조,절단,
    화성)는 태그 이름 자체가 "도장"이라 실제 내용과 무관하게 "도장"(→"미도장") 키워드에 걸리는 오매칭이
    있었다(v18.0에서 발견) — 태그를 먼저 떼어내면 그런 우연한 충돌을 줄일 수 있다."""
    return _TAG_PREFIX_RE.sub("", text)

# (공정, 제품군) 블록 순서 + 그 블록에 속한 (소분류, [세부내역, ...]) 리프 행 — 시트 원문 그대로.
_BLOCKS = [
    ("재공품", "1연도", [
        ("생산중 순수불량", ["수집부", "로봇 적재", "기타"]),
        ("초기시운전", ["초기 시운전", "부자재 교체", "기타"]),
        # (v24.6) 사용자 확인: 구분=시료인데 부적합 내용이 비어있는 행(1P 흐름숙성/고온/터널 라인 시료 등)은
        # "시료" 블록의 기타로 — 사용자 양식 화면에도 시료 아래 기타 행이 있음.
        ("시료", ["기술 시료", "품질 시료", "기타"]),
        ("기타(보정)", [""]),
    ]),
    ("재공품", "2연도(W/F)", [
        ("생산중 순수불량", ["차입", "파이프 막힘", "심금 절단", "기판 변형", "기타"]),
        ("초기시운전", ["초기 시운전", "시료"]),
        ("시료", ["검사시료", "개발시료", "기타(보정)"]),
    ]),
    ("재공품", "2연도(건조,절단,화성)", [
        ("생산중 순수불량", ["목락", "변형", "미도장", "미미크랙", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("조립", "AGM", [
        ("생파", ["공정 생파", "1세대 극판"]),
        ("생산중 순수불량", ["스태커", "콤비", "COS", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    # (v24.4) 지게차/EBGC/고정형의 "1세대 극판" 리프는 양식 원문에 없다 — v14.6에 값이 새지 않도록 임의로
    # 추가했었는데, 사용자 엑셀 계산본과 대조하니 이 제품군의 1세대 극판은 "공정 생파"에 합산돼 있어
    # (고정형 공정 생파 45.95 = MVP 공정 생파 39.78 + 1세대 극판 5.9) 양식 그대로 되돌리고
    # _map_assembly_or_finished()에서 공정 생파로 보낸다.
    ("조립", "지게차", [
        ("생파", ["공정 생파"]),
        ("생산중 순수불량", ["스태커", "COS", "콤비", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("조립", "EBGC", [
        ("생파", ["공정 생파"]),
        ("생산중 순수불량", ["스태커", "COS", "콤비", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("조립", "고정형", [
        ("생파", ["공정 생파"]),
        ("생산중 순수불량", ["스태커", "COS", "콤비", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("완성", "AGM", [
        ("생산중 순수불량", ["저전압", "저전류", "외관 파손", "핀홀", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("완성", "EBGC", [
        ("생산중 순수불량", ["저전압", "외관 부적합", "저전류", "핀홀", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    ("완성", "고정형", [
        ("생산중 순수불량", ["저전압", "저전류", "외관 파손", "핀홀", "기타"]),
        ("조건설정", ["시료"]),
        ("시료", ["검사시료", "개발시료"]),
    ]),
    # (v15.4) 완성 V-type은 부적합 자체가 없다는 사용자 확인에 따라 칸 삭제 — 혹시 관련 데이터가
    # 들어오면 _snap_to_leaf()가 갈 곳을 못 찾아 get_unmapped_summary()에 잡히므로 조용히 빠지지 않음.
]

_MAJOR_LABEL = {"재공품": "재공품", "조립": "조립극판", "완성": "완성공정"}

# 재공품 원본 텍스트 → 이 양식의 "생산중 순수불량" 세부내역 키워드(가족별). 먼저 걸리는 것을 채택.
_RECON_ABNORMAL_KEYWORDS = {
    "1연도": [("수집부", "수집부"), ("적재", "로봇 적재")],
    "2연도(W/F)": [
        # (v24.6) "변형" 키워드 삭제 — [충진]_극판변형·[자재]_튜브 변형은 사용자 확인대로 "기타"로
        # ([주조]_기판 변형은 v24.5 재용해 제외 규칙으로 아예 집계 안 됨).
        ("차입", "차입"), ("파이프", "파이프 막힘"), ("심금", "심금 절단"),
    ],
    "2연도(건조,절단,화성)": [
        ("목락", "목락"), ("변형", "변형"), ("도장", "미도장"), ("크랙", "미미크랙"),
    ],
}
# 재공품 원본 텍스트 → "초기시운전" 세부내역 키워드(1연도/2연도(W/F)만 해당).
_RECON_SETUP_KEYWORDS = {
    "1연도": [("초기 시운전", "초기 시운전"), ("교체", "부자재 교체")],
    # (v18.0) "조건 설정_형명(금형) 교체"가 기본값("시료")으로 빠지고 있었는데, 사용자가 매핑
    # 검토 파일에서 2연도(W/F)는 "초기 시운전"으로 확정해줬다(1연도는 그대로 "부자재 교체" 유지 —
    # 사용자가 그 행은 안 건드림, 패밀리별로 다르게 두는 게 맞는 것으로 확인됨).
    "2연도(W/F)": [("초기 시운전", "초기 시운전"), ("형명", "초기 시운전")],
}


def _map_recon(family: str, subtype: str, raw_text: str):
    """재공품 원본 텍스트를 양식의 (소분류, 세부내역)으로 근사 매칭한다. 매칭 대상이 없으면 None.

    (v18.0) 예전엔 subtype이 "생파"면 무조건 None(제외)이었다 — 원본 양식에 재공품용 "생파" 칸이
    없기 때문. 그런데 classify.classify_subtype()이 재공품에 대해서는 원본 '생파' 플래그를 아예 안 쓰고
    텍스트 내용으로 조건설정/공정부적합만 판정하도록 바뀌어서(사용자 확인: "1연도 재공품에는 생파가
    없고 시료/취파로만 나뉜다"), 여기 도달하는 subtype은 이제 "공정부적합"/"조건설정" 둘 중 하나뿐이라
    이 분기 자체가 더 이상 발생하지 않는다(죽은 코드라 삭제)."""
    text = _strip_tag_prefix(raw_text or "")
    if subtype == "조건설정":
        if text == "기타(보정)":
            if family == "1연도":
                return ("기타(보정)", "")
            if family == "2연도(W/F)":
                return ("시료", "기타(보정)")
            return ("조건설정", "시료")  # 2연도(건조,절단,화성)엔 기타(보정) 자리가 없어 조건설정으로
        # (v19.0) 시료 세분화 — 2연도(W/F)/2연도(건조,절단,화성)의 "시료" 소분류는 템플릿에 검사시료/
        # 개발시료 리프가 있다(1연도는 "기술 시료"/"품질 시료"라는 자기만의 이름을 그대로 쓰므로 여기서
        # 건드리지 않음 — raw_text가 이미 그 이름 그대로 들어오면 아래 생산중 순수불량 분기 이전에
        # _snap_to_leaf()가 알아서 맞는 리프로 붙여준다).
        if family == "2연도(W/F)":
            sample_leaf = _SAMPLE_LEAF_MAP.get(text.replace(" ", ""))
            if sample_leaf:
                return ("시료", sample_leaf)
        elif family == "2연도(건조,절단,화성)":
            # (v24.4) 사용자 엑셀 계산본은 이 제품군의 검사/품질 시료를 "조건설정/시료"에 집계하고
            # "시료/검사시료·개발시료" 행은 비워둔다(4월 검사 시료 0.086 = 엑셀 조건설정/시료 4월 0.07)
            # — 엑셀과 같은 칸에 넣는다.
            return ("조건설정", "시료")
        else:
            # (v24.4) 1연도는 양식의 "시료" 블록이 "기술 시료"/"품질 시료"라는 자기 이름을 그대로 쓴다 —
            # 예전 주석은 _snap_to_leaf()가 알아서 붙여줄 거라 했지만 실제로는 아래 초기시운전 분기가
            # 먼저 걸려 "초기시운전/기타"로 빠지고 있었음(사용자 엑셀 대조: 기술 시료 12.8백만원이 MVP 0).
            if text.replace(" ", "") in ("기술시료", "품질시료"):
                return ("시료", "기술 시료" if text.startswith("기술") else "품질 시료")
            if text in ("", "시료"):
                return ("시료", "기타")  # (v24.6) 내용 없는 시료 행 → 시료/기타(사용자 확인)
            # (v24.4) 1연도의 "조건 설정_형명(금형) 교체"는 사용자 엑셀 계산본에서 "기타"로 집계됨
            # (엑셀 기타 23.0 ≈ MVP 형명 26.3/1.13, 부자재 교체는 SPOOL·연도지 교체만으로 일치) —
            # "교체" 키워드로 부자재 교체에 섞이기 전에 먼저 기타로 보낸다. 2연도(W/F)는 사용자 확정대로
            # 초기 시운전 유지.
            if "형명" in text:
                return ("초기시운전", "기타")
        if family in ("1연도", "2연도(W/F)"):
            for kw, label in _RECON_SETUP_KEYWORDS.get(family, []):
                if kw in text:
                    return ("초기시운전", label)
            return ("초기시운전", "시료" if family == "2연도(W/F)" else "기타")
        return ("조건설정", "시료")  # 2연도(건조,절단,화성)
    # subtype == "공정부적합" → "생산중 순수불량" 그룹
    for kw, label in _RECON_ABNORMAL_KEYWORDS.get(family, []):
        if kw in text:
            return ("생산중 순수불량", label)
    return ("생산중 순수불량", "기타")


# (v19.0) "부적합 유형 기준 정보.xlsx"("총수율 유형" 컬럼) 자체가 담고 있는 회사 공식 대응관계 —
# 품질/검사 시료는 "검사시료", 기술 시료는 "개발시료"로 귀결된다(사용자가 준 원본 파일 그대로,
# 추측 아님). 공백 유무 표기가 섞여 있어(품질 시료/품질시료 등) 공백을 지우고 비교한다.
_SAMPLE_LEAF_MAP = {
    "검사시료": "검사시료", "품질시료": "검사시료",
    "개발시료": "개발시료", "기술시료": "개발시료",
}


# (v24.4) 완성 EBGC는 완성용 매핑표(부적합 유형 기준 정보.xlsx, AGM 중심)에 없는 원본 텍스트가 많아
# 그대로 "기타"로 빠지고 있었다. 사용자 엑셀 계산본과 대조해 확정한 대응: "[충전]_…"/"[외관]_…" 계열은
# 양식의 "외관 부적합" 행, "[고율방전]_방전 전압 부적합"은 "저전류" 행(엑셀 2월 0.14·6월 0.06 = MVP 2월
# 0.139·6월 0.072로 월별 금액까지 일치). 그 외 미매핑 텍스트는 기존대로 "기타".
_FINISHED_EBGC_PREFIX_MAP = (("[충전]", "외관 부적합"), ("[외관]", "외관 부적합"))
_FINISHED_EBGC_EXACT_MAP = {"[고율방전]_방전 전압 부적합": "저전류", "외관 파손": "외관 부적합"}


def _map_assembly_or_finished(subtype: str, detail: str, process: str = None, family: str = None):
    """조립/완성은 이미 우리 분류가 이 양식과 대부분 일치한다 — 소분류만 이름을 맞추고,
    세부내역은 "[스태커]"처럼 완전히 대괄호로 감싸인 경우만 안쪽 텍스트를 꺼내 쓴다("[충전]_복배"처럼
    매핑표에 없어 원본 텍스트가 그대로 온 값은 앞쪽 "["만 어설프게 잘리지 않도록 건드리지 않음 —
    이런 값은 아래 _snap_to_leaf()가 해당 그룹의 "기타"로 정리한다).

    (v19.0) 시료 세분화(검사시료/개발시료) 반영 — 예전엔 subtype=="조건설정"이면 무조건 "조건설정/시료"
    로 뭉뚱그려서 원본 양식의 별도 "시료"(검사시료/개발시료) 소분류 행이 항상 0이었다. 이제 세부내역이
    "검사 시료"/"품질 시료"/"기술 시료" 등으로 구체적이면 "시료" 소분류의 해당 리프로 보내고, "조건
    설정"처럼 더 세분화 안 되는 값만 기존처럼 "조건설정/시료"로 남긴다.

    (v24.4, 사용자 엑셀 계산본 대조로 확정) ① 시료 이름("검사 시료"/"품질 시료")은 소구분이 취파/생파로
    찍혀 있어도 "시료" 블록으로(엑셀 EBGC 검사시료 0.73 = 취파 품질 시료 0.58 + 시료 0.25 합) ② 완성엔
    양식에 "생파" 블록이 없어 조용히 빠지던 생파 행([자재]_전조 부적합 등)을 생산중 순수불량 그룹으로
    보내 "기타"에 잡히게 함 ③ 조립 AGM 외 제품군의 "1세대 극판"은 양식에 칸이 없고 엑셀은 공정 생파에
    합산 ④ 완성 EBGC 전용 라벨 대응(_FINISHED_EBGC_*_MAP)."""
    if detail and detail.startswith("[") and detail.endswith("]"):
        detail_clean = detail[1:-1]
    else:
        detail_clean = detail or ""
    detail_clean = {"핀홀 부적합": "핀홀"}.get(detail_clean, detail_clean)  # 양식 라벨과 표기 통일

    sample_leaf = _SAMPLE_LEAF_MAP.get(detail_clean.replace(" ", ""))
    if sample_leaf:
        return ("시료", sample_leaf)

    if process == "완성":
        if subtype == "생파":
            # 완성 양식엔 생파 블록이 없고, 사용자 엑셀은 생파 행([자재]_전조 부적합 등)을 "기타"에 집계
            # (EBGC 기타 4월 0.48·5월 0.09 = MVP 생파 4월 0.59·5월 0.11) — 조용히 빠뜨리지 않고 기타로.
            return ("생산중 순수불량", "기타")
        # "[완성]_저전압"처럼 매핑표에 없어 태그가 붙은 채 온 원본 텍스트는 태그를 떼고 양식 리프와 대조
        # (ES완성_MSB/UXL 라인의 "[완성]_저전압" 2.7백만원이 "기타"로 빠지던 것).
        stripped = {"핀홀 부적합": "핀홀"}.get(_strip_tag_prefix(detail_clean), _strip_tag_prefix(detail_clean))
        if detail_clean.startswith("[") and stripped in {"저전압", "저전류", "핀홀", "외관 파손", "외관 부적합", "기타"}:
            detail_clean = stripped
        if family == "EBGC":
            if detail_clean in _FINISHED_EBGC_EXACT_MAP:
                detail_clean = _FINISHED_EBGC_EXACT_MAP[detail_clean]
            else:
                for prefix, label in _FINISHED_EBGC_PREFIX_MAP:
                    if detail_clean.startswith(prefix):
                        detail_clean = label
                        break

    if subtype == "생파":
        if process == "조립" and family != "AGM" and detail_clean == "1세대 극판":
            return ("생파", "공정 생파")
        return ("생파", detail_clean)
    if subtype == "조건설정":
        return ("조건설정", "시료")  # "조건 설정"처럼 더는 안 나뉘는 값
    return ("생산중 순수불량", detail_clean)


# {(공정,제품군): {소분류: {세부내역, ...}}} — _BLOCKS에서 뽑은 "공식 양식에 실제로 있는 칸" 목록.
_LEAF_INFO: dict = {}
for _process, _family, _blocks in _BLOCKS:
    _fam_info = _LEAF_INFO.setdefault((_process, _family), {})
    for _subtype, _details in _blocks:
        _fam_info.setdefault(_subtype, set()).update(_details)


def _snap_to_leaf(process: str, family: str, subtype: str, label: str):
    """계산된 (소분류,세부내역)이 실제 양식 칸에 없으면(예: 매핑표에 없는 부적합유형 원본 텍스트,
    또는 그 공정군엔 없는 소분류) "기타" 칸으로 모으거나, 그마저 없으면 None(표에서 빠짐 —
    get_unmapped_summary()로 확인)을 반환해 값이 조용히 새지 않도록 한다."""
    fam_info = _LEAF_INFO.get((process, family))
    if fam_info is None:
        return None
    labels = fam_info.get(subtype)
    if labels is None:
        return None
    if label in labels:
        return (subtype, label)
    if "기타" in labels:
        return (subtype, "기타")
    return None


def _template_key(row):
    process, family = row["공정"], row["제품군"]
    subtype, detail = row["소분류"], row["세부내역"]
    if process == "재공품":
        mapped = _map_recon(family, subtype, detail)
    else:
        mapped = _map_assembly_or_finished(subtype, detail, process, family)
    if mapped is None:
        return None
    snapped = _snap_to_leaf(process, family, mapped[0], mapped[1])
    if snapped is None:
        return None
    return (process, family, snapped[0], snapped[1])


# (v24.1) 목표 시트("25년 실적 및 목표")의 소분류 어휘 → 이 양식의 소분류 어휘. 조건설정은 제품군에 따라
# 갈 곳이 달라서(재공품 1연도/2연도(W/F)는 "초기시운전" 블록) map_targets_to_template()에서 따로 다룬다.
_TARGET_SUBTYPE = {"공정부적합": "생산중 순수불량", "생파": "생파", "조건설정": "조건설정"}
TARGET_COLS = ["월 목표", "누계 목표", "달성률(%)"]


def map_targets_to_template(leaf: pd.DataFrame):
    """목표 시트 리프 행(공정/제품군/소분류/세부내역/월목표)을 이 양식의 칸에 붙인다 — 사용자 지침
    "동일한 항목 빼고는 기타에 넣으면 될 것": 이름이 같은 칸이 있으면 그대로, 없으면 그 그룹의 "기타"로
    합산, 기타조차 없는 그룹(예: 양식에서 삭제된 완성 V-type)은 제외한다.
    반환: ({(공정,제품군,소분류,세부내역): 월목표 합}, 기타 합산/제외 내역 DataFrame)"""
    mapped: dict = {}
    notes = []
    for _, r in leaf.iterrows():
        process, family, sub, label = r["공정"], r["제품군"], r["소분류"], r["세부내역"]
        if sub == "조건설정" and process == "재공품":
            if family == "1연도" and label == "기타(보정)":
                key_sub, key_label = "기타(보정)", ""
            elif family in ("1연도", "2연도(W/F)"):
                key_sub, key_label = "초기시운전", label
            else:
                key_sub, key_label = "조건설정", label
        else:
            key_sub, key_label = _TARGET_SUBTYPE.get(sub, sub), label
        snapped = _snap_to_leaf(process, family, key_sub, key_label)
        monthly = float(r["월목표"])
        base = {"공정": process, "제품군": family, "목표 시트 항목": f"{sub} / {label}", "월 목표": monthly}
        if snapped is None:
            notes.append({**base, "처리": "제외(양식에 해당 칸 없음)"})
            continue
        key = (process, family, snapped[0], snapped[1])
        mapped[key] = mapped.get(key, 0.0) + monthly
        if snapped[1] != key_label or snapped[0] != key_sub:
            notes.append({**base, "처리": f"'{snapped[0]} / {snapped[1] or '-'}'에 합산"})
    return mapped, pd.DataFrame(notes)


def _with_target(row: dict, monthly_target, months_elapsed: int) -> None:
    """행에 목표 3열을 채운다. 목표가 0이면 달성률은 계산 불가(NaN)로 둔다 — 임의값 금지."""
    cum = monthly_target * months_elapsed
    row["월 목표"] = monthly_target
    row["누계 목표"] = cum
    row["달성률(%)"] = row["합계"] / cum * 100 if cum > 0 else float("nan")


def build_fixed_report(data: pd.DataFrame, target_map: dict = None, months_elapsed: int = 0) -> pd.DataFrame:
    """전체 데이터(필터 적용 전) 기준으로 고정 양식 종합표를 만든다. 반환되는 DataFrame의 행/열은
    화면 필터와 무관하게 항상 동일하고(총수율 표 양식.xlsx 기준), 값만 실제 계산된 Q-cost(백만원)로
    채워진다. 매핑 안 되는 값(재공품 생파 등)은 여기서 빠지며 get_unmapped_summary()로 확인 가능.
    (v24.1) target_map(map_targets_to_template() 결과)을 주면 "월 목표 / 누계 목표(=월 목표×
    months_elapsed) / 달성률(%)" 3열을 리프·소계·합계 행에 덧붙인다."""
    df = data.copy()
    df["_key"] = df.apply(_template_key, axis=1)
    mapped = df[df["_key"].notna()]

    agg = mapped.groupby(["_key", "월"])["Q_cost_백만원"].sum()
    use_target = target_map is not None

    out_rows = []
    grand = {}
    grand_target = {}
    for process, family, blocks in _BLOCKS:
        major = _MAJOR_LABEL[process]
        block_month_totals = {m: 0.0 for m in MONTHS}
        block_target = 0.0
        for subtype, details in blocks:
            for label in details:
                key = (process, family, subtype, label)
                row = {"공정": major, "제품군": family, "소분류": subtype, "세부내역": label}
                row_total = 0.0
                for m in MONTHS:
                    v = float(agg.get((key, m), 0.0))
                    row[f"{m}월"] = v
                    row_total += v
                    block_month_totals[m] += v
                row["합계"] = row_total
                if use_target:
                    t = float(target_map.get(key, 0.0))
                    block_target += t
                    _with_target(row, t, months_elapsed)
                out_rows.append(row)
        subtotal = {"공정": major, "제품군": family, "소분류": "소계", "세부내역": ""}
        subtotal_total = 0.0
        for m in MONTHS:
            subtotal[f"{m}월"] = block_month_totals[m]
            subtotal_total += block_month_totals[m]
        subtotal["합계"] = subtotal_total
        if use_target:
            _with_target(subtotal, block_target, months_elapsed)
        out_rows.append(subtotal)

        g = grand.setdefault(major, {m: 0.0 for m in MONTHS})
        for m in MONTHS:
            g[m] += block_month_totals[m]
        grand_target[major] = grand_target.get(major, 0.0) + block_target

    result_by_major: dict[str, list] = {}
    for r in out_rows:
        result_by_major.setdefault(r["공정"], []).append(r)

    final_rows = []
    for process in ["재공품", "조립", "완성"]:
        major = _MAJOR_LABEL[process]
        final_rows.extend(result_by_major.get(major, []))
        totals = grand.get(major, {m: 0.0 for m in MONTHS})
        grand_row = {"공정": f"{major} 합계", "제품군": "", "소분류": "", "세부내역": ""}
        for m in MONTHS:
            grand_row[f"{m}월"] = totals[m]
        grand_row["합계"] = sum(totals.values())
        if use_target:
            _with_target(grand_row, grand_target.get(major, 0.0), months_elapsed)
        final_rows.append(grand_row)

    final = pd.DataFrame(final_rows)
    cols = ["공정", "제품군", "소분류", "세부내역"] + [f"{m}월" for m in MONTHS] + ["합계"]
    if use_target:
        cols += TARGET_COLS
    return final[cols]


def _escape(s) -> str:
    return str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _compute_rowspans(df: pd.DataFrame, merge_cols: list) -> tuple:
    """merge_cols를 계층 순서(예: 공정→제품군→소분류)로 보고, 상위 컬럼들까지 전부 같은 값으로
    이어지는 연속 구간만 하위 컬럼도 병합한다(가족 경계를 넘어 우연히 같은 소분류명이 나와도
    안 섞이도록). 반환: {컬럼: [그 행에서 시작하는 rowspan, ...]}, {컬럼: [그 행을 그릴지 여부]}."""
    n = len(df)
    spans = {c: [1] * n for c in merge_cols}
    hidden = {c: [False] * n for c in merge_cols}
    for ci, col in enumerate(merge_cols):
        parents = merge_cols[:ci]
        i = 0
        while i < n:
            j = i + 1
            while j < n and df.iloc[j][col] == df.iloc[i][col] and all(
                df.iloc[j][p] == df.iloc[i][p] for p in parents
            ):
                j += 1
            spans[col][i] = j - i
            for k in range(i + 1, j):
                hidden[col][k] = True
            i = j
    return spans, hidden


def to_merged_html(df: pd.DataFrame) -> str:
    """`build_fixed_report()` 결과를 실제 엑셀 양식처럼 같은 공정/제품군/소분류가 이어지는 구간을
    세로로 셀 병합해서 HTML 표로 그린다(사용자 요청: "같은 공정끼리는 셀 병합해서 깔끔하게 보이게").
    `st.dataframe`은 셀 병합을 지원하지 않아 직접 HTML을 만든다. 폭이 너무 넓어 화면 우측이 잘리던
    문제(월 12개+합계 열)는 폰트·여백을 줄이고 가로 스크롤 컨테이너로 감싸 해결한다."""
    merge_cols = ["공정", "제품군", "소분류"]
    value_cols = [f"{m}월" for m in MONTHS] + ["합계"] + [c for c in TARGET_COLS if c in df.columns]
    spans, hidden = _compute_rowspans(df, merge_cols)

    def _num_cell(col, v):
        if pd.isna(v):
            return '<td class="frt-num"></td>'
        if col == "달성률(%)":
            cls = "frt-num frt-over" if v > 100 else "frt-num frt-under"
            return f'<td class="{cls}">{v:,.0f}%</td>'
        if col in TARGET_COLS:
            return f'<td class="frt-num frt-target">{f"{v:,.1f}" if v else ""}</td>'
        return f'<td class="frt-num">{f"{v:,.1f}" if v else ""}</td>'

    n = len(df)
    body_rows = []
    for i in range(n):
        row = df.iloc[i]
        is_grand = row["제품군"] == "" and row["소분류"] == "" and row["세부내역"] == ""
        is_subtotal = row["소분류"] == "소계"
        row_cls = ' class="frt-summary"' if (is_grand or is_subtotal) else ""

        cells = []
        if not hidden["공정"][i]:
            cells.append(f'<td class="frt-label" rowspan="{spans["공정"][i]}">{_escape(row["공정"])}</td>')
        if is_grand:
            cells.append('<td class="frt-label" colspan="3"></td>')
        else:
            if not hidden["제품군"][i]:
                cells.append(f'<td class="frt-label" rowspan="{spans["제품군"][i]}">{_escape(row["제품군"])}</td>')
            if not hidden["소분류"][i]:
                cells.append(f'<td class="frt-label" rowspan="{spans["소분류"][i]}">{_escape(row["소분류"])}</td>')
            cells.append(f'<td class="frt-label">{_escape(row["세부내역"])}</td>')
        for col in value_cols:
            cells.append(_num_cell(col, row[col]))
        body_rows.append(f"<tr{row_cls}>" + "".join(cells) + "</tr>")

    header_cols = ["공정", "제품군", "소분류", "세부내역"] + value_cols
    header_html = "".join(f"<th>{_escape(c)}</th>" for c in header_cols)
    return (
        '<div class="frt-wrap"><table class="frt-table">'
        f"<thead><tr>{header_html}</tr></thead>"
        f"<tbody>{''.join(body_rows)}</tbody>"
        "</table></div>"
    )


def get_unmapped_summary(data: pd.DataFrame) -> pd.DataFrame:
    """이 고정 양식에 자리가 없어 표에서 빠진 행(예: 재공품 생파)을 공정/제품군별 건수·금액으로 요약한다."""
    df = data.copy()
    df["_key"] = df.apply(_template_key, axis=1)
    unmapped = df[df["_key"].isna()]
    if unmapped.empty:
        return unmapped
    return (
        unmapped.groupby(["공정", "제품군", "소분류"], observed=True)
        .agg(건수=("Q_cost_백만원", "size"), Q_cost_백만원=("Q_cost_백만원", "sum"))
        .reset_index()
    )
