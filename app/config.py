"""공통 상수 / 설정값. 실제 값은 PRD 부속문서(.docs/04_prd_appendix.md) 3장 TBD와 연동된다."""
import os

# 회수율: 공정별 고정값 (확정, 04_prd_appendix.md 3장 / 총폐기비용 산출기준.xlsx와 일치 확인됨)
RECOVERY_RATE = {"재공품": 0.61, "조립": 0.61, "완성": 0.39}

# LME/환율 기본값: '26년 경영계획 기준(필요자료/총폐기비용 산출기준.xlsx) — 2.76원/g으로 검산 확인됨
DEFAULT_LME = 2000.0    # US$/톤
DEFAULT_FX = 1380.0     # 원/US$

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
# 매년/수시로 갱신되는 파일들 — 기본으로 번들해두고, 화면에서 새 파일로 교체 가능해야 함
# (v22.0) 사용자가 극판중량/제품중량/제조원가를 한 파일(시트 3개)로 정리해서 제공 — 예전엔 3개 파일로
# 나뉘어 있었는데(제조원가는 '25년 기준, 극판·제품중량은 별도 파일), 이제 이 통합 파일 하나로 셋 다
# 가리킨다. `cost.load_cost_lookup()`/`weight.load_plate_weight_lookup()`/`load_product_weight_lookup()`
# 전부 파일 안에서 이름에 해당 키워드("제조원가"/"극판중량"/"제품중량")가 들어간 시트를 자동으로 찾아
# 읽으므로 세 상수가 같은 파일을 가리켜도 문제없다.
# (v23.0) 사용자가 `필요자료/`를 정리하면서 안 쓰는 파일을 지우고, 실제로 쓰는 기준정보 파일들을
# "기본정보(극판중량, 제품중량, 제조원가 등)" 폴더 하나로 모아둠 — 그 폴더 안 경로로 갱신.
_BASE_INFO_DIR = os.path.join(BASE_DIR, "필요자료", "기본정보(극판중량, 제품중량, 제조원가 등)")
_BASE_INFO_XLSX = os.path.join(_BASE_INFO_DIR, "제품중량, 제조원가, 극판중량_260922.xlsx")
COST_XLSX = _BASE_INFO_XLSX
PLATE_WEIGHT_XLS = _BASE_INFO_XLSX   # 재공품코드→Wet/Dry중량(g), 시트 "극판중량"
PRODUCT_WEIGHT_XLSX = _BASE_INFO_XLSX  # 자재코드→총중량(kg), 시트 "제품중량"
# 조립(EBGC/고정형/지게차)·완성(AGM/EBGC/고정형) 세부내역 공통 매핑표 (담당자 개별관리 엑셀에서 추출)
DETAIL_TYPE_XLSX = os.path.join(_BASE_INFO_DIR, "부적합 유형 기준 정보.xlsx")
# (v24.0) Q-cost 목표(경영계획) — 총수율 표 양식 파일 안의 "25년 실적 및 목표" 시트(26년 목표 월별 컬럼,
# 단위 백만원). 전체 현황 KPI 카드의 "목표 대비"에 쓰인다. targets.load_targets() 참고.
TARGET_XLSX = os.path.join(_BASE_INFO_DIR, "총수율 표 양식.xlsx")
# 부적합중량 산출 기준(확정): 재공품=Wet중량, 조립=Dry중량, 완성=제품 총중량 — weight.compute_defect_weight_g()에 반영됨

# 사용자가 중량표를 새로 교체 업로드하거나(코드&중량만 추출) 누락값 팝업에서 직접 입력한 값을
# 코드&값만 저장해 둔다(용량 최소화) — 다음 실행 때도 재업로드 없이 자동으로 반영됨(overrides.py 참고)
OVERRIDE_DIR = os.path.join(BASE_DIR, "output", "data_store", "overrides")
COST_OVERRIDE_PARQUET = os.path.join(OVERRIDE_DIR, "cost_overrides.parquet")
PLATE_WEIGHT_OVERRIDE_PARQUET = os.path.join(OVERRIDE_DIR, "plate_weight_overrides.parquet")
PRODUCT_WEIGHT_OVERRIDE_PARQUET = os.path.join(OVERRIDE_DIR, "product_weight_overrides.parquet")

# Q-cost 상세표(피벗)·엑셀 다운로드의 공정/제품군 행 순서(총수율 표 양식.xlsx 기준 표 형식 고정) —
# 필터/정렬을 걸어도 다운로드본은 이 기본 순서를 유지해야 한다는 요구사항 반영.
# 목록에 없는 값(신규 제품군, "미분류" 등)은 뒤에 알파벳순으로 자동 추가된다(app._ordered_category 참고).
PROCESS_ORDER = ["재공품", "조립", "완성"]
FAMILY_ORDER = [
    "1연도", "2연도(W/F)", "2연도(건조,절단,화성)",  # 재공품
    "AGM", "EBGC", "고정형", "지게차",                # 조립·완성 공용(완성엔 지게차가 없어 자연히 빠짐)
]
