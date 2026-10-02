import os
from io import BytesIO
import datetime as dt
import pandas as pd
import streamlit as st

from config import (
    DEFAULT_LME, DEFAULT_FX, RECOVERY_RATE, COST_XLSX, PLATE_WEIGHT_XLS, PRODUCT_WEIGHT_XLSX,
    DETAIL_TYPE_XLSX, TARGET_XLSX,
)
from data import load_defect as _load_defect_raw, read_table as _read_table_raw
from classify import (
    assign_family_and_detail, load_plate_generation_lookup, load_line_family_rules,
    load_detail_type_rules, detect_process,
)
from calc import add_qcost_columns, lead_price_per_g
from cost import load_cost_lookup
from weight import load_plate_weight_lookup, load_product_weight_lookup
import charts
import store
import overrides
import fixed_report
import trends
import targets

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DUMMY_DIR = os.path.join(BASE_DIR, "Schema", "dummies")
_BASE_INFO_DIR = os.path.join(BASE_DIR, "필요자료", "기본정보(극판중량, 제품중량, 제조원가 등)")
PLATE_XLSX = os.path.join(_BASE_INFO_DIR, "AGM조립 극판기준정보.xlsx")
LINE_FAMILY_XLSX = os.path.join(_BASE_INFO_DIR, "생산량 부적합량 라인명.xlsx")

DEFAULT_FILES = {
    "재공품": "재공품 부적합_더미데이터.csv",
    "조립": "조립 부적합_더미데이터.csv",
    "완성": "완성 부적합_더미데이터.csv",
}

st.set_page_config(page_title="품질비용(Q-cost) 현황", layout="wide")

SEBANG_CSS = """
<style>
/* (v15.0) 사용자가 마음에 들어한 발표 자료 사진(다크 배경 + 오렌지/청록 포인트)을 참고해 다크 테마로
   전환 — design/design.md 2.5절 "다크 모드가 필요한 프로젝트는 배경/텍스트를 반전하되 포인트 컬러는
   500 유지, 어두운 배경에서는 400 단계(더 밝은 톤)를 사용해 대비를 확보"를 그대로 따른 것이라
   임의 배색이 아니라 SEBANG 디자인 시스템이 이미 정의해 둔 다크모드 규칙의 적용이다. WCAG 대비도
   전부 계산해서 확인함(본문 텍스트 9~17:1, 포인트 컬러 400단계 5~6.5:1 — 전부 4.5:1 이상 통과).
   실제 위젯(사이드바/버튼/셀렉트박스 등)의 다크 테마는 app/.streamlit/config.toml에서 설정한다. */
:root{
  --sebang-dark-gray:#333F48; --sebang-gray:#A2AAAD; --sebang-light-gray:#D0D0CE;
  --sebang-green-700:#006A76; --sebang-green-400:#33ACBA;
  --sebang-orange:#EB3300; --sebang-orange-700:#A42400; --sebang-orange-400:#EF5C33;
  --color-bg-base:#14191D; --color-bg-surface:#1C2328; --color-bg-surface-2:#242C32;
  --color-border:#2B363D;
  --color-text-primary:#FBFBFB; --color-text-secondary:#B5BBBD; --color-text-disabled:#5C656D;
  --color-accent-primary:var(--sebang-orange-400); --color-accent-primary-strong:var(--sebang-orange);
  --color-accent-secondary:var(--sebang-green-400);
}
html, body, [class*="css"]{
  font-family:"Pretendard","Noto Sans KR","Inter",system-ui,-apple-system,"Segoe UI",Roboto,sans-serif;
}
.stApp{ background:var(--color-bg-base); }
h1, h2, h3 { color:var(--color-text-primary) !important; }
.badge{
  display:inline-block; font-size:0.7rem; padding:2px 10px; border-radius:999px;
  background:var(--color-bg-surface-2); color:var(--color-accent-secondary); font-weight:700; margin-left:8px;
}
div.stButton > button{
  background:var(--color-accent-primary-strong); color:#fff; border:none; border-radius:8px; font-weight:600;
}
div.stButton > button:hover{ background:var(--sebang-orange-700); color:#fff; }
.kpi-card{
  background:var(--color-bg-surface); border:1px solid var(--color-border); border-radius:8px;
  padding:12px 16px;
}
.kpi-label{ font-size:0.7rem; color:var(--color-text-secondary); font-weight:600; }
.kpi-value{ font-size:1.5rem; font-weight:800; color:var(--color-text-primary); margin-top:4px;
  font-variant-numeric:tabular-nums; letter-spacing:-0.01em; }
/* (v24.0) 전체 현황 숫자카드 4장 + 기간 비교 카드 3장 */
.kpi-grid{ display:grid; grid-template-columns:repeat(4, minmax(0,1fr)); gap:12px; margin:4px 0 14px; }
.kpi-grid-3{ grid-template-columns:repeat(3, minmax(0,1fr)); }
@media (max-width:900px){ .kpi-grid, .kpi-grid-3{ grid-template-columns:repeat(2, minmax(0,1fr)); } }
.kpi-unit{ font-size:0.8rem; font-weight:600; color:var(--color-text-secondary); margin-left:4px; }
.kpi-delta{ font-size:0.82rem; margin-top:4px; font-variant-numeric:tabular-nums; font-weight:600; }
.kpi-sub{ font-size:0.72rem; color:var(--color-text-secondary); margin-top:6px; }
.kpi-good{ color:var(--color-accent-secondary); }  /* 절감·감소 = 청록 */
.kpi-bad{ color:var(--color-accent-primary); }     /* 초과·증가 = 주황 */
.kpi-muted{ color:var(--color-text-secondary); }
.foot-note{
  /* 실제 에러(st.error, 빨강)와 헷갈리지 않도록 중립 회색톤 사용 — 이 박스는 항상 뜨는
     안내문(분류 규칙·매칭률 설명)일 뿐, 문제가 있다는 뜻이 아니다 */
  margin-top:8px; padding:10px 14px; background:var(--color-bg-surface); border:1px solid var(--color-border);
  border-radius:8px; font-size:0.8rem; color:var(--color-text-secondary); line-height:1.6;
}
.data-mgmt-row{ font-size:0.8rem; padding:2px 0; }
.frt-wrap{
  overflow-x:auto; overflow-y:auto; max-height:520px;
  border:1px solid var(--color-border); border-radius:8px;
}
.frt-table{ border-collapse:collapse; font-size:0.72rem; width:100%; }
.frt-table th, .frt-table td{
  border:1px solid var(--color-border); padding:3px 6px; white-space:nowrap;
}
.frt-table thead th{
  background:var(--color-bg-surface-2); color:var(--color-text-primary); font-weight:600; text-align:center;
  position:sticky; top:0; z-index:1;
}
.frt-table td.frt-label{ color:var(--color-text-primary); text-align:center; background:var(--color-bg-surface); }
.frt-table td.frt-num{ color:var(--color-text-primary); text-align:right; font-variant-numeric:tabular-nums; }
.frt-table tr.frt-summary td{ background:var(--color-bg-surface-2); font-weight:700; }
/* (v24.1) 목표 3열 — 목표값은 살짝 흐리게, 달성률은 100% 초과 주황 / 이하 청록 */
.frt-table td.frt-target{ color:var(--color-text-secondary); }
.frt-table td.frt-over{ color:var(--color-accent-primary); font-weight:700; }
.frt-table td.frt-under{ color:var(--color-accent-secondary); font-weight:700; }
</style>
"""
st.markdown(SEBANG_CSS, unsafe_allow_html=True)

st.markdown(
    '<h2>품질지표 현황 — 품질비용(Q-cost) 종합표 <span class="badge">MVP · Streamlit</span></h2>',
    unsafe_allow_html=True,
)
st.caption("PRD 부속문서(.docs/04_prd_appendix.md) 기준 구현 · 최초 실행 시 Schema/dummies로 자동 시드, 이후 사이드바에서 파일 업로드/월별 삭제로 데이터 관리")


def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str, index: bool = True) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=index)
    return buf.getvalue()


def load_defect(process: str, file_or_path):
    """(파싱된 df, 월별저장 기준 날짜컬럼) 반환. 파일명이 아니라 데이터 안 날짜로 월을 구분한다.
    A타입/B타입 판별은 data.load_defect가 실제 컬럼을 보고 자동으로 한다."""
    return _load_defect_raw(file_or_path, process), "일자_dt"


if not store.has_ever_seeded():
    # 최초 실행 시 딱 한 번만 Schema/dummies로 저장소를 시드한다.
    # (is_empty() 기준으로 재시드하면 사용자가 전부 지운 뒤에도 더미데이터가 되살아나는 문제가 있어 마커 파일로 관리)
    # (v18.1) 어느 (공정,월)에 더미데이터를 채웠는지 mark_seed_months()로 기록 — 사용자가 그 달에
    # 실제 데이터를 처음 올리면 병합이 아니라 교체되도록 해서 더미데이터가 영구히 안 섞이게 한다.
    # 더미 파일 자체를 저장소(배포본 등)에 안 넣어둔 경우에도 앱이 에러 화면으로 죽지 않고
    # "업로드된 데이터가 없습니다" 안내로 자연스럽게 넘어가도록 예외를 흡수한다 — 단, 시드 성공
    # 여부와 무관하게 mark_seeded()는 호출해 재실행마다 반복 시도하지 않게 한다.
    try:
        for process, name in DEFAULT_FILES.items():
            df, date_col = load_defect(process, os.path.join(DUMMY_DIR, name))
            saved_months = store.save_by_month(df, process, date_col)
            store.mark_seed_months(process, saved_months)
    except Exception:
        pass
    store.mark_seeded()

# ── 사이드바: 업로드(월별 누적) + 데이터 관리 ────────────────────
st.sidebar.header("데이터 업로드 (파일을 아래 칸에 드래그)")
st.sidebar.caption(
    "부적합 파일을 **재공품/조립/완성 구분 없이** 한 번에 올리면 됩니다(생산량 불필요) — "
    "파일 내용(컬럼 구성·라인명)으로 공정을 자동 인식합니다. 여러 파일을 한 번에 올릴 수 있고, "
    "파일명에 월이 없어도 데이터 안의 날짜로 월을 자동 구분합니다. "
    "파일을 선택한 뒤 아래 **📊 분석 시작** 버튼을 눌러야 저장·반영됩니다."
)

uploaded_files = st.sidebar.file_uploader(
    "부적합 파일", type=["xlsx", "xls", "csv"], accept_multiple_files=True, key="up_defect_files",
)

if st.sidebar.button("📊 분석 시작", type="primary"):
    if not uploaded_files:
        st.sidebar.warning("먼저 위 칸에 업로드할 파일을 선택해주세요.")
    else:
        line_family_rules_for_detect = load_line_family_rules(LINE_FAMILY_XLSX)
        by_process = {"재공품": [], "조립": [], "완성": []}
        unknown_files = []
        load_errors = []
        for f in uploaded_files:
            try:
                raw = _read_table_raw(f)
                proc = detect_process(raw, line_family_rules_for_detect)
                if proc == "미상":
                    unknown_files.append(f.name)
                    continue
                f.seek(0)  # 판별을 위해 한 번 읽었으니 실제 파싱 전에 되감기
                parsed = _load_defect_raw(f, proc)
            except Exception as e:
                load_errors.append(f"{f.name}: {e}")
                continue
            by_process[proc].append(parsed)

        saved_msgs = []
        for process, parsed_list in by_process.items():
            if not parsed_list:
                continue
            combined = pd.concat(parsed_list, ignore_index=True)
            saved_months = store.save_by_month(combined, process, "일자_dt")
            saved_msgs.append(
                f"{process}(자동 인식): {len(parsed_list)}개 파일 → {', '.join(saved_months) or '해당 월 없음'} 저장됨"
            )
        # st.rerun()은 즉시 스크립트를 재시작해 지금 이 시점에 그려둔 메시지를 전부 지워버리므로
        # (예: unknown_files 에러가 뜨자마자 사라지는 버그가 있었음), 세션스테이트에 담아뒀다가
        # 재실행 이후 화면에서 읽어서 표시한다.
        st.session_state["_upload_result"] = {
            "unknown_files": unknown_files, "load_errors": load_errors, "saved_msgs": saved_msgs,
        }
        st.rerun()

_upload_result = st.session_state.pop("_upload_result", None)
if _upload_result:
    for msg in _upload_result["saved_msgs"]:
        st.sidebar.success(msg)
    if _upload_result["load_errors"]:
        st.sidebar.error("파일을 읽는 중 오류:\n" + "\n".join(_upload_result["load_errors"]))
    if _upload_result["unknown_files"]:
        st.sidebar.error(
            "⚠️ 공정을 자동으로 인식하지 못한 파일: " + ", ".join(_upload_result["unknown_files"])
            + " — 라인명이 매핑표(생산량 부적합량 라인명.xlsx)의 키워드와 전혀 안 걸립니다. 파일을 확인해주세요."
        )

st.sidebar.markdown("**데이터 관리**")
for process in ["재공품", "조립", "완성"]:
    months = store.list_months(process)
    mcol1, mcol2 = st.sidebar.columns([3, 1])
    mcol1.markdown(
        f'<div class="data-mgmt-row"><b>{process}</b>: {", ".join(months) if months else "데이터 없음"}</div>',
        unsafe_allow_html=True,
    )
    if months:
        with mcol2.popover("삭제"):
            del_month = st.selectbox("삭제할 월", months, key=f"delsel_{process}")
            if st.button("삭제 확인", key=f"delok_{process}"):
                store.delete_month(process, del_month)
                st.rerun()

st.sidebar.header("Q-cost 계산 파라미터")
lme = st.sidebar.number_input("LME (US$/톤)", value=DEFAULT_LME, step=10.0)
fx = st.sidebar.number_input("환율 (원/US$)", value=DEFAULT_FX, step=10.0)
st.sidebar.caption(
    f"회수율(고정): 재공품 {RECOVERY_RATE['재공품']:.0%} · 조립 {RECOVERY_RATE['조립']:.0%} · 완성 {RECOVERY_RATE['완성']:.0%}"
)
st.sidebar.caption(f"g당 납가격(자동계산) = {lead_price_per_g(lme, fx):.2f} 원/g")

cost_file = st.sidebar.file_uploader(
    "제조원가 단가표 교체 (선택, xlsx)", type=["xlsx", "xls"], key="cost_xlsx_upload",
    help="기본은 필요자료/기본정보(극판중량, 제품중량, 제조원가 등)/제품중량, 제조원가, 극판중량_260922.xlsx(제조원가 시트). "
         "매년 갱신되면 새 파일로 교체하세요.",
)
cost_source = cost_file if cost_file is not None else COST_XLSX
cost_source_label = getattr(cost_file, "name", None) or "제품중량, 제조원가, 극판중량_260922.xlsx (기본)"


@st.cache_data(show_spinner=False)
def get_cost_lookup(_source, source_sig):
    return load_cost_lookup(_source)


cost_overrides = overrides.load_cost_overrides()
cost_lookup = {**get_cost_lookup(cost_source, cost_source_label), **cost_overrides}
st.sidebar.caption(
    f"단가표: {cost_source_label} · {len(cost_lookup):,}개 코드 매칭 가능"
    + (f" (수동 입력으로 저장된 보정값 {len(cost_overrides):,}개 포함)" if cost_overrides else "")
)

plate_weight_file = st.sidebar.file_uploader(
    "재공품/조립 중량표 교체 (선택, xls/xlsx)", type=["xlsx", "xls"], key="plate_weight_upload",
    help="기본은 필요자료/기본정보(극판중량, 제품중량, 제조원가 등)/제품중량, 제조원가, 극판중량_260922.xlsx(극판중량 시트). "
         "재공품코드별 Wet/Dry중량(g) — 재공품은 Wet, 조립은 Dry 기준으로 사용됩니다. "
         "업로드하면 코드&중량만 저장되어 다음 실행부터는 다시 올리지 않아도 자동 반영됩니다.",
)
if plate_weight_file is not None:
    _sig = f"{plate_weight_file.name}:{plate_weight_file.size}"
    if st.session_state.get("_plate_weight_upload_sig") != _sig:
        _new_lookup = load_plate_weight_lookup(plate_weight_file)
        if _new_lookup:
            overrides.merge_plate_weight_overrides(_new_lookup)
            st.sidebar.success(f"재공품/조립 중량표 {len(_new_lookup):,}개 코드 저장 완료 — 다음 실행부터 자동 반영됩니다.")
        else:
            st.sidebar.error("중량표 형식을 인식하지 못했습니다(컬럼명 확인 필요).")
        st.session_state["_plate_weight_upload_sig"] = _sig

product_weight_file = st.sidebar.file_uploader(
    "완성 제품 중량표 교체 (선택, xlsx)", type=["xlsx", "xls"], key="product_weight_upload",
    help="기본은 필요자료/기본정보(극판중량, 제품중량, 제조원가 등)/제품중량, 제조원가, 극판중량_260922.xlsx(제품중량 시트). "
         "자재코드별 총중량(kg). "
         "업로드하면 코드&중량만 저장되어 다음 실행부터는 다시 올리지 않아도 자동 반영됩니다.",
)
if product_weight_file is not None:
    _sig = f"{product_weight_file.name}:{product_weight_file.size}"
    if st.session_state.get("_product_weight_upload_sig") != _sig:
        _new_lookup = load_product_weight_lookup(product_weight_file)
        if _new_lookup:
            overrides.merge_product_weight_overrides(_new_lookup)
            st.sidebar.success(f"완성 제품 중량표 {len(_new_lookup):,}개 코드 저장 완료 — 다음 실행부터 자동 반영됩니다.")
        else:
            st.sidebar.error("중량표 형식을 인식하지 못했습니다(컬럼명 확인 필요).")
        st.session_state["_product_weight_upload_sig"] = _sig


@st.cache_data(show_spinner=False)
def get_base_plate_weight_lookup():
    return load_plate_weight_lookup(PLATE_WEIGHT_XLS)


@st.cache_data(show_spinner=False)
def get_base_product_weight_lookup():
    return load_product_weight_lookup(PRODUCT_WEIGHT_XLSX)


plate_weight_overrides = overrides.load_plate_weight_overrides()
product_weight_overrides = overrides.load_product_weight_overrides()
plate_weight_lookup = overrides.merge_plate_lookup(get_base_plate_weight_lookup(), plate_weight_overrides)
product_weight_lookup = {**get_base_product_weight_lookup(), **product_weight_overrides}
st.sidebar.caption(
    f"중량표: 재공품/조립 {len(plate_weight_lookup):,}개"
    + (f"(보정 {len(plate_weight_overrides):,}개 포함)" if plate_weight_overrides else "")
    + f" / 완성 {len(product_weight_lookup):,}개"
    + (f"(보정 {len(product_weight_overrides):,}개 포함)" if product_weight_overrides else "")
)


@st.cache_data(show_spinner="데이터를 불러와 Q-cost를 계산하는 중입니다... 잠시만 기다려주세요")
def build_dataset(lme, fx, store_sig, cost_lookup, plate_weight_lookup, product_weight_lookup):
    """(수정) 매개변수 전부 언더스코어 없이 그대로 hashing 대상으로 둔다 — 전부 언더스코어(_)를
    붙였던 예전 버전은 dict 등 인자값이 바뀌어도 캐시가 절대 갱신되지 않는 버그였음(캐시 키가
    함수 자체로 고정돼 최초 1회 실행 결과가 계속 재사용됨). LME/환율/업로드/보정값 변경이
    실제로 화면에 반영되려면 이 인자들이 hashing되어야 한다(Streamlit이 dict 내용까지 해싱함).
    store_sig(=store.store_fingerprint())는 (v14.11) 세션마다 0부터 다시 세던 data_version
    카운터를 대체한 것 — st.cache_data 캐시는 세션이 아니라 서버 프로세스 전체 공유라, 세션별
    카운터를 쓰면 새 세션(F5 등)이 과거의 우연히 같은 카운터값 캐시를 잘못 재사용하는 버그가
    있었다(디스크 실제 파일 상태 기반 지문을 쓰면 이 문제가 근본적으로 안 생김)."""
    # 업로드된 데이터가 전혀 없으면(배포본에 분류 매핑표가 아직 없는 경우 포함) 분류 매핑표를
    # 로딩하기 전에 먼저 빠져나간다 — 어차피 분류할 데이터가 없고, 매핑표 파일이 없을 때
    # "데이터가 없습니다" 안내 대신 오류 배너가 뜨는 것을 막는다(매핑표 자체가 깨졌는데 데이터는
    # 있는 정상 케이스는 여전히 아래에서 바로 오류로 보고됨 — 조용히 삼키지 않음).
    frames = [store.load_all(p) for p in ["재공품", "완성", "조립"]]
    frames = [f for f in frames if not f.empty]
    if not frames:
        return pd.DataFrame()

    plate_gen_lookup = load_plate_generation_lookup(PLATE_XLSX)
    line_family_rules = load_line_family_rules(LINE_FAMILY_XLSX)
    detail_type_rules = load_detail_type_rules(DETAIL_TYPE_XLSX)

    all_defects = pd.concat(frames, ignore_index=True)
    all_defects = assign_family_and_detail(
        all_defects, plate_gen_lookup=plate_gen_lookup, line_family_rules=line_family_rules,
        detail_type_rules=detail_type_rules,
    )
    all_defects = add_qcost_columns(
        all_defects, lme, fx, cost_lookup=cost_lookup,
        plate_weight_lookup=plate_weight_lookup, product_weight_lookup=product_weight_lookup,
    )
    return all_defects


try:
    data = build_dataset(
        lme, fx, store.store_fingerprint(), cost_lookup, plate_weight_lookup, product_weight_lookup
    )
except Exception as e:
    st.error(f"데이터 로딩/계산 중 오류: {e}")
    st.stop()

if data.empty:
    st.info("업로드된 데이터가 없습니다. 사이드바에서 파일을 올려주세요.")
    st.stop()

# (v24.5) 집계 제외 규칙(classify.assign_family_and_detail의 "집계제외사유") — 사용자 확인: 2연도(W/F)의
# [주조] 태그 부적합은 재용해라 총수율 Q-cost에 안 들어감. 조용히 빼지 않고 건수·금액을 안내한다.
if "집계제외사유" in data.columns:
    _excluded = data[data["집계제외사유"].notna()]
    data = data[data["집계제외사유"].isna()]
    if len(_excluded):
        _ex_sum = _excluded.groupby("집계제외사유", observed=True)["Q_cost_백만원"].agg(["size", "sum"])
        st.caption(
            "ℹ️ 집계 제외: "
            + " · ".join(f"{reason} {int(r['size']):,}건 / {r['sum']:,.2f}백만원" for reason, r in _ex_sum.iterrows())
            + " (사용자 확정 규칙)"
        )

unmatched = data[~data["제품군_매칭됨"]]
if len(unmatched):
    line_col = unmatched.apply(
        lambda r: r.get("Line명") if isinstance(r.get("Line명"), str) and r.get("Line명") else r.get("라인명", ""),
        axis=1,
    )
    unmatched_names = sorted(set(zip(unmatched["공정"], line_col)))
    preview = ", ".join(f"[{p}] {n}" for p, n in unmatched_names[:20])
    st.warning(
        f"⚠️ 라인명 매핑표(생산량 부적합량 라인명.xlsx)에 속하지 않는 데이터 {len(unmatched)}건 발견 "
        f"({len(unmatched_names)}개 라인명). 매핑표에 키워드 추가가 필요합니다.\n\n{preview}"
        + (" ..." if len(unmatched_names) > 20 else "")
    )


def _missing_weight_context(df: pd.DataFrame) -> dict:
    """{매칭키: {공정, ...}} — 이 코드가 어느 공정에서 왔는지(재공품=Wet, 조립=Dry, 완성=총중량 입력란 결정용)."""
    sub = df.loc[~df["부적합중량_실측여부"], ["매칭키", "공정"]].dropna(subset=["매칭키"])
    ctx: dict = {}
    for key, proc in zip(sub["매칭키"], sub["공정"]):
        ctx.setdefault(key, set()).add(proc)
    return ctx


@st.dialog("단가/중량표 누락 항목", width="large")
def show_missing_dialog(missing_cost_codes, missing_weight_codes, weight_context):
    st.write(
        "정확한 값을 알 수 없는 항목은 임의값(0 포함)으로 대체하지 않고, 해당 행의 Q-cost 계산을 하지 않습니다. "
        "아래에서 값을 직접 입력하고 저장하면 코드&값만 저장되어 **다음 실행 때도 남아있습니다**. "
        "테스트 등으로 지금 당장 안 채워도 된다면 맨 아래 **🙈 무시하고 진행**을 누르세요 — "
        "그 코드들은 Q-cost가 계산되지 않을 뿐 화면 진행에는 지장이 없고, 다음부터 다시 묻지 않습니다."
    )

    cost_inputs = {}
    if missing_cost_codes:
        st.markdown(f"**제조원가**를 찾지 못한 코드 {len(missing_cost_codes)}개 — 알고 있으면 직접 입력(원/개)")
        for code in missing_cost_codes:
            cost_inputs[code] = st.number_input(
                code, key=f"manual_cost_{code}", min_value=0.0, value=None, placeholder="표준원가(원)",
            )
        st.caption("모르는 코드는 비워두세요. 단가표에 코드를 추가하거나 사이드바에서 최신 단가표로 교체해도 됩니다.")

    weight_inputs = {}
    if missing_weight_codes:
        st.markdown(f"**부적합중량**을 찾지 못한 코드 {len(missing_weight_codes)}개 — 알고 있으면 직접 입력(g)")
        for code in missing_weight_codes:
            procs = weight_context.get(code, set())
            field_specs = []
            if "재공품" in procs:
                field_specs.append(("wet", "Wet중량(g)"))
            if "조립" in procs:
                field_specs.append(("dry", "Dry중량(g)"))
            if "완성" in procs:
                field_specs.append(("product", "총중량(g)"))
            if not field_specs:  # 방어적 fallback(이론상 발생 안 함)
                field_specs = [("wet", "중량(g)")]
            st.caption(f"{code}  ({'/'.join(sorted(procs)) or '공정 미상'})")
            entry = {}
            for col, (field, label) in zip(st.columns(len(field_specs)), field_specs):
                entry[field] = col.number_input(
                    label, key=f"manual_weight_{code}_{field}", min_value=0.0, value=None, placeholder=label,
                )
            weight_inputs[code] = entry
        st.caption("모르는 코드는 비워두세요. 중량표에 코드를 추가하거나 사이드바에서 최신 중량표로 교체해도 됩니다.")

    col_a, col_b, col_c = st.columns(3)
    if col_a.button("💾 저장하고 다시 계산", key="save_manual_missing", type="primary"):
        saved_cost = {c: v for c, v in cost_inputs.items() if v is not None}
        saved_plate = {
            c: e for c, e in weight_inputs.items()
            if e.get("wet") is not None or e.get("dry") is not None
        }
        saved_product = {c: e["product"] for c, e in weight_inputs.items() if e.get("product") is not None}
        if saved_cost:
            overrides.merge_cost_overrides(saved_cost)
        if saved_plate:
            overrides.merge_plate_weight_overrides(saved_plate)
        if saved_product:
            overrides.merge_product_weight_overrides(saved_product)
        saved_n = len(saved_cost) + len(saved_plate) + len(saved_product)
        if saved_n:
            st.session_state["_manual_save_msg"] = f"직접 입력한 값 {saved_n}건을 저장했습니다."
        st.rerun()
    if col_b.button("🙈 무시하고 진행", key="ignore_missing"):
        st.session_state["ignored_cost_codes"] |= set(missing_cost_codes)
        st.session_state["ignored_weight_codes"] |= set(missing_weight_codes)
        st.session_state["_manual_save_msg"] = (
            f"제조원가 {len(missing_cost_codes)}개·중량 {len(missing_weight_codes)}개 코드를 무시했습니다 — "
            "해당 코드는 Q-cost가 계산되지 않고, 이후 다시 묻지 않습니다."
        )
        st.rerun()
    if col_c.button("나중에 입력", key="ack_missing"):
        st.rerun()


if "ignored_cost_codes" not in st.session_state:
    st.session_state["ignored_cost_codes"] = set()
if "ignored_weight_codes" not in st.session_state:
    st.session_state["ignored_weight_codes"] = set()

missing_cost_all = sorted(data.loc[data["제조원가_단가"].isna(), "매칭키"].dropna().unique().tolist())
missing_weight_all = sorted(
    data.loc[~data["부적합중량_실측여부"], "매칭키"].dropna().unique().tolist()
)
missing_cost = [c for c in missing_cost_all if c not in st.session_state["ignored_cost_codes"]]
missing_weight = [c for c in missing_weight_all if c not in st.session_state["ignored_weight_codes"]]
_manual_save_msg = st.session_state.pop("_manual_save_msg", None)
if _manual_save_msg:
    st.success(f"✅ {_manual_save_msg}")
if missing_cost or missing_weight:
    weight_context = _missing_weight_context(data)
    sig = f"COST:{','.join(missing_cost)}|WEIGHT:{','.join(missing_weight)}"
    if st.session_state.get("_last_missing_sig") != sig:
        st.session_state["_last_missing_sig"] = sig
        show_missing_dialog(missing_cost, missing_weight, weight_context)
    if missing_cost:
        st.error(
            f"⚠️ 제조원가 단가표에 없는 코드 **{len(missing_cost)}개** — 해당 행은 Q-cost 계산에서 제외됩니다: "
            + ", ".join(missing_cost[:15]) + (" ..." if len(missing_cost) > 15 else "")
        )
    if missing_weight:
        st.error(
            f"⚠️ 부적합중량표에 없는 코드 **{len(missing_weight)}개** — 해당 행은 Q-cost 계산에서 제외됩니다: "
            + ", ".join(missing_weight[:15]) + (" ..." if len(missing_weight) > 15 else "")
        )
    bcol1, bcol2 = st.columns([1, 3])
    if bcol1.button("✏️ 누락값 직접 입력", key="reopen_missing_dialog"):
        show_missing_dialog(missing_cost, missing_weight, weight_context)
    if bcol1.button("🙈 무시하고 진행", key="ignore_missing_inline"):
        st.session_state["ignored_cost_codes"] |= set(missing_cost)
        st.session_state["ignored_weight_codes"] |= set(missing_weight)
        st.rerun()

ignored_n = len(st.session_state["ignored_cost_codes"]) + len(st.session_state["ignored_weight_codes"])
if ignored_n:
    ic1, ic2 = st.columns([4, 1])
    ic1.caption(f"🙈 무시한 코드 {ignored_n}개는 화면에 다시 표시하지 않습니다(Q-cost는 계속 미산출).")
    if ic2.button("다시 표시", key="reset_ignored"):
        st.session_state["ignored_cost_codes"] = set()
        st.session_state["ignored_weight_codes"] = set()
        st.session_state["_last_missing_sig"] = None
        st.rerun()

# (여러 섹션에서 공통으로 쓰는 날짜 범위) — 데이터 전체의 최소/최대 날짜
valid_dates = data["일자_dt"].dropna()
min_date = valid_dates.min().date() if len(valid_dates) else dt.date.today()
max_date = valid_dates.max().date() if len(valid_dates) else dt.date.today()


def _date_range_inputs(key_prefix: str, default_start=None, default_end=None):
    """시작일/종료일을 각각 별도 칸으로 받는다(사용자 요청 — 기간이 길면 달력 하나로 양 끝을
    한 번에 찍기 번거로우니 시작/종료를 따로 고를 수 있게). 시작일이 종료일보다 늦으면 경고 후
    종료일을 시작일로 맞춰 되돌린다. (v24.0) 기간 비교처럼 기본값을 전체 기간이 아닌 특정 달로
    두고 싶을 때 default_start/default_end를 넘긴다."""
    dcol1, dcol2 = st.columns(2)
    s = dcol1.date_input(
        "시작일", value=default_start or min_date, min_value=min_date, max_value=max_date,
        key=f"{key_prefix}_start",
    )
    e = dcol2.date_input(
        "종료일", value=default_end or max_date, min_value=min_date, max_value=max_date,
        key=f"{key_prefix}_end",
    )
    if s > e:
        st.warning("시작일이 종료일보다 늦어 종료일을 시작일로 맞췄습니다.")
        e = s
    return s, e


def _in_range(df: pd.DataFrame, s: dt.date, e: dt.date) -> pd.DataFrame:
    return df[df["일자_dt"].notna() & (df["일자_dt"].dt.date >= s) & (df["일자_dt"].dt.date <= e)]


# ── 공통 계산(전체 현황 KPI·세부분석 기본값에서 같이 씀) ─────────────
_monthly = trends.monthly_totals(data)
_mom = trends.compute_mom(_monthly)
_ytd = trends.ytd_total(_monthly)
# (v24.2) 무거운 계산(총수율 종합표 등)의 캐시 키 — data 전체를 해싱하지 않고 행수+Q-cost 열 해시로 지문을 만든다.
_data_sig = (len(data), int(pd.util.hash_pandas_object(data["Q_cost_백만원"], index=False).sum()))


@st.cache_data(show_spinner=False)
def get_targets(mtime):
    """mtime(파일 수정시각)을 캐시 키로 써서 목표 시트를 고치면 다음 실행에 자동 반영되게 한다."""
    return targets.load_targets(TARGET_XLSX)


_targets = get_targets(os.path.getmtime(TARGET_XLSX) if os.path.exists(TARGET_XLSX) else 0)


@st.cache_data(show_spinner="총수율 종합표를 계산하는 중입니다...")
def get_fixed_report(data_sig, _data, target_items, months_elapsed):
    """(v24.2) 총수율 종합표는 3만 행을 행 단위로 매핑하느라 매 실행마다 2~3초씩 걸려 화면 전체를
    느리게 만들던 주범 — 데이터 지문(data_sig)·목표·개월 수가 같으면 다시 계산하지 않는다.
    _data는 언더스코어라 해싱하지 않고(크기 때문) data_sig가 대신 캐시 키 역할을 한다."""
    target_map = dict(target_items) if target_items is not None else None
    fixed = fixed_report.build_fixed_report(_data, target_map=target_map, months_elapsed=months_elapsed)
    return fixed, fixed_report.get_unmapped_summary(_data)


def _kpi_card(label: str, value: str, unit: str = "", delta_html: str = "", sub: str = "") -> str:
    return (
        f'<div class="kpi-card"><div class="kpi-label">{label}</div>'
        f'<div class="kpi-value">{value}<span class="kpi-unit">{unit}</span></div>'
        f'{delta_html}<div class="kpi-sub">{sub}</div></div>'
    )


def _delta_html(delta: float, pct, suffix: str = "", up_word: str = "", down_word: str = "") -> str:
    """증가(=비용 악화)는 주황, 감소(=절감)는 청록. pct가 None이면 금액만 표시."""
    if delta > 0:
        cls, arrow, word = "kpi-bad", "▲", up_word
    elif delta < 0:
        cls, arrow, word = "kpi-good", "▼", down_word
    else:
        cls, arrow, word = "kpi-muted", "―", "변화 없음"
    pct_txt = f" ({pct:+.1f}%)" if pct is not None else ""
    return f'<div class="kpi-delta {cls}">{arrow} {abs(delta):,.1f}백만원{pct_txt} {word}{suffix}</div>'


def _style_delta(df: pd.DataFrame, delta_cols: list):
    """(v24.2) 증감 열을 눈에 띄게 — 증가(악화)는 붉은 포인트색, 감소는 청록, 굵게(사용자 요청)."""
    def _color(v):
        if pd.isna(v):
            return ""
        if v > 0:
            return "color:#EF5C33; font-weight:700"
        if v < 0:
            return "color:#33ACBA; font-weight:700"
        return ""
    fmt = {c: "{:,.2f}" for c in df.select_dtypes("number").columns}
    for c in delta_cols:
        fmt[c] = "{:+,.1f}%" if "%" in c else "{:+,.2f}"
    styler = df.style.format(fmt, na_rep="—")
    return styler.map(_color, subset=delta_cols) if hasattr(styler, "map") else styler.applymap(_color, subset=delta_cols)


def _overview_kpi_html() -> str:
    """(v24.0) 사용자 요청 "전체 현황 숫자카드 4장": ① 당월 Q-cost(+전월 대비) ② 당월 목표 대비
    ③ 연간 누계(YTD) ④ 누계 목표 대비. 목표는 총수율 표 양식.xlsx의 "25년 실적 및 목표" 시트
    (26년 목표 월별, 백만원)이며 없으면 임의값 대신 "목표 미등록"으로 비워둔다."""
    cards = []
    if len(_monthly) == 0:
        return ""
    cur_period, cur = str(_monthly.index[-1]), float(_monthly.iloc[-1])

    if _mom:
        if _mom["pct"] is None:
            d_html = f'<div class="kpi-delta kpi-muted">전월(0원) 대비 {_mom["status"]}</div>'
        else:
            d_html = _delta_html(_mom["delta"], _mom["pct"], up_word="증가", down_word="감소", suffix=" · 전월 대비")
        sub1 = f"전월({_mom['prev_period']}) {_mom['prev']:,.1f}백만원"
    else:
        d_html, sub1 = '<div class="kpi-delta kpi-muted">전월 데이터 없음</div>', "비교할 전월 데이터가 없습니다"
    cards.append(_kpi_card(f"당월 Q-cost ({cur_period})", f"{cur:,.1f}", "백만원", d_html, sub1))

    month_target = _targets.get("monthly_total") if _targets else None
    tgt_year = _targets.get("year") if _targets else None
    if month_target:
        diff = cur - month_target
        cards.append(_kpi_card(
            "당월 목표 대비", f"{cur / month_target * 100:,.1f}", "%",
            _delta_html(diff, None, up_word="초과", down_word="절감"),
            f"월 목표 {month_target:,.1f}백만원" + (f" ({tgt_year}년 경영계획)" if tgt_year else ""),
        ))
    else:
        cards.append(_kpi_card(
            "당월 목표 대비", "—", "", '<div class="kpi-delta kpi-muted">목표 미등록</div>',
            "총수율 표 양식.xlsx의 '실적 및 목표' 시트에서 26년 목표를 읽지 못했습니다",
        ))

    avg = _ytd["total"] / _ytd["months"] if _ytd["months"] else 0.0
    annual_target = _targets.get("annual_total") if _targets else None
    cards.append(_kpi_card(
        f"연간 누계 (YTD, {_ytd['year']}년 {_ytd['months']}개월)", f"{_ytd['total']:,.1f}", "백만원",
        f'<div class="kpi-delta kpi-muted">월평균 {avg:,.1f}백만원</div>',
        f"연간 목표 {annual_target:,.1f}백만원" if annual_target else "연간 목표 미등록",
    ))

    if month_target and _ytd["months"]:
        ytd_target = month_target * _ytd["months"]
        diff = _ytd["total"] - ytd_target
        cards.append(_kpi_card(
            "누계 목표 대비", f"{_ytd['total'] / ytd_target * 100:,.1f}", "%",
            _delta_html(diff, None, up_word="초과", down_word="절감"),
            f"누계 목표({_ytd['months']}개월) {ytd_target:,.1f}백만원",
        ))
    else:
        cards.append(_kpi_card("누계 목표 대비", "—", "", '<div class="kpi-delta kpi-muted">목표 미등록</div>', ""))
    return '<div class="kpi-grid">' + "".join(cards) + "</div>"


# ── ① 전체 현황(KPI 카드 4장 + 월별 추이) ────────────────────────
# 사용자 요청(2026-09-16): "선택 조건 요약" 자리에 전반적인 흐름을 볼 수 있는 꺾은선 그래프를 넣자 —
# 화면 필터와 무관하게 항상 전체 데이터 기준 월별 추이(총수율 종합표와 같은 취지로 "전체" 개요용).
# (v24.0) 사용자 요청 "전체 현황 숫자카드 4장(전월대비/목표대비 당월/누계/누계 목표대비)"을 맨 위에 추가.
with st.expander(
    "📈 전체 현황 — 목적: 얼마나 썼고 목표 대비 어디쯤인지, 월별 흐름과 전월 대비 증감을 한눈에 확인",
    expanded=True,
):
    _kpi_html = _overview_kpi_html()
    if _kpi_html:
        st.markdown(_kpi_html, unsafe_allow_html=True)
        if _targets and _targets.get("year") and _ytd["year"] and _targets["year"] != _ytd["year"]:
            st.caption(f"⚠️ 목표는 {_targets['year']}년 경영계획 기준인데 실적은 {_ytd['year']}년 데이터입니다 — 목표 시트를 확인하세요.")
    # (v21.0) 여러 달 데이터가 쌓이면서 3개월 이동평균을 보조선으로 함께 표시(사용자 시계열 비교 제안).
    st.plotly_chart(charts.render_monthly_trend(data, show_ma=True), key="trend_overall")

    # (v20.0) 사용자 요청 "문장형 자동 요약" — 전체 총액/최대 비중 공정·세부내역은 데이터가 1개월이든
    # 여러 달이든 항상 계산 가능. (v21.0) 이제 2개월 이상 쌓이면 전월 대비(MoM)와 증가 기여 1순위까지
    # 같은 문장에 자동으로 이어붙인다(사용자 예시: "...전월 대비 +128만 원(+9.7%) 증가. 증가분 중
    # 조립–AGM–생파–접합불량이 70만 원(전체 증가의 54%)을 차지함." 형식).
    _total_qcost_all = data["Q_cost_백만원"].sum()
    if _total_qcost_all > 0:
        _proc_sum = data.groupby("공정")["Q_cost_백만원"].sum().sort_values(ascending=False)
        _top_proc, _top_proc_amt = _proc_sum.index[0], _proc_sum.iloc[0]
        _detail_sum = data.groupby("세부내역")["Q_cost_백만원"].sum()
        _detail_sum = _detail_sum[_detail_sum > 0].sort_values(ascending=False)
        _months = sorted(data["일자_dt"].dropna().dt.to_period("M").astype(str).unique())
        _month_label = f"{_months[0]}~{_months[-1]}" if len(_months) > 1 else (_months[0] if _months else "")
        _summary = (
            f"{_month_label} 총 Q-cost는 <b>{_total_qcost_all:,.1f}백만원</b>이며, "
            f"<b>{_top_proc}</b>이 전체의 <b>{_top_proc_amt / _total_qcost_all:.0%}</b>로 가장 큰 비중을 차지합니다."
        )
        if len(_detail_sum):
            _top_detail, _top_detail_amt = _detail_sum.index[0], _detail_sum.iloc[0]
            _summary += (
                f" 세부적으로는 <b>{_top_detail}</b>이 {_top_detail_amt:,.1f}백만원"
                f"(전체의 {_top_detail_amt / _total_qcost_all:.0%})으로 가장 큽니다."
            )

        if _mom:
            if _mom["pct"] is None:
                _summary += f" <b>{_mom['cur_period']}</b>은 전월(0원)과 비교해 <b>{_mom['status']}</b>했습니다."
            else:
                _arrow = "증가" if _mom["delta"] > 0 else ("감소" if _mom["delta"] < 0 else "변화 없음")
                _summary += (
                    f" <b>{_mom['cur_period']}</b>은 전월 대비 <b>{_mom['delta']:+,.1f}백만원"
                    f"({_mom['pct']:+.1f}%) {_arrow}</b>했습니다."
                )
                if _mom["delta"] > 0:
                    _contrib = trends.contribution_top(data, ["공정", "제품군", "소분류"], top_n=1)
                    if len(_contrib):
                        _row = _contrib.iloc[0]
                        _label = f"{_row['공정']}-{_row['제품군']}-{_row['소분류']}"
                        _contrib_pct = _row["증감액"] / _mom["delta"] * 100 if _mom["delta"] else 0
                        _summary += (
                            f" 증가분 중 <b>{_label}</b>이 {_row['증감액']:,.1f}백만원"
                            f"(전체 증가의 {_contrib_pct:.0f}%)을 차지합니다."
                        )
        st.markdown(f'<div class="foot-note">{_summary}</div>', unsafe_allow_html=True)

        _anomaly = trends.detect_anomaly(_monthly)
        if _anomaly.get("available") and _anomaly.get("is_anomaly"):
            _direction = "급증" if _anomaly["cur"] > _anomaly["mean"] else "급감"
            st.warning(
                f"⚠️ 이상탐지: {_mom['cur_period'] if _mom else ''} Q-cost({_anomaly['cur']:,.1f}백만원)가 "
                f"이전 달들 평균({_anomaly['mean']:,.1f}백만원) 대비 **{_direction}**했습니다 "
                f"(Z-score {_anomaly['z']:+.2f}, 기준 ±2.0)."
            )

        if _mom:
            _c1, _c2 = st.columns(2)
            with _c1:
                st.markdown(f"**전월 대비 증감 기여 Top5** ({_mom['prev_period']} → {_mom['cur_period']})")
                _contrib5 = trends.contribution_top(data, ["공정", "제품군", "소분류"], top_n=5)
                if len(_contrib5):
                    # (v24.2) 사용자 요청: 증감률(%) 추가 + 증감 열 붉은색 강조. 전월이 0이면 증감률은 "—"(신규 발생).
                    _contrib5["증감률(%)"] = _contrib5["증감액"] / _contrib5["전월"].where(_contrib5["전월"] > 0) * 100
                    _show = _contrib5.rename(columns={"당월": _mom["cur_period"], "전월": _mom["prev_period"]})
                    st.dataframe(_style_delta(_show, ["증감액", "증감률(%)"]), hide_index=True, width="stretch")
                    st.caption("전월이 0이었던 항목(신규 발생)은 증감률을 표시하지 않습니다.")
            with _c2:
                # (v24.2) 예전 "재발 빈도"(최근 3개월 중 발생 개월 수) 표는 여러 달 데이터가 쌓이자 전부 3/3으로만
                # 나와 의미가 없었음(사용자: "어떤 걸 보여주려는지 모르겠음") → "3개월 연속 증가 경보"로 교체.
                _win = min(3, len(_monthly))
                st.markdown(f"**연속 증가 경보** (최근 {_win}개월 매달 늘어난 항목, 공정·제품군·세부내역 기준 상위 10개)")
                _rising = trends.consecutive_increase(data, ["공정", "제품군", "세부내역"], window=_win, top_n=10)
                if len(_rising):
                    _inc_col = [c for c in _rising.columns if c.startswith("증가폭")]
                    st.dataframe(_style_delta(_rising, _inc_col), hide_index=True, width="stretch")
                    st.caption("월별 값이 매달 전월보다 커진 항목만 골랐습니다 — 한 번 튀는 것이 아니라 계속 나빠지는 중인 항목입니다.")
                else:
                    st.caption(f"최근 {_win}개월 연속으로 증가한 항목이 없습니다.")

# ── ② Q-cost 분석(파이 / 소분류·제품군 막대 / 파레토) ─────────────
# (v24.2) st.fragment로 분리 — 막대 클릭·기간 변경 때 이 섹션만 다시 그린다(예전엔 페이지 전체가 다시
# 돌아 2~3초씩 걸렸고 그 사이 클릭이 씹혀 "2~3번 눌러야 바뀐다"는 불편의 원인이었음).
@st.fragment
def _section_qcost_analysis():
    st.caption("👆 공정 탭을 누르면 그 공정의 제품군별 비중이, 막대를 클릭하면 우측에 해당 제품군의 세부내역 파레토가 나타납니다.")
    qa_start, qa_end = _date_range_inputs("qa")
    filtered = _in_range(data, qa_start, qa_end)

    if "chart_proc_tab" not in st.session_state:
        st.session_state["chart_proc_tab"] = "전체"

    col_pie, col_bar, col_pareto = st.columns([1, 1, 1.6])

    with col_pie:
        st.markdown("**공정별 비중**")
        st.plotly_chart(charts.render_pie(filtered), key="pie_chart")

    with col_bar:
        proc_tab = st.radio(
            "공정 탭", ["전체", "재공품", "조립", "완성"], horizontal=True,
            key="chart_proc_tab", label_visibility="collapsed",
        )
        is_process_mode = proc_tab != "전체"
        bar_scope = filtered[filtered["공정"] == proc_tab] if is_process_mode else filtered
        category_col = "제품군" if is_process_mode else "소분류"
        st.markdown(f"**{proc_tab} · {'제품군/라인별 비중' if is_process_mode else '소분류별 비중'}**")

        if st.session_state.get("_last_proc_tab") != proc_tab:
            st.session_state["chart_selected_family"] = None
            st.session_state["_last_proc_tab"] = proc_tab

        selected_family = st.session_state.get("chart_selected_family")
        bar_fig, bar_grp = charts.render_bar(bar_scope, category_col, selected=selected_family)
        bar_event = st.plotly_chart(
            bar_fig, on_select="rerun",
            selection_mode="points", key=f"bar_chart_{proc_tab}",
        )

        # (v14.13) 예전엔 "전체" 탭(첫 화면 기본값)에서는 막대 클릭을 아예 무시하도록 돼 있었다
        # (공정 탭을 먼저 선택해야만 드릴다운이 동작) — 사용자가 처음 화면에서 바로 소분류 막대를
        # 클릭했을 때 파레토가 안 바뀌는 것처럼 보이던 원인. 탭 종류와 무관하게 항상 클릭을 반영한다.
        if bar_event and bar_event.selection and bar_event.selection.get("points"):
            clicked = bar_event.selection["points"][0].get("y")
            if clicked:
                st.session_state["chart_selected_family"] = clicked
                selected_family = clicked
        if not selected_family and len(bar_grp):
            selected_family = bar_grp.index[-1]  # 가장 큰 막대(오름차순 정렬이라 마지막)
            st.session_state["chart_selected_family"] = selected_family

    with col_pareto:
        if selected_family:
            pareto_scope = bar_scope[bar_scope[category_col] == selected_family]
            label = f"{proc_tab} · {selected_family}" if is_process_mode else selected_family
            st.markdown(f"**세부내역 파레토도 — {label}**")
        else:
            pareto_scope = filtered
            st.markdown("**세부내역 파레토도**")
        st.plotly_chart(charts.render_pareto(pareto_scope), key="pareto_chart")
        # (v20.0) 사용자 요청 "비용 집중도(파레토): 상위 20% 현상이 전체의 몇 % 차지하는지"
        _stats = charts.pareto_concentration_stats(pareto_scope)
        if _stats["n"] > 0:
            st.caption(
                f"상위 3개 항목이 전체의 **{_stats['top3_pct']:.0f}%**({_stats['top3_amount']:,.1f}백만원), "
                f"누적 80%에 도달하려면 **{_stats['n_to_80']}개** 항목이 필요합니다 "
                f"(전체 {_stats['n']}개 중 상위 {_stats['n_20pct']}개(20%)가 전체의 {_stats['top20pct_pct']:.0f}%)."
            )


with st.expander(
    "🥧 Q-cost 분석 — 목적: 공정·제품군·세부내역별 비중과 비용 집중도(파레토)를 드릴다운으로 확인",
    expanded=True,
):
    _section_qcost_analysis()

# ── ③ 세부분석 — 기간 비교 ──────────────────────────────────
# (v24.0) 사용자 요청: "9월 1~5일 데이터와 6~10일 데이터를 비교해보고 싶다 할 때 AGM조립공정의
# 공정부적합 - 스태커 - [스태커]_작업중 극판 파손 정도로 세분화 분석을 할 수 있는 칸" — 기간 A/B를
# 각각 고르고 공정→제품군→소분류→세부내역→상세(원본 텍스트) 5단계로 좁혀 두 기간을 비교한다.
# (v24.2) 사용자 요청: 날짜를 바꿀 때마다 바로 도는 대신 "분석 실행" 버튼을 눌렀을 때만 계산하고,
# 계산 중엔 스피너를 띄운다. 섹션 자체도 st.fragment라 조건을 바꿔도 페이지 전체가 다시 돌지 않는다.
_CMP_LEVELS = [
    ("공정", "공정"), ("제품군", "제품군"), ("소분류", "소분류"),
    ("세부내역", "세부내역"), ("상세내역_원본", "상세(원본)"),
]


def _month_bounds(period) -> tuple:
    s, e = period.start_time.date(), period.end_time.date()
    return max(s, min_date), min(e, max_date)


def _run_period_compare(pool: pd.DataFrame, a: tuple, b: tuple, breakdown_col: str, breakdown_label: str) -> dict:
    scope_a, scope_b = _in_range(pool, *a), _in_range(pool, *b)
    sum_a, sum_b = float(scope_a["Q_cost_백만원"].sum()), float(scope_b["Q_cost_백만원"].sum())
    qty_a = float(scope_a["불량수량"].sum()) if "불량수량" in scope_a.columns else 0.0
    qty_b = float(scope_b["불량수량"].sum()) if "불량수량" in scope_b.columns else 0.0
    tbl = None
    if not (scope_a.empty and scope_b.empty):
        ga = scope_a.groupby(breakdown_col)["Q_cost_백만원"].sum()
        gb = scope_b.groupby(breakdown_col)["Q_cost_백만원"].sum()
        tbl = pd.DataFrame({"기간 A": ga, "기간 B": gb}).fillna(0.0)
        tbl["차이(B−A)"] = tbl["기간 B"] - tbl["기간 A"]
        tbl["증감률(%)"] = tbl["차이(B−A)"] / tbl["기간 A"].where(tbl["기간 A"] > 0) * 100
        tbl = tbl.sort_values("기간 B", ascending=False)
        total = pd.DataFrame({
            "기간 A": [tbl["기간 A"].sum()], "기간 B": [tbl["기간 B"].sum()],
            "차이(B−A)": [tbl["차이(B−A)"].sum()],
            "증감률(%)": [tbl["차이(B−A)"].sum() / tbl["기간 A"].sum() * 100 if tbl["기간 A"].sum() else float("nan")],
        }, index=["합계"])
        tbl = pd.concat([tbl, total])
        tbl.index.name = breakdown_label
    return {
        "sum_a": sum_a, "sum_b": sum_b, "qty_a": qty_a, "qty_b": qty_b, "tbl": tbl,
        "fig": charts.render_period_compare(scope_a, scope_b, "기간 A", "기간 B") if tbl is not None else None,
    }


@st.fragment
def _section_period_compare():
    st.caption(
        "기간 A·B를 각각 고르고, 공정 → 제품군 → 소분류 → 세부내역 → 상세(원본) 순으로 좁혀가며 비교합니다. "
        "중간 단계에서 **(전체)**를 고르면 그 단계의 항목별로 나눠서 표를 보여줍니다. "
        "기본값은 A=전월, B=최신월이며, 조건을 고른 뒤 **분석 실행**을 눌러야 계산합니다."
    )
    if len(_monthly) >= 2:
        def_a, def_b = _month_bounds(_monthly.index[-2]), _month_bounds(_monthly.index[-1])
    elif len(_monthly) == 1:
        def_a = def_b = _month_bounds(_monthly.index[-1])
    else:
        def_a = def_b = (min_date, max_date)
    ca, cb = st.columns(2)
    with ca:
        st.markdown("**기간 A**")
        a_s, a_e = _date_range_inputs("cmp_a", default_start=def_a[0], default_end=def_a[1])
    with cb:
        st.markdown("**기간 B**")
        b_s, b_e = _date_range_inputs("cmp_b", default_start=def_b[0], default_end=def_b[1])

    # (v24.3) 드롭다운 후보를 "기간 A∪B 안에 실제로 있는 값"으로 제한 — 예전엔 전체 기간 기준이라
    # 다른 달에 1건만 있는 항목도 목록에 떠서, 고르고 나면 "두 기간 모두 데이터 없음"만 나오는 문제가 있었음
    # (사용자 제보: 재공품·1연도·조건설정에 '[도장]_극판 중량 부적합'이 떴는데 7월엔 0건).
    pool = _in_range(data, min(a_s, b_s), max(a_e, b_e))
    chosen = []
    breakdown_col, breakdown_label = None, None
    for (col, lbl), c in zip(_CMP_LEVELS, st.columns(len(_CMP_LEVELS))):
        opts = sorted(pool[col].dropna().astype(str).unique().tolist())
        if col == "공정":
            opts = [p for p in ["재공품", "조립", "완성"] if p in opts] or opts
        else:
            opts = ["(전체)"] + opts
        key = f"cmp_{col}"
        if opts and st.session_state.get(key) not in opts:
            st.session_state[key] = opts[0]  # 기간을 바꿔 이전 선택값이 목록에서 사라지면 첫 항목으로
        pick = c.selectbox(lbl, opts, key=key) if opts else None
        if pick and pick != "(전체)":
            pool = pool[pool[col].astype(str) == pick]
            chosen.append(pick)
        elif breakdown_col is None:
            breakdown_col, breakdown_label = col, lbl
    if breakdown_col is None:
        breakdown_col, breakdown_label = "상세내역_원본", "상세(원본)"

    n_a, n_b = len(_in_range(pool, a_s, a_e)), len(_in_range(pool, b_s, b_e))
    st.caption(f"선택한 조합의 데이터: 기간 A **{n_a:,}건** · 기간 B **{n_b:,}건** (드롭다운에는 두 기간 안에 실제로 있는 항목만 나옵니다)")

    params = {"a": (a_s, a_e), "b": (b_s, b_e), "chosen": tuple(chosen), "breakdown": breakdown_col}
    if st.button("🔍 분석 실행", type="primary", key="cmp_run"):
        with st.spinner("분석 중입니다... 잠시만 기다려주세요"):
            st.session_state["cmp_result"] = {
                "params": params, "chosen": chosen, "breakdown_label": breakdown_label,
                **_run_period_compare(pool, params["a"], params["b"], breakdown_col, breakdown_label),
            }
    res = st.session_state.get("cmp_result")
    if res is None:
        st.info("조건을 고른 뒤 **🔍 분석 실행**을 누르면 결과가 여기에 나타납니다.")
        return
    if res["params"] != params:
        st.warning("조건이 바뀌었습니다 — 아래 결과는 이전 조건 기준입니다. **분석 실행**을 다시 누르면 갱신됩니다.")

    (ra_s, ra_e), (rb_s, rb_e) = res["params"]["a"], res["params"]["b"]
    label_a, label_b = f"{ra_s:%Y-%m-%d} ~ {ra_e:%Y-%m-%d}", f"{rb_s:%Y-%m-%d} ~ {rb_e:%Y-%m-%d}"
    diff = res["sum_b"] - res["sum_a"]
    pct = diff / res["sum_a"] * 100 if res["sum_a"] else None
    st.markdown(f"**{' · '.join(res['chosen'])}**")
    st.markdown(
        '<div class="kpi-grid kpi-grid-3">'
        + _kpi_card("기간 A 합계", f"{res['sum_a']:,.2f}", "백만원", "",
                    f"{label_a} · 불량수량 {res['qty_a']:,.0f}개 · {(ra_e - ra_s).days + 1}일")
        + _kpi_card("기간 B 합계", f"{res['sum_b']:,.2f}", "백만원", "",
                    f"{label_b} · 불량수량 {res['qty_b']:,.0f}개 · {(rb_e - rb_s).days + 1}일")
        + _kpi_card("차이 (B − A)", f"{diff:+,.2f}", "백만원",
                    _delta_html(diff, pct, up_word="증가", down_word="감소") if (res["sum_a"] or res["sum_b"]) else "",
                    "A가 0이면 증감률은 표시하지 않습니다" if not res["sum_a"] else "")
        + "</div>",
        unsafe_allow_html=True,
    )
    if res["tbl"] is None:
        st.info("두 기간 모두 선택한 조합에 해당하는 데이터가 없습니다.")
        return
    tc, cc = st.columns([1.2, 1])
    with tc:
        bl = res["breakdown_label"]
        st.markdown(f"**{bl}별 비교** (선택한 조합 안에서 {bl} 기준으로 나눔)")
        tbl = res["tbl"]
        st.dataframe(
            _style_delta(tbl, ["차이(B−A)", "증감률(%)"]), width="stretch",
            column_config={
                "기간 A": st.column_config.Column(f"기간 A ({label_a})"),
                "기간 B": st.column_config.Column(f"기간 B ({label_b})"),
            },
        )
        st.download_button(
            "📥 이 비교표 엑셀 다운로드",
            data=df_to_excel_bytes(tbl.reset_index(), "기간 비교"),
            file_name=f"세부분석_기간비교_{dt.date.today().isoformat()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="cmp_download",
        )
    with cc:
        st.markdown("**기간별 추이** (기간 폭에 따라 일/주/월 단위 자동)")
        st.plotly_chart(res["fig"], key="cmp_chart")


with st.expander(
    "🔍 세부분석 — 기간 비교 — 목적: 두 기간을 골라 특정 항목(원본 부적합 내역까지)의 Q-cost 변화를 비교",
    expanded=True,
):
    _section_period_compare()

# ── ④ 총수율 종합표(고정 양식) ──────────────────────────────
# 사용자 요청(2026-09-11): "총수율 표 양식.xlsx"의 실제 회사 리포트 구조를 그대로 재현 — 위 상세표와
# 달리 화면 상단 필터 영향을 받지 않고 항상 전체 데이터 기준이며, 행(공정→제품군→소분류→세부내역)과
# 열(1~12월+합계)이 업로드된 데이터와 무관하게 항상 고정으로 나온다(데이터 없는 칸은 0).
# (v24.0) 사용자 요청으로 자유 피벗 대신 이 표를 4번째 자리에(기본 펼침) 두고 엑셀 다운로드 유지.
with st.expander(
    "📋 총수율 종합표 (고정 양식) — 목적: 사내 공식 리포트 양식 그대로 전체 데이터 기준 값을 확인·엑셀 다운로드",
    expanded=True,
):
    today = dt.date.today().isoformat()

    # (v24.1) 목표 시트를 양식 칸에 붙여 "월 목표 / 누계 목표 / 달성률(%)" 3열 추가 — 이름이 같은 항목은
    # 그대로, 다른 항목은 그 그룹의 "기타"로 합산(사용자 지침), 칸이 없는 그룹(V-type)은 제외.
    _months_elapsed = int(data["월"].nunique())
    if _targets:
        _target_map, _target_notes = fixed_report.map_targets_to_template(_targets["leaf"])
        _target_items = tuple(sorted(_target_map.items()))
    else:
        _target_map, _target_notes, _target_items = None, None, None
    st.caption(
        "`총수율 표 양식.xlsx`(폐기비용(창원) 보정전 시트) 구조 그대로 — 위 필터와 무관하게 항상 전체 데이터 기준, "
        "행·열 틀은 고정이고 실제 계산된 Q-cost(백만원)만 채워집니다."
        + (f" 우측 3열은 26년 목표(월 목표 · 누계 목표 = 월 목표 × 데이터가 있는 {_months_elapsed}개월 · 달성률 = 합계 ÷ 누계 목표, "
           "100% 초과는 주황·이하는 청록)입니다." if _target_map is not None else " (목표 시트를 읽지 못해 목표 열은 생략)")
    )
    fixed, unmapped = get_fixed_report(_data_sig, data, _target_items, _months_elapsed)
    if _target_notes is not None and len(_target_notes):
        with st.popover(f"목표 매핑 안내 — 양식과 이름이 달라 '기타'로 합산/제외한 목표 항목 {len(_target_notes)}건"):
            st.caption("목표 시트가 우리 양식보다 세분화돼 있어, 동일 항목 외에는 해당 그룹의 '기타'에 더했습니다(사용자 지침). 소계·합계는 영향 없습니다.")
            st.dataframe(_target_notes.round(2), hide_index=True, width="stretch")
    st.download_button(
        "📊 총수율 종합표 엑셀 다운로드",
        data=df_to_excel_bytes(fixed, "총수율 종합표", index=False),
        file_name=f"총수율_종합표_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.markdown(fixed_report.to_merged_html(fixed), unsafe_allow_html=True)

    if len(unmapped):
        st.caption(
            f"⚠️ 이 양식에 자리가 없어 표에서 빠진 데이터 {len(unmapped)}건 그룹 있음(예: 제품군 자체가 "
            "라인명 매핑표에 없어 '미분류'로 남은 재공품 데이터). 합계는 영향 없이 그대로 집계됩니다."
        )
        with st.popover("빠진 데이터 상세 보기"):
            st.dataframe(unmapped)

# ── ⑤ Q-cost 교차 비교(귀책부서·라인/설비별) ──────────────────
# 사용자 요청(2026-09-22, 시계열 비교 제안 중 "집단·교차 비교" 항목) — 공정/제품군 외에 어떤
# 귀책부서·라인(설비)이 비용을 주도하는지 확인. 항목 수가 많아(귀책부서 29개, 라인 61개) 상위
# 15개만 그린다(전체 대비 %는 항상 유지 — charts.render_bar의 top_n 옵션 참고).
@st.fragment
def _section_cross_compare():
    st.caption("귀책부서·라인(설비) 등 다른 관점에서 어떤 그룹이 비용을 주도하는지 확인합니다(상위 15개).")
    cx_start, cx_end = _date_range_inputs("cross")
    cx_scope = _in_range(data, cx_start, cx_end)
    cross_dim = st.radio(
        "비교 기준", ["귀책부서", "라인(설비)"], horizontal=True, key="cross_dim", label_visibility="collapsed",
    )
    cross_col = "귀책부서명" if cross_dim == "귀책부서" else "라인명_통합"
    cx_data = cx_scope[cx_scope[cross_col].notna()] if cross_col in cx_scope.columns else cx_scope.iloc[0:0]
    if cx_data.empty:
        st.info(f"'{cross_dim}' 데이터가 없습니다(해당 공정에 없는 컬럼일 수 있습니다 — 예: 조립은 귀책부서 정보가 없습니다).")
    else:
        cross_fig, _ = charts.render_bar(cx_data, cross_col, top_n=15)
        st.plotly_chart(cross_fig, key=f"cross_bar_{cross_dim}")


with st.expander(
    "🏭 Q-cost 교차 비교 — 목적: 귀책부서·라인(설비) 등 다른 기준으로 비용을 주도하는 그룹 확인",
    expanded=False,
):
    _section_cross_compare()

# ── ⑥ Q-cost 세부분석 — 월별 추이(공정·제품군·소분류·세부내역 조합) ────────
# 사용자 요청(2026-09-16): "공정-소분류-상세내역의 경향을 볼 수 있는 꺾은선 그래프. 예를 들어
# 1연도 생산중 순수불량 - 수집부의 1~8월 꺾은선" — 계층 선택(공정→제품군→소분류→세부내역) 후
# 그 조합 하나의 월별 Q-cost 추이만 그린다. 이 섹션 전용 기간 필터도 따로 둔다.
# (v24.0) 기간 비교 세부분석(③)이 생기면서 맨 아래로 내리고 기본 접힘으로 둠(기능은 그대로).
@st.fragment
def _section_detail_trend():
    st.caption("공정 · 제품군 · 소분류 · 세부내역을 차례로 고르면 그 조합의 월별 Q-cost 추이가 나타납니다.")
    da_start, da_end = _date_range_inputs("detail")

    p1, p2, p3, p4 = st.columns(4)
    detail_proc = p1.selectbox("공정", sorted(data["공정"].unique().tolist()), key="detail_proc")
    fam_pool2 = data[data["공정"] == detail_proc]
    detail_fam = p2.selectbox("제품군", sorted(fam_pool2["제품군"].unique().tolist()), key="detail_fam")
    sub_pool = fam_pool2[fam_pool2["제품군"] == detail_fam]
    detail_subtype = p3.selectbox("소분류", sorted(sub_pool["소분류"].unique().tolist()), key="detail_subtype")
    detail_pool = sub_pool[sub_pool["소분류"] == detail_subtype]
    detail_leaf = p4.selectbox("세부내역", sorted(detail_pool["세부내역"].unique().tolist()), key="detail_leaf")

    detail_scope = _in_range(
        data[
            (data["공정"] == detail_proc) & (data["제품군"] == detail_fam)
            & (data["소분류"] == detail_subtype) & (data["세부내역"] == detail_leaf)
        ],
        da_start, da_end,
    )
    st.markdown(f"**{detail_proc} · {detail_fam} · {detail_subtype} · {detail_leaf}**")
    if detail_scope.empty:
        st.info("선택한 조합·기간에 해당하는 데이터가 없습니다.")
    else:
        st.plotly_chart(charts.render_monthly_trend(detail_scope), key="trend_detail")


with st.expander(
    "📉 Q-cost 세부분석 — 월별 추이 — 목적: 공정-제품군-소분류-세부내역 조합 하나를 골라 월별 추이를 추적",
    expanded=False,
):
    _section_detail_trend()
