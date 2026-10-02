"""월별 누적 저장소. 담당자가 매월 업로드한 부적합 원본을 (공정, 연월) 단위 parquet로 보관하고,
필요 시 특정 월만 삭제 후 재업로드할 수 있게 한다 (04_prd_appendix.md 4장 '데이터 반영 방식').
생산량은 더 이상 쓰지 않으므로 공정당 부적합 하나만 관리한다."""
import json
import os
import pandas as pd

BASE_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
STORE_DIR = os.path.join(BASE_DIR, "output", "data_store")
SEEDED_MARKER = os.path.join(STORE_DIR, ".seeded")


def _dir(process: str) -> str:
    path = os.path.join(STORE_DIR, f"{process}_부적합")
    os.makedirs(path, exist_ok=True)
    return path


def _sanitize_for_parquet(df: pd.DataFrame) -> pd.DataFrame:
    """실제 엑셀 원본은 한 컬럼 안에 문자열/숫자가 섞여 들어오는 경우가 있다(예: 불량명에 숫자 코드가 섞임).
    pyarrow는 그런 혼합 타입 object 컬럼을 저장하지 못해 에러가 나므로, 저장 직전 문자열로 통일한다
    (NaN은 그대로 유지 — str()로 바꾸면 "nan" 문자열이 되어 이후 결측 판정이 깨짐)."""
    df = df.copy()
    for col in df.columns:
        if df[col].dtype == object:
            df[col] = df[col].map(lambda v: v if pd.isna(v) else str(v))
    return df


def _seed_months_path() -> str:
    """STORE_DIR이 테스트 등에서 바뀔 수 있어(patch) SEEDED_MARKER처럼 모듈 로드 시점에 고정하지
    않고 호출 시점마다 현재 STORE_DIR 기준으로 계산한다."""
    return os.path.join(STORE_DIR, ".seeded_months.json")


def _load_seed_months() -> dict:
    path = _seed_months_path()
    if not os.path.exists(path):
        return {}
    with open(path, "r", encoding="utf-8") as f:
        return json.load(f)


def _save_seed_months(data: dict):
    os.makedirs(STORE_DIR, exist_ok=True)
    with open(_seed_months_path(), "w", encoding="utf-8") as f:
        json.dump(data, f)


def mark_seed_months(process: str, months: list):
    """최초 1회 더미데이터 시드가 어느 (공정, 월)에 무엇을 채웠는지 기록해둔다 — save_by_month()가
    이 기록을 보고 "그 칸이 아직 더미데이터뿐"이면 사용자의 첫 실제 업로드 때 병합이 아니라 교체를
    하도록 한다(아래 save_by_month 설명 참고)."""
    data = _load_seed_months()
    data.setdefault(process, [])
    for m in months:
        if m not in data[process]:
            data[process].append(m)
    _save_seed_months(data)


def save_by_month(df: pd.DataFrame, process: str, date_col: str) -> list:
    """date_col(파일명이 아니라 데이터 안의 날짜 컬럼) 기준으로 월별로 나눠 저장한다.

    (v14.10 정정) 예전엔 같은 달에 다시 저장하면 그 달 파일을 통째로 덮어썼는데(재업로드=교체),
    반복 테스트 중 일부 파일만 다시 올리거나 파일 하나가 에러로 빠지면 그 순간 이미 저장돼 있던
    나머지 데이터까지 사라지는 문제가 있었다(사용자 제보: "테스트할 때마다 재공품이 사라지는 것 같다").
    이제는 기존에 저장된 그 달 데이터에 새로 올라온 데이터를 **병합**하고, 완전히 똑같은 행(중복
    재업로드)만 하나로 합친다 — 같은 파일을 두 번 올려도 중복 집계되지 않으면서, 새 파일을 추가로
    올려도 기존 데이터가 사라지지 않는다. 정말 통째로 갈아끼우고 싶으면 "데이터 관리"에서 그 달을
    먼저 삭제한 뒤 다시 올리면 된다(기존 방식 그대로 사용 가능).

    (v18.1 버그 수정) 이 "병합" 방식이 최초 1회 자동 시드된 더미데이터와도 그대로 합쳐져 버리는
    부작용이 있었다 — 사용자가 실제 파일을 처음 업로드해도 더미데이터가 영구히 안 지워지고 섞여
    남아있었음(예: 더미 재공품 CSV에 우연히 들어있던 "AGM완성_..." 라인명이 실제 재공품 데이터에
    섞여 "미분류"로 나타남 — 실데이터 문제가 아니라 이 버그였다). 이제 어떤 (공정,월)이 "아직
    더미데이터뿐"인지 기록해뒀다가, 그 칸에 처음으로 실제 데이터가 들어오면 병합이 아니라 완전히
    교체하고 더미 기록에서 지운다(그 다음부터는 정상적으로 병합됨)."""
    folder = _dir(process)
    saved = []
    seed_months = _load_seed_months()
    seed_for_process = set(seed_months.get(process, []))
    valid = _sanitize_for_parquet(df[df[date_col].notna()])
    for month, sub in valid.groupby(valid[date_col].dt.strftime("%Y-%m")):
        path = os.path.join(folder, f"{month}.parquet")
        if os.path.exists(path) and month not in seed_for_process:
            existing = pd.read_parquet(path)
            if "월" in existing.columns and existing["월"].dtype != "Int64":
                existing["월"] = existing["월"].astype("Int64")
            combined = pd.concat([existing, sub], ignore_index=True).drop_duplicates()
        else:
            combined = sub
        combined.to_parquet(path, index=False)
        saved.append(month)
        if month in seed_for_process:
            seed_for_process.discard(month)
    if process in seed_months:
        seed_months[process] = sorted(seed_for_process)
        _save_seed_months(seed_months)
    return saved


def list_months(process: str) -> list:
    folder = _dir(process)
    return sorted(f[:-8] for f in os.listdir(folder) if f.endswith(".parquet"))


def load_all(process: str) -> pd.DataFrame:
    folder = _dir(process)
    files = sorted(f for f in os.listdir(folder) if f.endswith(".parquet"))
    if not files:
        return pd.DataFrame()
    df = pd.concat([pd.read_parquet(os.path.join(folder, f)) for f in files], ignore_index=True)
    if "월" in df.columns and df["월"].dtype != "Int64":
        # data.load_defect()가 Int64로 저장하기 전(v14.1 이전)에 쌓인 parquet는 float64(예: 8.0)로
        # 남아있을 수 있어, 불러올 때마다 정규화한다 — 재업로드 없이도 "8.0" 표기가 바로 고쳐짐.
        df["월"] = df["월"].astype("Int64")
    return df


def delete_month(process: str, month: str) -> bool:
    path = os.path.join(_dir(process), f"{month}.parquet")
    if os.path.exists(path):
        os.remove(path)
        return True
    return False


def store_fingerprint() -> str:
    """output/data_store 안의 부적합 parquet 파일들의 (경로,수정시각,크기)를 모은 지문 문자열.

    (v14.11 도입) `build_dataset()`의 st.cache_data 캐시 키로 예전엔 세션마다 0부터 다시 세는
    `st.session_state["data_version"]` 카운터를 썼는데, `st.cache_data`의 캐시는 세션이 아니라
    **서버 프로세스 전체에서 공유**된다 — 새 브라우저 세션(F5, 새 탭 등)이 열려 카운터가 다시 0으로
    시작하면, 이 프로세스에서 과거(예: 파일 일부만 올렸을 때) 이미 실행됐던 `data_version=0` 호출의
    캐시 결과와 키가 우연히 겹쳐 그 오래된 결과가 그대로 재사용되는 심각한 버그가 있었다(사용자 제보:
    새로고침 후 재공품 비중이 사라져 보임 — 실제로는 디스크엔 데이터가 멀쩡히 남아있었음).
    세션과 무관하게 디스크의 실제 파일 상태만으로 지문을 만들면 이 충돌이 원천적으로 생기지 않는다."""
    parts = []
    for process in ("재공품", "조립", "완성"):
        folder = _dir(process)
        for name in sorted(os.listdir(folder)):
            if name.endswith(".parquet"):
                info = os.stat(os.path.join(folder, name))
                parts.append(f"{process}/{name}:{info.st_mtime_ns}:{info.st_size}")
    return "|".join(parts) or "empty"


def has_ever_seeded() -> bool:
    return os.path.exists(SEEDED_MARKER)


def mark_seeded():
    os.makedirs(STORE_DIR, exist_ok=True)
    with open(SEEDED_MARKER, "w") as f:
        f.write("1")
