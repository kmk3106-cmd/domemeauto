# -*- coding: utf-8 -*-
"""
네이버 카테고리 ID 자동 분류기 (고속버전: asyncio 동시처리 + 백오프 + 캐시)
- 카테고리맵 엑셀(A:ID, B:대, C:중, D:소)
- 4번_통합상품명.xlsx: A열=상품ID, B열=상품명
- rapidfuzz로 후보 TOP N 추린 뒤, ChatGPT 응답을 '동시에' 받아 속도 개선
- 429/timeout 등 에러는 지수백오프로 재시도
- 동일 상품명은 결과 캐시로 API 호출 생략

필요 패키지:
pip install pandas openpyxl rapidfuzz tqdm openai
"""

import os
import re
import json
import math
import asyncio
import random
from typing import List, Dict, Any, Tuple

import pandas as pd
from rapidfuzz import process, fuzz
from tqdm import tqdm

try:
    from openai import AsyncOpenAI
except Exception:
    AsyncOpenAI = None


# ===================== (중요) API 키/모델/동시성 설정 =====================
OPENAI_API_KEY   = os.environ.get("OPENAI_API_KEY", "")   # .env 의 OPENAI_API_KEY
MODEL_NAME       = "gpt-4o-mini"                # 경제형 모델(원하면 변경)
MAX_CONCURRENCY  = 20                           # 동시 배치 요청 개수(속도 개선)
BATCH_SIZE       = 8                            # 한 번에 분류할 상품 수(API 호출 1/BATCH_SIZE로 감소, 8개=12.5% 호출)
MAX_RETRIES      = 5                            # 재시도 횟수(429/네트워크 오류)
BASE_BACKOFF_SEC = 0.8                          # 백오프 기본(지수*랜덤지터)
REQUEST_TIMEOUT  = 45                           # 요청당 타임아웃(초)


# ===================== 파일 경로/설정 =====================
CATEGORIES_XLSX = r"C:\Users\USER\Documents\국내위탁\마이박스\네이버카테고리.xlsx"
PRODUCTS_XLSX   = r"C:\Users\USER\Documents\국내위탁\마이박스\26년3월3주차\6번사업자\6번_통합상품명_keywords.xlsx"
OUTPUT_XLSX     = r"C:\Users\USER\Documents\국내위탁\마이박스\26년3월3주차\6번사업자\6번_통합상품명_카테고리매핑_결과.xlsx"


PRODUCT_SHEET_NAME_OR_INDEX = 0    # 시트명 또는 인덱스(기본 0)
PRODUCT_ID_COLUMN_LETTER    = "A"  # 상품ID
PRODUCT_NAME_COLUMN_LETTER  = "B"  # 상품명

TOP_K_CANDIDATES = 8               # 퍼지 후보 개수(줄일수록 토큰/비용↓, 정확도 영향 주의)


# ===================== 유틸 =====================
def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        return ""
    s = s.strip()
    s = re.sub(r"\s+", " ", s)
    return s


def load_category_map(xlsx_path: str) -> pd.DataFrame:
    """
    카테고리 엑셀 스키마(헤더 없어도 위치 기준으로 읽음):
    A: category_id, B: d1(대), C: d2(중), D: d3(소)
    """
    df_raw = pd.read_excel(xlsx_path, header=None, dtype=str)
    if df_raw.shape[1] < 4:
        raise ValueError("카테고리 파일에 최소 4개 컬럼(A-D)이 필요합니다.")

    df = df_raw.iloc[:, :4].copy()
    df.columns = ["category_id", "d1", "d2", "d3"]
    for c in ["category_id", "d1", "d2", "d3"]:
        df[c] = df[c].astype(str).map(normalize_text)

    # 헤더 문자열/공란/비숫자 제거
    header_candidates = {"카테고리ID", "category_id", "카테고리 ID", "ID"}
    df = df[~df["category_id"].isin(header_candidates)]
    df = df[df["category_id"].str.fullmatch(r"\d+").fillna(False)]

    df["path"] = df[["d1", "d2", "d3"]].fillna("").apply(
        lambda r: " > ".join([x for x in r if x]),
        axis=1
    )
    df = df[df["category_id"].str.len() > 0].reset_index(drop=True)
    # 제외 카테고리:
    # - 서적/DVD/CD음반 (정가정책상 할인율 개별설정 필요)
    # - 전자기기 계열 전체 (노트북/배터리/충전기/PC/태블릿 등)
    exclude_keywords = (
        "도서", "서적", "DVD", "CD음반",
        "전자기기", "디지털", "가전", "노트북", "랩탑", "컴퓨터", "PC", "태블릿",
        "모니터", "키보드", "마우스", "배터리", "밧데리", "충전기", "보조배터리",
        "이어폰", "헤드폰", "스피커", "카메라", "프린터", "저장장치", "SSD", "HDD",
    )
    pattern = "|".join(re.escape(k) for k in exclude_keywords)
    df = df[~df["path"].str.contains(pattern, na=False, regex=True)].reset_index(drop=True)
    return df


def _col_index(letter: str) -> int:
    return ord(letter.upper()) - ord('A')


def load_product_records(xlsx_path: str,
                         sheet=0,
                         id_col="A",
                         name_col="B") -> List[Tuple[str, str]]:
    """
    4번_통합상품명.xlsx에서 (상품ID, 상품명) 튜플 리스트 반환.
    흔한 헤더 문자열 감지 시 첫 행 자동 스킵.
    """
    df = pd.read_excel(xlsx_path, sheet_name=sheet, header=None, dtype=str)
    id_idx = _col_index(id_col)
    name_idx = _col_index(name_col)
    if max(id_idx, name_idx) >= df.shape[1] or min(id_idx, name_idx) < 0:
        raise ValueError(f"시트에 {id_col}/{name_col} 열이 없습니다.")

    sub = df.iloc[:, [id_idx, name_idx]].copy()
    sub.columns = ["상품ID", "상품명"]
    sub["상품ID"] = sub["상품ID"].map(normalize_text)
    sub["상품명"] = sub["상품명"].map(normalize_text)

    # 첫 행 헤더 자동 제거(있을 때만)
    id_header_candidates = {"상품ID", "상품번호", "ID", "product_id", "상품id"}
    name_header_candidates = {"상품명", "제품명", "product_name", "상품명(필수)"}
    if len(sub) > 0:
        first_id = (sub.iloc[0]["상품ID"] or "").strip()
        first_nm = (sub.iloc[0]["상품명"] or "").strip()
        if first_id in id_header_candidates or first_nm in name_header_candidates:
            sub = sub.iloc[1:, :]

    # 빈 상품명 제거
    sub = sub[(sub["상품명"].astype(str).str.len() > 0)]
    sub["상품ID"] = sub["상품ID"].fillna("")

    return list(sub.itertuples(index=False, name=None))  # List[(상품ID, 상품명)]


def shortlist_candidates(name: str, cat_paths: List[str], k: int = 8) -> List[Tuple[str, float]]:
    """
    rapidfuzz로 name과 cat_paths 간 유사도 계산하여 상위 k개 반환
    각 아이템은 (카테고리경로, 점수) 튜플
    """
    return [(m[0], float(m[1])) for m in process.extract(name, cat_paths, scorer=fuzz.WRatio, limit=k)]


def build_prompt(product_name: str,
                 candidates: List[Dict[str, Any]]) -> List[Dict[str, str]]:
    """
    Chat Completions messages payload 구성 (JSON 출력 강제) - 단일 상품용
    """
    system = (
        "당신은 네이버 스마트스토어 상품 카테고리 분류 전문가입니다. "
        "입력된 상품명과 후보 카테고리 목록(ID, 경로)을 보고 가장 적절한 '단 하나의' 카테고리ID를 고르세요. "
        "가능한 한 네이버 쇼핑 카테고리 매핑 관행에 맞춰 현실적으로 판단하세요. "
        "출력은 반드시 JSON으로만 내보냅니다."

    )
    user = {
        "product_name": product_name,
        "candidates": [{"id": c["id"], "path": c["path"]} for c in candidates],
        "instructions": (
            "반드시 아래 JSON 스키마를 따르세요.\n"
            "{\n"
            '  "category_id": "문자열(선택한 카테고리ID)",\n'
            '  "confidence": 0.0~1.0,\n'
            '  "reason": "짧은 근거(한국어)"\n'
            "}\n"
            "주의: 후보 중 하나를 반드시 선택하고, ID는 candidates의 id 중 하나여야 합니다."
        )
    }
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": json.dumps(user, ensure_ascii=False)}
    ]


def build_batch_prompt(batch: List[Tuple[str, List[Dict[str, Any]]]]) -> List[Dict[str, str]]:
    """
    상품 N개를 한 번에 분류하는 배치용 프롬프트. 응답은 JSON 배열(길이 N).
    batch: [(product_name, candidates), ...]
    """
    system = (
        "당신은 네이버 스마트스토어 상품 카테고리 분류 전문가입니다. "
        "아래 products 배열의 각 상품에 대해, 해당 후보 카테고리 중 가장 적절한 '단 하나의' 카테고리ID를 고르세요. "
        "출력은 반드시 JSON 배열로만 내보냅니다."
    )
    products_payload = [
        {"product_name": name, "candidates": [{"id": c["id"], "path": c["path"]} for c in cand]}
        for name, cand in batch
    ]
    user = {
        "products": products_payload,
        "instructions": (
            "반드시 아래 형태의 JSON 배열로 응답하세요. 배열 길이는 입력 상품 개수와 동일해야 합니다.\n"
            "[ {\"category_id\": \"선택한ID\", \"confidence\": 0.0~1.0, \"reason\": \"짧은 근거\"}, ... ]\n"
            "각 항목의 category_id는 해당 상품의 candidates의 id 중 하나여야 합니다."
        )
    }
    return [
        {"role": "system", "content": system},
        {"role": "user", "content": json.dumps(user, ensure_ascii=False)}
    ]


def extract_json(text: str) -> Dict[str, Any]:
    """
    모델 응답에서 JSON만 뽑아 파싱. 코드펜스/텍스트 섞임 대비.
    """
    m = re.search(r"\{.*\}", text or "", flags=re.DOTALL)
    raw = (m.group(0) if m else (text or "")).strip().strip("`").strip()
    try:
        return json.loads(raw)
    except Exception:
        return {}


def extract_json_array(text: str) -> List[Dict[str, Any]]:
    """
    모델 응답에서 JSON 배열만 뽑아 파싱. 배치 응답용.
    """
    m = re.search(r"\[.*\]", text or "", flags=re.DOTALL)
    raw = (m.group(0) if m else (text or "")).strip().strip("`").strip()
    try:
        arr = json.loads(raw)
        return arr if isinstance(arr, list) else []
    except Exception:
        return []


# ===================== 비동기 API 호출 =====================
def _parse_single_result(data: Dict, candidates: List[Dict], fallback_id: str) -> Dict[str, Any]:
    pred_id = str(data.get("category_id", "")).strip()
    confidence = float(data.get("confidence", 0.0)) if isinstance(data.get("confidence"), (int, float)) else 0.0
    reason = str(data.get("reason", "")).strip()
    valid_ids = {c["id"] for c in candidates}
    if pred_id not in valid_ids:
        pred_id = fallback_id
        reason = (reason + " | 후보 외 ID여서 퍼지 1위로 대체").strip()
    return {
        "pred_id": pred_id,
        "confidence": confidence,
        "reason": reason,
        "raw": json.dumps(data, ensure_ascii=False)
    }


async def call_openai_batch_with_backoff(
    client: AsyncOpenAI, model: str,
    batch: List[Tuple[str, str, List[Dict[str, Any]], str]],  # (name, key, candidates, fallback_id)
    semaphore: asyncio.Semaphore
) -> List[Dict[str, Any]]:
    """
    배치 한 번에 API 호출. batch 길이만큼 결과 리스트 반환(순서 유지).
    """
    prompt_batch = [(name, cand) for name, _key, cand, _fb in batch]
    fallbacks = [fb for _n, _k, _c, fb in batch]

    async with semaphore:
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                resp = await client.chat.completions.create(
                    model=model,
                    messages=build_batch_prompt(prompt_batch),
                    temperature=0.1,
                    timeout=REQUEST_TIMEOUT
                )
                content = resp.choices[0].message.content if resp.choices else ""
                arr = extract_json_array(content)
                if not arr and "[" in (content or ""):
                    arr = extract_json_array(content)
                if not arr:
                    m = re.search(r"\[.*\]", content or "", flags=re.DOTALL)
                    if m:
                        try:
                            arr = json.loads(m.group(0).strip().strip("`"))
                            if not isinstance(arr, list):
                                arr = []
                        except Exception:
                            arr = []

                results = []
                for i, (_, _key, candidates, fallback_id) in enumerate(batch):
                    if i < len(arr) and isinstance(arr[i], dict):
                        results.append(_parse_single_result(arr[i], candidates, fallback_id))
                    else:
                        results.append({
                            "pred_id": fallback_id,
                            "confidence": 0.0,
                            "reason": "배치 파싱 실패로 퍼지 1위 대체",
                            "raw": ""
                        })
                return results

            except Exception as e:
                if attempt >= MAX_RETRIES:
                    return [
                        {"pred_id": fb, "confidence": 0.0, "reason": f"오류로 대체. err={str(e)[:100]}", "raw": ""}
                        for fb in fallbacks
                    ]
                backoff = BASE_BACKOFF_SEC * (2 ** (attempt - 1)) * (1 + random.random() * 0.25)
                await asyncio.sleep(backoff)
        return [
            {"pred_id": fb, "confidence": 0.0, "reason": "재시도 초과", "raw": ""}
            for fb in fallbacks
        ]


# ===================== 메인(비동기) =====================
async def async_main():
    if AsyncOpenAI is None:
        raise RuntimeError("openai 패키지(AsyncOpenAI)를 찾을 수 없습니다.  pip install openai  후 최신 버전 사용하세요.")
    if not OPENAI_API_KEY or not OPENAI_API_KEY.startswith("sk-"):
        raise RuntimeError("OPENAI_API_KEY 값이 비어있거나 올바르지 않습니다. 자신의 키로 교체하세요.")

    client = AsyncOpenAI(api_key=OPENAI_API_KEY)

    # 1) 카테고리/경로 로드
    cat_df = load_category_map(CATEGORIES_XLSX)
    cat_paths = cat_df["path"].tolist()

    # 2) 상품 (ID, 이름) 로드
    products = load_product_records(PRODUCTS_XLSX,
                                    PRODUCT_SHEET_NAME_OR_INDEX,
                                    PRODUCT_ID_COLUMN_LETTER,
                                    PRODUCT_NAME_COLUMN_LETTER)

    # 3) 각 상품의 후보 준비 (동일 상품명은 1회만 API 호출)
    ordered_jobs = []  # [(pid, name, candidates, fallback_id), ...] 원본 순서 유지
    unique_items: Dict[str, Tuple[List[Dict], str]] = {}  # name -> (candidates, fallback_id)
    seen_names = []

    for pid, name in products:
        key = name
        s_list = shortlist_candidates(name, cat_paths, k=TOP_K_CANDIDATES)
        cand_df = pd.DataFrame(s_list, columns=["path", "score"]).merge(
            cat_df[["category_id", "path"]], on="path", how="left"
        ).dropna(subset=["category_id"]).sort_values("score", ascending=False)
        candidates = [{"id": str(r["category_id"]), "path": r["path"]} for _, r in cand_df.iterrows()]
        fallback_id = str(cand_df.iloc[0]["category_id"]) if not cand_df.empty else ""
        ordered_jobs.append((pid, name, candidates, fallback_id))
        if key not in unique_items:
            unique_items[key] = (candidates, fallback_id)
            seen_names.append(key)

    # 4) 배치별 비동기 분류 (API 호출 수 = ceil(고유상품수/BATCH_SIZE))
    semaphore = asyncio.Semaphore(MAX_CONCURRENCY)
    name_to_result: Dict[str, Dict[str, Any]] = {}

    async def process_batch(batch_items: List[Tuple[str, List[Dict], str]]) -> int:
        """배치 1회 API 호출 후 name_to_result에 저장, 처리 건수 반환"""
        if not batch_items:
            return 0
        full_batch = [(name, name, cand, fb) for name, cand, fb in batch_items]
        res_list = await call_openai_batch_with_backoff(
            client, MODEL_NAME, full_batch, semaphore
        )
        for (name, _, _), res in zip(batch_items, res_list):
            name_to_result[name] = res
        return len(batch_items)

    # 배치 생성 (고유 상품명만 API 호출)
    batches = []
    for i in range(0, len(seen_names), BATCH_SIZE):
        chunk = seen_names[i : i + BATCH_SIZE]
        batches.append([(name, unique_items[name][0], unique_items[name][1]) for name in chunk])

    pbar = tqdm(total=len(seen_names), desc="분류(배치동시처리)", ncols=100)
    tasks = [process_batch(b) for b in batches]
    for coro in asyncio.as_completed(tasks):
        n = await coro
        pbar.update(n)
    pbar.close()

    # 5) 원본 순서대로 결과 복원
    results = []
    for pid, name, candidates, fallback_id in ordered_jobs:
        res = name_to_result.get(name, {"pred_id": fallback_id, "confidence": 0.0, "reason": "캐시미스", "raw": ""})
        results.append((pid, name, res, candidates))

    # 6) 결과 가공/저장
    out_rows = []
    for pid, name, res, candidates in results:
        pred_id = res["pred_id"]
        pred_path = ""
        if pred_id:
            r = cat_df.loc[cat_df["category_id"] == pred_id]
            if not r.empty:
                pred_path = r.iloc[0]["path"]

        out_rows.append({
            "상품ID": pid,
            "상품명": name,
            "예측ID": pred_id,
            "예측카테고리": pred_path,
            "신뢰도": res.get("confidence", 0.0),
            "근거": res.get("reason", ""),
            "후보목록": json.dumps(candidates, ensure_ascii=False),
            "모델원문": res.get("raw", "")
        })

    out_df = pd.DataFrame(out_rows)
    desired = ["상품ID", "상품명", "예측ID", "예측카테고리", "신뢰도", "근거", "후보목록", "모델원문"]
    out_df = out_df[[c for c in desired if c in out_df.columns]]

    os.makedirs(os.path.dirname(OUTPUT_XLSX) or ".", exist_ok=True)
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="분류결과", index=False)

    print(f"[완료] 저장 경로: {OUTPUT_XLSX}")


if __name__ == "__main__":
    asyncio.run(async_main())