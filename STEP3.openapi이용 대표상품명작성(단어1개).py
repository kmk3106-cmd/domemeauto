#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
대표키워드_추출기_batch.py

여러 개의 원본 상품명 파일(A=상품ID, B=상품명)을 순차 처리하여
각 파일마다 [상품ID, 대표키워드] 결과 파일을 생성합니다.

특징
- 1행=헤더, 2행부터 데이터 가정 (A: 상품ID, B: 상품명)
- 대표키워드 1개만 산출: 명사+명사 또는 형용사+명사, 모든 공백 제거
- 모드: openai(기본) / offline (규칙 기반 백업)
- 응답 파싱 견고화 + API 실패 시 규칙 기반 대체
- 노이즈/메모/빈값 필터, 시작/종료/소요시간, per-row 진행률
- 각 입력 파일 이름에 _keywords 접미사를 붙여 같은 폴더 혹은 지정 폴더에 저장

사용법
1) INPUT_PATHS 리스트에 처리할 파일 경로를 넣습니다.
2) OUTPUT_DIR를 원하면 지정(비우면 입력과 같은 폴더에 저장)
3) USE_MODE, OPENAI_MODEL, API_KEY(또는 환경변수 OPENAI_API_KEY) 설정
4) 실행: python 대표키워드_추출기_batch.py
"""
from __future__ import annotations
import os
import re
import sys
import time
import datetime
import random
from typing import List, Optional, Dict, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

# -----------------------------
# 하드코딩: 여러 입력 파일 경로들 (A=상품ID, B=상품명)
# -----------------------------
INPUT_PATHS = [
    r"C:\Users\USER\Documents\국내위탁\마이박스\26년1월\1번사업자\1번_통합상품명.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_2.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_3.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_4.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_5.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_6.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_7.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_8.xlsx",
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251007\5번사업자\\5_9.xlsx",
    # r"C:\\Users\\USER\\Documents\\국내위탁\\마이박스\\250923\\상품목록.xlsx",
]

# 출력 폴더(선택): 비우면 입력 파일과 같은 폴더에 저장
OUTPUT_DIR = r""  # 예: r"C:\\Users\\USER\\Documents\\국내위탁\\결과"

# 모드/모델/API 키
USE_MODE = "openai"               # "openai" 또는 "offline"
OPENAI_MODEL = "gpt-4o-mini"
API_KEY = os.environ.get("OPENAI_API_KEY", "")  # .env 의 OPENAI_API_KEY

# 429 대응: 스로틀링 + 지수 백오프
MAX_RETRIES = 6                  # 429 시 최대 재시도 횟수
BASE_BACKOFF_SEC = 2.0           # 기본 대기(초)
MAX_BACKOFF_SEC = 120            # 최대 대기(초)
MAX_CONCURRENCY = 4              # 동시 요청 수(429 방지용 낮게)
REQUEST_DELAY_SEC = 0.15         # 요청 간 최소 간격(초)

# -----------------------------
# OpenAI SDK
# -----------------------------
try:
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None  # type: ignore

# -----------------------------
# 규칙 기반 유틸 (offline 모드/백업용)
# -----------------------------
KOREAN = re.compile(r"[가-힣]+")
NOISE_PATTERNS = [
    re.compile(r"\b(\d+\.?\d*)\s*(cm|mm|m|L|l|ml|g|kg|oz|inch|in|호|세트|p|P|개|매|장|쌍)\b", re.I),
    re.compile(r"\b(\d+\s*[xX×]\s*\d+(?:\.\d+)?)\b"),
]
BRACKET_CONTENT = re.compile(r"[\(\[\{＜《〈【].*?[\)\]\}＞》〉】]")
SPECIALS = re.compile(r"[~`!@#$%^&*\-=_+\|\\:;\"'<>/?·•…，、\.,]")
MULTISPACE = re.compile(r"\s+")
STOPWORDS = set(
    "색상 색상랜덤 랜덤 무료배송 당일발송 국내배송 해외배송 정품 새상품 이벤트 한정 특가 사은품 세트 구성 옵션 택일 선택형 "
    "대형 소형 미니 대 중 소 남성 여성 유니섹스 남녀공용 여름 겨울 봄 가을 신상 인기 베스트 베이직 고급 프리미엄 기본형 업소용 가정용 "
    "어린이 성인 학생 유아 요가 헬스 캠핑 등산 낚시 골프 러닝 자전거 자동차 차량용 캠핑용 여름용 겨울용 사계절용 업그레이드 리뉴얼 강화형"
    .split()
)

SYSTEM_PROMPT = (
    """
당신은 한국 쇼핑 오픈마켓용 상품명 가공 보조자입니다.
규칙을 철저히 따르세요.
1) 입력: 원본 상품명(한국어/영어 혼합 가능)
2) 출력: 대표키워드 1개만. 공백/개행/구두점 없이.
3) 형태: "명사+명사" 또는 "형용사+명사" (예: 무타공걸이, 대리석시트지, 방수시트지, 튼튼도어후크)
4) 브랜드/색상/사이즈/수량/옵션 표기는 제외.
5) 가능한 한 구체적이고 검색 핵심어에 가까운 조합을 선택.
6) 오직 최종 문자열만 출력(설명/따옴표/코멘트 금지).
    """
).strip()

# -----------------------------
# 공통 유틸
# -----------------------------

def _clean_name(name: str) -> str:
    s = str(name)
    s = BRACKET_CONTENT.sub(" ", s)
    for pat in NOISE_PATTERNS:
        s = pat.sub(" ", s)
    s = SPECIALS.sub(" ", s)
    s = s.replace("/", " ")
    s = MULTISPACE.sub(" ", s)
    return s.strip()


def _tokenize(name: str) -> List[str]:
    toks: List[str] = []
    for t in name.split():
        t = re.sub(r"[^가-힣A-Za-z0-9]", "", t)
        if t:
            toks.append(t)
    return toks


def _is_trash(tok: str) -> bool:
    if tok in STOPWORDS:
        return True
    if len(tok) == 1 and not tok.isdigit():
        return True
    if sum(ch.isdigit() for ch in tok) >= max(2, len(tok) - 1):
        return True
    return False


def _stitch(tokens: List[str]) -> Optional[str]:
    core = [t for t in tokens if not _is_trash(t)]
    if not core:
        return None
    kor = [t for t in core if KOREAN.search(t)]
    other = [t for t in core if t not in kor]
    ordered = kor + other
    if len(ordered) >= 2:
        return ordered[0] + ordered[1]
    return ordered[0]

# -----------------------------
# 429 백오프 + 캐시 유틸
# -----------------------------

def _parse_retry_after(err_msg: str) -> Optional[float]:
    """오류 메시지에서 'try again in X.XXs' 파싱"""
    m = re.search(r"try again in ([\d.]+)s?", err_msg, re.I)
    if m:
        return float(m.group(1)) + 0.5
    return None


def _llm_keyword_raw(client: "OpenAI", model: str, title: str, hint: str = "") -> str:
    """API 호출만 수행 (재시도 없음)"""
    user = "원본 상품명: " + title + (f"\n핵심후보: {hint}" if hint else "") + "\n대표키워드:"
    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user},
        ],
    )
    text = _extract_text_from_response(resp).strip()
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[\"'`.,]", "", text)
    return text or "대표키워드"


# -----------------------------
# OpenAI 호출/파싱
# -----------------------------

def _extract_text_from_response(resp) -> str:
    text = getattr(resp, "output_text", None)
    if text:
        return str(text)
    try:
        parts: List[str] = []
        for item in getattr(resp, "output", []) or []:
            if getattr(item, "type", None) == "message":
                for c in getattr(item, "content", []) or []:
                    if getattr(c, "type", None) == "text":
                        t = getattr(c, "text", None)
                        if t:
                            parts.append(str(t))
        if parts:
            return "".join(parts)
    except Exception:
        pass
    try:
        return resp.choices[0].message.content
    except Exception:
        return ""


def _llm_keyword(client: "OpenAI", model: str, title: str, hint: str = "",
                  cache: Optional[Dict[Tuple[str, str], str]] = None) -> str:
    """지수 백오프 + 캐시 적용"""
    key = (title, hint)
    if cache is not None and key in cache:
        return cache[key]
    last_err = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            time.sleep(REQUEST_DELAY_SEC)
            kw = _llm_keyword_raw(client, model, title, hint)
            if cache is not None:
                cache[key] = kw
            return kw
        except Exception as e:
            last_err = e
            err_str = str(e)
            if "429" in err_str or "rate_limit" in err_str.lower():
                retry_after = _parse_retry_after(err_str)
                delay = retry_after if retry_after else BASE_BACKOFF_SEC * (2 ** attempt)
                delay = min(delay * (1 + random.random() * 0.25), MAX_BACKOFF_SEC)
                if attempt < MAX_RETRIES:
                    print(f"[429] 대기 {delay:.1f}초 후 재시도 ({attempt + 1}/{MAX_RETRIES})...", file=sys.stderr)
                    time.sleep(delay)
                else:
                    raise
            else:
                raise
    raise last_err or RuntimeError("API 호출 실패")

# -----------------------------
# 파일 입출력
# -----------------------------

def _read_input(path: str) -> pd.DataFrame:
    ext = os.path.splitext(path.lower())[1]
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(path, dtype=str)
    elif ext in [".csv", ".txt"]:
        df = pd.read_csv(path, dtype=str, sep=None, engine="python")
    else:
        raise ValueError("지원 확장자: xlsx/xls/csv/txt")
    if df.shape[1] < 2:
        raise ValueError("입력에는 최소 2컬럼(상품ID, 상품명)이 필요합니다.")
    out = df.iloc[:, :2].copy()
    out.columns = ["상품ID", "상품명"]

    # 노이즈/이상치 필터
    out["상품ID"] = out["상품ID"].astype(str).str.strip()
    out["상품명"] = out["상품명"].astype(str).str.strip()
    out = out[out["상품명"].str.len() > 0]
    out = out[~out["상품ID"].str.contains(r"\[|필수|삭제", regex=True)]
    out = out[out["상품ID"].str.len() > 0]
    out = out.reset_index(drop=True)
    return out


def _ensure_out_path(in_path: str) -> str:
    in_dir, in_file = os.path.split(in_path)
    stem, ext = os.path.splitext(in_file)
    out_dir = OUTPUT_DIR if OUTPUT_DIR else in_dir
    os.makedirs(out_dir, exist_ok=True)
    if ext.lower() in [".xlsx", ".xls"]:
        return os.path.join(out_dir, f"{stem}_keywords{ext}")
    else:
        return os.path.join(out_dir, f"{stem}_keywords.csv")


def _write_output(df: pd.DataFrame, out_path: str) -> None:
    ext = os.path.splitext(out_path.lower())[1]
    if ext in [".xlsx", ".xls"]:
        df.to_excel(out_path, index=False)
    else:
        df.to_csv(out_path, index=False, encoding="utf-8-sig")

# -----------------------------
# 단일 파일 처리
# -----------------------------

def _process_one_row(args: Tuple) -> Tuple[int, str, str]:
    """단일 행 처리 (스레드 풀용). (idx, pid, kw) 반환"""
    idx, pid, raw_title, title, base, client, model, cache = args
    if client is not None and USE_MODE == "openai":
        try:
            kw = _llm_keyword(client, model, title, hint=base, cache=cache)
        except Exception as e:
            print(f"[경고] API 오류: {pid} → {e}", file=sys.stderr)
            kw = base
    else:
        kw = base
    kw = re.sub(r"\s+", "", kw)
    return (idx, pid, kw)


def process_one_file(client: Optional["OpenAI"], model: str, in_path: str) -> str:
    print(f"\n[파일시작] {in_path}")
    file_start = time.time()

    df = _read_input(in_path)
    total = len(df)
    cache: Dict[Tuple[str, str], str] = {}

    tasks = []
    for idx, row in enumerate(df.itertuples(index=False), start=1):
        pid = str(row.상품ID)
        raw_title = str(row.상품명)
        title = _clean_name(raw_title)
        tokens = _tokenize(title)
        base = _stitch(tokens) or "대표키워드"
        tasks.append((idx, pid, raw_title, title, base, client, model, cache))

    results: List[Tuple[int, str]] = []
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        futures = {ex.submit(_process_one_row, t): t[0] for t in tasks}
        done = 0
        for f in as_completed(futures):
            try:
                idx, pid, kw = f.result()
                results.append((idx, kw))
                done += 1
                print(f"[진행중] {done}/{total} ({done/total*100:.1f}%) : {pid} → 완료")
            except Exception as e:
                print(f"[경고] 처리 실패: {e}", file=sys.stderr)

    results.sort(key=lambda x: x[0])
    kw_list = [r[1] for r in results]

    out_df = pd.DataFrame({
        "상품ID": df["상품ID"],
        "대표키워드": kw_list,
    })

    out_path = _ensure_out_path(in_path)
    _write_output(out_df, out_path)

    mm, ss = divmod(int(time.time() - file_start), 60)
    print(f"[파일종료] {out_path} | 소요: {mm}분 {ss}초 | 건수: {total}")
    return out_path

# -----------------------------
# 메인 (배치)
# -----------------------------

def main() -> int:
    start_ts = time.time()
    print(f"[배치시작] {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    client = None
    if USE_MODE == "openai":
        if OpenAI is None:
            print("[오류] openai 라이브러리가 설치되어 있지 않습니다. `pip install openai` 후 다시 시도하세요.", file=sys.stderr)
            return 2
        key = (API_KEY or os.environ.get("OPENAI_API_KEY", "")).strip()
        if not key:
            print("[오류] OPENAI_API_KEY가 설정되지 않았습니다. API_KEY 또는 환경변수로 지정하세요.", file=sys.stderr)
            return 2
        client = OpenAI(api_key=key)

    outputs: List[str] = []
    for path in INPUT_PATHS:
        if not path or not os.path.exists(path):
            print(f"[건너뜀] 입력 파일을 찾을 수 없습니다: {path}")
            continue
        try:
            out_path = process_one_file(client, OPENAI_MODEL, path)
            outputs.append(out_path)
        except Exception as e:
            print(f"[에러] 파일 처리 실패: {path} → {e}", file=sys.stderr)

    elapsed = int(time.time() - start_ts)
    mm, ss = divmod(elapsed, 60)
    print(f"[배치종료] {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"[총소요시간] {mm}분 {ss}초 | 총 파일: {len(INPUT_PATHS)} | 성공: {len(outputs)}")
    if outputs:
        print("[생성결과]")
        for p in outputs:
            print(" -", p)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
