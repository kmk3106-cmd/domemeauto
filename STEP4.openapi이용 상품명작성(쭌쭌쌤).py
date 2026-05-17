#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
쭌쭌쌤_상품명_생성기_batch.py

여러 개 입력 파일을 순차 처리하여 각각 결과 파일을 생성합니다.
- 각 입력: 1행 헤더, 2행부터 데이터 / A열=상품ID, B열=대표키워드(1개 단어)
- 각 출력: [상품ID, 대표키워드, 쿠팡상품명, 검수결과, 위반규칙, 단어수, 바이트수]
- 규칙: 쭌쭌쌤 노출 극대화 규칙 + 검수/보정 + 80바이트 제한 + 7단어 이상
- 진행 로그: 작업 시작/종료 시각, 총 소요시간, per-file, per-row 진행률

사용법
1) 아래 INPUT_PATHS 리스트에 처리할 파일 경로들을 넣으세요.
2) OUTPUT_DIR를 지정하면 그 폴더에 저장, 비우면 입력 파일과 같은 폴더에 저장.
3) 실행: python 쭌쭌쌤_상품명_생성기_batch.py
"""
from __future__ import annotations
import os
import re
import sys
import time
import datetime
import random
from typing import List, Tuple, Dict, Optional
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd

# -----------------------------
# 하드코딩: 여러 입력 파일 경로들
# -----------------------------
INPUT_PATHS = [
    r"C:\Users\USER\Documents\국내위탁\마이박스\26년1월\1번사업자\통합 문서1.xlsx",


#
    # r"C:\\Users\\USER\\Documents\\국내위탁\\마이박스\\250924\\대표키워드.xlsx",
]

# 출력 폴더(선택): 비우면 입력과 같은 폴더에 저장
OUTPUT_DIR = r""  # 예: r"C:\\Users\\USER\\Documents\\국내위탁\\결과"

OPENAI_MODEL = "gpt-4o-mini"
API_KEY = os.environ.get("OPENAI_API_KEY", "")  # .env 의 OPENAI_API_KEY

# 429 대응: 스로틀링 + 지수 백오프
MAX_RETRIES = 6
BASE_BACKOFF_SEC = 2.0
MAX_BACKOFF_SEC = 120
MAX_CONCURRENCY = 4
REQUEST_DELAY_SEC = 0.15

# -----------------------------
# OpenAI SDK
# -----------------------------
try:
    from openai import OpenAI  # type: ignore
except Exception:
    OpenAI = None  # type: ignore

# -----------------------------
# 규칙 프롬프트(시스템)
# -----------------------------
SYSTEM_PROMPT = (
    """
당신은 한국 쇼핑몰 쿠팡 상품명 작성 전문가입니다. 
아래 규칙을 반드시 지켜서 대표키워드를 확장된 상품명으로 바꾸세요.

[쭌쭌쌤 쿠팡 키워드 만들기 규칙]
1) 맨 앞 단어는 반드시 대표키워드와 동일하게 시작한다.
2) 단어 간 공백은 1칸만 사용한다. 탭/콤마 금지.
3) 형용사, 유사단어, 검색 키워드를 자연스럽게 삽입하되 문장처럼 만들지 않는다.
4) 중복 단어 금지. 맨 앞 단어(대표키워드)는 1회만 등장.
5) 길이는 최대 80바이트(UTF-8) 이내로 작성(가능하면 꽉 채움).
6) 마지막 단어는 반드시 명사.
7) 7~8 단어 이상 풍성하게 표현.
8) 브랜드/지재권 의심 단어 절대 금지.
9) 갯수표현금지 (10개, 5개입 등)
10) 쿠팡 스타일(용도 + 기능 + 사용처 + 대상 + 형태 등)을 참고.
출력은 오직 최종 상품명 문자열만 반환하라.
    """
).strip()

# -----------------------------
# 금지어(간단 필터)
# -----------------------------
BANNED_WORDS = {
    "샤넬","루이비통","나이키","아디다스","뉴발란스","구찌","프라다","에르메스","애플","아이폰",
    "갤럭시","닌텐도","플스","디즈니","마블","스파이더맨","짭","레플리카","정품아님","짝퉁","유니섹스","친환경","링티",
    "레이벤",

}

SENTENCE_ENDERS = re.compile(r"[.!?]+|。|！|？")
MULTISPACE = re.compile(r"\s+")

# -----------------------------
# 응답 파싱(Responses API 호환)
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

# -----------------------------
# 검수/보정 유틸
# -----------------------------

def _looks_noun(word: str) -> bool:
    return not re.search(r"(을|를|이|가|은|는|에|에서|으로|하게|적인|스러운|합니다|하세요|하다|하는|한)$", word)


def _dedup_keep_order(words: List[str]) -> List[str]:
    seen = set()
    out: List[str] = []
    for w in words:
        if w and w not in seen:
            seen.add(w)
            out.append(w)
    return out


def _byte_len(s: str) -> int:
    return len(s.encode("utf-8"))


def _is_word_list_not_sentence(text: str) -> bool:
    if SENTENCE_ENDERS.search(text):
        return False
    if "," in text or "\t" in text:
        return False
    return True


def _validate_title(keyword: str, title: str):
    words = [w for w in title.split(" ") if w]
    issues: Dict[str, bool] = {
        "시작키워드": title.startswith(keyword + " ") or title == keyword,
        "단어나열": _is_word_list_not_sentence(title),
        "콤마탭금지": ("," not in title and "\t" not in title),
        "공백1칸": ("  " not in title and "\t" not in title),
        "중복없음": (len(words) == len(set(words))),
        "마지막명사": _looks_noun(words[-1]) if words else False,
        "바이트80이내": (_byte_len(title) <= 150),
        "최소7단어": (len(words) >= 7),
    }
    ok = all(issues.values())
    metrics = {"단어수": len(words), "바이트수": _byte_len(title)}
    return ok, issues, metrics


def _sanitize_and_enforce(keyword: str, title: str) -> str:
    t = title.replace(",", " ")
    t = MULTISPACE.sub(" ", t).strip()
    if not t.startswith(keyword + " ") and t != keyword:
        t = f"{keyword} {t}"
    words = [w for w in t.split(" ") if w and w not in BANNED_WORDS]
    words = [words[0]] + [w for w in words[1:] if w != keyword] if words else [keyword]
    words = _dedup_keep_order(words)
    filler = ["차량용","가정용","휴대용","간편설치","다용도","강화내구성","안전","편리","거치대","도구","용품"]
    i = 0
    while len(words) < 7 and i < len(filler):
        if filler[i] not in words:
            words.append(filler[i])
        i += 1
    if not _looks_noun(words[-1]):
        for cand in ["용품","도구","거치대","장갑","가방","세트","보관함","케이스"]:
            if cand not in words:
                words[-1] = cand
                break
    out = " ".join(words)
    while _byte_len(out) > 150 and len(words) > 1:
        words.pop()
        out = " ".join(words)
    return out


def _fallback_title(keyword: str) -> str:
    base = [keyword, "차량용", "긴급", "탈출", "안전", "다용도", "휴대용", "도구"]
    out = " ".join(base)
    while _byte_len(out) > 150 and len(base) > 1:
        base.pop()
        out = " ".join(base)
    return out

# -----------------------------
# 파일 입출력
# -----------------------------

def _read_input(input_path: str) -> pd.DataFrame:
    ext = os.path.splitext(input_path.lower())[1]
    if ext in [".xlsx", ".xls"]:
        df = pd.read_excel(input_path, dtype=str)
    elif ext in [".csv", ".txt"]:
        df = pd.read_csv(input_path, dtype=str, sep=None, engine="python")
    else:
        raise ValueError("지원 확장자: xlsx/xls/csv/txt")
    if df.shape[1] < 2:
        raise ValueError("입력에는 최소 2컬럼(상품ID, 대표상품키워드)이 필요합니다.")
    out = df.iloc[:, :2].copy()  # A,B 컬럼만 사용
    out.columns = ["상품ID", "대표키워드"]
    return out


def _ensure_output_path(input_path: str) -> str:
    in_dir, in_file = os.path.split(input_path)
    stem, ext = os.path.splitext(in_file)
    out_dir = OUTPUT_DIR if OUTPUT_DIR else in_dir
    os.makedirs(out_dir, exist_ok=True)
    if ext.lower() in [".xlsx", ".xls"]:
        return os.path.join(out_dir, f"{stem}_쭌쭌쌤{ext}")
    else:
        return os.path.join(out_dir, f"{stem}_쭌쭌쌤.csv")


def _write_output(df: pd.DataFrame, output_path: str) -> None:
    ext = os.path.splitext(output_path.lower())[1]
    if ext in [".xlsx", ".xls"]:
        df.to_excel(output_path, index=False)
    else:
        df.to_csv(output_path, index=False, encoding="utf-8-sig")

# -----------------------------
# 429 백오프 + 캐시 유틸
# -----------------------------

def _parse_retry_after(err_msg: str) -> Optional[float]:
    m = re.search(r"try again in ([\d.]+)s?", err_msg, re.I)
    if m:
        return float(m.group(1)) + 0.5
    return None


def _call_llm_raw(client: "OpenAI", model: str, keyword: str) -> str:
    user = f"대표키워드: {keyword}\n위 규칙을 적용하여 쿠팡 상품명을 작성하세요."
    resp = client.responses.create(
        model=model,
        input=[
            {"role": "system", "content": SYSTEM_PROMPT},
            {"role": "user", "content": user},
        ],
    )
    return _extract_text_from_response(resp).strip()


def _call_llm(client: "OpenAI", model: str, keyword: str, cache: Optional[Dict[str, str]] = None) -> str:
    """지수 백오프 + 캐시 적용"""
    if cache is not None and keyword in cache:
        return cache[keyword]
    last_err = None
    for attempt in range(MAX_RETRIES + 1):
        try:
            time.sleep(REQUEST_DELAY_SEC)
            raw = _call_llm_raw(client, model, keyword)
            if cache is not None:
                cache[keyword] = raw
            return raw
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
# 단일 행 처리 (스레드 풀용)
# -----------------------------

def _process_one_row(args: Tuple) -> Tuple[int, Dict]:
    idx, pid, keyword, client, model, cache = args
    if not keyword:
        return (idx, {
            "상품ID": pid,
            "대표키워드": keyword,
            "쿠팡상품명": "생성실패",
            "검수결과": "FAIL",
            "위반규칙": "대표키워드없음",
            "단어수": 0,
            "바이트수": 0,
        })
    try:
        raw = _call_llm(client, model, keyword, cache=cache)
    except Exception as e:
        print(f"[경고] API 오류: {pid}/{keyword} → {e}", file=sys.stderr)
        raw = ""
    title1 = raw if raw else _fallback_title(keyword)
    title2 = _sanitize_and_enforce(keyword, title1)
    ok, issues, metrics = _validate_title(keyword, title2)
    if not ok:
        title3 = _sanitize_and_enforce(keyword, title2)
        ok2, issues2, metrics2 = _validate_title(keyword, title3)
        if ok2:
            title2, ok, issues, metrics = title3, ok2, issues2, metrics2
        else:
            title2 = _fallback_title(keyword)
            ok, issues, metrics = _validate_title(keyword, title2)
    final = title2
    return (idx, {
        "상품ID": pid,
        "대표키워드": keyword,
        "쿠팡상품명": final,
        "검수결과": "OK" if ok else "FAIL",
        "위반규칙": "" if ok else ",".join([k for k, v in issues.items() if not v]),
        "단어수": metrics.get("단어수", 0),
        "바이트수": metrics.get("바이트수", 0),
    })


# -----------------------------
# 단일 파일 처리
# -----------------------------

def process_one_file(client: "OpenAI", model: str, input_path: str) -> str:
    print(f"\n[파일시작] {input_path}")
    file_start = time.time()
    df = _read_input(input_path)
    total = len(df)
    cache: Dict[str, str] = {}

    tasks = [(idx, str(row.상품ID).strip(), str(row.대표키워드).strip(), client, model, cache)
             for idx, row in enumerate(df.itertuples(index=False), start=1)]

    rows: List[Tuple[int, Dict]] = []
    with ThreadPoolExecutor(max_workers=MAX_CONCURRENCY) as ex:
        futures = {ex.submit(_process_one_row, t): t[0] for t in tasks}
        done = 0
        for f in as_completed(futures):
            try:
                idx, row_data = f.result()
                rows.append((idx, row_data))
                done += 1
                print(f"[진행중] {done}/{total} ({done/total*100:.1f}%) : {row_data['상품ID']} → 완료")
            except Exception as e:
                print(f"[경고] 처리 실패: {e}", file=sys.stderr)

    rows.sort(key=lambda x: x[0])
    row_dicts = [r[1] for r in rows]

    out_df = pd.DataFrame(row_dicts)
    out_path = _ensure_output_path(input_path)
    _write_output(out_df, out_path)

    mm, ss = divmod(int(time.time() - file_start), 60)
    print(f"[파일종료] {out_path} | 소요: {mm}분 {ss}초 | 건수: {total}")
    return out_path

# -----------------------------
# 메인 (배치 실행)
# -----------------------------

def main() -> int:
    start_ts = time.time()
    print(f"[배치시작] {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    if OpenAI is None:
        print("[오류] openai 라이브러리가 설치되어 있지 않습니다. `pip install openai` 후 다시 시도하세요.", file=sys.stderr)
        return 2

    key = (API_KEY or os.environ.get("OPENAI_API_KEY", "")).strip()
    if not key:
        print("[오류] OPENAI_API_KEY가 설정되지 않았습니다. API_KEY 또는 환경변수로 지정하세요.", file=sys.stderr)
        return 2

    client = OpenAI(api_key=key)

    outputs: List[str] = []
    for i, path in enumerate(INPUT_PATHS, start=1):
        if not path:
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
