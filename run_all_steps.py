#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
STEP1(xls변환) ~ STEP6(최종합치기) 한 번에 순차 실행

사용법:
  python run_all_steps.py                    # 전체 실행
  python run_all_steps.py --only 9           # STEP6 최종합치기만 실행 (해당 부분만)
  STEP1에서 Excel이 멈추면: STEP1_EXCEL_VISIBLE=1 (Excel 창 표시·원인 확인)

실행 순서 (폴더명이 "N번사업자"이면 결과물은 모두 "N번_..." 로 생성):
  1) STEP1.xls변환    - 폴더 내 .xls → .xlsx
  2) STEP1-1         - 엑셀 합치기 → {N번}_통합상품명.xlsx
  3) STEP2           - 이미지 변환 → {N번}_통합상품명_이미지 폴더
  4) STEP3           - 대표키워드 → {N번}_통합상품명_keywords.xlsx
  5) STEP4           - 쭌쭌쌤 상품명 → {N번}_통합상품명_keywords_쭌쭌쌤.xlsx
  6) STEP5           - 네이버클라우드 이미지 업로드
  7) STEP5-1         - 링크 목록 → {N번}_domeme_links.xlsx
  8) STEP6 카테고리  - 카테고리 매칭 → {N번}_통합상품명_카테고리매핑_결과.xlsx
  9) STEP6 최종합치기 - 최종 합침 → {N번}_통합상품명_최종.xlsx
"""

import argparse
import os
import sys
import re
import asyncio
import importlib.util
from pathlib import Path

# cp949 등 콘솔에서 한글·치환문자 출력 시 print() 크래시 방지
for _sn in ("stdout", "stderr"):
    try:
        getattr(sys, _sn).reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass

# .env 로드: STEP 모듈을 importlib 로 로드하기 전에 환경변수를 채워둔다.
try:
    from dotenv import load_dotenv

    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

# ============== 여기만 수정하세요 (run_all_steps.py 단독 실행 시 사용) ==============
BASE_DIR = r"C:\Users\USER\Documents\국내위탁\마이박스\26년3월2주차\5#ll번사업자"
# ==============================================

# 폴더명(예: 1번사업자, 2번사업자)에서 "N번" 추출 → 결과물 파일명에 사용 (단독 실행 시)
_folder_name = os.path.basename(BASE_DIR.rstrip(os.sep))
_match = re.match(r"(\d+)번", _folder_name)
PREFIX = f"{_match.group(1)}번" if _match else "1번"

# STEP9 최종 파일명: domeme에서 호출 시 엑셀생성규칙_최종 적용 ({user_id}_{키워드}_{주차}_{회차}회_최종.xlsx)
FINAL_OUTPUT_FILENAME = None

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))


def load_module(name: str, filepath: str):
    """파일 경로로 모듈 로드"""
    spec = importlib.util.spec_from_file_location(name, filepath)
    mod = importlib.util.module_from_spec(spec)
    sys.modules[name] = mod
    spec.loader.exec_module(mod)
    return mod


def run_step1_xls():
    """STEP1: .xls → .xlsx 변환"""
    path = os.path.join(PROJECT_DIR, "STEP1.xls변환.py")
    print("[run_all] STEP1 모듈 로드…", flush=True)
    mod = load_module("step1_xls", path)
    print("[run_all] STEP1 변환 시작 (Excel COM, 멈추면 STEP1_EXCEL_VISIBLE=1 로 재시도)", flush=True)
    mod.convert_xls_to_xlsx(BASE_DIR)
    print("[run_all] STEP1 xls변환 완료\n", flush=True)


def run_step1_1():
    """STEP1-1: 엑셀 파일 합치기 → {PREFIX}_통합상품명.xlsx"""
    path = os.path.join(PROJECT_DIR, "STEP1-1엑셀파일합치기.py")
    mod = load_module("step1_1", path)
    mod.INPUT_DIR = BASE_DIR
    mod.OUTPUT_XLSX = os.path.join(BASE_DIR, f"{PREFIX}_통합상품명.xlsx")
    mod.INPUT_FILES = []
    mod.main()
    print("[run_all] STEP1-1 엑셀합치기 완료\n")


def run_step2():
    """STEP2: 이미지 다운로드/변환"""
    path = os.path.join(PROJECT_DIR, "STEP2.이미지변환_2nd.py")
    mod = load_module("step2", path)
    mod.BASE_DIR = BASE_DIR
    mod.GLOB_PATTERN = f"{PREFIX}_통합상품명.xlsx"
    mod.main()
    print("[run_all] STEP2 이미지변환 완료\n")


def run_step3():
    """STEP3: 대표키워드(OpenAPI) → {PREFIX}_통합상품명_keywords.xlsx"""
    path = os.path.join(PROJECT_DIR, "STEP3.openapi이용 대표상품명작성(단어1개).py")
    mod = load_module("step3", path)
    mod.INPUT_PATHS = [os.path.join(BASE_DIR, f"{PREFIX}_통합상품명.xlsx")]
    mod.OUTPUT_DIR = ""
    exit_code = mod.main()
    if exit_code != 0:
        raise SystemExit(exit_code)
    print("[run_all] STEP3 대표키워드 완료\n")


def run_step4():
    """STEP4: 쭌쭌쌤 상품명(OpenAPI) → {PREFIX}_통합상품명_keywords_쭌쭌쌤.xlsx"""
    path = os.path.join(PROJECT_DIR, "STEP4.openapi이용 상품명작성(쭌쭌쌤).py")
    mod = load_module("step4", path)
    mod.INPUT_PATHS = [os.path.join(BASE_DIR, f"{PREFIX}_통합상품명_keywords.xlsx")]
    mod.OUTPUT_DIR = ""
    exit_code = mod.main()
    if exit_code != 0:
        raise SystemExit(exit_code)
    print("[run_all] STEP4 쭌쭌쌤 상품명 완료\n")


def run_step5():
    """STEP5: 네이버클라우드 이미지 업로드"""
    path = os.path.join(PROJECT_DIR, "STEP5.네이버클라우드이미지업로드.py")
    mod = load_module("step5", path)
    mod.LOCAL_DIR = os.path.join(BASE_DIR, f"{PREFIX}_통합상품명_이미지")
    mod.main()
    print("[run_all] STEP5 네이버클라우드 업로드 완료\n")


def run_step5_1():
    """STEP5-1: 버킷 목록 → {PREFIX}_domeme_links.xlsx"""
    path = os.path.join(PROJECT_DIR, "STEP5-1.네이버이미지Link추출.py")
    mod = load_module("step5_1", path)
    mod.OUTPUT_XLSX = os.path.join(BASE_DIR, f"{PREFIX}_domeme_links.xlsx")
    mod.main()
    print("[run_all] STEP5-1 링크추출 완료\n")


def run_step6_category():
    """STEP6: 네이버 카테고리 매칭 (Playwright 등 기존 이벤트 루프와 충돌 방지: 별도 스레드에서 asyncio.run)"""
    path = os.path.join(PROJECT_DIR, "STEP6.네이버카테고리매칭.py")
    mod = load_module("step6_cat", path)
    mod.PRODUCTS_XLSX = os.path.join(BASE_DIR, f"{PREFIX}_통합상품명_keywords.xlsx")
    mod.OUTPUT_XLSX = os.path.join(BASE_DIR, f"{PREFIX}_통합상품명_카테고리매핑_결과.xlsx")

    import concurrent.futures
    def _run_async():
        asyncio.run(mod.async_main())
    try:
        loop = asyncio.get_running_loop()
    except RuntimeError:
        loop = None
    if loop is not None:
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as pool:
            pool.submit(_run_async).result()
    else:
        asyncio.run(mod.async_main())
    print("[run_all] STEP6 카테고리매칭 완료\n")


def run_step6_merge():
    """STEP6: 상품명/이미지/카테고리 최종 합치기. 엑셀생성규칙_최종 적용 시 {user_id}_{키워드}_{주차}_{회차}회_최종.xlsx"""
    path = os.path.join(PROJECT_DIR, "STEP6.상품명,이미지링크최종합치기.py")
    mod = load_module("step6_merge", path)
    mod.BASE = Path(BASE_DIR)
    mod.FILE1 = mod.BASE / f"{PREFIX}_통합상품명.xlsx"
    mod.FILE2 = mod.BASE / f"{PREFIX}_통합상품명_keywords_쭌쭌쌤.xlsx"
    mod.FILE3 = mod.BASE / f"{PREFIX}_domeme_links.xlsx"
    mod.FILE4 = mod.BASE / f"{PREFIX}_통합상품명_카테고리매핑_결과.xlsx"
    out_name = FINAL_OUTPUT_FILENAME if FINAL_OUTPUT_FILENAME else f"{PREFIX}_통합상품명_최종.xlsx"
    mod.OUTPUT = mod.BASE / out_name
    mod.main()
    print("[run_all] STEP6 최종합치기 완료\n")


STEPS = [
    ("STEP1 xls변환", run_step1_xls),
    ("STEP1-1 엑셀합치기", run_step1_1),
    ("STEP2 이미지변환", run_step2),
    ("STEP3 대표키워드", run_step3),
    ("STEP4 쭌쭌쌤 상품명", run_step4),
    ("STEP5 네이버클라우드 업로드", run_step5),
    ("STEP5-1 링크추출", run_step5_1),
    ("STEP6 카테고리매칭", run_step6_category),
    ("STEP6 최종합치기", run_step6_merge),
]


def run_all_steps_for_dir(base_dir, prefix, only_step=None, final_output_name=None):
    """
    지정한 사업자 폴더에 대해 STEP1 ~ STEP6(최종)까지 순차 실행.
    domeme_auto_login_temp.py에서 엑셀 다운로드 후 호출용.

    :param base_dir: 사업자 폴더 경로 (예: .../26년3월3주차/3회차/1번사업자)
    :param prefix: 파일명 접두사 (예: "1번")
    :param only_step: 1~9 중 하나만 실행하려면 지정. None이면 전체.
    :param final_output_name: STEP9 최종 파일명 (엑셀생성규칙_최종). 예: "r6_런닝화_26년5월3주차_3회_최종.xlsx". None이면 {PREFIX}_통합상품명_최종.xlsx
    """
    global BASE_DIR, PREFIX, FINAL_OUTPUT_FILENAME
    base_dir = Path(base_dir).resolve()
    BASE_DIR = str(base_dir)
    PREFIX = str(prefix)
    FINAL_OUTPUT_FILENAME = final_output_name

    if not os.path.isdir(BASE_DIR):
        print(f"[run_all_steps] 폴더 없음, 건너뜀: {BASE_DIR}")
        return

    print("=" * 60)
    print(f"STEP1~6 통합 실행: {PREFIX} 사업자")
    print(f"BASE_DIR: {BASE_DIR}")
    if only_step:
        print(f"※ {only_step}번째 단계만 실행")
    print("=" * 60)

    to_run = [(i, label, fn) for i, (label, fn) in enumerate(STEPS, 1)]
    if only_step:
        to_run = [(i, label, fn) for i, label, fn in to_run if i == only_step]

    for i, label, fn in to_run:
        print(f"\n>>> [{i}/9] {label}")
        try:
            fn()
        except Exception as e:
            print(f"[실패] {label}: {e}")
            raise

    # === 최종 완료 후 _통합상품명_이미지 폴더 내 파일 전부 완전삭제 ===
    img_dir = Path(BASE_DIR) / f"{PREFIX}_통합상품명_이미지"
    if img_dir.is_dir():
        cnt = 0
        for f in img_dir.rglob("*"):
            if f.is_file():
                try:
                    f.unlink()
                    cnt += 1
                except OSError as e:
                    print(f"[이미지삭제] {f.name} 실패: {e}")
        if cnt > 0:
            print(f"[이미지삭제] {img_dir.name} 내 {cnt}개 파일 완전삭제 완료")

    print("=" * 60)
    print(f"{PREFIX} 사업자 처리 완료.")
    print("=" * 60)


def main():
    print(f"실행 경로: {Path(__file__).resolve()}")
    parser = argparse.ArgumentParser(description="STEP1~STEP6 통합 실행 (특정 단계만 실행 가능)")
    parser.add_argument("--only", type=int, choices=range(1, 10), metavar="N",
                        help="N번째 단계만 실행 (1~9). 예: --only 9 → STEP6 최종합치기만")
    args = parser.parse_args()

    print("=" * 60)
    print("STEP1 ~ STEP6 통합 실행")
    print(f"BASE_DIR: {BASE_DIR}")
    print(f"사업자 번호(파일명 접두사): {PREFIX}")
    if args.only:
        print(f"※ {args.only}번째 단계만 실행")
    print("=" * 60)

    if not os.path.isdir(BASE_DIR):
        print(f"[오류] 폴더가 없습니다: {BASE_DIR}")
        sys.exit(1)

    to_run = [(i, label, fn) for i, (label, fn) in enumerate(STEPS, 1)]
    if args.only:
        to_run = [(i, label, fn) for i, label, fn in to_run if i == args.only]

    for i, label, fn in to_run:
        print(f"\n>>> [{i}/9] {label}")
        try:
            fn()
        except Exception as e:
            print(f"[실패] {label}: {e}")
            raise

    print("=" * 60)
    print("완료.")
    print("=" * 60)


if __name__ == "__main__":
    main()
