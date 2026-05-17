#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
여러 엑셀 파일을 '아래로' 단순 이어 붙여 한 파일로 통합
- 각 파일의 첫 번째 시트만 사용(필요 시 SHEET_NAME 지정)
- 첫 파일의 컬럼 순서를 기준으로, 이후 파일에서 없는 컬럼은 빈칸으로 채움
- 파일마다 컬럼 구성이 달라도 합쳐짐(outer join 효과)
필요: pip install pandas openpyxl xlsxwriter
- 저장 시 xlsxwriter 사용 → 매크로 불가 형식
- '인터넷에서 가져온 파일' 표시 제거(Unblock) → Excel 제한된 보기/자동 xlsm 생성 방지
"""

import os
import re
import shutil
import stat
import sys
import tempfile
import pandas as pd

# ===== 설정 =====
# 1) 합칠 파일들을 '순서대로' 나열 (권장)
INPUT_FILES = [

    # 추가 파일이 있다면 여기에 나열
    # r"C:\Users\USER\Documents\국내위탁\마이박스\251014\5번_할로윈고가_3번째.xlsx",
]

# 2) 또는 폴더에서 자동으로 찾기(입력 목록이 비어있을 때만 사용)
INPUT_DIR = r"C:\Users\USER\Documents\국내위탁\마이박스\26년2월\260218\2번사업자"
GLOB_PATTERN = "*.xlsx"   # 폴더 자동수집 시 필터
USE_NATURAL_SORT = True   # 파일명에 숫자 있으면 사람식 정렬

# 출력 경로 (수정됨) — 반드시 .xlsx로 저장되도록 아래에서 확장자 강제
OUTPUT_XLSX = r"C:\Users\USER\Documents\국내위탁\마이박스\26년2월\260218\2번사업자\2번_통합상품명.xlsx"

# 시트 이름: None이면 첫 시트 사용, 아니면 이름/인덱스 지정 가능
SHEET_NAME = 0  # 예) "Sheet1" 또는 0

# True면 임시 폴더에 먼저 저장 후 최종 경로로 복사 (동일 증상 시 True 권장)
SAVE_VIA_TEMP = True
# 저장 후 '인터넷에서 가져온 파일' 표시 제거 (True 권장)
UNBLOCK_OUTPUT_FILE = True
# 합치기 전에 입력 파일들도 Unblock (2_1, 2_2 등 열 때 경고 안 뜨게 하려면 True)
UNBLOCK_INPUT_FILES = False
# 출력 파일 읽기 전용으로 설정 (True면 Excel에서 '저장' 대신 '다른 이름으로 저장'만 가능 → 형식으로 .xlsx 선택 가능)
MAKE_OUTPUT_READONLY = True
# ==============


def natural_key(s: str):
    """사람이 보는 숫자 정렬용 키"""
    return [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)]


def ensure_xlsx_path(path: str) -> str:
    """출력 경로를 항상 .xlsx 확장자로 통일 (xlsm 등으로 바뀌지 않게)"""
    base, _ = os.path.splitext(path)
    return base + ".xlsx"


def unblock_file_win(path: str) -> bool:
    """Windows: Zone.Identifier 제거 → '인터넷에서 가져온 파일' 표시 제거 (Excel 제한된 보기/xlsm 방지)"""
    if sys.platform != "win32":
        return False
    path = os.path.abspath(path)
    if not os.path.isfile(path):
        return False
    stream_path = path + ":Zone.Identifier"
    try:
        os.remove(stream_path)
        return True
    except FileNotFoundError:
        return False  # 스트림이 없으면 이미 정상
    except OSError:
        return False


def collect_files():
    if INPUT_FILES:
        return INPUT_FILES
    files = [os.path.join(INPUT_DIR, f) for f in os.listdir(INPUT_DIR)
             if f.lower().endswith(".xlsx")]
    if USE_NATURAL_SORT:
        files.sort(key=lambda p: natural_key(os.path.basename(p)))
    else:
        files.sort()
    return files


def main():
    files = collect_files()
    if not files:
        raise SystemExit("[오류] 합칠 엑셀 파일이 없습니다.")

    # 출력은 항상 .xlsx로 저장 (xlsm 방지)
    out_path = ensure_xlsx_path(OUTPUT_XLSX)

    master_cols = None
    frames = []

    for i, fp in enumerate(files, 1):
        if UNBLOCK_INPUT_FILES:
            unblock_file_win(fp)
        print(f"[{i}/{len(files)}] 읽는 중: {fp}")
        df = pd.read_excel(fp, sheet_name=SHEET_NAME, dtype=object)  # 값 보존
        # 첫 파일의 컬럼 순서 기억
        if master_cols is None:
            master_cols = list(df.columns)
        else:
            # 이전에 없던 새 컬럼이 나오면 뒤에 추가
            for c in df.columns:
                if c not in master_cols:
                    master_cols.append(c)
        frames.append(df)

    # 통합: 컬럼 유니온 + 첫 파일 순서 유지
    unified = pd.concat(
        (f.reindex(columns=master_cols) for f in frames),
        ignore_index=True,
        sort=False
    )

    base, _ = os.path.splitext(out_path)
    xlsm_path = base + ".xlsm"
    os.makedirs(os.path.dirname(out_path) or ".", exist_ok=True)
    # 같은 이름의 기존 .xlsm 삭제
    if os.path.isfile(xlsm_path):
        try:
            os.remove(xlsm_path)
            print(f"[삭제] 기존 파일 제거: {xlsm_path}")
        except OSError as e:
            print(f"[경고] 기존 xlsm 삭제 실패 (파일 열려 있을 수 있음): {e}")

    if SAVE_VIA_TEMP:
        # 임시 폴더에 먼저 저장(보안 표시 안 붙는 경로) → 그다음 최종 위치로 복사
        with tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False) as tmp:
            tmp_path = tmp.name
        try:
            with pd.ExcelWriter(tmp_path, engine="xlsxwriter") as xw:
                unified.to_excel(xw, sheet_name="merged", index=False)
            shutil.copy2(tmp_path, out_path)
            print(f"[복사] 임시 → 최종: {out_path}")
        finally:
            try:
                os.remove(tmp_path)
            except OSError:
                pass
    else:
        with pd.ExcelWriter(out_path, engine="xlsxwriter") as xw:
            unified.to_excel(xw, sheet_name="merged", index=False)

    # Zone.Identifier 제거 → '인터넷에서 가져온 파일' 제거 (제한된 보기/xlsm 자동 생성 방지)
    if UNBLOCK_OUTPUT_FILE and unblock_file_win(out_path):
        print(f"[Unblock] 출력 파일 보안 표시 제거됨")

    # 읽기 전용으로 설정 → Excel에서 '저장' 시 덮어쓰기 안 되고 '다른 이름으로 저장'만 가능 (형식에서 .xlsx 선택 유도)
    if MAKE_OUTPUT_READONLY:
        try:
            os.chmod(out_path, stat.S_IREAD)
            print(f"[읽기 전용] 출력 파일이 읽기 전용으로 설정되었습니다.")
            print("  → Excel에서 수정 후 저장 시: [다른 이름으로 저장] → 파일 형식에서 'Excel 통합문서 (*.xlsx)' 선택")
        except OSError:
            pass

    print(f"[완료] {len(files)}개 파일 통합 → {out_path}")
    print(f" - 총 행수: {len(unified)}")
    print(f" - 컬럼수: {len(unified.columns)}")


if __name__ == "__main__":
    main()
