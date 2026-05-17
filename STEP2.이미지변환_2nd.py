#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import re
import glob
import math
import time
import json
import random
import pandas as pd
import requests
import numpy as np
from io import BytesIO
from PIL import Image
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

# ============================
# 사용자 설정
# ============================
# 1) 단일 파일만 처리하려면 SINGLE_XLSX에 경로를 넣고, BASE_DIR는 무시됨.
SINGLE_XLSX = None
# SINGLE_XLSX = r"C:\Users\USER\Documents\국내위탁\마이박스\251004\3번사업자\3번_엠에스유통협력사_1005_2번째.xlsx"

# 2) 폴더 전체 처리 (재귀). SINGLE_XLSX가 None일 때만 사용.
BASE_DIR        = r"C:\Users\USER\Documents\국내위탁\마이박스\26년1월\1번사업자"
GLOB_PATTERN    = "1번_통합상품명.xlsx"            # 하위폴더까지 전부: "**/*.xlsx" / 현재폴더만: "*.xlsx"
IGNORE_PREFIXES = ("~$",)                # 임시/잠금 파일 무시
REQUIRE_COLS    = ("도매매 상품번호", "대표이미지링크")

# 3) 출력 폴더 규칙: 각 엑셀파일 옆에 [파일명_이미지] 폴더 생성 (예: 3번_엠에스...xlsx → 3번_엠에스..._이미지)
FOLDER_SUFFIX   = "_이미지"

# 4) 처리 옵션
MAX_WORKERS         = 10        # 병렬 갯수
ADD_NOISE           = False     # 노이즈 추가 여부(투명 유지 안 함, RGB에서만)
NOISE_STD           = 5.0       # 가우시안 노이즈 표준편차 (0~10 정도 권장)
ROTATE_DEGREE       = -7        # 시계방향(+), 반시계(-)
SKIP_IF_EXISTS      = True      # 파일이 이미 있으면 건너뛰기

# 5) 저장 형식
SAVE_FORMAT         = "PNG"     # PNG 권장(투명 유지)
SAVE_EXT            = ".png"    # .png / .jpg

# 6) 네트워크 설정
TIMEOUT_SEC         = 15
RETRY               = 2
USER_AGENT          = "Mozilla/5.0 (compatible; ImageWorker/1.0; +https://example.local)"

# ============================
# 공통 유틸
# ============================
def safe_stem(path: str) -> str:
    """파일명 스템(확장자 제거)에서 폴더명으로 안전하게 변경"""
    stem = os.path.splitext(os.path.basename(path))[0]
    stem = re.sub(r"[\\/:*?\"<>|]", "_", stem)  # 금지문자 -> _
    stem = stem.strip()
    return stem or "images"

def ensure_dir(path: str):
    if path and not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def guess_ext_from_url(url: str) -> str:
    p = urlparse(url)
    name = os.path.basename(p.path).lower()
    for ext in (".png",".jpg",".jpeg",".webp",".gif",".bmp"):
        if name.endswith(ext):
            return ext
    return SAVE_EXT

def to_rgb_if_needed(img: Image.Image) -> Image.Image:
    """JPEG로 저장해야 한다면 RGB 필요"""
    if SAVE_EXT.lower() in (".jpg", ".jpeg"):
        return img.convert("RGB")
    return img

# ============================
# 이미지 가공
# ============================
def add_noise_rgb(img: Image.Image, std: float = 5.0) -> Image.Image:
    """
    간단 가우시안 노이즈 (RGB 전용). std=0이면 영향 없음.
    투명 유지가 필요하면 PNG+RGBA로 두고 노이즈는 생략하는 것을 권장.
    """
    if std <= 0:
        return img
    if img.mode != "RGB":
        img = img.convert("RGB")
    arr = np.array(img).astype(np.float32)
    noise = np.random.normal(0.0, std, arr.shape).astype(np.float32)
    arr = np.clip(arr + noise, 0, 255).astype(np.uint8)
    return Image.fromarray(arr, mode="RGB")

def transform_image(img: Image.Image) -> Image.Image:
    """
    회전(투명 유지) + (옵션)노이즈
    - PNG로 저장 시 RGBA로 유지
    - JPEG로 저장 시 RGB 변환
    """
    # 1) RGBA로
    img = img.convert("RGBA")

    # 2) 회전
    rotated = img.rotate(ROTATE_DEGREE, expand=True)

    # 3) 저장 형식별 변환/노이즈
    if SAVE_EXT.lower() in (".jpg", ".jpeg"):
        # JPEG: 투명 불가 → 흰 배경 합성 후 RGB
        bg = Image.new("RGB", rotated.size, (255, 255, 255))
        bg.paste(rotated, mask=rotated.split()[-1])  # alpha 합성
        out = bg
        if ADD_NOISE:
            out = add_noise_rgb(out, NOISE_STD)
        return out
    else:
        # PNG: 투명 유지(RGBA), 노이즈는 생략(투명 채널 망가질 수 있음)
        return rotated

# ============================
# 다운로드 + 처리
# ============================
def make_session() -> requests.Session:
    s = requests.Session()
    s.headers.update({"User-Agent": USER_AGENT})
    adapter = requests.adapters.HTTPAdapter(pool_connections=MAX_WORKERS, pool_maxsize=MAX_WORKERS, max_retries=0)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    return s

def download_transform_save(session: requests.Session, url: str, save_path: str) -> bool:
    try:
        # 이미 존재하면 스킵
        if SKIP_IF_EXISTS and os.path.exists(save_path):
            return True

        # 다운로드 + 간단 재시도
        last_exc = None
        for attempt in range(RETRY + 1):
            try:
                resp = session.get(url, timeout=TIMEOUT_SEC)
                resp.raise_for_status()
                img = Image.open(BytesIO(resp.content))
                break
            except Exception as e:
                last_exc = e
                time.sleep(0.4 * (attempt + 1))
        else:
            raise last_exc or RuntimeError("download failed")

        # 변환
        out = transform_image(img)

        # 저장 폴더 보장
        ensure_dir(os.path.dirname(save_path))

        # 저장
        fmt = SAVE_FORMAT
        if SAVE_EXT.lower() in (".jpg",".jpeg"):
            fmt = "JPEG"
        out.save(save_path, format=fmt)

        # 정리
        try:
            img.close()
            out.close()
        except Exception:
            pass
        return True

    except Exception as e:
        print(f"[이미지 실패] {url} → {e}")
        return False

# ============================
# 엑셀 처리
# ============================
def process_excel(xlsx_path: str):
    try:
        df = pd.read_excel(xlsx_path)
    except Exception as e:
        print(f"[엑셀 로드 실패] {xlsx_path} → {e}")
        return

    for col in REQUIRE_COLS:
        if col not in df.columns:
            print(f"[건너뜀] '{xlsx_path}' : 필수 컬럼 누락 → {REQUIRE_COLS}")
            return

    # 파일별 전용 폴더: [파일명_이미지]
    base_dir = os.path.dirname(xlsx_path)
    folder_name = safe_stem(xlsx_path) + FOLDER_SUFFIX
    image_dir = os.path.join(base_dir, folder_name)
    ensure_dir(image_dir)

    session = make_session()

    tasks = []
    with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
        for _, row in df.iterrows():
            pid = str(row["도매매 상품번호"]).strip()
            url = str(row["대표이미지링크"]).strip()
            if not pid or not url or url.lower() == "nan":
                continue
            # 저장 경로: [전용폴더]/상품번호.png
            save_path = os.path.join(image_dir, f"{pid}{SAVE_EXT}")
            tasks.append(ex.submit(download_transform_save, session, url, save_path))

        done, total = 0, len(tasks)
        for fut in as_completed(tasks):
            ok = fut.result()
            done += 1
            if done % 25 == 0 or done == total:
                print(f"  - 진행 {done}/{total} (in {folder_name})")

    print(f"[완료] {os.path.basename(xlsx_path)} → 저장폴더: {image_dir}")

# ============================
# 메인
# ============================
def main():
    targets = []
    if SINGLE_XLSX:
        if os.path.isfile(SINGLE_XLSX):
            targets = [SINGLE_XLSX]
        else:
            print(f"[오류] SINGLE_XLSX 경로가 올바르지 않습니다: {SINGLE_XLSX}")
            return
    else:
        pattern = os.path.join(BASE_DIR, GLOB_PATTERN)
        for path in glob.glob(pattern, recursive=True):
            name = os.path.basename(path)
            if any(name.startswith(pfx) for pfx in IGNORE_PREFIXES):
                continue
            if not name.lower().endswith(".xlsx"):
                continue
            targets.append(path)

    if not targets:
        print("[알림] 처리할 엑셀 파일이 없습니다.")
        return

    print(f"[대상 파일 수] {len(targets)}")
    for x in targets:
        print(" -", x)

    for xlsx in targets:
        process_excel(xlsx)

if __name__ == "__main__":
    # 재현성(노이즈)
    random.seed(42)
    np.random.seed(42)
    main()
