#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
NCP Object Storage: 버킷 목록 → [이름, 폴더경로, 최상위폴더, 폴더깊이, LINK, Size, LastModified(KST)] 엑셀 저장
- 버킷이 전체공개면 presigned 불필요 → 공개 직접 URL 사용
- 폴더(의사 디렉터리) 인식하여 경로/최상위폴더/깊이 컬럼 분리
- Excel이 tz-aware datetime을 못 받아서 LastModified는 KST로 변환 후 tz 제거
필요: pip install boto3 pandas openpyxl tzdata  (Windows에서 zoneinfo용 tzdata 권장)
"""

import os
import pandas as pd
import boto3
from urllib.parse import quote
from botocore.config import Config
from botocore.exceptions import ClientError

# 표준 라이브러리 zoneinfo (Python 3.9+). Windows에선 tzdata 패키지 필요할 수 있음.
try:
    from zoneinfo import ZoneInfo
    ZONE_KST = ZoneInfo("Asia/Seoul")
except Exception:
    ZONE_KST = None

# ================== 설정 ==================
ACCESS_KEY   = os.environ.get("NCLOUD_ACCESS_KEY", "")  # .env
SECRET_KEY   = os.environ.get("NCLOUD_SECRET_KEY", "")  # .env
ENDPOINT_URL = "https://kr.object.ncloudstorage.com"
REGION       = "kr-standard"

BUCKET       = "domeme"
PREFIX       = "4번사업자_0201"  # 특정 폴더만: 예) "imgs/"; 전체면 ""

ONLY_IMAGES  = False
IMG_EXTS     = {".png",".jpg",".jpeg",".webp",".gif",".bmp",".tif",".tiff"}

OUTPUT_XLSX  = r"C:\Users\USER\Documents\국내위탁\마이박스\26년2월\4번사업자\domeme_links.xlsx"
# =========================================


def is_image_key(key: str) -> bool:
    if not ONLY_IMAGES:
        return True
    lower = key.lower()
    return any(lower.endswith(ext) for ext in IMG_EXTS)


def direct_url(bucket: str, key: str) -> str:
    # 키를 URL 인코딩(경로 구분 '/'는 유지)
    return f"{ENDPOINT_URL}/{bucket}/{quote(key, safe='/~()!*.\'')}"


def split_folder_info(key: str):
    # 폴더 경로/파일명 분리
    if "/" in key:
        folder_path, name = key.rsplit("/", 1)
        folder_path += "/"  # 보기 좋게 끝에 '/' 유지
    else:
        folder_path, name = "", key
    # 최상위 폴더 + 깊이
    if folder_path:
        top = folder_path.split("/", 1)[0] + "/"
        depth = folder_path.count("/")
    else:
        top = ""
        depth = 0
    return name, folder_path, top, depth


def to_kst_naive(dt_utc):
    """UTC tz-aware datetime → KST naive(datetime without tz)로 변환. 실패 시 문자열 반환."""
    if dt_utc is None:
        return None
    try:
        # boto3가 주는 LastModified는 tz-aware UTC
        if ZONE_KST is not None and hasattr(dt_utc, "astimezone"):
            return dt_utc.astimezone(ZONE_KST).replace(tzinfo=None)
        # zoneinfo가 없으면 그냥 문자열로
        return str(dt_utc)
    except Exception:
        return str(dt_utc)


def main():
    s3 = boto3.client(
        "s3",
        endpoint_url=ENDPOINT_URL,
        region_name=REGION,
        aws_access_key_id=ACCESS_KEY,
        aws_secret_access_key=SECRET_KEY,
        config=Config(signature_version="s3v4", s3={"addressing_style": "path"}),
    )

    rows = []
    paginator = s3.get_paginator("list_objects_v2")
    try:
        pages = paginator.paginate(Bucket=BUCKET, Prefix=PREFIX)
    except ClientError as e:
        print("[ERROR] list_objects_v2 실패:",
              e.response.get("Error", {}).get("Code"),
              "-", e.response.get("Error", {}).get("Message"))
        return

    total = 0
    for page in pages:
        for obj in page.get("Contents", []) or []:
            key = obj["Key"]
            # 폴더 플레이스홀더 제외
            if key.endswith("/"):
                continue
            if not is_image_key(key):
                continue

            name, folder_path, top, depth = split_folder_info(key)
            name_no_ext = os.path.splitext(name)[0]  # ← 확장자 제거
            link = direct_url(BUCKET, key)

            size = obj.get("Size", 0)
            lastmod_utc = obj.get("LastModified")
            lastmod_kst_naive = to_kst_naive(lastmod_utc)

            rows.append({
                "이름": name_no_ext,
                "폴더경로": folder_path,     # 예: imgs/sub/
                "최상위폴더": top,          # 예: imgs/
                "폴더깊이": depth,          # 0=루트
                "LINK": link,               # 공개 직접 URL
                "Key": key,                 # 전체 키(경로+이름)
                "Size": size,               # 바이트
                "LastModified(KST)": lastmod_kst_naive,  # tz 없는 datetime (Excel 호환)
            })
            total += 1

    if not rows:
        print("[INFO] 조건에 맞는 객체가 없습니다.")
        return

    # 정렬: 폴더경로 -> 파일명
    df = pd.DataFrame(rows)
    df.sort_values(by=["폴더경로", "이름"], inplace=True)

    # 보조 요약 시트(폴더별 개수/용량)
    summary = (
        df.groupby("폴더경로", dropna=False)
          .agg(개수=("이름","count"), 총용량=("Size","sum"))
          .reset_index()
          .sort_values(by="폴더경로", na_position="first")
    )

    os.makedirs(os.path.dirname(OUTPUT_XLSX) or ".", exist_ok=True)
    with pd.ExcelWriter(OUTPUT_XLSX, engine="openpyxl") as xw:
        df.to_excel(xw, sheet_name="objects", index=False)
        summary.to_excel(xw, sheet_name="폴더별요약", index=False)

    print(f"[OK] {total}개 항목을 엑셀로 저장: {OUTPUT_XLSX}")


if __name__ == "__main__":
    main()
