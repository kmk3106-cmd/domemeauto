#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
로컬 폴더 → NCP Object Storage 일괄 업로드 스크립트 (하드코딩 버전)
- 하위 디렉터리까지 재귀 업로드
- S3 키는 로컬 경로를 '/' 기준으로 바꿔 prefix 뒤에 붙입니다
- 최신 boto3에서 체크섬 기본값으로 인한 오류 회피 (when_required)
- 기본은 ACL 헤더 미전송(= private). 필요 시 ACL_MODE="private" 로만 사용 권장
"""

import os
import sys
import mimetypes
import boto3
from botocore.config import Config
from boto3.s3.transfer import TransferConfig
from botocore.exceptions import ClientError

# ============== 여기만 채우세요 ==============
ENDPOINT   = "https://kr.object.ncloudstorage.com"
REGION     = "kr-standard"
ACCESS_KEY   = os.environ.get("NCLOUD_ACCESS_KEY", "")  # .env
SECRET_KEY   = os.environ.get("NCLOUD_SECRET_KEY", "")  # .env

BUCKET     = "domeme"              # 업로드할 버킷명
LOCAL_DIR  = r"C:\Users\USER\Documents\국내위탁\마이박스\26년2월\4번사업자\4번_통합상품명_이미지"  # ☆ 업로드할 '로컬 폴더' (Windows 예시)
PREFIX     = "4번사업자_0201/"            # 버킷 안에서 저장될 '프리픽스' (폴더처럼 보이게)

# ACL/SSE 설정 (필요할 때만)
ACL_MODE   = "public-read"                  # None(권장) 또는 "private" 만 사용 권장
SSE_MODE   = None                  # None 또는 "AES256"
# =============================================


def make_client():
    cfg = Config(
        signature_version="s3v4",
        s3={"addressing_style": "path"},
        request_checksum_calculation="when_required",
        response_checksum_validation="when_required",
        connect_timeout=15,
        read_timeout=60,
        retries={"max_attempts": 5, "mode": "standard"},
    )
    return boto3.client(
        "s3",
        endpoint_url=ENDPOINT,
        region_name=REGION,
        aws_access_key_id=ACCESS_KEY,
        aws_secret_access_key=SECRET_KEY,
        config=cfg,
    )


def iter_files(base_dir: str):
    base_dir = os.path.abspath(base_dir)
    for root, dirs, files in os.walk(base_dir):
        for name in files:
            full = os.path.join(root, name)
            rel  = os.path.relpath(full, base_dir)
            yield full, rel


def build_key(relpath: str) -> str:
    # 윈도우 백슬래시 → S3는 항상 슬래시
    rel = relpath.replace(os.sep, "/")
    if PREFIX:
        return f"{PREFIX.rstrip('/')}/{rel}"
    return rel


def build_extra_args(local_path: str):
    extra = {}
    # Content-Type 추정
    ctype, _ = mimetypes.guess_type(local_path)
    if ctype:
        extra["ContentType"] = ctype
    # ACL/SSE (필요할 때만)
    if ACL_MODE:
        extra["ACL"] = ACL_MODE                # 권장: None(미전송) 또는 "private"
    if SSE_MODE:
        extra["ServerSideEncryption"] = SSE_MODE  # "AES256" 필요 시만
    return extra


def main():
    if not os.path.isdir(LOCAL_DIR):
        print(f"[ERR] 로컬 폴더가 없습니다: {LOCAL_DIR}")
        sys.exit(1)

    s3 = make_client()

    # 버킷 접근 확인
    try:
        s3.head_bucket(Bucket=BUCKET)
        print(f"[OK] bucket access: {BUCKET}")
    except ClientError as e:
        print("[ERR] head_bucket:", e.response.get("Error"))
        sys.exit(1)

    # 멀티파트 전송 설정(기본값도 충분하지만 예시로 명시)
    tcfg = TransferConfig(multipart_threshold=8 * 1024 * 1024,
                          multipart_chunksize=8 * 1024 * 1024,
                          max_concurrency=4,
                          use_threads=True)

    total = ok = fail = 0
    failed = []

    print(f"[INFO] 업로드 시작: LOCAL_DIR={LOCAL_DIR} → s3://{BUCKET}/{PREFIX}")
    for full, rel in iter_files(LOCAL_DIR):
        total += 1
        key = build_key(rel)
        extra = build_extra_args(full)
        try:
            s3.upload_file(full, BUCKET, key, ExtraArgs=extra if extra else None, Config=tcfg)
            print(f"[OK] {full}  ->  s3://{BUCKET}/{key}")
            ok += 1
        except ClientError as e:
            err = e.response.get("Error", {})
            print(f"[FAIL] {full}  ->  s3://{BUCKET}/{key} | {err.get('Code')} - {err.get('Message')}")
            fail += 1
            failed.append((full, err.get("Code"), err.get("Message")))
        except Exception as e:
            print(f"[FAIL] {full}  ->  s3://{BUCKET}/{key} | {repr(e)}")
            fail += 1
            failed.append((full, "Exception", repr(e)))

    print("\n=== 업로드 요약 ===")
    print(f"총 파일: {total}  성공: {ok}  실패: {fail}")
    if failed:
        print("실패 목록:")
        for path, code, msg in failed[:20]:
            print(f" - {path} | {code} - {msg}")
        if len(failed) > 20:
            print(f" ... (외 {len(failed)-20}건 더 있음)")


if __name__ == "__main__":
    main()