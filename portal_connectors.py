# -*- coding: utf-8 -*-
"""
통합 포털 - 마켓별 커넥터 (사업자별 공식 OpenAPI).

MVP 신호 2개:
  - shipment_holds()        : 고객 출고중지 요청건 (미발송 처리 대상)
  - unanswered_inquiries()  : 미응대 고객문의

설계 원칙:
  - 자격증명은 .env 에서만 로드 (커밋 금지). 미설정/오류 시 빈 리스트 반환 → 포털은 절대 죽지 않음.
  - 마켓별 Connector 를 BaseConnector 인터페이스로 통일. 신규 마켓은 클래스만 추가.
  - 쿠팡 HMAC 서명은 공식 문서 표준 그대로. 엔드포인트 경로는 계정/문서에 맞춰 .env 로 덮어쓸 수 있게 분리.

.env 키 (사업자 rank 1~6):
  COUPANG_{rank}_VENDOR_ID, COUPANG_{rank}_ACCESS_KEY, COUPANG_{rank}_SECRET_KEY
  (선택) COUPANG_SHIPMENT_HOLD_PATH, COUPANG_INQUIRY_PATH  ← 계정 API 문서대로 경로 지정
"""
from __future__ import annotations

import hashlib
import hmac
import os
import time
from datetime import datetime, timezone
from typing import Any, Dict, List

try:
    from dotenv import load_dotenv
    from pathlib import Path
    load_dotenv(Path(__file__).resolve().parent / ".env")
except ImportError:
    pass

try:
    import requests
except ImportError:
    requests = None


class BaseConnector:
    """마켓 커넥터 인터페이스. 모든 메서드는 실패 시 [] 반환(포털 무중단)."""

    market = "base"

    def __init__(self, rank: int, biz_id: str = ""):
        self.rank = rank
        self.biz_id = biz_id

    @property
    def configured(self) -> bool:
        return False

    def shipment_holds(self) -> List[Dict[str, Any]]:
        return []

    def unanswered_inquiries(self) -> List[Dict[str, Any]]:
        return []


class CoupangConnector(BaseConnector):
    """쿠팡 마켓플레이스 OpenAPI (WING). HMAC-SHA256 서명(공식 표준)."""

    market = "coupang"
    BASE = "https://api-gateway.coupang.com"
    # 계정/문서에 맞춰 .env 로 덮어쓰기 (확정 전 기본값은 자리표시; 반드시 검증 후 사용)
    SHIPMENT_HOLD_PATH = os.environ.get(
        "COUPANG_SHIPMENT_HOLD_PATH",
        "/v2/providers/openapi/apis/api/v4/vendors/{vendorId}/shipment-hold",
    )
    INQUIRY_PATH = os.environ.get(
        "COUPANG_INQUIRY_PATH",
        "/v2/providers/openapi/apis/api/v1/vendors/{vendorId}/onlineInquiries",
    )

    def __init__(self, rank: int, biz_id: str = ""):
        super().__init__(rank, biz_id)
        self.vendor_id = os.environ.get(f"COUPANG_{rank}_VENDOR_ID", "").strip()
        self.access_key = os.environ.get(f"COUPANG_{rank}_ACCESS_KEY", "").strip()
        self.secret_key = os.environ.get(f"COUPANG_{rank}_SECRET_KEY", "").strip()

    @property
    def configured(self) -> bool:
        return bool(self.vendor_id and self.access_key and self.secret_key and requests)

    def _auth_header(self, method: str, path: str, query: str) -> str:
        # 공식 스펙: signed-date = yyMMdd'T'HHmmss'Z' (UTC)
        signed_date = datetime.now(timezone.utc).strftime("%y%m%dT%H%M%SZ")
        message = signed_date + method + path + query
        signature = hmac.new(
            self.secret_key.encode("utf-8"),
            message.encode("utf-8"),
            hashlib.sha256,
        ).hexdigest()
        return (
            f"CEA algorithm=HmacSHA256, access-key={self.access_key}, "
            f"signed-date={signed_date}, signature={signature}"
        )

    def _get(self, path_tmpl: str, query: str = "") -> Any:
        if not self.configured:
            return None
        path = path_tmpl.replace("{vendorId}", self.vendor_id)
        try:
            headers = {
                "Authorization": self._auth_header("GET", path, query),
                "Content-Type": "application/json;charset=UTF-8",
            }
            url = self.BASE + path + (("?" + query) if query else "")
            r = requests.get(url, headers=headers, timeout=15)
            if r.status_code != 200:
                print(f"[coupang {self.rank}번] {path} HTTP {r.status_code}: {r.text[:200]}")
                return None
            return r.json()
        except Exception as e:
            print(f"[coupang {self.rank}번] 요청 실패 {path}: {e}")
            return None

    def shipment_holds(self) -> List[Dict[str, Any]]:
        """고객 출고중지 요청건. 응답 스키마는 계정 API 문서에 맞춰 _adapt 에서 매핑."""
        data = self._get(self.SHIPMENT_HOLD_PATH)
        return self._adapt(data, kind="shipment_hold")

    def unanswered_inquiries(self) -> List[Dict[str, Any]]:
        data = self._get(self.INQUIRY_PATH)
        return self._adapt(data, kind="inquiry")

    def _adapt(self, data: Any, kind: str) -> List[Dict[str, Any]]:
        """쿠팡 응답 → 포털 공통 스키마. (실제 응답 키는 계정 문서대로 여기만 수정)"""
        if not data:
            return []
        rows = data.get("data") if isinstance(data, dict) else data
        if not isinstance(rows, list):
            return []
        out = []
        for it in rows:
            if not isinstance(it, dict):
                continue
            out.append({
                "market": self.market,
                "rank": self.rank,
                "kind": kind,
                "id": it.get("orderId") or it.get("inquiryId") or it.get("id") or "",
                "title": it.get("content") or it.get("title") or it.get("productName") or "",
                "created": it.get("createdAt") or it.get("inquiryAt") or it.get("orderedAt") or "",
                "raw": it,
            })
        return out


# 신규 마켓은 여기에 클래스 추가 (스마트스토어/11번가 등) — .env 키 줄 때 구현
MARKET_CONNECTORS = {
    "coupang": CoupangConnector,
}


def build_connectors(accounts: List[str], markets: List[str] | None = None) -> Dict[int, List[BaseConnector]]:
    """rank(1~N) → [Connector,...]. accounts = 사업자 ID 리스트(라벨용)."""
    markets = markets or list(MARKET_CONNECTORS.keys())
    result: Dict[int, List[BaseConnector]] = {}
    for idx, biz_id in enumerate(accounts, start=1):
        conns = []
        for m in markets:
            cls = MARKET_CONNECTORS.get(m)
            if cls:
                conns.append(cls(idx, biz_id))
        result[idx] = conns
    return result
