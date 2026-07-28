#!/usr/bin/env python3
"""ECOUNT(이카운트) OAPI V2 클라이언트.

오프라인 재고 시스템의 데이터 소스. 온라인(Cafe24)과 완전히 분리된 별도 재고를
이카운트에서 읽어온다.

인증 흐름 (이카운트 OAPI V2):
  1. Zone 조회   : COM_CODE(회사코드) → ZONE 문자열
  2. 로그인      : COM_CODE + USER_ID + ZONE + API_CERT_KEY → SESSION_ID
  3. API 호출    : 이후 모든 요청에 ?SESSION_ID=... 쿼리로 인증

도메인:
  - 테스트(sandbox): https://sboapi{ZONE}.ecount.com
  - 운영(production): https://oapi{ZONE}.ecount.com
  Zone 조회는 ZONE이 붙기 전이라 각각 https://sboapi.ecount.com / https://oapi.ecount.com 사용.

인증정보 소스 (환경변수 우선, 없으면 gitignore된 로컬 파일):
  - ECOUNT_COM_CODE      회사코드
  - ECOUNT_USER_ID       API용 사용자 ID
  - ECOUNT_API_CERT_KEY  API 인증키 (Self-Customizing → 정보관리 → API 인증키발급)
  - ECOUNT_MODE          "test"(기본) 또는 "prod"
  - ECOUNT_ZONE          (선택) 이미 아는 경우 Zone 조회를 건너뜀
  로컬 파일 fallback: .secrets/ecount.json  (COM_CODE/USER_ID/API_CERT_KEY/MODE 키)

주의: 인증키는 절대 커밋하지 않는다. .secrets/ 는 .gitignore 처리됨.
"""
from __future__ import annotations

import json
import os
from pathlib import Path
from typing import Any

import requests

_SECRETS_FILE = Path(__file__).parent / ".secrets" / "ecount.json"
_DEFAULT_TIMEOUT = 30


class EcountError(RuntimeError):
    """이카운트 API 호출 실패. 원본 응답을 함께 담는다."""

    def __init__(self, message: str, *, raw: Any = None):
        super().__init__(message)
        self.raw = raw


def _load_secrets() -> dict[str, str]:
    """환경변수 → 로컬 파일 순으로 인증정보를 모은다."""
    creds: dict[str, str] = {}
    if _SECRETS_FILE.exists():
        try:
            creds.update({k: str(v) for k, v in json.loads(_SECRETS_FILE.read_text()).items()})
        except Exception:
            pass
    # 환경변수가 있으면 파일값을 덮어쓴다.
    for env_key, cfg_key in (
        ("ECOUNT_COM_CODE", "COM_CODE"),
        ("ECOUNT_USER_ID", "USER_ID"),
        ("ECOUNT_API_CERT_KEY", "API_CERT_KEY"),
        ("ECOUNT_MODE", "MODE"),
        ("ECOUNT_ZONE", "ZONE"),
    ):
        v = os.environ.get(env_key, "").strip()
        if v:
            creds[cfg_key] = v
    return creds


class EcountClient:
    def __init__(
        self,
        com_code: str | None = None,
        user_id: str | None = None,
        api_cert_key: str | None = None,
        mode: str | None = None,
        zone: str | None = None,
        lan_type: str = "ko-KR",
    ):
        creds = _load_secrets()
        self.com_code = com_code or creds.get("COM_CODE", "")
        self.user_id = user_id or creds.get("USER_ID", "")
        self.api_cert_key = api_cert_key or creds.get("API_CERT_KEY", "")
        self.mode = (mode or creds.get("MODE") or "test").lower()
        self.zone = zone or creds.get("ZONE") or ""
        self.lan_type = lan_type
        self.session_id: str = ""

        if not self.com_code:
            raise EcountError("ECOUNT_COM_CODE(회사코드)가 없습니다. .secrets/ecount.json 또는 환경변수 설정 필요.")
        if not self.api_cert_key:
            raise EcountError("ECOUNT_API_CERT_KEY(인증키)가 없습니다.")

    # ---- 도메인 구성 ----
    @property
    def _prefix(self) -> str:
        return "oapi" if self.mode == "prod" else "sboapi"

    def _zone_lookup_url(self) -> str:
        return f"https://{self._prefix}.ecount.com/OAPI/V2/Zone"

    def _base_url(self) -> str:
        if not self.zone:
            raise EcountError("ZONE이 아직 설정되지 않았습니다. login() 또는 fetch_zone()을 먼저 호출하세요.")
        return f"https://{self._prefix}{self.zone}.ecount.com/OAPI/V2"

    # ---- 저수준 호출 ----
    @staticmethod
    def _post(url: str, payload: dict) -> dict:
        resp = requests.post(url, json=payload, timeout=_DEFAULT_TIMEOUT)
        try:
            data = resp.json()
        except ValueError as e:
            resp.raise_for_status()  # 200인데 JSON 파싱만 실패한 경우는 아래에서 별도 에러로
            raise EcountError(f"JSON 파싱 실패: {url}", raw=resp.text) from e
        if not resp.ok:
            # raise_for_status()는 응답 바디(이카운트가 왜 거부했는지)를 안 보여주고 그냥
            # HTTPError만 던져서 원인 파악이 안 됨 — 바디를 먼저 파싱해 메시지에 포함시킨다.
            body_preview = json.dumps(data, ensure_ascii=False)[:500]
            raise EcountError(f"HTTP {resp.status_code} {url} — {body_preview}", raw=data)
        return data

    # ---- 1) Zone 조회 ----
    def fetch_zone(self) -> str:
        if self.zone:
            return self.zone
        data = self._post(self._zone_lookup_url(), {"COM_CODE": self.com_code})
        zone = _dig(data, ["Data", "ZONE"]) or _dig(data, ["ZONE"])
        if not zone:
            raise EcountError("Zone 조회 응답에서 ZONE을 찾지 못했습니다.", raw=data)
        self.zone = str(zone)
        return self.zone

    # ---- 2) 로그인 ----
    def login(self) -> str:
        self.fetch_zone()
        url = f"{self._base_url()}/OAPILogin"
        payload = {
            "COM_CODE": self.com_code,
            "USER_ID": self.user_id,
            "API_CERT_KEY": self.api_cert_key,
            "LAN_TYPE": self.lan_type,
            "ZONE": self.zone,
        }
        data = self._post(url, payload)
        sid = (
            _dig(data, ["Data", "Datas", "SESSION_ID"])
            or _dig(data, ["Data", "SESSION_ID"])
            or _dig(data, ["SESSION_ID"])
        )
        if not sid:
            raise EcountError("로그인 응답에서 SESSION_ID를 찾지 못했습니다.", raw=data)
        self.session_id = str(sid)
        return self.session_id

    def ensure_session(self) -> None:
        if not self.session_id:
            self.login()

    # ---- 3) 인증된 API 호출 ----
    def call(self, path: str, payload: dict | None = None) -> dict:
        """SESSION_ID를 붙여 POST 호출. path 예: 'InventoryBalance/GetListInventoryBalanceStatusByLocation'"""
        self.ensure_session()
        url = f"{self._base_url()}/{path.lstrip('/')}?SESSION_ID={self.session_id}"
        return self._post(url, payload or {})

    # ---- 편의 메서드 ----
    def inventory_balance_by_location(self, base_date: str, **extra) -> dict:
        """창고 + 품목별 재고현황. base_date: 'YYYYMMDD'."""
        payload = {"BASE_DATE": base_date}
        payload.update(extra)
        return self.call("InventoryBalance/GetListInventoryBalanceStatusByLocation", payload)


def _dig(obj: Any, path: list[str]) -> Any:
    """중첩 dict에서 안전하게 값 꺼내기."""
    cur = obj
    for key in path:
        if isinstance(cur, dict) and key in cur:
            cur = cur[key]
        else:
            return None
    return cur
