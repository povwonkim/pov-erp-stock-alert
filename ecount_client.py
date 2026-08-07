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
_SESSION_CACHE_FILE = Path(__file__).parent / ".secrets" / "ecount_session.json"
_DEFAULT_TIMEOUT = 30

# 이카운트 실서버(prod) 전송기준(2026-07-28 공식 문서 확인): 로그인(Zone+로그인)과 조회
# (발주서조회/품목조회/재고현황/창고별재고현황)는 각각 종류별로 1회/10분. 스크립트를 실행할
# 때마다 새로 로그인하면 이 한도를 금방 넘긴다 — 그래서 세션ID를 파일에 캐시해두고 재사용한다.
# (이카운트 문서: "세션ID는 설정된 시간 동안 재사용 가능하며, API를 한 번 호출하면 다시 그
# 시간 동안 사용할 수 있다" — 즉 계속 쓰는 한 만료되지 않는다.)


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
        self._session_from_cache = False

        if not self.com_code:
            raise EcountError("ECOUNT_COM_CODE(회사코드)가 없습니다. .secrets/ecount.json 또는 환경변수 설정 필요.")
        if not self.api_cert_key:
            raise EcountError("ECOUNT_API_CERT_KEY(인증키)가 없습니다.")

        # zone/session_id를 명시적으로 안 받았으면 캐시에서 복원 시도 — 로그인 API도 1회/10분
        # 제한이라, 스크립트 실행할 때마다 새로 로그인하면 금방 막힌다.
        if not zone or not self.session_id:
            cached = self._load_session_cache()
            if cached:
                self.zone = self.zone or cached.get("zone", "")
                if cached.get("session_id"):
                    self.session_id = cached["session_id"]
                    self._session_from_cache = True

    def _load_session_cache(self) -> dict | None:
        if not _SESSION_CACHE_FILE.exists():
            return None
        try:
            data = json.loads(_SESSION_CACHE_FILE.read_text())
        except Exception:
            return None
        # 다른 계정/모드의 캐시를 잘못 재사용하지 않게 회사코드+모드가 일치할 때만 쓴다.
        if data.get("com_code") != self.com_code or data.get("mode") != self.mode:
            return None
        return data

    def _save_session_cache(self) -> None:
        _SESSION_CACHE_FILE.parent.mkdir(parents=True, exist_ok=True)
        _SESSION_CACHE_FILE.write_text(json.dumps({
            "com_code": self.com_code, "mode": self.mode,
            "zone": self.zone, "session_id": self.session_id,
        }, ensure_ascii=False, indent=2))

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
        except ValueError:
            # 바디가 JSON이 아님 — 성공(200)인데 파싱만 실패한 경우와, 애초에 실패 응답(빈 바디
            # 등)이라 파싱할 게 없는 경우를 구분해서 보여준다. 게이트웨이/WAF가 이카운트 앱까지
            # 가기 전에 막았을 가능성이 있어 응답 헤더도 같이 보여준다.
            if resp.ok:
                raise EcountError(f"JSON 파싱 실패(200인데 바디가 JSON 아님): {url}", raw=resp.text)
            raise EcountError(
                f"HTTP {resp.status_code} {url} — 바디가 JSON 아님(빈 바디 또는 게이트웨이 응답으로 "
                f"추정). 바디 미리보기: {resp.text[:300]!r} / 응답헤더: {dict(resp.headers)}",
                raw=resp.text,
            )
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
        self._session_from_cache = False
        self._save_session_cache()
        return self.session_id

    def ensure_session(self) -> None:
        if not self.session_id:
            self.login()

    # ---- 3) 인증된 API 호출 ----
    def call(self, path: str, payload: dict | None = None) -> dict:
        """SESSION_ID를 붙여 POST 호출. path 예: 'InventoryBalance/GetListInventoryBalanceStatusByLocation'

        캐시된 세션이 만료돼서 실패한 경우에만 한 번 재로그인 후 재시도한다(로그인도 1회/10분
        제한이라 매번 재시도하면 오히려 더 막히므로, 캐시 세션을 썼을 때만 시도).
        """
        self.ensure_session()
        url = f"{self._base_url()}/{path.lstrip('/')}?SESSION_ID={self.session_id}"
        try:
            return self._post(url, payload or {})
        except EcountError:
            if not self._session_from_cache:
                raise
            self.session_id = ""
            self.login()
            url = f"{self._base_url()}/{path.lstrip('/')}?SESSION_ID={self.session_id}"
            return self._post(url, payload or {})

    # ---- 편의 메서드 ----
    def inventory_balance_by_location(self, base_date: str, **extra) -> dict:
        """창고 + 품목별 재고현황. base_date: 'YYYYMMDD'."""
        payload = {"BASE_DATE": base_date}
        payload.update(extra)
        return self.call("InventoryBalance/GetListInventoryBalanceStatusByLocation", payload)

    # 품목조회 API 경로 후보. OAPI V2 조회 API 4개 중 하나로 문서에 존재하는 것은 확인했으나
    # (README '조회 API 전수 확인 결과'), 이 계정은 품목조회 인증이 안 되어 있어 실제 응답으로
    # path를 확정하지 못했다. probe_item_list()로 후보를 하나씩 때려보고 되는 것을 찾는다.
    ITEM_LIST_PATH_CANDIDATES = (
        "InventoryBasic/GetBasicProductsList",
        "InventoryBasic/GetBasicProduct",
        "Inventory/GetBasicProductsList",
        "AccountBasic/GetBasicProductsList",
    )

    def item_list(self, *, path: str = "", **extra) -> dict:
        """품목등록 전량 조회.

        이 시스템의 품목마스터가 '재고·판매에 등장한 품목'의 부산물이 아니라 '이카운트에
        등록된 품목 전량'의 미러가 되려면 이 API가 있어야 한다. 재고현황 API는 해당 창고에
        재고 기록이 있는 품목만 돌려주기 때문에, 등록만 되고 아직 안 움직인 신상품은 어떤
        경로로도 안 잡힌다(2026-08-07 확인: 카페24에 있는데 시트에 없는 16건의 원인).

        path를 안 주면 ITEM_LIST_PATH_CANDIDATES의 첫 번째를 쓴다. 확정되면 그 값을
        기본값으로 올릴 것.
        """
        return self.call(path or self.ITEM_LIST_PATH_CANDIDATES[0], dict(extra))

    def probe_item_list(self, **extra) -> tuple[str, dict]:
        """품목조회 path 후보를 순서대로 호출해 처음 성공하는 것을 (path, 응답)으로 돌려준다.

        조회 API는 종류당 1회/10분 제한이라(README 참고) 후보를 연달아 때리면 뒤쪽은
        제한에 걸릴 수 있다. 그래서 실패 사유를 그대로 모아 마지막에 전부 보여준다 —
        인증 문제인지 path 문제인지 전송기준 문제인지 구분할 수 있어야 하기 때문.
        """
        failures = []
        for path in self.ITEM_LIST_PATH_CANDIDATES:
            try:
                return path, self.call(path, dict(extra))
            except EcountError as exc:
                failures.append(f"  {path}\n    → {exc}")
        raise EcountError(
            "품목조회 path 후보를 전부 실패했습니다. 이카운트 로그인 후 API 문서에서 정확한 "
            "path를 확인하거나, Self-Customizing → 정보관리 → API 인증키발급에서 품목조회 "
            "권한이 켜져 있는지 보세요.\n" + "\n".join(failures)
        )


def _dig(obj: Any, path: list[str]) -> Any:
    """중첩 dict에서 안전하게 값 꺼내기."""
    cur = obj
    for key in path:
        if isinstance(cur, dict) and key in cur:
            cur = cur[key]
        else:
            return None
    return cur
