#!/usr/bin/env python3
"""ECOUNT API 탐사 스크립트.

인증키를 받은 뒤 실행하면:
  1. Zone 조회 → 로그인(SESSION_ID) 이 정상 동작하는지 검증
  2. 재고 관련 엔드포인트를 호출해 **실제 응답 구조(컬럼/필드명)** 를 덤프

여기서 확인한 실제 필드명으로 파서(ecount_stock_report.py 등)를 만든다.
추측으로 짜지 않기 위한 단계.

사용법:
  # 1) 인증 검증만
  python3 ecount_probe.py --check

  # 2) 재고현황 덤프 (오늘 기준)
  python3 ecount_probe.py --inventory

  # 3) 임의 엔드포인트 호출 (재고변동표 정확한 path를 이카운트 API 문서에서 확인 후)
  python3 ecount_probe.py --endpoint "InventoryBalance/GetListInventoryBalanceStatusByLocation" --json '{"BASE_DATE":"20260727"}'

덤프는 cron_tracking/ecount/ 아래에 저장되며 gitignore 처리됨(원본에 품목/재고 데이터 포함).
"""
from __future__ import annotations

import argparse
import json
import sys
from datetime import datetime, timedelta, timezone
from pathlib import Path

from ecount_client import EcountClient, EcountError

KST = timezone(timedelta(hours=9))
DUMP_DIR = Path(__file__).parent / "cron_tracking" / "ecount"


def _dump(name: str, obj) -> Path:
    DUMP_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(KST).strftime("%Y%m%d_%H%M%S")
    path = DUMP_DIR / f"{name}_{ts}.json"
    path.write_text(json.dumps(obj, ensure_ascii=False, indent=2))
    return path


def _summarize(obj) -> None:
    """응답 최상위 구조와 첫 데이터 레코드의 필드명을 출력."""
    if isinstance(obj, dict):
        print(f"  최상위 키: {list(obj.keys())}")
        # 흔한 이카운트 응답 경로에서 리스트 찾기
        for path in (["Data", "Result"], ["Data", "Datas"], ["Data"], ["Result"]):
            cur = obj
            for k in path:
                cur = cur.get(k) if isinstance(cur, dict) else None
            if isinstance(cur, list) and cur:
                print(f"  레코드 리스트 위치: {' > '.join(path)}  (총 {len(cur)}건)")
                if isinstance(cur[0], dict):
                    print(f"  첫 레코드 필드: {list(cur[0].keys())}")
                return
        print("  (레코드 리스트를 자동으로 못 찾음 — 덤프 파일을 직접 확인하세요)")


def main() -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--check", action="store_true", help="Zone+로그인 검증만")
    ap.add_argument("--inventory", action="store_true", help="재고현황(창고별) 덤프")
    ap.add_argument("--endpoint", help="임의 엔드포인트 path")
    ap.add_argument("--json", default="{}", help="임의 엔드포인트 payload (JSON 문자열)")
    ap.add_argument("--base-date", default=datetime.now(KST).strftime("%Y%m%d"),
                    help="기준일 YYYYMMDD (기본: 오늘 KST)")
    args = ap.parse_args()

    try:
        client = EcountClient()
    except EcountError as e:
        print(f"[probe] 인증정보 오류: {e}", file=sys.stderr)
        print("  → .secrets/ecount.json 을 만들거나 ECOUNT_* 환경변수를 설정하세요.", file=sys.stderr)
        return 2

    print(f"[probe] mode={client.mode}  com_code={client.com_code}")

    try:
        zone = client.fetch_zone()
        print(f"[probe] ZONE = {zone}")
        sid = client.login()
        print(f"[probe] 로그인 성공. SESSION_ID = {sid[:8]}…")
    except EcountError as e:
        print(f"[probe] 인증 실패: {e}", file=sys.stderr)
        if e.raw is not None:
            p = _dump("login_error", e.raw)
            print(f"  원본 응답 덤프: {p}", file=sys.stderr)
        return 1
    except Exception as e:  # 네트워크/HTTP
        print(f"[probe] 호출 오류: {e}", file=sys.stderr)
        return 1

    if args.check:
        print("[probe] 인증 검증 완료 ✅")
        return 0

    try:
        if args.inventory:
            print(f"[probe] 재고현황 조회 (BASE_DATE={args.base_date}) …")
            data = client.inventory_balance_by_location(args.base_date)
            p = _dump("inventory_balance", data)
            print(f"[probe] 덤프 저장: {p}")
            _summarize(data)
        elif args.endpoint:
            payload = json.loads(args.json)
            print(f"[probe] {args.endpoint} 호출 … payload={payload}")
            data = client.call(args.endpoint, payload)
            safe = args.endpoint.replace("/", "_")
            p = _dump(safe, data)
            print(f"[probe] 덤프 저장: {p}")
            _summarize(data)
        else:
            print("[probe] --check / --inventory / --endpoint 중 하나를 지정하세요.")
    except EcountError as e:
        print(f"[probe] API 오류: {e}", file=sys.stderr)
        if e.raw is not None:
            print(f"  원본: {json.dumps(e.raw, ensure_ascii=False)[:500]}", file=sys.stderr)
        return 1
    except Exception as e:
        print(f"[probe] 호출 오류: {e}", file=sys.stderr)
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
