#!/usr/bin/env python3
"""이카운트 웹 리포트(판매현황) 페이지 구조를 확인하기 위한 1회성 진단 스크립트.

자동알림 이메일엔 첨부파일이 없고 "수신문서보기" 팝업 링크만 있다(2026-07-28 확인).
이 링크는 로그인이 필요하고, 매일 SEND_CM_ID가 달라져 하드코딩할 수 없다 — 그래서 매번
Gmail에서 그날의 링크를 다시 찾아야 한다(ecount_gmail_fetch.find_view_link).

로그인 폼/표의 실제 HTML 구조를 모르는 채로 스크래퍼를 완성할 수 없어서, 이 스크립트는
1) 링크로 들어가서 2) 필요하면 로그인 시도 후 3) 스크린샷 + 렌더된 HTML을 저장만 한다.
이 결과를 보고 ecount_sales_scraper.py의 실제 선택자(selector)를 확정한다.

사용법 (서버에서):
  pip install playwright && playwright install chromium --with-deps
  export ECOUNT_WEB_PASSWORD='...'   # .secrets/ecount.json에 WEB_PASSWORD로 넣어도 됨
  .venv/bin/python ecount_web_diag.py
"""
from __future__ import annotations

import sys
from pathlib import Path

from ecount_gmail_fetch import _load_creds, list_messages, get_html_body, find_view_link
from ecount_client import _load_secrets

DUMP_DIR = Path(__file__).parent / "cron_tracking" / "ecount"
GMAIL_QUERY_DEFAULT = "from:ecountnotice@ecount.com newer_than:2d"


def get_today_view_link(query: str = GMAIL_QUERY_DEFAULT) -> str:
    from googleapiclient.discovery import build

    creds = _load_creds()
    service = build("gmail", "v1", credentials=creds)
    messages = list_messages(service, query)
    if not messages:
        raise SystemExit(f"[diag] Gmail 쿼리에 매치되는 메일이 없습니다: {query!r}")
    html = get_html_body(service, messages[0]["id"])
    link = find_view_link(html)
    if not link:
        DUMP_DIR.mkdir(parents=True, exist_ok=True)
        dump_path = DUMP_DIR / "diag_00_email_body.html"
        dump_path.write_text(html)
        raise SystemExit(
            f"[diag] 메일 본문에서 '수신문서보기' 링크를 못 찾았습니다. "
            f"본문 원본을 저장했으니 확인해보세요: {dump_path}"
        )
    return link


def main() -> int:
    secrets = _load_secrets()
    web_user_id = secrets.get("WEB_USER_ID", "POV_API")
    web_password = secrets.get("WEB_PASSWORD", "")
    if not web_password:
        print("[diag] .secrets/ecount.json에 WEB_PASSWORD가 없습니다. 먼저 넣어주세요:", file=sys.stderr)
        print('  {"WEB_USER_ID": "POV_API", "WEB_PASSWORD": "..."} 형태로 기존 파일에 필드만 추가', file=sys.stderr)
        return 2

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("[diag] playwright가 설치되어 있지 않습니다: pip install playwright && playwright install chromium", file=sys.stderr)
        return 2

    print("[diag] Gmail에서 오늘의 '수신문서보기' 링크 찾는 중...")
    link = get_today_view_link()
    print(f"[diag] 링크: {link}")

    DUMP_DIR.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page()
        print("[diag] 페이지 로드 중...")
        page.goto(link, wait_until="networkidle", timeout=60000)

        page.screenshot(path=str(DUMP_DIR / "diag_01_initial.png"), full_page=True)
        (DUMP_DIR / "diag_01_initial.html").write_text(page.content())
        print(f"[diag] 초기 화면 저장: {DUMP_DIR / 'diag_01_initial.png'} / .html")

        # 로그인 폼이 있는지 확인 (비밀번호 입력창 존재 여부로 판단).
        pw_inputs = page.locator('input[type="password"]')
        if pw_inputs.count() > 0:
            print(f"[diag] 로그인 폼 발견 (password input {pw_inputs.count()}개) — 로그인 시도...")
            id_inputs = page.locator('input[type="text"], input[type="email"]')
            print(f"[diag]   텍스트/이메일 입력창 {id_inputs.count()}개 발견")
            if id_inputs.count() > 0:
                try:
                    id_inputs.first.fill(web_user_id)
                except Exception as e:
                    print(f"[diag]   아이디 입력 실패(이미 채워져 있을 수 있음): {e}")
            pw_inputs.first.fill(web_password)

            page.screenshot(path=str(DUMP_DIR / "diag_02_filled.png"), full_page=True)
            print(f"[diag] 입력 후 화면 저장: {DUMP_DIR / 'diag_02_filled.png'}")

            login_btn = page.get_by_text("로그인", exact=False)
            if login_btn.count() > 0:
                login_btn.first.click()
                page.wait_for_load_state("networkidle", timeout=30000)
            else:
                print("[diag]   '로그인' 텍스트를 가진 버튼을 못 찾음 — 수동 확인 필요")
        else:
            print("[diag] 로그인 폼 없음 — 이미 인증된 상태이거나 팝업 구조가 다름")

        page.screenshot(path=str(DUMP_DIR / "diag_03_after_login.png"), full_page=True)
        (DUMP_DIR / "diag_03_after_login.html").write_text(page.content())
        print(f"[diag] 로그인 후 화면 저장: {DUMP_DIR / 'diag_03_after_login.png'} / .html")

        tables = page.locator("table")
        print(f"[diag] <table> 태그 {tables.count()}개 발견")
        rows = page.locator("table tr")
        print(f"[diag] <table> 안 <tr> 총 {rows.count()}개 발견")
        if rows.count() > 1:
            print("[diag] 첫 3개 행 텍스트 미리보기:")
            for i in range(min(3, rows.count())):
                print(f"   [{i}] {rows.nth(i).inner_text()[:200]!r}")

        browser.close()

    print("\n[diag] 완료. 아래 파일들을 확인해서 실제 로그인 폼/표 구조를 파악하세요:")
    for f in sorted(DUMP_DIR.glob("diag_*")):
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
