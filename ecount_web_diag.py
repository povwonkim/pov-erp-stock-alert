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
# 브라우저 로그인 상태(쿠키 등)를 저장해뒀다가 재사용 — 안 그러면 매번 새 브라우저로 접속할
# 때마다 "새로운 기기 로그인 알림" 팝업이 뜨는 걸 실제로 확인함(2026-07-28).
STORAGE_STATE_PATH = Path(__file__).parent / ".secrets" / "ecount_web_state.json"


def dismiss_new_device_modal(page) -> bool:
    """'새로운 기기 로그인 알림' 모달이 보이면 [등록]을 눌러 이 브라우저를 신뢰하게 만든다.
    처리했으면 True. 등록해두면 STORAGE_STATE_PATH에 저장된 쿠키로 다음 실행부턴 이 모달
    자체가 안 뜬다."""
    register_btn = page.get_by_role("button", name="등록", exact=True)
    try:
        if register_btn.count() > 0 and register_btn.first.is_visible():
            print("[diag] '새로운 기기 로그인 알림' 모달 발견 — [등록] 클릭")
            register_btn.first.click()
            page.wait_for_timeout(1500)
            return True
    except Exception as e:
        print(f"[diag]   모달 처리 중 예외(무시): {e}")
    return False


def open_browser_context(p):
    """저장된 로그인 상태(있으면)를 재사용하는 브라우저/컨텍스트를 연다.
    ecount_sales_scraper.py도 동일한 로직을 쓰므로 여기서 공유한다."""
    # --disable-dev-shm-usage: 리소스가 작은 VPS/컨테이너에서 /dev/shm 용량 제한 때문에
    # 큰 페이지를 렌더링할 때 Chromium이 멈추거나 죽는 문제의 표준 우회법(Playwright 권장).
    browser = p.chromium.launch(headless=True, args=["--disable-dev-shm-usage"])
    context_kwargs = {}
    if STORAGE_STATE_PATH.exists():
        context_kwargs["storage_state"] = str(STORAGE_STATE_PATH)
    context = browser.new_context(**context_kwargs)
    return browser, context


def login_if_needed(page, web_user_id: str, web_password: str) -> None:
    """이미 로그인 폼이 없으면(=쿠키로 인증됨) 아무것도 안 하고, 있으면 로그인 후
    '새로운 기기' 모달까지 처리한다. ecount_sales_scraper.py도 이 함수를 그대로 쓴다."""
    dismiss_new_device_modal(page)

    # 2026-07-28 실제 폼 구조 확인: id="txtUserId"/id="txtPass"/버튼 id="save"
    # (onclick="excuteLogin()"). 텍스트 기반 탐색은 엉뚱한 요소를 집을 수 있어 id로 직접 지정.
    pw_inputs = page.locator("#txtPass")
    if pw_inputs.count() == 0:
        return

    id_input = page.locator("#txtUserId")
    if id_input.count() > 0:
        id_input.fill(web_user_id)
    pw_inputs.fill(web_password)

    save_btn = page.locator("#save")
    if save_btn.count() > 0:
        save_btn.click()
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        page.wait_for_timeout(2000)  # SPA/JS 리다이렉트 여유

    # 로그인 직후에도 "새로운 기기" 모달이 뜰 수 있어 한 번 더 확인.
    dismiss_new_device_modal(page)


def save_storage_state(context) -> None:
    STORAGE_STATE_PATH.parent.mkdir(parents=True, exist_ok=True)
    context.storage_state(path=str(STORAGE_STATE_PATH))


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
        if STORAGE_STATE_PATH.exists():
            print(f"[diag] 저장된 로그인 상태 재사용: {STORAGE_STATE_PATH}")
        browser, context = open_browser_context(p)
        page = context.new_page()
        print("[diag] 페이지 로드 중...")
        page.goto(link, wait_until="networkidle", timeout=60000)

        page.screenshot(path=str(DUMP_DIR / "diag_01_initial.png"), full_page=True)
        (DUMP_DIR / "diag_01_initial.html").write_text(page.content())
        print(f"[diag] 초기 화면 저장: {DUMP_DIR / 'diag_01_initial.png'} / .html")

        had_login_form = page.locator("#txtPass").count() > 0
        login_if_needed(page, web_user_id, web_password)
        if had_login_form:
            page.screenshot(path=str(DUMP_DIR / "diag_02_filled.png"), full_page=True)
            print(f"[diag] 로그인 시도 후 현재 URL: {page.url}")
            alert_box = page.locator(".alert, [class*='alert']")
            if alert_box.count() > 0:
                for i in range(min(alert_box.count(), 3)):
                    txt = alert_box.nth(i).inner_text().strip()
                    if txt:
                        print(f"[diag]   ⚠️ 알림창 텍스트[{i}]: {txt!r}")
        else:
            print("[diag] 로그인 폼(#txtPass) 없음 — 이미 인증된 상태이거나 팝업 구조가 다름")

        # 로그인 상태(쿠키)를 저장해서 다음 실행에서 재사용.
        save_storage_state(context)
        print(f"[diag] 로그인 상태 저장: {STORAGE_STATE_PATH}")

        # 로그인 후 페이지는 표가 1000행 넘게 나오기도 해서(2026-07-28 확인: 1004개 <tr>),
        # full_page=True 스크린샷이 그 긴 페이지 전체를 렌더링하려다 헤드리스 브라우저가
        # 메모리를 과도하게 먹어 서버 전체가 응답 불가 상태가 된 적이 있다 — 뷰포트만 찍는다.
        page.screenshot(path=str(DUMP_DIR / "diag_03_after_login.png"), full_page=False)
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
