"""
Download fund price data from KFR K-FROMS.

Setup once:
  python -m pip install playwright
  python -m playwright install chromium

Credentials are read from environment variables so they are not stored in this
file:
  set KFROM_ID=caesar4th64
  set KFROM_PASSWORD=<your password>

Run:
  python download_fund_price.py

The downloaded Excel file is saved as "펀드 기준가.xlsx" in this script's folder.
"""

from __future__ import annotations

import argparse
import os
import shutil
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo


BASE_DIR = Path(__file__).resolve().parent
DOWNLOAD_NAME = "펀드 기준가.xlsx"
SITE_URL = "https://kfroms.kfr.co.kr/home"
DEFAULT_START_DATE = "2025-01-01"
DEBUG_TEXT = BASE_DIR / "kfroms_debug_text.txt"
DEBUG_SCREENSHOT = BASE_DIR / "kfroms_debug_screen.png"
DEBUG_HTML = BASE_DIR / "kfroms_debug.html"


def previous_business_day(today: date | None = None) -> date:
    current = today or datetime.now(ZoneInfo("Asia/Seoul")).date()
    current -= timedelta(days=1)
    while current.weekday() >= 5:
        current -= timedelta(days=1)
    return current


def normalize_date(value: str) -> str:
    return datetime.strptime(value, "%Y-%m-%d").strftime("%Y-%m-%d")


def env_required(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        raise RuntimeError(f"환경변수 {name}가 설정되어 있지 않습니다.")
    return value


def first_visible(page, selectors: list[str], timeout: int = 2500):
    last_error = None
    for selector in selectors:
        try:
            locator = page.locator(selector).first
            locator.wait_for(state="visible", timeout=timeout)
            return locator
        except Exception as exc:  # noqa: BLE001 - keep selector fallback simple
            last_error = exc
    raise RuntimeError(f"화면 요소를 찾지 못했습니다: {selectors}") from last_error


def click_text(page, text: str, timeout: int = 8000) -> None:
    candidates = [
        f"role=button[name='{text}']",
        f"role=link[name='{text}']",
        f"text={text}",
        f"xpath=//*[normalize-space()='{text}']",
    ]
    first_visible(page, candidates, timeout=timeout).click()


def visible_text_box(page, text: str):
    return page.evaluate(
        """
        (text) => {
            const nodes = Array.from(document.querySelectorAll('body *'));
            for (const node of nodes) {
                if ((node.textContent || '').trim() !== text) continue;
                const style = window.getComputedStyle(node);
                const rect = node.getBoundingClientRect();
                if (
                    style.visibility === 'hidden' ||
                    style.display === 'none' ||
                    rect.width === 0 ||
                    rect.height === 0
                ) continue;
                return { x: rect.x, y: rect.y, width: rect.width, height: rect.height };
            }
            return null;
        }
        """,
        text,
    )


def is_text_visible(page, text: str) -> bool:
    return visible_text_box(page, text) is not None


def click_visible_text(page, text: str, right_offset: float = 0, timeout: int = 8000) -> None:
    deadline = datetime.now().timestamp() + timeout / 1000
    box = visible_text_box(page, text)
    while box is None and datetime.now().timestamp() < deadline:
        page.wait_for_timeout(200)
        box = visible_text_box(page, text)
    if box is None:
        raise RuntimeError(f"화면 요소를 찾지 못했습니다: {text}")
    page.mouse.click(box["x"] + box["width"] / 2 + right_offset, box["y"] + box["height"] / 2)


def click_visible_menu_anchor(page, text: str, timeout: int = 8000) -> None:
    deadline = datetime.now().timestamp() + timeout / 1000
    clicked = False
    while not clicked and datetime.now().timestamp() < deadline:
        clicked = page.evaluate(
            """
            (text) => {
                const nodes = Array.from(document.querySelectorAll('body *'));
                for (const node of nodes) {
                    if ((node.textContent || '').trim() !== text) continue;
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    if (
                        style.visibility === 'hidden' ||
                        style.display === 'none' ||
                        rect.width === 0 ||
                        rect.height === 0
                    ) continue;
                    const anchor = node.closest('a');
                    if (!anchor) continue;
                    anchor.click();
                    return true;
                }
                return false;
            }
            """,
            text,
        )
        if not clicked:
            page.wait_for_timeout(200)
    if not clicked:
        raise RuntimeError(f"화면 요소를 찾지 못했습니다: {text}")


def fill_date(page, label_text: str, value: str) -> None:
    candidates = [
        f"input[aria-label*='{label_text}']",
        f"input[placeholder*='{label_text}']",
        f"xpath=//*[contains(normalize-space(), '{label_text}')]/following::input[1]",
    ]
    field = first_visible(page, candidates)
    field.click()
    field.press("Control+A")
    field.fill(value)


def login(page, user_id: str, password: str) -> None:
    id_field = first_visible(
        page,
        [
            "input[name='id']",
            "input[name='userId']",
            "input[name='loginId']",
            "input[type='text']",
            "input:not([type])",
        ],
        timeout=10000,
    )
    id_field.fill(user_id)

    password_field = first_visible(
        page,
        [
            "input[name='password']",
            "input[name='passwd']",
            "input[name='pwd']",
            "input[type='password']",
        ],
    )
    password_field.fill(password)

    for text in ["로그인", "Login", "LOGIN"]:
        try:
            click_text(page, text, timeout=2500)
            return
        except Exception:
            continue
    password_field.press("Enter")


def navigate_to_report(page) -> None:
    try:
        if not is_text_visible(page, "하나은행"):
            click_text(page, "위탁평가")
            page.wait_for_timeout(700)
        if not is_text_visible(page, "전체펀드 기준가추이"):
            click_visible_text(page, "하나은행", right_offset=135)
            page.wait_for_timeout(1000)
        click_visible_menu_anchor(page, "전체펀드 기준가추이")
        page.wait_for_timeout(700)
    except Exception:
        DEBUG_TEXT.write_text(page.locator("body").inner_text(timeout=5000), encoding="utf-8")
        DEBUG_HTML.write_text(page.content(), encoding="utf-8")
        page.screenshot(path=DEBUG_SCREENSHOT, full_page=True)
        raise


def results_table_box(page, timeout: int = 20000):
    headers = ["기준일", "유형", "예탁원펀드코드", "투자일", "펀드명", "일수익률"]
    deadline = datetime.now().timestamp() + timeout / 1000
    while datetime.now().timestamp() < deadline:
        box = page.evaluate(
            """
            (headers) => {
                const visibleRect = (node) => {
                    const style = window.getComputedStyle(node);
                    const rect = node.getBoundingClientRect();
                    if (
                        style.visibility === 'hidden' ||
                        style.display === 'none' ||
                        rect.width === 0 ||
                        rect.height === 0
                    ) return null;
                    return rect;
                };

                const matches = [];
                const nodes = Array.from(document.querySelectorAll('body *'));
                for (const header of headers) {
                    const node = nodes.find((candidate) => {
                        const text = (candidate.textContent || '').replace(/\\s+/g, '').trim();
                        return text === header && visibleRect(candidate);
                    });
                    if (!node) continue;

                    const grid = node.closest('.ag-root, [role="grid"], table, [class*="Grid"], [class*="grid"]');
                    const rect = visibleRect(grid || node);
                    if (rect) {
                        matches.push({
                            x: rect.x,
                            y: rect.y,
                            width: rect.width,
                            height: rect.height,
                            right: rect.right,
                            bottom: rect.bottom
                        });
                    }
                }

                if (matches.length < 4) return null;

                const largest = matches
                    .slice()
                    .sort((a, b) => (b.width * b.height) - (a.width * a.height))[0];
                if (largest.width > 250 && largest.height > 80) return largest;

                const left = Math.min(...matches.map((rect) => rect.x));
                const top = Math.min(...matches.map((rect) => rect.y));
                const right = Math.max(...matches.map((rect) => rect.right));
                const bottom = Math.max(...matches.map((rect) => rect.bottom));
                return { x: left, y: top, width: right - left, height: Math.max(120, bottom - top) };
            }
            """,
            headers,
        )
        if box:
            return box
        page.wait_for_timeout(300)
    return None


def open_context_menu_excel_download(page, download_dir: Path) -> Path:
    if not results_table_box(page):
        DEBUG_TEXT.write_text(page.locator("body").inner_text(timeout=5000), encoding="utf-8")
        DEBUG_HTML.write_text(page.content(), encoding="utf-8")
        page.screenshot(path=DEBUG_SCREENSHOT, full_page=True)
        raise RuntimeError(
            "기준일/유형/예탁원펀드코드/투자일/펀드명/일수익률 헤더가 있는 결과 테이블을 찾지 못했습니다."
        )

    cell = first_visible(
        page,
        [
            "td[data-column-name='TRADEDAY']",
            "td[data-column-name='FUNDKNAME']",
            "td[data-column-name='RET']",
            ".tui-grid-cell",
        ],
        timeout=10000,
    )
    cell.click(button="right")
    page.wait_for_timeout(500)

    with page.expect_download(timeout=30000) as download_info:
        for text in ["엑셀다운로드", "엑셀 다운로드", "Excel 다운로드", "Excel Export", "Export to Excel", "엑셀"]:
            try:
                click_text(page, text, timeout=2500)
                break
            except Exception:
                continue
        else:
            DEBUG_TEXT.write_text(page.locator("body").inner_text(timeout=5000), encoding="utf-8")
            DEBUG_HTML.write_text(page.content(), encoding="utf-8")
            page.screenshot(path=DEBUG_SCREENSHOT, full_page=True)
            raise RuntimeError("우클릭 메뉴에서 엑셀 다운로드 항목을 찾지 못했습니다.")

    download = download_info.value
    suggested_name = download.suggested_filename or DOWNLOAD_NAME
    temp_path = download_dir / f"__download_{datetime.now():%Y%m%d_%H%M%S}_{suggested_name}"
    download.save_as(temp_path)
    return temp_path


def run(start_date: str, end_date: str, headless: bool) -> Path:
    try:
        from playwright.sync_api import sync_playwright
    except ModuleNotFoundError as exc:
        raise RuntimeError(
            "playwright가 설치되어 있지 않습니다. "
            "`python -m pip install playwright` 및 "
            "`python -m playwright install chromium`을 먼저 실행하세요."
        ) from exc

    user_id = env_required("KFROM_ID")
    password = env_required("KFROM_PASSWORD")

    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(headless=headless)
        context = browser.new_context(accept_downloads=True)
        page = context.new_page()
        page.goto(SITE_URL, wait_until="domcontentloaded", timeout=60000)

        login(page, user_id, password)
        page.wait_for_load_state("networkidle", timeout=60000)

        navigate_to_report(page)
        fill_date(page, "시작", start_date)
        fill_date(page, "종료", end_date)
        click_text(page, "검색")
        page.wait_for_load_state("networkidle", timeout=60000)
        page.wait_for_timeout(1500)

        downloaded = open_context_menu_excel_download(page, BASE_DIR)
        context.close()
        browser.close()

    output = BASE_DIR / DOWNLOAD_NAME
    if output.exists():
        backup = BASE_DIR / f"{output.stem}_backup_{datetime.now():%Y%m%d_%H%M%S}{output.suffix}"
        shutil.move(str(output), str(backup))
    shutil.move(str(downloaded), str(output))
    return output


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-date", default=DEFAULT_START_DATE)
    parser.add_argument("--end-date", default=previous_business_day().strftime("%Y-%m-%d"))
    parser.add_argument("--headed", action="store_true", help="브라우저 화면을 보면서 실행합니다.")
    args = parser.parse_args()

    output = run(
        start_date=normalize_date(args.start_date),
        end_date=normalize_date(args.end_date),
        headless=not args.headed,
    )
    print(output)


if __name__ == "__main__":
    main()
