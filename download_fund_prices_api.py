"""Download one daily KFR fund-prices JSON snapshot."""

from __future__ import annotations

import argparse
import json
import os
import urllib.error
import urllib.parse
import urllib.request
from datetime import date, datetime, timedelta
from pathlib import Path
from zoneinfo import ZoneInfo


BASE_URL = "https://apiservice.kfr.co.kr"
OUTPUT_DIR = Path(__file__).resolve().parent / "data" / "kfr"
KST = ZoneInfo("Asia/Seoul")
REQUIRED_FIELDS = {
    "trade_day", "cm_seq", "composite_name", "fund_ksd_code", "invest_day", "fund_k_name",
    "ret", "price", "prev_price", "share_rate", "cul_ret", "kospi", "kosdaq", "sp500",
    "nasdaq", "kobi30", "kobi120", "econo_idx",
}


def previous_business_day() -> date:
    current = datetime.now(KST).date() - timedelta(days=1)
    while current.weekday() >= 5:
        current -= timedelta(days=1)
    return current


def request_json(request: urllib.request.Request) -> dict:
    try:
        with urllib.request.urlopen(request, timeout=90) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"KFR HTTP {exc.code}: {detail[:1000]}") from exc
    if not isinstance(payload, dict):
        raise RuntimeError("KFR response is not a JSON object")
    return payload


def validate_prices_payload(payload: dict, trade_day: str) -> list[dict]:
    rows = payload.get("content")
    if not isinstance(rows, list) or any(not isinstance(row, dict) for row in rows):
        raise RuntimeError("KFR prices content is not a list of objects")
    if not rows or {str(row.get("trade_day") or "")[:10] for row in rows} != {trade_day}:
        raise RuntimeError(f"KFR prices has no complete business data for {trade_day}")
    for index, row in enumerate(rows, start=1):
        missing = REQUIRED_FIELDS - set(row)
        if missing:
            raise RuntimeError(
                f"KFR prices row {index} fields missing: {', '.join(sorted(missing))}"
            )
    if payload.get("total_elements") not in (None, len(rows)):
        raise RuntimeError("KFR prices total_elements does not match content")
    return rows


def token() -> str:
    app_key = os.environ.get("KFR_APP_KEY_ID", "").strip()
    app_secret = os.environ.get("KFR_APP_KEY_SECRET", "").strip()
    if not app_key or not app_secret:
        raise RuntimeError("KFR_APP_KEY_ID and KFR_APP_KEY_SECRET are required")
    body = json.dumps({"app_key_id": app_key, "app_key_secret": app_secret}).encode("utf-8")
    payload = request_json(urllib.request.Request(
        f"{BASE_URL}/v1/auth/token", data=body, method="POST",
        headers={"Accept": "application/json", "Content-Type": "application/json"},
    ))
    access_token = payload.get("access_token")
    if not isinstance(access_token, str) or not access_token:
        raise RuntimeError("Token response did not contain access_token")
    return access_token


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--date", default=previous_business_day().isoformat())
    parser.add_argument("--output-dir", type=Path, default=OUTPUT_DIR)
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="Exit successfully only when the local daily JSON already exists and is valid",
    )
    args = parser.parse_args()
    date.fromisoformat(args.date)
    target = args.output_dir / f"prices_{args.date}.json"
    if args.check_only:
        if not target.is_file():
            raise FileNotFoundError(f"missing KFR prices JSON: {target}")
        payload = json.loads(target.read_text(encoding="utf-8-sig"))
        if not isinstance(payload, dict):
            raise RuntimeError(f"KFR prices JSON is not an object: {target}")
        rows = validate_prices_payload(payload, args.date)
        print(f"valid {target}: rows={len(rows)}")
        return

    query = urllib.parse.urlencode({"tradeDay": args.date})
    payload = request_json(urllib.request.Request(
        f"{BASE_URL}/v1/hbank/funds/prices?{query}", method="GET",
        headers={"Accept": "application/json", "Authorization": f"Bearer {token()}"},
    ))
    rows = payload.get("content")
    if not isinstance(rows, list):
        raise RuntimeError("KFR prices content is not a list")
    if not rows or {str(row.get('trade_day') or '')[:10] for row in rows} != {args.date}:
        print(f"no KFR business data for {args.date}; download skipped")
        return
    rows = validate_prices_payload(payload, args.date)
    args.output_dir.mkdir(parents=True, exist_ok=True)
    target.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"created {target}: rows={len(rows)}")


if __name__ == "__main__":
    main()
