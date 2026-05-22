#!/usr/bin/env python3
"""
test_idcard_api.py — Manual one-shot test for api/Open/IdcardAdd

Intentionally omits ValidityPeriod and Mobile to probe which fields
the server actually enforces as required.

Usage:
    pip install requests
    python test_idcard_api.py
"""

import base64
import hashlib
import json
import random
import string
import time

import requests

# ── Credentials ───────────────────────────────────────────────────────────────
CONSIGNOR_CODE = "HZ20250346"
TOKEN          = "b8f13aeb-7135-2633-8b15-cd5d38d3ae27"
API_BASE       = "http://open-api.auodexpress.com"

# ── Test data ─────────────────────────────────────────────────────────────────
TRACKING        = "JR25135351E"
FULL_NAME       = "崔俊珏"
ID_CARD_CODE    = "140107198905113921"
FRONT_IMAGE     = r"C:\Users\Lynn\Documents\JJ\身份证\崔俊珏.jpg"
BACK_IMAGE      = r"C:\Users\Lynn\Documents\JJ\身份证\崔俊珏1.jpg"
VALIDITY_PERIOD = "长期"
MOBILE          = "13466865054"


# ── Signature ─────────────────────────────────────────────────────────────────

def _make_nonce() -> str:
    """11-character random alphanumeric (matches JS: Math.random().toString(36).substr(2))."""
    return ''.join(random.choices(string.ascii_lowercase + string.digits, k=11))


def _sign(consignor_code: str, token: str, timestamp: str, nonce: str, body_json: str) -> str:
    """
    Algorithm (from API docs):
      1. concat = consignorCode + token + timestamp + nonce + body_json
      2. sort every character in concat
      3. MD5(sorted_string).toUpperCase()

    body_json must be the exact JSON string being sent (compact, no spaces).
    """
    combined  = consignor_code + token + timestamp + nonce + body_json
    sorted_str = ''.join(sorted(combined))
    return hashlib.md5(sorted_str.encode("utf-8")).hexdigest().upper()


def build_headers(consignor_code: str, token: str, body_json: str) -> dict:
    timestamp = str(int(time.time() * 1000))   # 13-digit ms timestamp
    nonce     = _make_nonce()
    sig       = _sign(consignor_code, token, timestamp, nonce, body_json)
    return {
        "Content-Type":  "application/json",
        "apiType":       "consignor",
        "consignorCode": consignor_code,
        "token":         token,
        "timestamp":     timestamp,
        "nonce":         nonce,
        "signature":     sig,
    }


# ── Helpers ───────────────────────────────────────────────────────────────────

def image_to_base64(path: str) -> str:
    with open(path, "rb") as f:
        return base64.b64encode(f.read()).decode("ascii")


# ── Main ──────────────────────────────────────────────────────────────────────

def main():
    print("=== IdcardAdd API Test ===\n")

    # Build payload — doc field names with camelCase (lowercase first letter)
    payload: dict = {
        "fullName":          FULL_NAME,
        "idCardCode":        ID_CARD_CODE,
        "positiveServerUrl": image_to_base64(FRONT_IMAGE),
        "negativeServerUrl": image_to_base64(BACK_IMAGE),
        "validityPeriod":    VALIDITY_PERIOD,
        "wayBillCode":       TRACKING,
    }
    if MOBILE:
        payload["mobile"] = MOBILE

    # Serialize once — same string used for both signature and request body
    body_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))

    print("Payload fields sent:", [k for k in payload if k not in ("positiveServerUrl", "negativeServerUrl")],
          "(+ front/back images as base64)")
    print(f"Front image: {len(payload['positiveServerUrl'])} base64 chars")
    print(f"Back  image: {len(payload['negativeServerUrl'])} base64 chars\n")

    headers = build_headers(CONSIGNOR_CODE, TOKEN, body_json)
    print("Request headers:")
    for k, v in headers.items():
        print(f"  {k}: {v}")

    url = f"{API_BASE}/api/Open/IdcardAdd"
    print(f"\nPOST {url}\n")

    try:
        resp = requests.post(
            url,
            data=body_json.encode("utf-8"),   # send exact bytes used for signature
            headers=headers,
            timeout=30,
        )
        print(f"HTTP {resp.status_code}")
        print(f"Raw response: {resp.text}\n")

        try:
            decoded = resp.content.decode("gbk")
            parsed  = json.loads(decoded)
            print(f"Result     : {parsed.get('result')}")
            print(f"Msg        : {parsed.get('msg')}")
            print(f"statusCode : {parsed.get('statusCode')}")
        except Exception:
            print("(Could not parse response as GBK JSON)")

    except requests.exceptions.ConnectionError as exc:
        print(f"Connection error: {exc}")
    except Exception as exc:
        print(f"Request failed: {exc}")


if __name__ == "__main__":
    main()
