#!/usr/bin/env python3
"""
ID Document Downloader + WMS Uploader
======================================
Reads a waybill range (--start / --end) or xlsx file, looks up buyer name +
mobile from phones.txt, downloads face + back ID images to a daily subfolder,
scrapes the ID number and validity period from the ID viewer page, then
uploads everything to the WMS API (api/Open/IdcardAdd).

Quick start
-----------
1.  pip install -r requirements.txt
2.  playwright install chromium
3.  Edit phones.txt with this run's waybill → name + mobile data
4.  python automation.py --start JR25135001E --end JR25135050E
"""

import argparse
import asyncio
import base64
import sys
import hashlib
import json
import re
import os
import logging
import random
import string
import time
from pathlib import Path
from datetime import datetime

import openpyxl
import requests as _sync_requests
from playwright.async_api import async_playwright, TimeoutError as PlaywrightTimeout


# ════════════════════════════════════════════════════════════════════════════
# CONFIGURATION  — edit before first run
# ════════════════════════════════════════════════════════════════════════════

_AUTOMATION_DIR = Path(__file__).resolve().parent

XLSX_PATH          = _AUTOMATION_DIR / "testorders.xlsx"
OUTPUT_DIR         = r"C:\Users\Lynn\Documents\JJ\身份证"
CHROME_PROFILE_DIR = _AUTOMATION_DIR / "chrome_profile"
PHONES_PATH        = _AUTOMATION_DIR / "phones.txt"
ORDER_SEARCH_URL   = "https://myseller.taobao.com/home.htm/trade-platform/tp/sold"

# Daily output subfolder created once at startup, e.g. "22APR"
_TODAY_LABEL = datetime.now().strftime("%d%b").upper()
OUTPUT_DAILY = Path(OUTPUT_DIR) / _TODAY_LABEL

# ── WMS API credentials ───────────────────────────────────────────────────────
API_BASE      = "http://open-api.auodexpress.com"
API_CONSIGNOR = "HZ20250346"
API_TOKEN     = "b8f13aeb-7135-2633-8b15-cd5d38d3ae27"


# ── Selectors ─────────────────────────────────────────────────────────────────
SELECTORS = {
    # Search page
    "search_input":  "input[aria-label='搜索']",
    "search_button": "span.next-btn-helper:has-text('搜索订单')",
    "view_id_link":  "a[href*='global-buy-idcards']",

    # Eye icons on the ID viewer page (姓名 / 证件号码)
    "reveal_name_icon":   ".next-row:has(.next-col-2:has-text('姓名')) span.icon[aria-haspopup='true']",
    "reveal_idcard_icon": ".next-row:has(.next-col-2:has-text('证件号码')) span.icon[aria-haspopup='true']",

    # Balloon that appears after clicking either eye icon
    "balloon_text": ".next-balloon-content",

    # Validity period — typically plain text (not hidden) in the 有效期 row
    "validity_period_row": ".next-row:has(.next-col-2:has-text('有效期'))",

    # Image viewer links
    "face_id_icon": "xpath=//div[contains(.,'查看正面') and contains(.,'查看反面')]//a[@rel='noreferrer'][1]",
    "back_id_icon": "xpath=//div[contains(.,'查看正面') and contains(.,'查看反面')]//a[@rel='noreferrer'][2]",
}

# Timeouts (milliseconds)
PAGE_TIMEOUT    = 30_000
ELEMENT_TIMEOUT = 15_000
CONCURRENCY     = 2


# ════════════════════════════════════════════════════════════════════════════
# internals
# ════════════════════════════════════════════════════════════════════════════

LOG_PATH = os.path.join(OUTPUT_DIR, f"log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt")

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s  %(levelname)-5s  %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

_new_tab_lock = asyncio.Lock()
_output_lock  = asyncio.Lock()


# ── Human-like interaction helpers ────────────────────────────────────────────

async def human_sleep(min_s: float, max_s: float) -> None:
    await asyncio.sleep(random.uniform(min_s, max_s))


async def human_click(locator, page, timeout: int = 15_000) -> None:
    await locator.wait_for(state="visible", timeout=timeout)
    box = await locator.bounding_box()
    if box:
        tx = box["x"] + box["width"]  * random.uniform(0.3, 0.7)
        ty = box["y"] + box["height"] * random.uniform(0.3, 0.7)
        await page.mouse.move(
            tx + random.uniform(-40, 40),
            ty + random.uniform(-20, 20),
            steps=random.randint(4, 10),
        )
        await human_sleep(0.05, 0.15)
        await page.mouse.move(tx, ty, steps=random.randint(3, 6))
        await human_sleep(0.04, 0.10)
        await page.mouse.click(tx, ty)
    else:
        await locator.click(timeout=timeout)


# ── Tracking number extraction ────────────────────────────────────────────────

def get_tracking_numbers(xlsx_path: str | Path) -> list[str]:
    pattern = re.compile(r"^JR251\d+E?$")
    found: list[str] = []
    wb = openpyxl.load_workbook(xlsx_path, read_only=True, data_only=True)
    for sheet in wb.worksheets:
        for row in sheet.iter_rows(values_only=True):
            for val in row:
                if isinstance(val, str) and pattern.match(val.strip()):
                    found.append(val.strip())
    wb.close()
    logger.info(f"Found {len(found)} tracking numbers in xlsx")
    return found


def generate_tracking_range(start: str, end: str) -> list[str]:
    pat = re.compile(r"^([A-Za-z]*)(\d+)([A-Za-z]*)$")
    ms, me = pat.match(start.strip()), pat.match(end.strip())
    if not ms or not me:
        raise ValueError(f"Cannot parse tracking numbers: {start!r}, {end!r}")
    prefix, start_digits, suffix = ms.group(1), ms.group(2), ms.group(3)
    prefix2, end_digits, suffix2 = me.group(1), me.group(2), me.group(3)
    if prefix != prefix2 or suffix != suffix2:
        raise ValueError(f"Start and end have different prefix/suffix: {start!r} vs {end!r}")
    n_start, n_end = int(start_digits), int(end_digits)
    if n_start > n_end:
        raise ValueError(f"Start {start!r} is after end {end!r}")
    width = len(start_digits)
    numbers = [f"{prefix}{n:0{width}d}{suffix}" for n in range(n_start, n_end + 1)]
    logger.info(f"Generated {len(numbers)} tracking numbers ({start} → {end})")
    return numbers


# ── phones.txt parser ─────────────────────────────────────────────────────────

_MOBILE_RE  = re.compile(r"^1[3-9]\d{9}$")
_JR_CODE_RE = re.compile(r"^JR\d+E?$")

def parse_phones_file(path: Path) -> dict[str, dict]:
    """
    Parse phones.txt into {waybill: {'name': str, 'mobile': str}}.

    Entries may or may not be separated by blank lines. Entry boundaries are
    detected by a short numeric sequence number whose next non-empty line is
    a JR waybill code. Mobile is identified by pattern (11-digit Chinese
    mobile starting with 1[3-9]); name is the line immediately before it.
    Robust to extra status lines (e.g. 已揽收) between name and mobile.
    """
    mapping: dict[str, dict] = {}
    try:
        text = Path(path).read_text(encoding="utf-8")
    except FileNotFoundError:
        logger.warning(f"phones.txt not found at {path} — name/mobile mapping unavailable")
        return mapping

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    # Find indices where each new entry starts (digit-only line followed by JR code)
    starts = [
        i for i in range(len(lines) - 1)
        if re.match(r"^\d{1,5}$", lines[i]) and _JR_CODE_RE.match(lines[i + 1])
    ]

    for j, start in enumerate(starts):
        end   = starts[j + 1] if j + 1 < len(starts) else len(lines)
        entry = lines[start:end]

        waybill = entry[1]
        mobile = name = ""
        # Name is always the line immediately after the print-status field
        for i, line in enumerate(entry):
            if line in ("已打印", "未打印") and i + 1 < len(entry):
                name = entry[i + 1]
                break
        # Mobile is the first 11-digit number starting with 1[3-9]
        for line in entry:
            if _MOBILE_RE.match(line):
                mobile = line
                break

        if waybill and name and mobile:
            mapping[waybill] = {"name": name, "mobile": mobile}
            logger.info(f"  phones  : {waybill} → {name} / {mobile}")
        else:
            logger.warning(
                f"  phones  : could not parse entry — "
                f"waybill={waybill!r} name={name!r} mobile={mobile!r}"
            )

    logger.info(f"Loaded {len(mapping)} entries from phones.txt")

    # ── Validation ────────────────────────────────────────────────────────────
    _STATUS_WORDS = {
        "已揽收", "已提交", "已打印", "直邮", "仓配",
        "已签收", "已发货", "运输中", "已入库", "已退回",
    }
    problems: list[str] = []
    for waybill, info in mapping.items():
        name, mobile = info["name"], info["mobile"]
        if not name:
            problems.append(f"  {waybill}: name is empty")
        elif name in _STATUS_WORDS:
            problems.append(f"  {waybill}: name={name!r} looks like a status/logistics word")
        elif any(ch.isdigit() for ch in name):
            problems.append(f"  {waybill}: name={name!r} contains digits")
        elif len(name) > 10:
            problems.append(f"  {waybill}: name={name!r} is suspiciously long (>{10} chars)")
        if not _MOBILE_RE.match(mobile):
            problems.append(f"  {waybill}: mobile={mobile!r} is not a valid 11-digit mobile")

    if problems:
        logger.warning("phones.txt has suspicious entries:")
        for p in problems:
            logger.warning(p)
        answer = input("Continue anyway? [Y/N]: ").strip().upper()
        if answer != "Y":
            logger.error("Aborted.")
            sys.exit(1)
        logger.info("Continuing with suspicious entries as-is.")

    return mapping


# ── Waybill resolution ───────────────────────────────────────────────────────

def _parse_tracking_parts(s: str) -> tuple[str, int, str]:
    m = re.match(r"^([A-Za-z]*)(\d+)([A-Za-z]*)$", s.strip())
    if not m:
        raise ValueError(f"Cannot parse tracking number: {s!r}")
    return m.group(1), int(m.group(2)), m.group(3)


def resolve_waybill(tracking_num: str, phones_map: dict) -> tuple[str, dict]:
    """
    Return (canonical_waybill, phones_entry).
    canonical_waybill is the exact key from phones.txt when a match exists
    (with or without trailing E). Use it for search, API upload, and logging.
    """
    for candidate in (tracking_num, tracking_num[:-1] if tracking_num.endswith("E") else tracking_num + "E"):
        if candidate in phones_map:
            return candidate, phones_map[candidate]
    return tracking_num, {}


def phones_waybills_in_range(phones_map: dict, start: str, end: str) -> list[str]:
    """Return phones.txt waybill keys whose numeric segment is within start..end."""
    prefix, n_start, _ = _parse_tracking_parts(start)
    prefix2, n_end, _ = _parse_tracking_parts(end)
    if prefix != prefix2:
        raise ValueError(f"Start and end have different prefix: {start!r} vs {end!r}")
    if n_start > n_end:
        raise ValueError(f"Start {start!r} is after end {end!r}")

    matched = []
    for wb in phones_map:
        p, n, _ = _parse_tracking_parts(wb)
        if p == prefix and n_start <= n <= n_end:
            matched.append(wb)
    matched.sort(key=lambda wb: _parse_tracking_parts(wb)[1])
    return matched


# ── Validity period parser ────────────────────────────────────────────────────

def parse_validity_end_date(text: str) -> str:
    """
    Extract the end/expiry date from various validity period formats and
    return it as YYYY-MM-DD.

    Handles: "20160913-20360913", "2016-09-13至2036-09-13",
             "2037-05-08", "20370508", "长期"
    """
    text = text.strip()
    if not text:
        return ""
    if "长期" in text:
        return "长期"

    # "YYYYMMDD-YYYYMMDD" — take the second
    m = re.search(r"\d{8}-(\d{8})$", text)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"

    # "YYYY-MM-DD至YYYY-MM-DD" — take the second
    m = re.search(r"\d{4}-\d{2}-\d{2}至(\d{4}-\d{2}-\d{2})", text)
    if m:
        return m.group(1)

    # Standalone "YYYY-MM-DD"
    m = re.search(r"\d{4}-\d{2}-\d{2}", text)
    if m:
        return m.group(0)

    # Standalone "YYYYMMDD"
    m = re.search(r"(\d{8})", text)
    if m:
        d = m.group(1)
        return f"{d[:4]}-{d[4:6]}-{d[6:8]}"

    return text  # fallback: return as-is


# ── Output path logic ──────────────────────────────────────────────────────────

def get_output_paths(name: str) -> tuple[Path, Path]:
    """Returns (front_path, back_path) under OUTPUT_DAILY. Overwrites if already exists."""
    base = OUTPUT_DAILY
    base.mkdir(parents=True, exist_ok=True)
    return base / f"{name}.jpg", base / f"{name}1.jpg"


async def get_output_paths_async(name: str) -> tuple[Path, Path]:
    async with _output_lock:
        return get_output_paths(name)


# ── New-tab helpers ───────────────────────────────────────────────────────────

async def find_content_frame(page, selector: str, timeout_ms: int = 5_000):
    for scan in range(2):
        try:
            await page.locator(selector).first.wait_for(state="attached", timeout=timeout_ms)
            return page
        except Exception:
            pass
        for frame in page.frames:
            if frame == page.main_frame:
                continue
            try:
                await frame.locator(selector).first.wait_for(state="attached", timeout=timeout_ms)
                return frame
            except Exception:
                continue
        if scan == 0:
            await asyncio.sleep(3.0)
    logger.warning(
        f"  '{selector}' not found in page or any of {len(page.frames)} frame(s) — using page"
    )
    return page


async def open_new_tab(ctx, page, selector: str):
    for attempt in range(3):
        new_tab = None
        url_before = page.url
        try:
            async with _new_tab_lock:
                pages_before = set(ctx.pages)
                await human_click(page.locator(selector).first, page, timeout=15_000)
                # Poll up to 10 s for a new tab
                for _ in range(100):
                    new_pages = set(ctx.pages) - pages_before
                    if new_pages:
                        new_tab = next(iter(new_pages))
                        break
                    await asyncio.sleep(0.1)

                if not new_tab:
                    # Taobao may have navigated in the same tab instead of opening a new one
                    await asyncio.sleep(0.5)
                    if page.url != url_before:
                        logger.info(f"  Link navigated same tab → {page.url[:80]}")
                        new_tab = page  # treat current page as id_page; caller must NOT close it

            if not new_tab:
                logger.warning(
                    f"  No new tab appeared (attempt {attempt + 1}/3) — "
                    f"pages in ctx: {len(ctx.pages)}, page url: {page.url[:60]}"
                )
                await human_sleep(1.5, 3.0)
                continue

            if new_tab is not page:
                logger.info(f"  New tab opened: {new_tab.url[:80]}")
                try:
                    await new_tab.wait_for_load_state("load", timeout=PAGE_TIMEOUT)
                except PlaywrightTimeout:
                    logger.warning("  Tab load timed out — continuing anyway")
            return new_tab

        except Exception as exc:
            logger.warning(f"  Open-tab attempt {attempt + 1}/3 failed: {str(exc)[:140]}")
            if attempt < 2:
                await human_sleep(1.5, 3.0)

    logger.error(f"  Failed to open new tab after 3 attempts ({selector})")
    return None


async def capture_image(ctx, content, selector: str, id_page) -> bytes | None:
    image_future: asyncio.Future[bytes] = asyncio.get_running_loop().create_future()

    async def on_page_response(response):
        ct = response.headers.get("content-type", "")
        if ct.startswith("image/jpeg") and not image_future.done():
            try:
                body = await response.body()
                if len(body) >= 20_000 and not image_future.done():
                    image_future.set_result(body)
            except Exception:
                pass

    id_page.on("response", on_page_response)
    new_tab = None

    try:
        async with _new_tab_lock:
            pages_before = set(ctx.pages)
            await human_click(content.locator(selector).first, id_page, timeout=15_000)
            # Poll up to 3 s; a timeout here just means the image loads inline (Mode B)
            new_tab = None
            for _ in range(30):
                new_pages = set(ctx.pages) - pages_before
                if new_pages:
                    new_tab = next(iter(new_pages))
                    break
                await asyncio.sleep(0.1)

        if new_tab:
            async def on_tab_response(response):
                ct = response.headers.get("content-type", "")
                if ct.startswith("image/") and not image_future.done():
                    try:
                        body = await response.body()
                        if len(body) >= 20_000 and not image_future.done():
                            image_future.set_result(body)
                    except Exception:
                        pass
            new_tab.on("response", on_tab_response)
            try:
                await new_tab.wait_for_load_state("load", timeout=PAGE_TIMEOUT)
            except PlaywrightTimeout:
                pass
        else:
            await human_sleep(3.5, 5.5)

        try:
            body = await asyncio.wait_for(asyncio.shield(image_future), timeout=4.0)
            return body
        except asyncio.TimeoutError:
            logger.warning("    Response interceptor: no image yet, trying img src…")

        img_targets = [new_tab] if new_tab else ([content, id_page] if content is not id_page else [id_page])
        for tgt in img_targets:
            try:
                img_el = tgt.locator("img").first
                if await img_el.count() > 0:
                    src = await img_el.get_attribute("src")
                    if src:
                        result = await id_page.evaluate(
                            """async (url) => {
                                const r = await fetch(url, {credentials: 'include'});
                                const ab = await r.arrayBuffer();
                                return Array.from(new Uint8Array(ab));
                            }""",
                            src,
                        )
                        result_bytes = bytes(result)
                        if len(result_bytes) >= 20_000:
                            return result_bytes
            except Exception:
                continue

        for tgt in img_targets:
            try:
                img_el = tgt.locator("img").first
                if await img_el.count() > 0:
                    return await img_el.screenshot()
            except Exception:
                continue

        return None

    finally:
        id_page.remove_listener("response", on_page_response)
        if new_tab:
            try:
                await new_tab.close()
            except Exception:
                pass


# ── WMS API upload ────────────────────────────────────────────────────────────

def _make_nonce() -> str:
    return "".join(random.choices(string.ascii_lowercase + string.digits, k=11))


def _sign(consignor: str, token: str, ts: str, nonce: str, body_json: str) -> str:
    combined = consignor + token + ts + nonce + body_json
    return hashlib.md5("".join(sorted(combined)).encode("utf-8")).hexdigest().upper()


def upload_idcard(
    tracking_num: str,
    name: str,
    mobile: str,
    idcard_code: str,
    validity_period: str,
    face_bytes: bytes,
    back_bytes: bytes,
) -> tuple[bool, str]:
    """Synchronous API call — run via run_in_executor from async code."""
    payload = {
        "fullName":          name,
        "idCardCode":        idcard_code,
        "positiveServerUrl": base64.b64encode(face_bytes).decode("ascii"),
        "negativeServerUrl": base64.b64encode(back_bytes).decode("ascii"),
        "validityPeriod":    validity_period,
        "mobile":            mobile,
        "wayBillCode":       tracking_num,
    }
    body_json = json.dumps(payload, ensure_ascii=False, separators=(",", ":"))
    ts    = str(int(time.time() * 1000))
    nonce = _make_nonce()
    headers = {
        "Content-Type":  "application/json",
        "apiType":       "consignor",
        "consignorCode": API_CONSIGNOR,
        "token":         API_TOKEN,
        "timestamp":     ts,
        "nonce":         nonce,
        "signature":     _sign(API_CONSIGNOR, API_TOKEN, ts, nonce, body_json),
    }
    try:
        resp = _sync_requests.post(
            f"{API_BASE}/api/Open/IdcardAdd",
            data=body_json.encode("utf-8"),
            headers=headers,
            timeout=30,
        )
    except Exception as exc:
        return False, f"HTTP request failed: {exc}"

    raw = resp.content
    try:
        parsed = json.loads(raw.decode("gbk"))
        return parsed.get("result", False), parsed.get("msg", "unknown")
    except Exception:
        # Server returned non-JSON — surface the status code and raw body
        preview = raw[:200].decode("utf-8", errors="replace").strip()
        return False, f"HTTP {resp.status_code}, non-JSON response: {preview!r}"


# ── Error logging ─────────────────────────────────────────────────────────────

def write_error(err_log, tracking_num: str, name: str, reason: str) -> None:
    line = f"{tracking_num}\t{name}\t{reason}"
    logger.warning(f"  SKIP  {line}")
    err_log.write(line + "\n")
    err_log.flush()


# ── Balloon reveal helper ─────────────────────────────────────────────────────

async def reveal_field(content, icon_selector: str, id_page, label: str) -> str:
    """
    Click the eye icon for `label`, read the balloon text, then dismiss.
    Retries up to 3×. Returns the revealed text, or "" on failure.
    """
    value = ""
    for attempt in range(3):
        try:
            await human_click(content.locator(icon_selector).first, id_page)
            value = (
                await content.locator(SELECTORS["balloon_text"]).text_content(timeout=8_000) or ""
            ).strip()
        except Exception:
            pass
        if value:
            break
        if attempt < 2:
            logger.warning(f"  {label} balloon didn't appear, retrying ({attempt + 2}/3)…")
            await human_sleep(1.2, 2.5)

    # Dismiss balloon so the next icon click works cleanly
    try:
        await id_page.keyboard.press("Escape")
        await human_sleep(0.3, 0.6)
    except Exception:
        pass

    return value


# ── Per-order workflow ────────────────────────────────────────────────────────

async def process_one(
    page, tracking_num: str, phones_map: dict
) -> tuple[bool | None, str]:
    """
    Returns (result, reason):
        (True,  "")       — success
        (False, reason)   — non-retryable skip (no id, already exists)
        (None,  reason)   — retryable failure
    Never writes to the error log — that is the worker's responsibility.
    """
    ctx = page.context
    effective_num, phones_entry = resolve_waybill(tracking_num, phones_map)
    display_name = phones_entry.get("name", "?")
    mobile       = phones_entry.get("mobile", "")
    if effective_num != tracking_num:
        logger.info(f"  waybill : {tracking_num} → {effective_num} (using phones.txt)")

    try:
        # 1 ── Search ────────────────────────────────────────────────────────
        await page.goto(ORDER_SEARCH_URL, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT)
        await page.wait_for_selector(
            SELECTORS["search_input"], state="visible", timeout=PAGE_TIMEOUT
        )
        await page.click(SELECTORS["search_input"])
        await page.keyboard.press("Control+a")
        await human_sleep(0.1, 0.25)
        await page.type(SELECTORS["search_input"], effective_num, delay=random.randint(60, 140))
        await human_sleep(0.3, 0.7)
        await human_click(page.locator(SELECTORS["search_button"]).first, page)
        await page.wait_for_load_state("load", timeout=PAGE_TIMEOUT)
        await human_sleep(1.2, 2.5)

        # 2 ── Check results ─────────────────────────────────────────────────
        id_link_loc = page.locator(SELECTORS["view_id_link"])
        if await id_link_loc.count() == 0:
            if await page.locator("text=请联系消费者完成实名认证").count() > 0:
                return False, "no id"
            return False, "请确认订单信息"

        # 3 ── Open ID viewer page (Tab 1) ───────────────────────────────────
        id_page = await open_new_tab(ctx, page, SELECTORS["view_id_link"])
        if not id_page:
            return None, "could not open ID viewer page"

        try:
            await human_sleep(1.2, 2.2)
            content = await find_content_frame(id_page, SELECTORS["reveal_name_icon"])

            # 4 ── Reveal buyer name ──────────────────────────────────────────
            idcard_name = await reveal_field(
                content, SELECTORS["reveal_name_icon"], id_page, "name"
            )
            if not idcard_name:
                return None, "failed to reveal name on ID page"
            if display_name != "?" and idcard_name != display_name:
                logger.warning(
                    f"  NAME MISMATCH  phones.txt={display_name!r}  id_page={idcard_name!r}"
                    f"  — using id_page name for upload"
                )
            logger.info(f"  name    : {idcard_name}")

            # 5 ── Reveal ID card number ──────────────────────────────────────
            idcard_code = await reveal_field(
                content, SELECTORS["reveal_idcard_icon"], id_page, "ID number"
            )
            if not idcard_code:
                return None, "failed to reveal ID card number"
            logger.info(f"  id code : {idcard_code}")

            # 6 ── Get validity period ────────────────────────────────────────
            validity_raw = ""
            try:
                row = content.locator(SELECTORS["validity_period_row"])
                if await row.count() > 0:
                    row_text = (await row.inner_text(timeout=5_000) or "").strip()
                    validity_raw = re.sub(r"^有效期\s*", "", row_text).strip()
            except Exception as exc:
                logger.debug(f"  validity period direct read failed: {exc}")

            validity_period = parse_validity_end_date(validity_raw) if validity_raw else ""
            if not validity_period:
                return None, "failed to read validity period"
            logger.info(f"  validity: {validity_period}  (raw: {validity_raw!r})")

            # 7 ── Resolve output paths ───────────────────────────────────────
            front_path, back_path = await get_output_paths_async(idcard_name)

            # 8 ── Face image ─────────────────────────────────────────────────
            id_viewer_url = id_page.url
            face_bytes = await capture_image(
                ctx, content, SELECTORS["face_id_icon"], id_page
            )
            if face_bytes:
                front_path.write_bytes(face_bytes)
                logger.info(f"  face    → {front_path.name}")
            else:
                return None, "face image not captured"

            # 9 ── Back image ─────────────────────────────────────────────────
            await human_sleep(1.2, 2.5)
            if id_page.url != id_viewer_url:
                logger.info("  Page navigated during face capture — returning to ID viewer…")
                try:
                    await id_page.goto(
                        id_viewer_url, wait_until="domcontentloaded", timeout=PAGE_TIMEOUT
                    )
                except Exception as exc:
                    logger.warning(f"  Could not return to ID viewer: {exc}")
            content = await find_content_frame(id_page, SELECTORS["back_id_icon"], timeout_ms=8_000)
            back_bytes = await capture_image(
                ctx, content, SELECTORS["back_id_icon"], id_page
            )
            if back_bytes:
                back_path.write_bytes(back_bytes)
                logger.info(f"  back    → {back_path.name}")
            else:
                return None, "back image not captured"

            # 10 ── Upload to WMS API ─────────────────────────────────────────
            if not mobile:
                return None, "no mobile in phones.txt"

            ok, msg = await asyncio.get_running_loop().run_in_executor(
                None,
                lambda: upload_idcard(
                    effective_num, idcard_name, mobile,
                    idcard_code, validity_period,
                    face_bytes, back_bytes,
                ),
            )
            if ok:
                logger.info(f"  API     : uploaded OK")
                return True, ""
            # 535 = ID already exists — non-retryable
            non_retryable = "已存在" in (msg or "") or "535" in (msg or "")
            return (False if non_retryable else None), f"API upload failed: {msg}"

        finally:
            if id_page is not page:
                await id_page.close()

    except PlaywrightTimeout as exc:
        return None, f"timeout: {str(exc)[:120]}"
    except Exception as exc:
        return None, f"error: {str(exc)[:120]}"


# ── Worker ────────────────────────────────────────────────────────────────────

async def worker(ctx, queue: asyncio.Queue, err_log, results: list, phones_map: dict) -> None:
    page = await ctx.new_page()
    page.set_default_timeout(PAGE_TIMEOUT)
    try:
        while True:
            try:
                idx, total, num = queue.get_nowait()
            except asyncio.QueueEmpty:
                break
            waybill, phones_entry = resolve_waybill(num, phones_map)
            display_name = phones_entry.get("name", "?")
            if waybill != num:
                logger.info(f"[{idx}/{total}]  {waybill}  ({display_name})  [input: {num}]")
            else:
                logger.info(f"[{idx}/{total}]  {waybill}  ({display_name})")

            result, reason = await process_one(page, waybill, phones_map)
            if result is None:
                logger.info(f"  retrying in 2 s…")
                await human_sleep(1.5, 3.0)
                result, reason = await process_one(page, waybill, phones_map)
                if result is None:
                    reason = f"failed after retry: {reason}"

            # One log entry per waybill, written only on failure
            if result is not True:
                write_error(err_log, waybill, display_name, reason)

            results.append(result is True)
            await human_sleep(0.4, 1.2)
    finally:
        await page.close()


# ── Main ──────────────────────────────────────────────────────────────────────

async def main() -> None:
    parser = argparse.ArgumentParser(description="ID Document Downloader + WMS Uploader")
    src = parser.add_mutually_exclusive_group()
    src.add_argument("--xlsx", metavar="FILE",
                     help="Path to xlsx containing tracking numbers (default: testorders.xlsx)")
    src.add_argument("--start", metavar="NUM",
                     help="First tracking number in a range, e.g. JR25135001E")
    parser.add_argument("--end", metavar="NUM",
                        help="Last tracking number in range (required with --start)")
    args = parser.parse_args()

    phones_map = parse_phones_file(PHONES_PATH)

    if args.start:
        if not args.end:
            parser.error("--end is required when --start is used")
        try:
            numbers = phones_waybills_in_range(phones_map, args.start, args.end)
            if numbers:
                logger.info(
                    f"Using {len(numbers)} waybill(s) from phones.txt "
                    f"({numbers[0]} → {numbers[-1]})"
                )
            else:
                numbers = generate_tracking_range(args.start, args.end)
                logger.warning(
                    "No phones.txt entries in range — using generated tracking numbers; "
                    "upload/log will use phones.txt format only when a match exists"
                )
        except ValueError as exc:
            parser.error(str(exc))
    else:
        xlsx_path = Path(args.xlsx) if args.xlsx else XLSX_PATH
        numbers = get_tracking_numbers(xlsx_path)

    numbers = list(dict.fromkeys(
        resolve_waybill(num, phones_map)[0] for num in numbers
    ))

    # Create output dirs
    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)
    OUTPUT_DAILY.mkdir(parents=True, exist_ok=True)
    logger.info(f"Output folder: {OUTPUT_DAILY}")

    if not numbers:
        logger.error("No tracking numbers found — check your input.")
        return

    CHROME_PROFILE_DIR.mkdir(parents=True, exist_ok=True)
    logger.info("Launching browser…")

    async with async_playwright() as p:
        ctx = await p.chromium.launch_persistent_context(
            user_data_dir=str(CHROME_PROFILE_DIR),
            headless=False,
            viewport={"width": 1920, "height": 1080},
            user_agent=(
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/132.0.0.0 Safari/537.36"
            ),
            args=[
                "--window-size=1920,1080",
                "--no-first-run",
                "--no-default-browser-check",
                "--disable-blink-features=AutomationControlled",
                "--disable-infobars",
            ],
        )
        login_page = ctx.pages[0] if ctx.pages else await ctx.new_page()
        login_page.set_default_timeout(PAGE_TIMEOUT)

        await asyncio.get_event_loop().run_in_executor(
            None,
            input,
            "\nBrowser is open. Log in to the ordering system if needed,\n"
            "then press Enter here to start processing… ",
        )
        print()

        queue: asyncio.Queue = asyncio.Queue()
        for idx, num in enumerate(numbers, 1):
            await queue.put((idx, len(numbers), num))

        n_workers = min(CONCURRENCY, len(numbers))
        logger.info(f"Starting {n_workers} parallel worker(s) for {len(numbers)} order(s)…")

        results: list[bool] = []
        with open(LOG_PATH, "a", encoding="utf-8") as err_log:
            err_log.write(f"\n=== Run started {datetime.now()} ===\n")
            await asyncio.gather(
                *[asyncio.create_task(worker(ctx, queue, err_log, results, phones_map))
                  for _ in range(n_workers)]
            )

        ok_count   = sum(results)
        fail_count = len(results) - ok_count
        logger.info(
            f"Finished — {ok_count} OK  ·  {fail_count} issues  ·  log: {LOG_PATH}"
        )
        await ctx.close()


if __name__ == "__main__":
    asyncio.run(main())
