# ID Document Downloader + WMS Uploader

Automates the daily workflow of collecting buyer ID card images and getting them into the WMS system for customs clearance.

Given a list of waybills — from an Excel file, or a `--start`/`--end` range matched against `phones.txt` — the script:
1. Searches each order on `myseller.taobao.com`
2. Opens the buyer's identity verification page
3. Reveals the buyer's name, ID number, and ID validity period
4. Downloads the front and back ID card images
5. Saves them as `<name>.jpg` / `<name>1.jpg` in a dated output subfolder
6. Uploads name, ID number, validity period, mobile, and both images to the WMS API (`api/Open/IdcardAdd`)

If a buyer hasn't completed Taobao's real-name verification (`no id`), the script first tries to reuse a photo pair already on file for that exact name. If none exists, it uploads a random **old** photo pair as a temporary placeholder (to be replaced later) — see [Same-name-match fallback](#same-name-match-fallback-no-id) and [Random placeholder fallback](#random-placeholder-fallback-no-id-no-name-match). If the WMS API returns **HTTP 413**, the photos just saved are retried through the website wizard instead — see [HTTP 413 website fallback](#http-413-website-fallback).

---

## Requirements

- Python 3.10 or later
- Google Chrome (installed normally — the script drives your real Chrome profile)
- Windows (tested on Windows 11)

---

## Setup (first time only)

**1. Install Python dependencies**

```
pip install -r requirements.txt
```

**2. Install the Playwright browser driver**

```
playwright install chromium
```

**3. Edit the configuration at the top of `automation.py`**

| Setting | Default | Description |
|---|---|---|
| `XLSX_PATH` | `testorders.xlsx` | Full path to the Excel file containing tracking numbers |
| `OUTPUT_DIR` | `C:\Users\Lynn\Documents\JJ\身份证` | Root folder where ID images are saved (dated subfolders inside) |
| `CHROME_PROFILE_DIR` | `chrome_profile\` | Folder where your login session is stored |
| `PHONES_PATH` | `phones.txt` | Waybill → name/mobile lookup, pasted in before each run |
| `ORDER_SEARCH_URL` | myseller.taobao.com | The order search page URL |
| `CONCURRENCY` | `2` | How many Taobao orders are processed at the same time |
| `OLD_FOLDER_MIN_AGE_DAYS` | `10` | Placeholder photos must come from a subfolder older than this, or from files sitting directly in `OUTPUT_DIR` |
| `RANDOM_PLACEHOLDER_MAX_RETRIES` | `2` | Extra placeholder pairs to try after the first one fails (`no id - retried x times`) |
| `PLACEHOLDER_MAX_NAME_LEN` | `3` | Random placeholder pairs whose filename name is longer than this (e.g. `欧阳春晓`) are skipped |
| `PLACEHOLDER_MIN_IDLE_DAYS` | `10` | Random placeholder skips a pair if it was website-uploaded within this many days |

WMS API credentials (`API_CONSIGNOR`, `API_TOKEN`) also live at the top of `automation.py`.

---

## Running

**From an Excel file (default):**
```
python automation.py
python automation.py --xlsx path\to\orders.xlsx
```

**From a waybill range** (looked up against `phones.txt` first; falls back to a generated sequence if no entries match):
```
python automation.py --start JR25135001E --end JR25135050E
```

**Replaying a previous log** (see [Replay mode](#replay-mode---from-log) below). Logs live in `OUTPUT_DIR`; pass the full path:
```
python automation.py --from-log "C:\Users\Lynn\Documents\JJ\身份证\log_20260918_233507.txt"
```

A Chrome window will open in every mode (the website-upload fallback needs it). For a normal `--start`/`--xlsx` run, **log in to the seller platform** if prompted, then press **Enter** in the terminal to start processing. `--from-log` skips Taobao entirely and starts as soon as the browser is up. On future runs your login session is remembered automatically — just press Enter straight away.

---

## phones.txt format

Before each run, paste in this run's batch of waybill → buyer data (name, mobile, address) as copied from the order source. Entries don't need blank-line separators — the parser detects boundaries by a short numeric sequence number immediately followed by a `JR…` waybill code, and pulls the name from the line right after a `已打印`/`未打印` status marker and the mobile from the first 11-digit `1[3-9]…` number in the entry. Suspicious entries (empty/numeric/status-word names, invalid mobiles) are flagged with a Y/N confirmation prompt before the run continues.

---

## Output

Images are saved inside a dated subfolder of `OUTPUT_DIR` (e.g. `OUTPUT_DIR\2026-09-18\`, ISO format so it sorts correctly across year boundaries):

| File | Contents |
|---|---|
| `张三.jpg` | Front ID card |
| `张三1.jpg` | Back ID card |

If the same buyer appears more than once **on the same day**, the next pair is numbered `张三2.jpg` / `张三3.jpg`, then `张三4.jpg` / `张三5.jpg`, and so on. The same name can also legitimately reappear across different dated subfolders (a repeat customer) — this is what the same-name-match fallback searches across.

A timestamped log file is also written to `OUTPUT_DIR` (e.g. `log_20260302_090000.txt`) listing every order that wasn't uploaded via the normal API path.

---

## Error log

Every waybill that doesn't go through the normal WMS API upload gets exactly **one** tab-separated line, `<waybill>\t<name>\t<reason>` (the in-progress "found a same name match" / "trying placeholder" states are printed to the console for visibility but are not written to the file — only the final outcome is).

**Unsuccessful waybills** (no ID uploaded in the end — unhandled API/timeout errors, `no id` after placeholder retries, same-name or 413 website upload still failing, `请确认订单信息`, etc.) are prefixed with `**` so they stand out and are easy to grep. Successful fallback uploads (`same name match uploaded successfully`, `no id, uploaded … - to be replaced later`) are **not** prefixed. After the duplicate-name summary, the log may also list placeholder files that failed during this run (`**<filename> has been tried and reported error: …`); those extra lines are not waybill rows.

| Reason | Meaning |
|---|---|
| `**…	no id` | Buyer hasn't completed identity verification, no same-name photo match was found, and there was no eligible old pair left to try as a placeholder |
| `**…	no id - retried x times` | Same as above, but 1–3 random old pairs were tried (first attempt plus up to 2 retries) and all failed |
| `no id, uploaded <name>(<folder>) - to be replaced later` | No same-name match; a random old pair (`<name>` from folder `<folder>`, e.g. `2026-04-23`) was uploaded as a placeholder (not `**`-prefixed) |
| `same name match uploaded successfully` | Buyer hasn't completed verification, but a same-name photo match was found and the website-upload fallback succeeded (not `**`-prefixed) |
| `**…	same name match fail to upload` | Same-name website upload failed — `**`-prefixed like every other unsuccessful waybill |
| `**…	请确认订单信息` | Order not found or status unclear |
| `**…	failed after retry: API upload failed: HTTP 413, …` | REST API rejected the images as too large, **and** the [website 413 fallback](#http-413-website-fallback) could not find today's/yesterday's photos or the website upload also failed. A successful 413 fallback writes **no** log line. |
| `failed after retry: could not open ID viewer page` | ID viewer tab never opened after one retry. **Treated as `no id`** (same-name match, then random placeholder) — the log will show that flow's outcome, not this original text. |
| `**…	failed after retry: ...` | Any other unexpected error after one automatic retry (timeout, other HTTP errors, etc.) — logged as-is with the `**` prefix, no extra upload |

---

## Same-name-match fallback (`no id`)

The WMS API (`IdcardAdd`) requires an ID **number**, which we only get by reading it off the Taobao verification page — if the buyer never completed verification, we have no ID number, so the API path is a dead end even if we had a photo. To get around this without OCR'ing the number ourselves (unreliable — watermarks on some scans obscure digits), the script instead drives auodexpress.com's own public ID-upload wizard, which does the OCR server-side.

The same flow also runs when Taobao never opened the ID viewer (`failed after retry: could not open ID viewer page`), because that usually means there is no usable ID page either.

1. Look up the buyer's exact name (from `phones.txt`) in the same-name photo index (see below) — every complete front/back photo pair filed under that name, across every dated subfolder.
2. If **no** pair exists anywhere, fall through to the [random placeholder fallback](#random-placeholder-fallback-no-id-no-name-match) instead of giving up.
3. If one or more pairs exist, **pick one at random** (deliberate — spreads reuse across candidates rather than always resubmitting the same photo for a repeat name). "found a same name match" is printed to the console at this point, but not written to the log file yet — see [Error log](#error-log).
4. Upload that pair through `https://www.auodexpress.com/user.html#/upload-card-id`: upload front → upload back → click 开始识别身份证 (OCR) → fill in 运单号 with the waybill → click 确认提交.
5. **One attempt only.** Any failure at any step (image rejected, OCR timeout/error, submit error) gives up on that waybill immediately and logs `**...same name match fail to upload` — no retry, no placeholder, no partial credit.

This flow always runs **single-tab / sequential**, even while the normal Taobao workers run with `CONCURRENCY` in parallel — a deliberate scope decision to keep it easy to watch and debug while the flow is still being tuned, not a hard technical limit.

### Same-name photo index (`name_index_cache.json`)

Searching all ~40k+ files under `OUTPUT_DIR` from scratch on every lookup doesn't scale, so the name → photo-pairs index is cached on disk (`name_index_cache.json`, next to `automation.py`, gitignored) and built once at the start of every run:

- Every dated subfolder is scanned **once**, ever — the cache treats past folders as immutable, which holds in normal use (photos only get added, never removed).
- **Today's** subfolder is always rescanned each run, since it can still be actively growing.
- Folders that no longer exist on disk are dropped from the cache automatically.
- If you ever reorganize or delete old photos and need a full rebuild, just delete `name_index_cache.json` — it'll be rebuilt (scanning everything once) on the next run.
- If a cached pair is missing on disk when the script tries to use it (or while building the placeholder pool), that **one** pair is removed from `name_index_cache.json` and `photo_upload_history.json`. The waybill is still logged as a failure the same as before. You do not need to delete the whole cache for a few missing files.

### Photo upload history (`photo_upload_history.json`)

Every successful website-wizard upload (same-name match, random placeholder, or HTTP 413 fallback) is recorded in `photo_upload_history.json` (next to `automation.py`, gitignored). Each front-file path stores:

- `count` — how many times that pair has been uploaded through the wizard
- `last_uploaded` — local timestamp of the most recent success

On startup the console prints how many pairs are tracked and which file has been used the most. This history is **not** the same as `name_index_cache.json`: deleting photos does not rewrite history, and deleting the history file only resets counts (it does not rebuild the name index).

### Random placeholder fallback (`no id`, no name match)

If the same-name lookup finds nothing (the case that used to log a plain `no id`), the script uploads a **random existing photo pair** as a temporary stand-in for that waybill, through the same auodexpress.com wizard:

1. Eligible pairs are files sitting **directly** in `OUTPUT_DIR`, plus complete pairs inside any subfolder whose Windows creation time is **at least 10 days ago**. Folders created within the last 10 days (including today's dated folder) are skipped.
2. Pick one pair at random. Names longer than 3 characters (e.g. `欧阳春晓`) are never chosen. A pair uploaded through the website wizard in the **last 10 days** is never chosen. A pair that already succeeded or already failed earlier in this run is never reused.
3. On success, log `<waybill>	<original name>	no id, uploaded <placeholder name>(<folder>) - to be replaced later`. The folder is the dated (or other) subfolder under `OUTPUT_DIR` the pair came from; omitted if the files sit directly in `OUTPUT_DIR`.
4. On any error, **do not log that attempt yet** — try a different pair. First attempt plus **up to 2 retries** (3 pairs max). If they all fail, log `**…	no id - retried x times` (`x` is how many extra attempts ran after the first). If the pool was empty and nothing was tried, log `**…	no id`.
5. At the **end of the whole run**, every pair that failed is listed once as `<filename> has been tried and reported error: <message>` (message included when the wizard/page gave one). Other error reasons (API, timeout, same-name upload fail, etc.) are unchanged.

This also runs in `--from-log` mode, for every source line whose reason is exactly `no id` or `failed after retry: could not open ID viewer page`.

### HTTP 413 website fallback

The WMS REST API (`IdcardAdd`) sometimes rejects a photo payload with **HTTP 413** (request too large) and a non-JSON body. The public auodexpress.com wizard usually accepts the same files.

This runs after the normal Taobao capture + API upload has already been retried once (when the script is about to log `failed after retry: API upload failed: HTTP 413, …`), and again in `--from-log` for those same log lines:

1. Look for a complete same-name front/back pair in **today's** dated folder (`OUTPUT_DIR\YYYY-MM-DD\`, relative to when the script runs). The name used is the one the files were just saved under (ID-page name) when available, falling back to the `phones.txt` / log name.
2. If nothing is there, look in **yesterday's** folder only — not older folders, and not the random placeholder pool.
3. Upload that pair through the same website wizard as the `no id` fallbacks (not the REST API). **One attempt.**
4. On success: write **nothing** to the error log (console still notes that the website upload worked).
5. If no pair is found, or the website upload fails: log the original `failed after retry: API upload failed: HTTP 413, …` line unchanged.

### Replay mode (`--from-log`)

```
python automation.py --from-log "C:\Users\Lynn\Documents\JJ\身份证\log_20260918_233507.txt"
```

Re-processes `no id` lines, `failed after retry: could not open ID viewer page` lines, and HTTP 413 lines from that log — it never touches Taobao at all, since a normal run already tried and failed on those. `no id` / ID-viewer-fail go through same-name match then the random placeholder; 413 lines use today's (then yesterday's) same-name website upload. Every other reason in the source log (other API errors, timeouts, prior `same name match uploaded successfully` / `**...fail to upload`, `no id, uploaded … - to be replaced later`, and `no id - retried x times`) is carried through **unchanged** into a new, separate timestamped log file. Feeding a log back through `--from-log` a second time is safe: only still-unmatched `no id` / ID-viewer-fail entries and still-failed 413 lines get retried. This mode exists mainly for testing/iterating on the website-upload flow without re-running an entire batch.

---

## How it works

- **`CONCURRENCY` parallel workers** each keep their own Taobao search tab open and pull orders from a shared queue; the website-upload fallback (same-name match, random placeholder, and HTTP 413) is serialized behind all of them (single tab, see above)
- Each order opens the ID viewer in a new tab, captures both images, then closes that tab
- Images are captured via network response interception (primary) with `<img src>` fetch and screenshot as fallbacks
- A minimum size check (20 KB) rejects placeholder/loading images before saving
- Orders that fail due to transient errors (timeouts, slow page loads) are automatically retried once before being logged as failed
- On success, name/ID number/validity period/mobile/images are uploaded to the WMS API with an MD5-signed request (`consignorCode + token + timestamp + nonce + body`, sorted then hashed)

---

## Installing on another PC

1. Install Python 3.10+ from [python.org](https://www.python.org)
2. Copy the `automation\` folder to the new PC
3. Run `pip install -r requirements.txt` and `playwright install chromium`
4. Update the paths in `automation.py` to match the new machine
5. Run `python automation.py` and log in — the session will be saved for future runs
