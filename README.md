# ID Document Downloader + WMS Uploader

Automates the daily workflow of collecting buyer ID card images and getting them into the WMS system for customs clearance.

Given a list of waybills — from an Excel file, or a `--start`/`--end` range matched against `phones.txt` — the script:
1. Searches each order on `myseller.taobao.com`
2. Opens the buyer's identity verification page
3. Reveals the buyer's name, ID number, and ID validity period
4. Downloads the front and back ID card images
5. Saves them as `<name>.jpg` / `<name>1.jpg` in a dated output subfolder
6. Uploads name, ID number, validity period, mobile, and both images to the WMS API (`api/Open/IdcardAdd`)

If a buyer hasn't completed Taobao's real-name verification (`no id`), and a photo pair for that exact name already exists from a previous order, the script falls back to uploading through auodexpress.com's public ID-upload wizard instead of giving up — see [Same-name-match fallback](#same-name-match-fallback-no-id) below.

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

**Replaying a previous log** (see [Replay mode](#replay-mode---from-log) below):
```
python automation.py --from-log "C:\Users\Lynn\Documents\JJ\身份证\log_20260918_190138.txt"
```

A Chrome window will open (except in `--from-log` mode, which doesn't need Taobao). **Log in to the seller platform** if prompted, then press **Enter** in the terminal to start processing. On future runs your login session is remembered automatically — just press Enter straight away.

---

## phones.txt format

Before each run, paste in this run's batch of waybill → buyer data (name, mobile, address) as copied from the order source. Entries don't need blank-line separators — the parser detects boundaries by a short numeric sequence number immediately followed by a `JR…` waybill code, and pulls the name from the line right after a `已打印`/`未打印` status marker and the mobile from the first 11-digit `1[3-9]…` number in the entry. Suspicious entries (empty/numeric/status-word names, invalid mobiles) are flagged with a Y/N confirmation prompt before the run continues.

---

## Output

Images are saved inside a dated subfolder of `OUTPUT_DIR` (e.g. `OUTPUT_DIR\18SEP\`):

| File | Contents |
|---|---|
| `张三.jpg` | Front ID card |
| `张三1.jpg` | Back ID card |

If the same buyer appears more than once **on the same day**, the next pair is numbered `张三2.jpg` / `张三3.jpg`, then `张三4.jpg` / `张三5.jpg`, and so on. The same name can also legitimately reappear across different dated subfolders (a repeat customer) — this is what the same-name-match fallback searches across.

A timestamped log file is also written to `OUTPUT_DIR` (e.g. `log_20260302_090000.txt`) listing every order that wasn't uploaded via the normal API path.

---

## Error log

Every waybill that doesn't go through the normal WMS API upload gets one tab-separated line, `<waybill>\t<name>\t<reason>`:

| Reason | Meaning |
|---|---|
| `no id` | Buyer hasn't completed identity verification, and no existing photo match was found either |
| `found a same name match` | Buyer hasn't completed verification, but a same-name photo pair was found — the website-upload fallback is being attempted |
| `same name match uploaded successfully` | The fallback upload above succeeded |
| `**...same name match fail to upload` | The fallback upload above failed at some step — `**`-prefixed so it's easy to grep for |
| `请确认订单信息` | Order not found or status unclear |
| `failed after retry: ...` | Unexpected error after one automatic retry (timeout, HTTP error, couldn't open ID viewer, etc.) |

---

## Same-name-match fallback (`no id`)

The WMS API (`IdcardAdd`) requires an ID **number**, which we only get by reading it off the Taobao verification page — if the buyer never completed verification, we have no ID number, so the API path is a dead end even if we had a photo. To get around this without OCR'ing the number ourselves (unreliable — watermarks on some scans obscure digits), the script instead drives auodexpress.com's own public ID-upload wizard, which does the OCR server-side:

1. Search `OUTPUT_DIR` and all of its dated subfolders for a complete front/back photo pair filed under the buyer's exact name (from `phones.txt`).
2. If **no** pair exists anywhere, log a plain `no id` — nothing else to try.
3. If one or more pairs exist, **pick one at random** (deliberate — spreads reuse across candidates rather than always resubmitting the same photo for a repeat name) and log `found a same name match`.
4. Upload that pair through `https://www.auodexpress.com/user.html#/upload-card-id`: upload front → upload back → click 开始识别身份证 (OCR) → fill in 运单号 with the waybill → click 确认提交.
5. **One attempt only.** Any failure at any step (image rejected, OCR timeout/error, submit error) gives up on that waybill immediately and logs `**...same name match fail to upload` — no retry, no partial credit.

This flow always runs **single-tab / sequential**, even while the normal Taobao workers run with `CONCURRENCY` in parallel — a deliberate scope decision to keep it easy to watch and debug while the flow is still being tuned, not a hard technical limit.

### Replay mode (`--from-log`)

```
python automation.py --from-log <path to a previous run's log file>
```

Re-processes only the `no id` lines from that log through the same-name-match flow above — it never touches Taobao at all, since a normal run already tried and failed on those. Every other reason in the source log (API errors, timeouts, `could not open ID viewer page`, etc.) is carried through **unchanged** into a new, separate timestamped log file, so nothing from the original run is lost. This mode exists mainly for testing/iterating on the website-upload flow without re-running an entire batch.

---

## How it works

- **`CONCURRENCY` parallel workers** each keep their own Taobao search tab open and pull orders from a shared queue; the website-upload fallback is serialized behind all of them (single tab, see above)
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
