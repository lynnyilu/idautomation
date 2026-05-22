# ID Document Downloader

Automates the daily workflow of downloading buyer ID card images from the Taobao seller platform.

Given a list of tracking numbers in an Excel file, the script:
1. Searches each order on `myseller.taobao.com`
2. Opens the buyer's identity verification page
3. Reveals and records the buyer's name
4. Downloads the front and back ID card images
5. Saves them as `<name>.jpg` / `<name>1.jpg` in the output folder

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
| `OUTPUT_DIR` | `C:\Users\Lynn\Documents\JJ\身份证` | Folder where ID images are saved |
| `CHROME_PROFILE_DIR` | `chrome_profile\` | Folder where your login session is stored |
| `ORDER_SEARCH_URL` | myseller.taobao.com | The order search page URL |
| `CONCURRENCY` | `3` | How many orders are processed at the same time |

---

## Running

```
python automation.py
```

A Chrome window will open. **Log in to the seller platform** if prompted, then press **Enter** in the terminal to start processing.

On future runs your login session is remembered automatically — just press Enter straight away.

---

## Excel file format

The script scans every sheet and every cell in the workbook. It picks up any value matching the pattern:

```
JR251<digits>          e.g.  JR25125594
JR251<digits>E         e.g.  JR25125594E
```

Column position and sheet name do not matter.

---

## Output

Images are saved flat (no subfolders) inside `OUTPUT_DIR`:

| File | Contents |
|---|---|
| `张三.jpg` | Front ID card |
| `张三1.jpg` | Back ID card |

If the same buyer appears more than once the next pair is numbered `张三2.jpg` / `张三3.jpg`, then `张三4.jpg` / `张三5.jpg`, and so on.

A timestamped log file is also written to `OUTPUT_DIR` (e.g. `log_20260302_090000.txt`) listing any orders that could not be processed.

---

## Error log

Orders that cannot be completed are written to the log file with one of these reasons:

| Reason | Meaning |
|---|---|
| `no id` | Buyer has not completed identity verification |
| `请确认订单信息` | Order not found or status unclear |
| `failed after retry` | Unexpected error after one automatic retry |

---

## How it works

- **3 parallel workers** each keep their own search tab open and pull orders from a shared queue
- Each order opens the ID viewer in a new tab, captures both images, then closes that tab
- Images are captured via network response interception (primary) with `<img src>` fetch and screenshot as fallbacks
- A minimum size check (20 KB) rejects placeholder/loading images before saving
- Orders that fail due to transient errors (timeouts, slow page loads) are automatically retried once before being logged as failed

---

## Installing on another PC

1. Install Python 3.10+ from [python.org](https://www.python.org)
2. Copy the `automation\` folder to the new PC
3. Run `pip install -r requirements.txt` and `playwright install chromium`
4. Update the paths in `automation.py` to match the new machine
5. Run `python automation.py` and log in — the session will be saved for future runs
