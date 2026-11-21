# WebCTRL BMS Context Extractor

A Python + Playwright automation tool that attaches to an existing authenticated Chrome session of Automated Logic **WebCTRL**, drills down through the navigation tree for every control program listed in an Excel loadsheet, and extracts both **Graphics** and **Logic** context into a structured local directory.

## Features

* **CDP Session Attachment:** Connects to your already signed-in Chrome tab over the Chrome DevTools Protocol (`http://localhost:9222`) without storing credentials or launching a separate headless login flow.
* **Tab-Snaking Navigation:** Detects whichever tab (`Graphics` or `Logic`) persisted from the previous control program and captures it first before switching, cutting tab transitions in half while guarding against missing tabs.
* **Zero-Popup Graphics Point Extraction:** Resolves every bound point on `#WorkspaceLayer` (both visible and hidden conditional/alarm points) via background `fetch()` calls using WebCTRL's `cpyses=true` session-cloning flag, capturing `RefName`, `Display Name`, `Description`, `point_type`, live values, and pixel coordinates (`region_coords` and `png_coords`) into `graphics_points.json`.
* **Full-Height Logic Capture:** Captures the raw `<div id="cjLogicOuter">` markup (`logic.html`) and scroll-stitches the full logic diagram across WebCTRL's scrollable container into a scrollbar-free `logic.png` (using Pillow locally, with an automatic offscreen browser canvas fallback).
* **Automatic Single Retry Pass:** Re-attempts any control program that failed navigation or missed a tab once at the end of the run, retrying only the tab(s) that failed.

## Prerequisites

* Python 3.9+
* Google Chrome

## Installation

```bash
python3 -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
pip install -r requirements.txt
```

> **Note:** Because the script attaches to an existing Chrome instance via CDP, you do not need to run `playwright install chromium`.

## Loadsheet Format

Provide an Excel workbook (`.xlsx`) containing at least these two columns:

| Column | Description | Example |
| :--- | :--- | :--- |
| `location` | Forward-slash (`/`) separated path of parent nodes in the WebCTRL navigation tree | `Site A / Building 1 / Roof` |
| `controlProgram` | Exact display name of the target control program node | `AHU-1 / Main Supply` |

Duplicate `(location, controlProgram)` rows and blank rows are automatically skipped.

## Usage

### 1. Launch Chrome with Remote Debugging

Start Chrome with `--remote-debugging-port=9222` (using a dedicated user data directory if Chrome is already running):

**Linux**
```bash
google-chrome --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-cdp
```

**macOS**
```bash
/Applications/Google\ Chrome.app/Contents/MacOS/Google\ Chrome --remote-debugging-port=9222 --user-data-dir=/tmp/chrome-cdp
```

**Windows (PowerShell)**
```powershell
& "C:\Program Files\Google\Chrome\Application\chrome.exe" --remote-debugging-port=9222 --user-data-dir="$env:TEMP\chrome-cdp"
```

Open WebCTRL in that Chrome window and sign in so an active tab has `webctrl` in its URL.

### 2. Run the Extractor

```bash
python3 bms_scraper.py
```

When prompted, enter the absolute path to your `.xlsx` loadsheet:

```text
Enter absolute file path to loadsheet: /path/to/site_loadsheet.xlsx
```

## Output Structure

A `bms_context/` directory is created in the **same directory as the submitted loadsheet**, with one sanitized subfolder per control program:

```text
<loadsheet_directory>/
├── site_loadsheet.xlsx
└── bms_context/
    ├── AHU-1 _ Main Supply/
    │   ├── graphics.png
    │   ├── graphics_points.json
    │   ├── logic.png
    │   └── logic.html
    └── AHU-2 _ East Wing/
        ├── graphics.png
        ├── graphics_points.json
        ├── logic.png
        └── logic.html
```

### Output Files

| File | Contents |
| :--- | :--- |
| `graphics.png` | Full-page screenshot of the control program's **Graphics** tab (`1920x1080` viewport). |
| `graphics_points.json` | Structured metadata for every bound `[primid]` / `[updateid]` element in `#WorkspaceLayer` (visible and hidden), including Summary popup fields (`ref_name`, `display_name`, `description`, `point_type`, `abspath`, `summary_fields`) plus `region_coords` (native `#WorkspaceLayer` `top`/`left`/`width`/`height`) and `png_coords` (rendered `(x, y)` pixels on `graphics.png`). |
| `logic.png` | Full-height, scrollbar-free stitched PNG of `#cjLogicOuter` from the **Logic** tab. |
| `logic.html` | Raw `outerHTML` of `<div id="cjLogicOuter">`, preserving live values and `mbObjArray` microblock coordinates. |

## Configuration

Top-level constants in `bms_scraper.py` can be adjusted as needed:

| Constant | Default | Purpose |
| :--- | :--- | :--- |
| `BMS_URL_SUBSTRING` | `"webctrl"` | Substring used to locate the open WebCTRL tab in Chrome. |
| `OUTPUT_DIR_NAME` | `"bms_context"` | Name of the output folder created next to the loadsheet. |
| `MAX_RETRY_PASSES` | `1` | Number of retry passes for failed or partially captured control programs. |

## Troubleshooting

* **`Failed to connect to Chrome via CDP`:** Confirm Chrome was launched with `--remote-debugging-port=9222` and that no other Chrome instance ignored the flag (use a separate `--user-data-dir` if needed).
* **`BMS tab not found`:** Make sure the WebCTRL tab is open, loaded, and its URL contains `BMS_URL_SUBSTRING` (`webctrl`).
* **Points with `"has_details": false` in `graphics_points.json`:** Expected for non-microblock bindings on WebCTRL graphic templates (such as floor/equipment header titles, the `Notes:` box, or display-only widgets that do not have a microblock Details popup).
* **Detailed diagnostics:** Warnings and errors from each run are written to `error_log.txt` in the working directory.
