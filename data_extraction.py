import base64
import io
import json
import logging
import os
import re
import time
import pandas as pd
from playwright.sync_api import sync_playwright

try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

BMS_URL_SUBSTRING = "webctrl"
OUTPUT_DIR_NAME = "bms_context"
MAX_RETRY_PASSES = 1

TAB_CONFIG = {
    "graphics": {
        "label": "Graphics",
        "title": "View graphics",
    },
    "logic": {
        "label": "Logic",
        "title": "View program for this control program",
    },
}


def attach_to_bms_page(playwright):
    """Attach to the existing BMS tab in Chrome via CDP.

    Returns:
        (browser, context, page) on success
        (None, None, None) on failure
    """
    try:
        browser = playwright.chromium.connect_over_cdp("http://localhost:9222")
    except Exception as e:
        logging.error("Failed to connect to Chrome via CDP: %s", str(e))
        return None, None, None

    try:
        if not browser.contexts:
            logging.error("No browser contexts found after CDP connection.")
            return None, None, None
        context = browser.contexts[0]
    except Exception as e:
        logging.error("Failed accessing browser context: %s", str(e))
        return None, None, None

    try:
        pages = context.pages or []
        page = next(
            (p for p in pages if BMS_URL_SUBSTRING in (p.url or "")),
            None
        )
    except Exception as e:
        logging.error("Error while scanning context pages: %s", str(e))
        return None, None, None

    if page is None:
        logging.error("BMS tab not found — is it open and loaded?")
        return None, None, None

    # ---- Set viewport size for high-res screenshots ----
    try:
        page.set_viewport_size({"width": 1920, "height": 1080})
    except Exception as e:
        logging.warning("Could not set viewport size: %s", str(e))

    # ---- Stabilization gate ----
    try:
        page.wait_for_load_state("domcontentloaded", timeout=15000)
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception as e:
        logging.warning("Initial page stabilization incomplete: %s", str(e))

    return browser, context, page


def data_transfer(file_path):
    """Loads an Excel file and returns a list of unique control program items:

    [(location_path_tuple, controlProgram), ...]
    """
    if not os.path.isfile(file_path):
        logging.error("Input file not found: %s", file_path)
        print(f"Error: File not found at '{file_path}'")
        return []

    try:
        df = pd.read_excel(file_path)
    except Exception as e:
        logging.error("Failed reading Excel file %s: %s", file_path, str(e))
        print("Error: Failed to read Excel file.")
        return []

    required_cols = ["location", "controlProgram"]
    missing_cols = [c for c in required_cols if c not in df.columns]
    if missing_cols:
        logging.error("Missing required columns: %s", missing_cols)
        print(f"Error: Missing required columns: {missing_cols}")
        return []

    control_programs = []
    seen = set()

    for idx, row in df.iterrows():
        try:
            location = row.get("location")
            control_program = row.get("controlProgram")

            if pd.isna(location) or pd.isna(control_program):
                continue

            if not isinstance(location, str):
                logging.warning("Non-string location at row %s: %s", idx, location)
                continue

            # Split location string by '/' and extract all non-empty nested levels
            path_parts = tuple(p.strip() for p in location.split("/") if p.strip())
            if not path_parts:
                logging.warning("Malformed/empty location string at row %s: %s", idx, location)
                continue

            control_program_str = str(control_program).strip()

            item_key = (path_parts, control_program_str)
            if item_key not in seen:
                seen.add(item_key)
                control_programs.append(item_key)

        except Exception as e:
            logging.warning("Row processing failed at index %s: %s", idx, str(e))
            continue

    return control_programs


def get_nav_frame(page):
    """Returns a Playwright FrameLocator for the nested WebCTRL navigation tree iframe.

    Using FrameLocator ensures Playwright automatically re-queries the frame if it reloads.
    """
    return page.frame_locator("#navTableFrame").frame_locator("#navContent")


def find_tree_node(nav_frame, name):
    """Finds a tree node in the WebCTRL nav frame by text.

    Handles non-breaking spaces (\xa0), extra whitespace, and slashes seamlessly.
    """
    words = str(name).strip().split()
    if not words:
        return nav_frame.locator("nobr.TreeCtrl-text", has_text=str(name).strip())

    # Build regex that allows \s or \xa0 between words
    pattern = re.compile(r'[\s\xa0]+'.join(map(re.escape, words)), re.IGNORECASE)
    return nav_frame.locator("nobr.TreeCtrl-text", has_text=pattern)


def expand_tree_node(page, name, timeout=5000):
    """Locates a nesting node in the WebCTRL tree and ensures it is expanded.

    Returns True if found and expanded (or already expanded/non-collapsible).
    """
    nav_frame = get_nav_frame(page)
    node_loc = find_tree_node(nav_frame, name)

    if node_loc.count() == 0:
        # Node not visible/found at this level (e.g. system container or already expanded root)
        logging.info("Node label not found in nav tree (may be root container or already expanded): '%s'", name)
        return True

    first_node = node_loc.first
    try:
        first_node.wait_for(state="attached", timeout=timeout)
        first_node.wait_for(state="visible", timeout=timeout)
    except Exception as e:
        logging.warning("Timed out waiting for node '%s': %s", name, str(e))
        return False

    # Check twisty image next to node
    twisty = first_node.locator("xpath=ancestor::div[1]//img[contains(@class,'TreeCtrl-twisty')]")

    if twisty.count() > 0:
        try:
            twisty_src = twisty.first.get_attribute("src") or ""
            if "clean_collapsed" in twisty_src or "collapsed" in twisty_src:
                twisty.first.click(force=True)
                page.wait_for_timeout(1000)  # wait for AJAX expansion
        except Exception as e:
            logging.warning("Failed clicking twisty for '%s': %s", name, str(e))
            try:
                first_node.click(force=True)
                page.wait_for_timeout(1000)
            except Exception:
                pass
    else:
        try:
            first_node.click(force=True)
            page.wait_for_timeout(1000)
        except Exception as e:
            logging.warning("Failed clicking label for '%s': %s", name, str(e))

    return True


def navigate_to_control_program(page, location_path, control_program):
    """Navigates the BMS tree step-by-step through all nested location levels,
    then selects the target control program once the drill down is complete.
    """
    path_str = " > ".join(location_path)
    print(f"   -> Drilling down path: {path_str}")

    for step_idx, node_name in enumerate(location_path, start=1):
        print(f"      [{step_idx}/{len(location_path)}] Expanding level: '{node_name}'")
        success = expand_tree_node(page, node_name)
        if not success:
            print(f"   [!] Failed expanding nesting level: '{node_name}'")
            return False

    print(f"   -> Selecting Control Program: '{control_program}'")
    nav_frame = get_nav_frame(page)
    control_loc = find_tree_node(nav_frame, control_program)

    if control_loc.count() == 0:
        logging.warning("Control program label not found: '%s'", control_program)
        print(f"   [!] Control program label not found in navigation tree: '{control_program}'")
        return False

    target = control_loc.first
    try:
        target.wait_for(state="attached", timeout=10000)
        target.wait_for(state="visible", timeout=10000)
        target.click(force=True, timeout=10000)
        page.wait_for_timeout(1200)  # Allow click event and action bar update to settle
        return True
    except Exception as e:
        logging.warning("Failed clicking control program '%s': %s", control_program, str(e))
        print(f"   [!] Error clicking control program '{control_program}': {e}")
        return False


def get_tab_button(page, tab_name):
    """Locates a top-level MainBarTR action tab button ('graphics' or 'logic')."""
    cfg = TAB_CONFIG[tab_name]
    label_pattern = re.compile(rf"^\s*{re.escape(cfg['label'])}\s*$", re.IGNORECASE)

    by_label = page.locator("#actButtonSpan > span.Action-WidgetTextButton-base").filter(
        has=page.locator("div", has_text=label_pattern)
    )
    if by_label.count() > 0:
        return by_label

    return page.locator(f'#actButtonSpan span.Action-WidgetTextButton-base[title="{cfg["title"]}"]')


def is_tab_active(button_loc):
    """Returns True if the given action button currently has the 'down' (active) class."""
    try:
        if button_loc.count() == 0:
            return False
        classes = button_loc.first.get_attribute("class") or ""
        return "Action-WidgetTextButton-down" in classes
    except Exception:
        return False


def get_active_tab(page):
    """Detects which of the target tabs ('graphics' or 'logic') is currently selected."""
    for tab_name in ("graphics", "logic"):
        if is_tab_active(get_tab_button(page, tab_name)):
            return tab_name
    return None


def switch_to_tab(page, tab_name, timeout=10000):
    """Switches to the requested action tab if it exists and is not already active.

    Returns:
        'ok'      - tab is active and ready
        'missing' - tab button does not exist / is not visible for this control program
        'error'   - tab button exists but failed to activate
    """
    btn = get_tab_button(page, tab_name)
    if btn.count() == 0:
        logging.warning("Tab '%s' not found in #actButtonSpan.", tab_name)
        return "missing"

    first_btn = btn.first
    try:
        if not first_btn.is_visible():
            logging.warning("Tab '%s' is present in DOM but not visible.", tab_name)
            return "missing"
    except Exception:
        return "missing"

    if is_tab_active(btn):
        return "ok"

    try:
        first_btn.click(force=True, timeout=timeout)
        deadline = time.monotonic() + (timeout / 1000.0)
        while time.monotonic() < deadline:
            if is_tab_active(get_tab_button(page, tab_name)):
                break
            page.wait_for_timeout(150)
        page.wait_for_timeout(1000)
        return "ok"
    except Exception as e:
        logging.error("Failed clicking '%s' tab: %s", tab_name, str(e))
        return "error"


def sanitize_filename(name):
    """Sanitizes a control program title for safe filesystem folder/file creation."""
    return re.sub(r'[\\/*?:"<>|]', "_", str(name)).strip()


def get_program_dir(control_program, output_root_dir):
    """Returns full directory path under bms_context/ for a control program."""
    folder_name = sanitize_filename(control_program)
    return os.path.join(output_root_dir, folder_name)


def find_graphics_frame(page, timeout=10000):
    """Finds the frame containing #WorkspaceLayer on the Graphics tab."""
    deadline = time.monotonic() + (timeout / 1000.0)
    while time.monotonic() < deadline:
        # Check actionContent frame first (standard WebCTRL layout)
        for frame in page.frames:
            if frame == page.main_frame:
                continue
            try:
                if frame.locator("#WorkspaceLayer").count() > 0:
                    return frame
            except Exception:
                continue

        if page.locator("#WorkspaceLayer").count() > 0:
            return page.main_frame

        page.wait_for_timeout(200)

    return None


def get_frame_page_offset(page, target_frame):
    """Returns the cumulative (x, y) offset of target_frame's inner origin in top-level page pixels."""
    off_x, off_y = 0.0, 0.0
    cur_frame = target_frame
    while cur_frame != page.main_frame:
        fe = cur_frame.frame_element()
        f_rect = fe.evaluate(
            """el => {
                const r = el.getBoundingClientRect();
                return { left: r.left + el.clientLeft, top: r.top + el.clientTop };
            }"""
        )
        off_x += float(f_rect["left"])
        off_y += float(f_rect["top"])
        cur_frame = cur_frame.parent_frame
    return off_x, off_y


def extract_graphics_points(page, program_dir, control_program):
    """Extracts Summary metadata for every bound point (visible and hidden) on #WorkspaceLayer
    via background fetch() against WebCTRL's launchLogicPopup servlet and saves graphics_points.json.
    Uses cpyses=true and sequential two-phase fetching so the main WebCTRL browser session is never mutated.
    """
    os.makedirs(program_dir, exist_ok=True)
    json_path = os.path.join(program_dir, "graphics_points.json")

    graphics_frame = find_graphics_frame(page, timeout=10000)
    if graphics_frame is None:
        logging.warning("No #WorkspaceLayer found on Graphics tab for '%s'.", control_program)
        print("   [!] No #WorkspaceLayer found for graphics point extraction.")
        return None

    frame_off_x, frame_off_y = get_frame_page_offset(page, graphics_frame)

    # 1) Collect all bound point elements (visible and hidden) from #WorkspaceLayer
    dom_points = graphics_frame.evaluate(
        """({ frameOffX, frameOffY }) => {
            const ws = document.getElementById('WorkspaceLayer');
            if (!ws) return [];

            const clean = s => (s || '').replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim();
            const textNoScripts = node => {
                if (!node) return '';
                const clone = node.cloneNode(true);
                clone.querySelectorAll('script, style').forEach(s => s.remove());
                return clean(clone.textContent);
            };
            const parsePx = val => {
                const n = parseFloat(val);
                return Number.isFinite(n) ? Math.round(n) : null;
            };

            const nodes = Array.from(ws.querySelectorAll('[primid], [updateid]'));
            const byPrim = new Map();

            for (const el of nodes) {
                const rawPrim = el.getAttribute('primid');
                const rawUpd = el.getAttribute('updateid');
                let primid = rawPrim;
                if (!primid && rawUpd) {
                    const m = rawUpd.match(/^(prim_\\d+)/);
                    if (m) primid = m[1];
                }
                if (!primid) continue;

                const cs = window.getComputedStyle(el);
                const visible = el.offsetParent !== null && cs.visibility !== 'hidden' && cs.display !== 'none';

                let value = null;
                if (el.tagName === 'TEXTAREA' || el.tagName === 'INPUT') {
                    value = clean(el.value);
                } else if (el.tagName === 'IMG') {
                    const src = el.getAttribute('src') || '';
                    value = src.split('/').pop() || null;
                } else {
                    const ta = el.querySelector('textarea, input');
                    value = ta ? clean(ta.value) : textNoScripts(el);
                }

                // Region ID and its native #WorkspaceLayer top/left/width/height coordinates
                const region = el.closest('div[id^="region"]');
                const regionId = region ? region.id : null;
                let regionCoords = null;
                if (region) {
                    const rcs = window.getComputedStyle(region);
                    regionCoords = {
                        top: parsePx(region.style.top || rcs.top),
                        left: parsePx(region.style.left || rcs.left),
                        width: parsePx(region.style.width || rcs.width),
                        height: parsePx(region.style.height || rcs.height)
                    };
                }

                // Rendered (x, y) pixel coordinates on graphics.png (accounts for iframe offset and scale)
                let pngCoords = null;
                if (visible) {
                    const rect = el.getBoundingClientRect();
                    pngCoords = {
                        x: Math.round(frameOffX + rect.left),
                        y: Math.round(frameOffY + rect.top),
                        width: Math.round(rect.width),
                        height: Math.round(rect.height)
                    };
                }

                // Walk to the nearest region row to capture adjacent label and units
                let label = null;
                let units = null;
                const outerRow = el.closest('table')?.parentElement?.closest('tr');
                if (outerRow) {
                    const directTds = Array.from(outerRow.children).filter(c => c.tagName === 'TD');
                    let elCellIdx = -1;
                    const cellTexts = directTds.map((td, i) => {
                        if (td.contains(el)) elCellIdx = i;
                        return textNoScripts(td);
                    });

                    if (elCellIdx > 0) {
                        for (let i = elCellIdx - 1; i >= 0; i--) {
                            if (cellTexts[i]) { label = cellTexts[i]; break; }
                        }
                    }
                    if (elCellIdx >= 0 && elCellIdx < cellTexts.length - 1) {
                        for (let i = elCellIdx + 1; i < cellTexts.length; i++) {
                            if (cellTexts[i]) { units = cellTexts[i]; break; }
                        }
                    }
                }

                const occurrence = {
                    updateid: rawUpd || null,
                    tag: el.tagName.toLowerCase(),
                    visible,
                    graphic_label: label,
                    graphic_value: value || null,
                    graphic_units: units,
                    region_id: regionId,
                    region_coords: regionCoords,
                    png_coords: pngCoords
                };

                if (!byPrim.has(primid)) {
                    byPrim.set(primid, {
                        primid,
                        visible,
                        graphic_label: label,
                        graphic_value: value || null,
                        graphic_units: units,
                        region_id: regionId,
                        region_coords: regionCoords,
                        png_coords: pngCoords,
                        occurrences: [occurrence]
                    });
                } else {
                    const existing = byPrim.get(primid);
                    existing.occurrences.push(occurrence);
                    // Prefer a text/labeled occurrence over an icon-only <img> for top-level fields
                    if (!existing.graphic_label && label) existing.graphic_label = label;
                    if ((!existing.graphic_value || existing.occurrences[0].tag === 'img') && el.tagName !== 'IMG' && value) {
                        existing.graphic_value = value;
                        if (regionId) existing.region_id = regionId;
                        if (regionCoords) existing.region_coords = regionCoords;
                        if (pngCoords) existing.png_coords = pngCoords;
                    }
                    if (!existing.graphic_units && units) existing.graphic_units = units;
                    if (visible) existing.visible = true;
                }
            }

            return Array.from(byPrim.values());
        }""",
        {"frameOffX": frame_off_x, "frameOffY": frame_off_y},
    )

    if not dom_points:
        output_payload = {
            "control_program": control_program,
            "dbid": None,
            "point_count": 0,
            "points": [],
        }
        with open(json_path, "w", encoding="utf-8") as f:
            json.dump(output_payload, f, indent=2)
        print(f"   [✓] Graphics points saved (0 points): {json_path}")
        return json_path

    # 2) Two-phase background fetch:
    #    Phase A: Resolve all primids -> /~abspath/... URLs sequentially while main session is intact.
    #    Phase B: Fetch each Summary page sequentially with cpyses=true so WebCTRL clones the session
    #             and never mutates the active tab's wBrowserSession state.
    enriched = page.evaluate(
        """async ({ points }) => {
            const wbs = window.wBrowserSession;
            const dbid = window.navWin && typeof window.navWin.getSelectedDbid === 'function'
                ? window.navWin.getSelectedDbid()
                : null;

            if (!wbs || !dbid) {
                return { dbid, error: 'Missing wBrowserSession or selected dbid', points };
            }

            const clean = s => (s || '').replace(/\\u00a0/g, ' ').replace(/\\s+/g, ' ').trim();

            function parseSummaryHtml(html, popupUrl) {
                const doc = new DOMParser().parseFromString(html, 'text/html');
                const cjPageMeta = doc.querySelector('meta[name="cjpage"]');
                const summaryJsp = cjPageMeta ? cjPageMeta.getAttribute('content') : null;
                const title = clean(doc.title);

                // Strip all <script> and <style> elements so setUpd("prim_..."); never leaks into textContent
                doc.querySelectorAll('script, style').forEach(el => el.remove());

                const ws = doc.getElementById('WorkspaceLayer') || doc.body;
                let pointType = null;
                let refName = null;
                let displayName = null;
                let description = null;
                const summaryFields = {};

                // Remove status/alarm boilerplate table before scanning summary rows
                const eventTable = ws.querySelector('#eventstatus_table');
                const tables = Array.from(ws.querySelectorAll('table')).filter(t => t !== eventTable && !eventTable?.contains(t));

                for (const table of tables) {
                    const rows = Array.from(table.querySelectorAll('tr'));
                    for (const row of rows) {
                        const cells = Array.from(row.children).filter(c => c.tagName === 'TD');
                        if (!cells.length) continue;

                        const cellTexts = cells.map(c => clean(c.textContent));

                        // Check for inline "RefName: <val>" cell (header row also holds pointType)
                        for (let i = 0; i < cells.length; i++) {
                            const txt = cellTexts[i];
                            if (/^RefName:/i.test(txt)) {
                                const valSpan = cells[i].querySelector('[primid], [updateid]');
                                refName = valSpan ? clean(valSpan.textContent) : clean(txt.replace(/^RefName:/i, ''));
                                // Earlier non-empty cell in the same row is the microblock point type
                                for (let j = i - 1; j >= 0; j--) {
                                    if (cellTexts[j]) {
                                        pointType = cellTexts[j];
                                        break;
                                    }
                                }
                            }
                        }

                        // Check for "Label:" | (spacer) | "Value" rows (e.g. Display Name, Description)
                        const nonEmpty = cells
                            .map((c, idx) => ({ el: c, text: cellTexts[idx] }))
                            .filter(x => x.text.length > 0);

                        if (nonEmpty.length >= 2 && /:$/.test(nonEmpty[0].text) && !/^RefName:/i.test(nonEmpty[0].text)) {
                            const key = nonEmpty[0].text.replace(/:$/, '').trim();
                            const valEl = nonEmpty[1].el.querySelector('input, textarea');
                            const val = valEl ? clean(valEl.value) : nonEmpty.slice(1).map(x => x.text).join(' ').trim();
                            if (/^Display Name$/i.test(key)) {
                                displayName = val;
                            } else if (/^Description$/i.test(key)) {
                                description = val;
                            } else if (key) {
                                summaryFields[key] = val;
                            }
                        } else if (nonEmpty.length >= 3 && nonEmpty[1].text === '=') {
                            // Rows like "Input | = | Off"
                            const key = nonEmpty[0].text;
                            const val = nonEmpty[2].text;
                            if (key && val !== '?') {
                                summaryFields[key] = val;
                            }
                        }
                    }
                }

                // Fallback refName from URL path if not found in HTML
                let abspath = null;
                if (popupUrl) {
                    const pathPart = popupUrl.split('?')[0].replace(/^\\/~abspath\\//, '');
                    try { abspath = decodeURIComponent(pathPart); } catch { abspath = pathPart; }
                    if (!refName && abspath) {
                        const segs = abspath.split('/').filter(Boolean);
                        refName = segs.length ? segs[segs.length - 1] : null;
                    }
                }

                return {
                    ref_name: refName,
                    display_name: displayName,
                    description: description,
                    point_type: pointType,
                    abspath,
                    popup_url: popupUrl,
                    summary_jsp: summaryJsp,
                    popup_title: title || null,
                    summary_fields: summaryFields
                };
            }

            // Phase A: Resolve all primids to their /~abspath/... URLs first
            const resolved = [];
            for (const pt of points) {
                try {
                    const servletUrl = `/_common/servlet/lvl5/action?wbs=${encodeURIComponent(wbs)}&action=launchLogicPopup&id=${encodeURIComponent(dbid)}&primid=${encodeURIComponent(pt.primid)}`;
                    const servletText = await (await fetch(servletUrl, { credentials: 'same-origin' })).text();
                    const m = servletText.match(/\\/~abspath\\/[^'"\\s)]+/);
                    resolved.push({ pt, popupUrl: m ? m[0] : null, err: null });
                } catch (err) {
                    resolved.push({ pt, popupUrl: null, err: String(err) });
                }
            }

            // Phase B: Fetch Summary HTML for each resolved URL using cpyses=true (cloned session)
            const results = [];
            for (const item of resolved) {
                const { pt, popupUrl, err } = item;
                if (err) {
                    results.push({ ...pt, has_details: false, details_error: err });
                    continue;
                }
                if (!popupUrl) {
                    results.push({ ...pt, has_details: false });
                    continue;
                }

                try {
                    const sep = popupUrl.includes('?') ? '&' : '?';
                    const summaryUrl = `${popupUrl}${sep}wbs=${encodeURIComponent(wbs)}&cpyses=true`;
                    const summaryHtml = await (await fetch(summaryUrl, { credentials: 'same-origin' })).text();

                    if (/darkblank\\.html/i.test(summaryHtml) && summaryHtml.length < 400) {
                        const pathPart = popupUrl.split('?')[0].replace(/^\\/~abspath\\//, '');
                        let abspath = pathPart;
                        try { abspath = decodeURIComponent(pathPart); } catch {}
                        const segs = abspath.split('/').filter(Boolean);
                        results.push({
                            ...pt,
                            has_details: false,
                            abspath,
                            ref_name: segs.length ? segs[segs.length - 1] : null,
                            popup_url: popupUrl,
                            details_error: 'Redirected to darkblank.html'
                        });
                        continue;
                    }

                    const parsed = parseSummaryHtml(summaryHtml, popupUrl);
                    results.push({ ...pt, has_details: true, ...parsed });
                } catch (fetchErr) {
                    results.push({ ...pt, has_details: false, popup_url: popupUrl, details_error: String(fetchErr) });
                }
            }

            return { dbid, error: null, points: results };
        }""",
        {"points": dom_points},
    )

    if enriched.get("error"):
        logging.warning("Graphics point enrichment warning for '%s': %s", control_program, enriched["error"])

    points = enriched.get("points", [])
    with_details = sum(1 for pt in points if pt.get("has_details"))

    output_payload = {
        "control_program": control_program,
        "dbid": enriched.get("dbid"),
        "point_count": len(points),
        "points_with_details": with_details,
        "points": points,
    }

    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(output_payload, f, indent=2)

    print(f"   [✓] Graphics points saved ({with_details}/{len(points)} with Summary details): {json_path}")
    return json_path


def capture_graphics_tab(page, program_dir, control_program):
    """Waits for the Graphics view to render, saves graphics.png, and extracts graphics_points.json."""
    os.makedirs(program_dir, exist_ok=True)
    screenshot_path = os.path.join(program_dir, "graphics.png")

    # Wait for the main action frame if present
    try:
        page.wait_for_selector("iframe#actionContent", state="attached", timeout=15000)
    except Exception as e:
        logging.warning("actionContent frame not attached before graphics screenshot: %s", str(e))

    # Give view time to render SVG graphics / trend graphs
    page.wait_for_timeout(2000)

    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    page.screenshot(path=screenshot_path, full_page=True)
    print(f"   [✓] Graphics screenshot saved: {screenshot_path}")

    json_path = extract_graphics_points(page, program_dir, control_program)
    return screenshot_path, json_path


def find_logic_container(page, timeout=15000):
    """Locates <div id='cjLogicOuter'> in the top-level page or any loaded frame.

    Returns:
        (host_frame, locator) on success
        (None, None) on timeout
    """
    deadline = time.monotonic() + (timeout / 1000.0)
    while time.monotonic() < deadline:
        top_loc = page.locator("#cjLogicOuter")
        if top_loc.count() > 0:
            return page.main_frame, top_loc.first

        for frame in page.frames:
            if frame == page.main_frame:
                continue
            try:
                frame_loc = frame.locator("#cjLogicOuter")
                if frame_loc.count() > 0:
                    return frame, frame_loc.first
            except Exception:
                continue

        page.wait_for_timeout(200)

    return None, None


def wait_for_logic_images(host_frame, timeout=15000):
    """Waits until all multiPartLogicImage* tiles inside #cjLogicOuter finish loading."""
    try:
        host_frame.wait_for_function(
            """() => {
                const container = document.getElementById('cjLogicOuter');
                if (!container) return false;
                const imgs = Array.from(container.querySelectorAll("img[id^='multiPartLogicImage']"));
                if (imgs.length === 0) return true;
                return imgs.every(img => img.complete && img.naturalWidth > 0);
            }""",
            timeout=timeout,
        )
    except Exception as e:
        logging.warning("Timed out waiting for all multiPartLogicImage tiles to load: %s", str(e))


def get_logic_scroll_geometry(page, host_frame):
    """Determines the scrollable container for #cjLogicOuter and its visible clip box
    in top-level page coordinates (excluding scrollbars).
    """
    geom = host_frame.evaluate(
        """() => {
            const logic = document.getElementById('cjLogicOuter');
            if (!logic) return null;

            // Find nearest scrollable ancestor inside this frame, or fall back to window
            let scroller = null;
            let cur = logic.parentElement;
            while (cur && cur !== document.documentElement && cur !== document.body) {
                const cs = window.getComputedStyle(cur);
                const canScrollY = (cs.overflowY === 'auto' || cs.overflowY === 'scroll') && cur.scrollHeight > cur.clientHeight + 5;
                const canScrollX = (cs.overflowX === 'auto' || cs.overflowX === 'scroll') && cur.scrollWidth > cur.clientWidth + 5;
                if (canScrollY || canScrollX) {
                    scroller = cur;
                    break;
                }
                cur = cur.parentElement;
            }

            // Tag scroller so we can address it cleanly during scroll steps
            document.querySelectorAll('[data-__bms-scroller]').forEach(el => {
                delete el.dataset.__bmsScroller;
            });

            const logicWidth = Math.ceil(Math.max(logic.scrollWidth, logic.offsetWidth));
            const logicHeight = Math.ceil(Math.max(logic.scrollHeight, logic.offsetHeight));

            if (scroller) {
                scroller.dataset.__bmsScroller = '1';
                const r = scroller.getBoundingClientRect();
                return {
                    isWindow: false,
                    viewLeft: r.left + scroller.clientLeft,
                    viewTop: r.top + scroller.clientTop,
                    viewWidth: scroller.clientWidth,
                    viewHeight: scroller.clientHeight,
                    totalWidth: Math.min(logicWidth, scroller.scrollWidth),
                    totalHeight: Math.min(logicHeight, scroller.scrollHeight),
                    origScrollLeft: scroller.scrollLeft,
                    origScrollTop: scroller.scrollTop
                };
            }

            const docEl = document.scrollingElement || document.documentElement;
            return {
                isWindow: true,
                viewLeft: 0,
                viewTop: 0,
                viewWidth: document.documentElement.clientWidth,
                viewHeight: document.documentElement.clientHeight,
                totalWidth: Math.min(logicWidth, docEl.scrollWidth),
                totalHeight: Math.min(logicHeight, docEl.scrollHeight),
                origScrollLeft: window.scrollX || docEl.scrollLeft || 0,
                origScrollTop: window.scrollY || docEl.scrollTop || 0
            };
        }"""
    )

    if not geom:
        raise RuntimeError("Could not read scroll geometry for #cjLogicOuter.")

    # Translate frame-local visible rect into top-level page coordinates
    clip_x = float(geom["viewLeft"])
    clip_y = float(geom["viewTop"])
    clip_w = float(geom["viewWidth"])
    clip_h = float(geom["viewHeight"])

    cur_frame = host_frame
    while cur_frame != page.main_frame:
        fe = cur_frame.frame_element()
        f_rect = fe.evaluate(
            """el => {
                const r = el.getBoundingClientRect();
                return {
                    left: r.left + el.clientLeft,
                    top: r.top + el.clientTop,
                    width: el.clientWidth,
                    height: el.clientHeight
                };
            }"""
        )
        # Intersect current clip rect with this iframe's visible inner area
        abs_left = f_rect["left"] + clip_x
        abs_top = f_rect["top"] + clip_y
        abs_right = min(f_rect["left"] + f_rect["width"], abs_left + clip_w)
        abs_bottom = min(f_rect["top"] + f_rect["height"], abs_top + clip_h)

        clip_x = max(f_rect["left"], abs_left)
        clip_y = max(f_rect["top"], abs_top)
        clip_w = max(0.0, abs_right - clip_x)
        clip_h = max(0.0, abs_bottom - clip_y)

        cur_frame = cur_frame.parent_frame

    # Clamp to top-level page viewport
    vp = page.evaluate("() => ({ width: window.innerWidth, height: window.innerHeight })")
    clip_right = min(float(vp["width"]), clip_x + clip_w)
    clip_bottom = min(float(vp["height"]), clip_y + clip_h)
    clip_x = max(0.0, clip_x)
    clip_y = max(0.0, clip_y)
    clip_w = int(max(1, round(clip_right - clip_x)))
    clip_h = int(max(1, round(clip_bottom - clip_y)))

    geom["clipX"] = int(round(clip_x))
    geom["clipY"] = int(round(clip_y))
    geom["clipWidth"] = clip_w
    geom["clipHeight"] = clip_h
    return geom


def scroll_logic_view(host_frame, is_window, target_x, target_y):
    """Scrolls the logic container to (target_x, target_y) and returns actual (scrollLeft, scrollTop)."""
    return host_frame.evaluate(
        """({ isWindow, x, y }) => {
            if (isWindow) {
                window.scrollTo(x, y);
                window.dispatchEvent(new Event('scroll'));
                const docEl = document.scrollingElement || document.documentElement;
                return {
                    x: Math.round(window.scrollX || docEl.scrollLeft || 0),
                    y: Math.round(window.scrollY || docEl.scrollTop || 0)
                };
            }
            const scroller = document.querySelector('[data-__bms-scroller="1"]');
            if (!scroller) return { x: 0, y: 0 };
            scroller.scrollLeft = x;
            scroller.scrollTop = y;
            scroller.dispatchEvent(new Event('scroll'));
            return {
                x: Math.round(scroller.scrollLeft),
                y: Math.round(scroller.scrollTop)
            };
        }""",
        {"isWindow": is_window, "x": target_x, "y": target_y},
    )


def stitch_slices_to_png(page, slices, total_w, total_h, png_path):
    """Stitches captured (x, y, png_bytes) viewport tiles into a single full-resolution PNG."""
    if HAS_PIL:
        canvas = Image.new("RGB", (total_w, total_h), (0, 0, 0))
        for x, y, png_bytes in slices:
            with Image.open(io.BytesIO(png_bytes)) as tile:
                canvas.paste(tile.convert("RGB"), (x, y))
        canvas.save(png_path, format="PNG")
        return

    # Pure-browser fallback if Pillow is not installed
    payload = [
        {"x": x, "y": y, "b64": base64.b64encode(png_bytes).decode("ascii")}
        for x, y, png_bytes in slices
    ]
    stitched_b64 = page.evaluate(
        """async ({ width, height, tiles }) => {
            const canvas = document.createElement('canvas');
            canvas.width = width;
            canvas.height = height;
            const ctx = canvas.getContext('2d');
            ctx.fillStyle = '#000000';
            ctx.fillRect(0, 0, width, height);

            for (const t of tiles) {
                await new Promise((resolve, reject) => {
                    const img = new Image();
                    img.onload = () => {
                        ctx.drawImage(img, t.x, t.y);
                        resolve();
                    };
                    img.onerror = reject;
                    img.src = 'data:image/png;base64,' + t.b64;
                });
            }
            return canvas.toDataURL('image/png').split(',')[1];
        }""",
        {"width": total_w, "height": total_h, "tiles": payload},
    )
    with open(png_path, "wb") as f:
        f.write(base64.b64decode(stitched_b64))


def capture_logic_tab(page, program_dir):
    """Captures logic.html (outerHTML of #cjLogicOuter) and logic.png (scroll-and-stitch capture)."""
    os.makedirs(program_dir, exist_ok=True)
    html_path = os.path.join(program_dir, "logic.html")
    png_path = os.path.join(program_dir, "logic.png")

    host_frame, logic_loc = find_logic_container(page, timeout=15000)
    if logic_loc is None:
        raise RuntimeError("#cjLogicOuter was not found on the Logic tab.")

    logic_loc.wait_for(state="visible", timeout=15000)
    wait_for_logic_images(host_frame, timeout=15000)

    # Allow live logic overlays (#logicgraphic) to finish updating
    page.wait_for_timeout(1500)
    try:
        page.wait_for_load_state("networkidle", timeout=5000)
    except Exception:
        pass

    # 1) Save raw outerHTML of #cjLogicOuter
    outer_html = logic_loc.evaluate("el => el.outerHTML")
    with open(html_path, "w", encoding="utf-8") as f:
        f.write(outer_html)
    print(f"   [✓] Logic HTML saved: {html_path}")

    # 2) Scroll-and-stitch #cjLogicOuter across its real scrollable container
    geom = get_logic_scroll_geometry(page, host_frame)
    total_w = int(geom["totalWidth"])
    total_h = int(geom["totalHeight"])
    clip_x = geom["clipX"]
    clip_y = geom["clipY"]
    step_w = geom["clipWidth"]
    step_h = geom["clipHeight"]
    is_window = geom["isWindow"]

    # Build scroll coordinates covering the full width and height
    x_targets = list(range(0, total_w, step_w))
    y_targets = list(range(0, total_h, step_h))

    slices = []
    seen_positions = set()

    try:
        for ty in y_targets:
            for tx in x_targets:
                pos = scroll_logic_view(host_frame, is_window, tx, ty)
                ax, ay = pos["x"], pos["y"]
                if (ax, ay) in seen_positions:
                    continue
                seen_positions.add((ax, ay))

                # Let multipart logic tiles & overlays repaint after scroll
                page.wait_for_timeout(200)
                wait_for_logic_images(host_frame, timeout=5000)

                png_bytes = page.screenshot(
                    clip={
                        "x": clip_x,
                        "y": clip_y,
                        "width": min(step_w, total_w),
                        "height": min(step_h, total_h),
                    },
                    scale="css",
                    animations="disabled",
                )
                slices.append((ax, ay, png_bytes))
    finally:
        # Restore original scroll position and clean up helper attribute
        try:
            scroll_logic_view(
                host_frame, is_window, geom["origScrollLeft"], geom["origScrollTop"]
            )
            host_frame.evaluate(
                """() => {
                    document.querySelectorAll('[data-__bms-scroller]').forEach(el => {
                        delete el.dataset.__bmsScroller;
                    });
                }"""
            )
        except Exception as e:
            logging.warning("Could not restore scroll position after logic capture: %s", str(e))

    stitch_slices_to_png(page, slices, total_w, total_h, png_path)
    print(f"   [✓] Logic screenshot saved ({total_w}x{total_h}, {len(slices)} slice(s)): {png_path}")
    return png_path, html_path


def process_control_program(
    page,
    location_path,
    control_program,
    output_root_dir,
    idx,
    total_count,
    tabs_to_capture=None,
    pass_label="",
):
    """Navigates to a control program and captures the requested tabs.

    Returns:
        dict with keys:
            'captured_tabs': set of tab names succeeded in this call
            'failed_tabs':   list of tab names that failed or were missing
            'issues':        list of human-readable failure descriptions
    """
    if tabs_to_capture is None:
        tabs_to_capture = ["graphics", "logic"]

    path_display = " > ".join(location_path)
    prefix = f"[{pass_label}] " if pass_label else ""
    print(f"{prefix}Processing ({idx}/{total_count} control programs): {path_display} > {control_program}")

    nav_success = navigate_to_control_program(page, location_path, control_program)
    if not nav_success:
        logging.warning("%sSkipping '%s' due to navigation failure.", prefix, control_program)
        print(f"   [!] Failed to navigate to: {control_program} ({idx}/{total_count})\n")
        return {
            "captured_tabs": set(),
            "failed_tabs": list(tabs_to_capture),
            "issues": ["Navigation failed"],
        }

    program_dir = get_program_dir(control_program, output_root_dir)
    os.makedirs(program_dir, exist_ok=True)

    # Snake through tabs: capture whichever requested tab is already active first
    active_tab = get_active_tab(page)
    if active_tab in tabs_to_capture:
        tab_order = [active_tab] + [t for t in tabs_to_capture if t != active_tab]
    else:
        tab_order = list(tabs_to_capture)

    print(f"   -> Active tab: {active_tab or 'other'} | Capture order: {' -> '.join(tab_order)}")

    captured_tabs = set()
    failed_tabs = []
    tab_issues = []

    for tab_name in tab_order:
        status = switch_to_tab(page, tab_name)
        if status == "missing":
            msg = f"{tab_name.capitalize()} tab not available"
            print(f"   [!] {msg} for '{control_program}' — skipping.")
            failed_tabs.append(tab_name)
            tab_issues.append(msg)
            continue
        if status != "ok":
            msg = f"Failed switching to {tab_name.capitalize()} tab"
            print(f"   [!] {msg} for '{control_program}' — skipping.")
            failed_tabs.append(tab_name)
            tab_issues.append(msg)
            continue

        try:
            if tab_name == "graphics":
                capture_graphics_tab(page, program_dir, control_program)
            else:
                capture_logic_tab(page, program_dir)
            captured_tabs.add(tab_name)
        except Exception as e:
            logging.error(
                "%sFailed capturing %s tab for '%s': %s", prefix, tab_name, control_program, str(e)
            )
            print(f"   [!] Error capturing {tab_name} tab for '{control_program}': {e}")
            failed_tabs.append(tab_name)
            tab_issues.append(f"{tab_name.capitalize()} capture error: {e}")

    if not failed_tabs:
        print(f"   [+] Done ({idx}/{total_count}).\n")
    elif captured_tabs:
        print(f"   [~] Partially captured ({idx}/{total_count}).\n")
    else:
        print(f"   [!] Failed ({idx}/{total_count}).\n")

    return {
        "captured_tabs": captured_tabs,
        "failed_tabs": failed_tabs,
        "issues": tab_issues,
    }


def main():
    logging.basicConfig(
        filename="error_log.txt",
        filemode="w",
        level=logging.WARNING,
        format="%(asctime)s [%(levelname)s] %(message)s"
    )

    excel_file_path = input("Enter absolute file path to loadsheet: ").strip().strip('"').strip("'")

    # Load control programs from Excel
    control_programs = data_transfer(excel_file_path)
    if not control_programs:
        print("No control programs to process.")
        return

    # Place bms_context/ in the same directory as the submitted loadsheet
    loadsheet_dir = os.path.dirname(os.path.abspath(excel_file_path))
    output_root_dir = os.path.join(loadsheet_dir, OUTPUT_DIR_NAME)
    os.makedirs(output_root_dir, exist_ok=True)

    total_count = len(control_programs)
    print(f"Found {total_count} control program(s) to capture.")
    print(f"Output directory: {output_root_dir}\n")

    with sync_playwright() as p:
        browser, context, page = attach_to_bms_page(p)
        if page is None:
            print("Error: Could not attach to BMS page.")
            return

        # Track cumulative state per control program across initial pass + retry pass
        program_states = {}

        # ---- Pass 1: Initial extraction ----
        for idx, (location_path, control_program) in enumerate(control_programs, start=1):
            result = process_control_program(
                page, location_path, control_program, output_root_dir, idx, total_count
            )
            program_states[(location_path, control_program)] = {
                "captured_tabs": set(result["captured_tabs"]),
                "failed_tabs": list(result["failed_tabs"]),
                "issues": list(result["issues"]),
            }

        # ---- Pass 2: Single retry pass for any failed or partial control programs ----
        recovered_count = 0
        for retry_round in range(1, MAX_RETRY_PASSES + 1):
            retry_items = [
                (item, state)
                for item, state in program_states.items()
                if state["failed_tabs"]
            ]
            if not retry_items:
                break

            retry_total = len(retry_items)
            print("\n==================================================")
            print(
                f"RETRY PASS ({retry_round}/{MAX_RETRY_PASSES}): "
                f"Retrying {retry_total} failed/partial control program(s)..."
            )
            print("==================================================\n")

            for r_idx, ((location_path, control_program), state) in enumerate(retry_items, start=1):
                tabs_to_retry = list(state["failed_tabs"])
                result = process_control_program(
                    page,
                    location_path,
                    control_program,
                    output_root_dir,
                    r_idx,
                    retry_total,
                    tabs_to_capture=tabs_to_retry,
                    pass_label=f"Retry {retry_round}/{MAX_RETRY_PASSES}",
                )
                state["captured_tabs"].update(result["captured_tabs"])
                state["failed_tabs"] = list(result["failed_tabs"])
                state["issues"] = list(result["issues"])

                if not state["failed_tabs"]:
                    recovered_count += 1

        # ---- Final Summary ----
        full_success_count = 0
        partial_control_programs = []
        failed_control_programs = []

        for (_, control_program), state in program_states.items():
            if not state["failed_tabs"]:
                full_success_count += 1
            elif state["captured_tabs"]:
                partial_control_programs.append((control_program, "; ".join(state["issues"])))
            else:
                failed_control_programs.append((control_program, "; ".join(state["issues"])))

        print("\n==================================================")
        retry_note = f" ({recovered_count} recovered on retry)" if recovered_count else ""
        print(
            f"SUMMARY: {full_success_count}/{total_count} fully captured{retry_note} | "
            f"{len(partial_control_programs)} partial | {len(failed_control_programs)} failed."
        )

        if partial_control_programs:
            print(f"\n[~] Partially Captured Control Programs ({len(partial_control_programs)} total):")
            for cp, reason in partial_control_programs:
                print(f"  • {cp} ({reason})")

        if failed_control_programs:
            print(f"\n[!] Failed Control Programs ({len(failed_control_programs)} total):")
            for cp, reason in failed_control_programs:
                print(f"  • {cp} ({reason})")

        if not partial_control_programs and not failed_control_programs:
            print("\n[✓] All control programs were captured successfully!")
        print("==================================================")


if __name__ == "__main__":
    main()