"""
Automates the MaxDB Activities export from spc.ondemand.com.

The "EU Non-Core MaxDB DB L2" view must be captured once via --setup:
the script saves the URL that was active when the view was selected
to maxdb_url.txt. Headless runs navigate to that URL directly so the
correct SAPUI5 view is loaded without any dropdown interaction.

A pre-export screenshot (maxdb_preexport.png) is saved on every run
so you can visually verify which view was loaded.

Usage:
  python3 maxdb_download.py --setup   # First time (or after session expires):
                                      # opens browser, select MaxDB view, press ENTER
  python3 maxdb_download.py           # Normal run: headless download → MaxDB.xlsx
"""

import argparse
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout, Error as PlaywrightError

SCRIPT_DIR       = Path(__file__).parent
SESSION_FILE     = SCRIPT_DIR / "maxdb_session.json"   # dedicated — never shared with spc_download.py
SPC_SESSION_FILE = SCRIPT_DIR / "spc_session.json"     # used as base for first-time --setup only
MAXDB_URL_FILE   = SCRIPT_DIR / "maxdb_url.txt"        # view-specific URL captured during --setup
OUTPUT_FILE      = SCRIPT_DIR / "MaxDB.xlsx"

BROWSER_CHANNEL = "chrome"

SPC_URL = (
    "https://spc.ondemand.com/sap/public/ap/ui/repository/SAP_UI/HTMLOBERON5/"
    "client.html#n=eyJpblBvcnQiOiIiLCJ0YXJnZXQiOiIvQTFTX1NQQy9JbmNpZGVudC9UaWNrZXQv"
    "VGlja2V0V29ya0NlbnRlclZpZXcuV0NWSUVXLnVpd29jdmlldyIsInBhcmFtcyI6e30sImxpc3RzIjp"
    "7fSwid2NDb250ZXh0Ijp7IndvY0lkIjoiL0ExU19TUEMvV29ya0NlbnRlci9Nb25pdG9yaW5nX1dDRi"
    "5XQ0YudWl3b2MiLCJ2aWV3SWQiOiIvQTFTX1NQQy9JbmNpZGVudC9UaWNrZXQvVGlja2V0V29ya0Nl"
    "bnRlclZpZXcuV0NWSUVXLnVpd29jdmlldyIsInN0YXRzSXNDb21tb25UYXNrIjpmYWxzZSwiaXNOYXZ"
    "Ob2RlIjpmYWxzZSwibmF2Tm9kZVRleHQiOiIifX0%3D"
)

MAXDB_VIEW = "EU Non-Core MaxDB DB L2"


def setup_session():
    """Open browser so you can select the MaxDB view, then saves the URL and session."""
    base = SESSION_FILE if SESSION_FILE.exists() else SPC_SESSION_FILE
    if not base.exists():
        print("No session found. Run spc_download.py --setup first to log in.")
        sys.exit(1)

    print(f"Opening browser (base session: {base.name}) …")
    print(f"Select the '{MAXDB_VIEW}' view from the dropdown, wait for the ticket list to load,")
    print("then press ENTER here.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(base))
        page = context.new_page()
        page.goto(SPC_URL, timeout=60_000)

        input(f"\n>>> Press ENTER once the '{MAXDB_VIEW}' ticket list is visible: ")

        # Capture the current URL — SAPUI5 Fiori apps encode the selected view in the URL hash.
        # Headless runs navigate to this URL directly so the correct view loads every time.
        maxdb_url = page.url
        MAXDB_URL_FILE.write_text(maxdb_url)
        if maxdb_url != SPC_URL:
            print(f"View URL captured (different from base URL) → {MAXDB_URL_FILE.name}")
        else:
            print(f"Warning: URL did not change after view selection ({MAXDB_URL_FILE.name})")
            print("Headless runs will fall back to session-state approach.")
        print(f"URL: {maxdb_url}")

        context.storage_state(path=str(SESSION_FILE))
        print(f"Session saved → {SESSION_FILE}")
        browser.close()


def download_report(extra_output_dir=None):
    if not SESSION_FILE.exists():
        print("No session found. Run: python3 maxdb_download.py --setup")
        sys.exit(1)

    # Use the view-specific URL captured during --setup if available
    if MAXDB_URL_FILE.exists():
        nav_url = MAXDB_URL_FILE.read_text().strip()
        print(f"Using MaxDB-specific URL from {MAXDB_URL_FILE.name}")
    else:
        nav_url = SPC_URL
        print("No maxdb_url.txt found — using base SPC URL. Run --setup to capture the correct URL.")

    print("Launching browser with saved session …")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(SESSION_FILE))
        page = context.new_page()

        print("Navigating to SPC …")
        for _attempt in range(1, 4):
            try:
                page.goto(nav_url, timeout=60_000)
                break
            except (PlaywrightTimeout, PlaywrightError):
                if _attempt < 3:
                    print(f"Navigation failed (attempt {_attempt}/3) — waiting 30s …")
                    time.sleep(30)
                    page = context.new_page()
                else:
                    print("Navigation failed after 3 attempts. VPN may not be connected.")
                    browser.close()
                    sys.exit(1)

        print(f"Waiting for page to load …")
        try:
            page.wait_for_selector("button", timeout=60_000)
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeout:
            page.screenshot(path=str(SCRIPT_DIR / "maxdb_debug.png"), full_page=True)
            print("Page took too long to load — screenshot → maxdb_debug.png")
            print("Session may have expired. Run: python3 maxdb_download.py --setup")
            browser.close()
            sys.exit(1)

        # Switch to the MaxDB saved view using the SAPUI5 JavaScript runtime directly.
        # core.mElements was deprecated/emptied in SAPUI5 1.100+; use
        # sap.ui.core.Element.registry (available since 1.67) with a DOM-based fallback.
        print(f"Switching to '{MAXDB_VIEW}' view via SAPUI5 API …")
        found_view = False

        js_result = page.evaluate(r"""
            (viewName) => {
                if (typeof sap === 'undefined' || !sap.ui) {
                    return {error: 'SAPUI5 not available'};
                }

                // Collect all registered UI5 elements — try three APIs in priority order
                let allElems = {};

                // 1. sap.ui.core.Element.registry (SAPUI5 >= 1.67, preferred)
                try {
                    const reg = sap.ui.core.Element.registry;
                    if (reg && reg.all) {
                        allElems = reg.all();
                    }
                } catch(e) {}

                // 2. sap.ui.getCore().mElements (older, may be empty in 1.100+)
                if (Object.keys(allElems).length === 0) {
                    try {
                        const core = sap.ui.getCore();
                        allElems = core.mElements || {};
                    } catch(e) {}
                }

                // 3. DOM-based: look up each [role="combobox"] element by its SAPUI5 id
                if (Object.keys(allElems).length === 0) {
                    try {
                        document.querySelectorAll('[role="combobox"]').forEach(el => {
                            if (el.id) {
                                try {
                                    const c = sap.ui.getCore().byId(el.id);
                                    if (c) allElems[el.id] = c;
                                } catch(e) {}
                            }
                        });
                    } catch(e) {}
                }

                const totalElements = Object.keys(allElems).length;

                // Find every sap.m.Select in the element registry
                const selects = [];
                for (const id in allElems) {
                    const c = allElems[id];
                    try {
                        let ok = false;
                        try { ok = c.isA && c.isA('sap.m.Select'); } catch(e) {}
                        if (!ok) {
                            try { ok = c.getMetadata && c.getMetadata().getName() === 'sap.m.Select'; } catch(e) {}
                        }
                        if (ok) selects.push(c);
                    } catch(e) {}
                }

                const diagnostics = {
                    totalElements,
                    totalSelects: selects.length,
                    selectDetails: selects.map(s => {
                        try {
                            const items = s.getItems ? s.getItems().map(i => i.getText ? i.getText() : '') : [];
                            const sel = s.getSelectedItem ? s.getSelectedItem() : null;
                            return {id: s.getId(), selected: sel ? sel.getText() : '', items: items.slice(0, 8)};
                        } catch(e) { return {id: '?', error: String(e)}; }
                    })
                };

                // Identify the view selector: it always has at least one item with ' | '
                // (service-area items never use that separator)
                for (const sel of selects) {
                    try {
                        const items = sel.getItems ? sel.getItems() : [];
                        const hasPipe = items.some(i => i.getText && i.getText().includes(' | '));
                        if (!hasPipe) continue;
                        const target = items.find(i => i.getText && i.getText() === viewName);
                        if (!target) continue;
                        sel.setSelectedItem(target);
                        sel.fireChange({selectedItem: target});
                        return {
                            success: true,
                            controlId: sel.getId(),
                            selectedText: target.getText(),
                            allItems: items.map(i => i.getText ? i.getText() : '').slice(0, 10),
                            diagnostics
                        };
                    } catch(e) {}
                }

                return {success: false, reason: 'view selector not found or view not in list', diagnostics};
            }
        """, MAXDB_VIEW)

        print(f"  SAPUI5 result: {js_result}")

        if isinstance(js_result, dict) and js_result.get("success"):
            print(f"  ✓ SAPUI5 fired change on '{MAXDB_VIEW}' (control: {js_result.get('controlId')})")
            print(f"    Items in view selector: {js_result.get('allItems')}")
            page.wait_for_load_state("networkidle", timeout=30_000)
            found_view = True
        else:
            reason = js_result.get("reason") or js_result.get("error") if isinstance(js_result, dict) else js_result
            print(f"  SAPUI5 approach failed: {reason}")
            diag = js_result.get("diagnostics") if isinstance(js_result, dict) else {}
            if diag:
                print(f"    totalElements={diag.get('totalElements')} totalSelects={diag.get('totalSelects')}")
                for d in (diag.get("selectDetails") or []):
                    print(f"    Select id={d.get('id')} selected='{d.get('selected')}' items={d.get('items')}")

        if not found_view:
            page.screenshot(path=str(SCRIPT_DIR / "maxdb_debug.png"), full_page=True)
            print(f"Could not switch to '{MAXDB_VIEW}' — screenshot → maxdb_debug.png")
            print("Run: python3 maxdb_download.py --setup to re-save the session.")
            browser.close()
            sys.exit(1)

        # Click Go to execute the query for the selected view.
        # Give SAPUI5 extra time to re-render the filter form after the view change.
        page.wait_for_timeout(3000)
        print("Loading ticket list …")
        try:
            go_btn = page.get_by_role("button", name="Go")
            if go_btn.count() > 0:
                go_btn.first.click(timeout=10_000, force=True)
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeout:
            print("Timed out waiting for ticket list to load.")
            browser.close()
            sys.exit(1)

        # Screenshot before export so the loaded view can be verified
        page.screenshot(path=str(SCRIPT_DIR / "maxdb_preexport.png"))
        print("Screenshot saved → maxdb_preexport.png (verify correct view is shown)")

        # Export — same flow as spc_download.py
        print("Clicking Export …")
        try:
            export_btn = (
                page.get_by_role("button", name="Export") or
                page.locator("button:has-text('Export')").first
            )
            export_btn.click(timeout=15_000)
        except PlaywrightTimeout:
            print("Could not find Export button.")
            browser.close()
            sys.exit(1)

        print("Waiting for export popup …")
        try:
            page.wait_for_selector("text=Microsoft Excel", timeout=15_000)
            with page.expect_download(timeout=60_000) as dl_info:
                page.get_by_text("Microsoft Excel").click()
            download = dl_info.value
        except PlaywrightTimeout:
            print("Export popup did not appear or download did not start.")
            browser.close()
            sys.exit(1)

        download.save_as(str(OUTPUT_FILE))
        print(f"Saved → {OUTPUT_FILE}")
        if extra_output_dir:
            import shutil
            shutil.copy2(str(OUTPUT_FILE), str(extra_output_dir / OUTPUT_FILE.name))

        context.storage_state(path=str(SESSION_FILE))
        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true",
                        help="Open browser to select MaxDB view, capture URL, and save session")
    args = parser.parse_args()

    if args.setup:
        setup_session()
    else:
        download_report()
