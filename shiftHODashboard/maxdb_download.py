"""
Automates the MaxDB Activities export from spc.ondemand.com.

Uses a dedicated session file (maxdb_session.json) that stores the SPC
session with the "EU Non-Core MaxDB DB L2" view already selected in
localStorage. SAPUI5 restores this view on every page load, so no
dropdown interaction is needed in headless mode.

First-time setup bootstraps from spc_session.json (so no re-login is
required) and saves to maxdb_session.json after the user selects the
MaxDB view manually.

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
    """Open browser so you can select the MaxDB view and save a dedicated session."""
    # Bootstrap from the existing SPC session if maxdb_session.json doesn't exist yet
    base = SESSION_FILE if SESSION_FILE.exists() else SPC_SESSION_FILE
    if not base.exists():
        print("No session found. Run spc_download.py --setup first to log in.")
        sys.exit(1)

    print(f"Opening browser (base session: {base.name}) …")
    print(f"Select the '{MAXDB_VIEW}' view from the dropdown, wait for the ticket list to load,")
    print("then press ENTER here to save the session.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(base))
        page = context.new_page()
        page.goto(SPC_URL, timeout=60_000)

        input(f"\n>>> Press ENTER once the '{MAXDB_VIEW}' ticket list is visible: ")

        context.storage_state(path=str(SESSION_FILE))
        print(f"Session saved → {SESSION_FILE}")
        browser.close()


def download_report(extra_output_dir=None):
    if not SESSION_FILE.exists():
        print("No session found. Run: python3 maxdb_download.py --setup")
        sys.exit(1)

    print("Launching browser with saved session …")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(SESSION_FILE))
        page = context.new_page()

        print("Navigating to SPC …")
        for _attempt in range(1, 4):
            try:
                page.goto(SPC_URL, timeout=60_000)
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

        # SAPUI5 restores the last selected view from localStorage stored in maxdb_session.json.
        # As long as --setup was run with the MaxDB view selected, the correct view loads here.
        print(f"Waiting for '{MAXDB_VIEW}' view to restore from session …")
        try:
            page.wait_for_selector("button", timeout=60_000)
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeout:
            page.screenshot(path=str(SCRIPT_DIR / "maxdb_debug.png"), full_page=True)
            print("Page took too long to load — screenshot → maxdb_debug.png")
            print("Session may have expired. Run: python3 maxdb_download.py --setup")
            browser.close()
            sys.exit(1)

        # Click Go to execute the query for the restored view
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
                        help="Open browser to select MaxDB view and save a dedicated session")
    args = parser.parse_args()

    if args.setup:
        setup_session()
    else:
        download_report()
