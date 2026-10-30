"""
Automates the MaxDB Activities export from spc.ondemand.com.

Uses the same saved session as spc_download.py (same domain / SSO).
After the page loads it switches to the "EU Non-Core MaxDB L2" saved view,
then follows the same Export → Microsoft Excel flow as SPC.

Usage:
  python3 maxdb_download.py --setup   # First time: verify session and navigate to MaxDB view
  python3 maxdb_download.py           # Normal run: headless download → MaxDB.xlsx

To find the exact view-selector if the default locator fails:
  playwright codegen https://spc.ondemand.com/...   (use SPC_URL below)
  Look for the element that shows the current view name (e.g. "My Open Tickets")
  and copy its selector here.
"""

import argparse
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout, Error as PlaywrightError

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "spc_session.json"   # shared with spc_download.py
OUTPUT_FILE  = SCRIPT_DIR / "MaxDB.xlsx"

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
    """Open browser with existing SPC session so you can verify the MaxDB view is reachable."""
    if not SESSION_FILE.exists():
        print("No session found. Run spc_download.py --setup first to log in.")
        sys.exit(1)

    print(f"Opening browser with saved SPC session …")
    print(f"Navigate to the '{MAXDB_VIEW}' view using the upper-left dropdown.")
    print("Once the MaxDB ticket list is visible, press ENTER here.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(SESSION_FILE))
        page = context.new_page()
        page.goto(SPC_URL, timeout=60_000)

        input(f"\n>>> Press ENTER once you can see the '{MAXDB_VIEW}' ticket list: ")

        # Refresh the shared session so it stays valid
        context.storage_state(path=str(SESSION_FILE))
        print(f"Session refreshed → {SESSION_FILE}")
        browser.close()


def download_report():
    """Use the saved SPC session to export MaxDB Activities and save as MaxDB.xlsx."""
    if not SESSION_FILE.exists():
        print("No session found. Run spc_download.py --setup first.")
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

        print("Waiting for page to load …")
        try:
            page.wait_for_selector("button", timeout=60_000)
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeout:
            print("Page took too long to load. Session may have expired — run spc_download.py --setup.")
            browser.close()
            sys.exit(1)

        # Switch to the MaxDB view via the upper-left view selector dropdown.
        # The dropdown renders items as <li class="sapMSelectListItemBase …"> after ~3s.
        print(f"Switching to '{MAXDB_VIEW}' view …")
        try:
            page.get_by_role("combobox").nth(1).click(timeout=10_000, force=True)
            page.wait_for_timeout(3000)
            item = page.locator("li.sapMSelectListItemBase").filter(has_text=MAXDB_VIEW)
            item.scroll_into_view_if_needed(timeout=10_000)
            item.click(timeout=10_000)
        except PlaywrightTimeout:
            page.screenshot(path=str(SCRIPT_DIR / "maxdb_debug.png"), full_page=True)
            print(f"Could not switch to '{MAXDB_VIEW}' view — screenshot → maxdb_debug.png")
            browser.close()
            sys.exit(1)

        # Click Go to execute the query for the selected view
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

        context.storage_state(path=str(SESSION_FILE))
        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true",
                        help="Verify session and navigate to MaxDB view manually")
    args = parser.parse_args()

    if args.setup:
        setup_session()
    else:
        download_report()
