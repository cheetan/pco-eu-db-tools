"""
Automates the SPC Report export from spc.ondemand.com.

Usage:
  python3 spc_download.py --setup   # First time: log in manually and save session
  python3 spc_download.py           # Normal run: use saved session and download

The downloaded file is saved as SPCReport.xlsx in the same folder as this script.
"""

import argparse
import sys
import time
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout, Error as PlaywrightError

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "spc_session.json"
OUTPUT_FILE  = SCRIPT_DIR / "SPCReport.xlsx"

# Use the installed Chrome instead of Playwright's bundled Chromium so that
# macOS Keychain grants certificate access without prompting (Chrome is already trusted).
BROWSER_CHANNEL = "chrome"  # change to "msedge" if you use Edge

SPC_URL = (
    "https://spc.ondemand.com/sap/public/ap/ui/repository/SAP_UI/HTMLOBERON5/"
    "client.html#n=eyJpblBvcnQiOiIiLCJ0YXJnZXQiOiIvQTFTX1NQQy9JbmNpZGVudC9UaWNrZXQv"
    "VGlja2V0V29ya0NlbnRlclZpZXcuV0NWSUVXLnVpd29jdmlldyIsInBhcmFtcyI6e30sImxpc3RzIjp"
    "7fSwid2NDb250ZXh0Ijp7IndvY0lkIjoiL0ExU19TUEMvV29ya0NlbnRlci9Nb25pdG9yaW5nX1dDRi"
    "5XQ0YudWl3b2MiLCJ2aWV3SWQiOiIvQTFTX1NQQy9JbmNpZGVudC9UaWNrZXQvVGlja2V0V29ya0Nl"
    "bnRlclZpZXcuV0NWSUVXLnVpd29jdmlldyIsInN0YXRzSXNDb21tb25UYXNrIjpmYWxzZSwiaXNOYXZ"
    "Ob2RlIjpmYWxzZSwibmF2Tm9kZVRleHQiOiIifX0%3D"
)


def setup_session():
    """Open browser so you can log in manually, then saves the session for future runs."""
    print("Opening browser — log in to SPC, then wait until the ticket list is fully loaded.")
    print("Once the page is loaded and you can see the tickets, press ENTER here to save the session.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel=BROWSER_CHANNEL)
        context = browser.new_context()
        page = context.new_page()
        page.goto(SPC_URL)

        input("\n>>> Press ENTER after you have logged in and the ticket list is visible: ")

        context.storage_state(path=str(SESSION_FILE))
        print(f"Session saved to {SESSION_FILE}")
        browser.close()


def download_report():
    """Use the saved session to export the SPC report and save it as SPCReport.xlsx."""
    if not SESSION_FILE.exists():
        print("No session found. Run with --setup first.")
        sys.exit(1)

    print("Launching browser with saved session …")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(SESSION_FILE))
        page = context.new_page()

        print(f"Navigating to SPC …")
        for _attempt in range(1, 4):
            try:
                page.goto(SPC_URL, timeout=60_000)
                break
            except (PlaywrightTimeout, PlaywrightError):
                if _attempt < 3:
                    print(f"Navigation failed (attempt {_attempt}/3) — VPN may still be connecting. Waiting 30s…")
                    time.sleep(30)
                    page = context.new_page()
                else:
                    print("Navigation failed after 3 attempts. VPN may not be connected.")
                    browser.close()
                    sys.exit(1)

        # Wait for the ticket list to load (SAPUI5 apps are slow to initialise)
        print("Waiting for page to load …")
        try:
            # Wait for any visible button or table row — adjust selector if needed
            page.wait_for_selector("button", timeout=60_000)
            page.wait_for_load_state("networkidle", timeout=60_000)
        except PlaywrightTimeout:
            print("Page took too long to load. Session may have expired — run --setup again.")
            browser.close()
            sys.exit(1)

        # Click the Export button
        # NOTE: if this fails, run: playwright codegen <SPC_URL>
        # to record the exact selector for the Export button on your screen.
        print("Clicking Export …")
        try:
            # Try common SAPUI5 button patterns
            export_btn = (
                page.get_by_role("button", name="Export") or
                page.locator("button:has-text('Export')").first
            )
            export_btn.click(timeout=15_000)
        except PlaywrightTimeout:
            print("Could not find Export button. Run: playwright codegen to record the selector.")
            browser.close()
            sys.exit(1)

        # Wait for popup and click "Microsoft Excel" option
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

        # Save refreshed session cookies so it stays valid longer
        context.storage_state(path=str(SESSION_FILE))
        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true",
                        help="First-time login: opens browser for manual login and saves session")
    args = parser.parse_args()

    if args.setup:
        setup_session()
    else:
        download_report()
