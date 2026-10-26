"""
Automates the SnowAlert / SnowCase / SnowInc Excel export from itsm.services.sap.

Usage:
  python3 snow_download.py --setup                    # First time: log in manually and save session
  python3 snow_download.py --alert                    # Download SnowAlert.xlsx
  python3 snow_download.py --case                     # Download SnowCase.xlsx
  python3 snow_download.py --inc                      # Download SnowInc.xlsx
  python3 snow_download.py --alert --case --inc       # Download all three
"""

import argparse
import sys
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

SCRIPT_DIR   = Path(__file__).parent
SESSION_FILE = SCRIPT_DIR / "snow_session.json"

BROWSER_CHANNEL = "chrome"  # change to "msedge" if you use Edge

SNOW_ALERT_URL = (
    "https://itsm.services.sap/now/sow/list/params/list-id/aab8952a2f69b69cb9d4fb2cbfa4e32a"
    "/__state__/b64-tiny~TuwBdaJx92hcyxoZ0JbseY0Ls82dzPdo"
)

SNOW_CASE_URL = (
    "https://itsm.services.sap/now/cwf/agent/list/params/list-id/364b19a22f2db69cb9d4fb2cbfa4e304"
    "/__state__/b64~eyIwYzY5ZGFlOGNjODQyMjEwZjg3Nzk5NTgzZTE2NjhiMiI6eyJsaXN0X2NvbnRyb2"
    "xsZXIiOnsiXyI6eyJncm91cEJ5IjoiIiwicXVlcnkiOiJhc3NpZ25tZW50X2dyb3VwLnN5c19pZElOZDZj"
    "MWY4Y2FlYmMyMjI5MDc0YzNmMjRlMGJkMGNkZjAsZjNkMWIwNGFlYmM2NjI5MDE1N2ZmYjkzZGFkMGNk"
    "YjAsZWFjMTNjY2FlYmMyMjI5MDc0YzNmMjRlMGJkMGNkN2QsZTJjMTNjY2FlYmMyMjI5MDc0YzNmMjRl"
    "MGJkMGNkODAsMjNkMWIwNGFlYmM2NjI5MDE1N2ZmYjkzZGFkMGNkMzMsMmVjMTNjY2FlYmMyMjI5MDc0"
    "YzNmMjRlMGJkMGNkZjQsZTdkMTcwNGFlYmM2NjI5MDE1N2ZmYjkzZGFkMGNkZTJeT1J1X2FwcF9jb21w"
    "b25lbnQudV9wc19mdW5jX25leElOamF2YXNjcmlwdDp4X3NhcGRhX2Nhc2VfYXBpLlNBUF9HZXRfVXNlcl"
    "9BcmVhKCk7XnN0YXRlSU5udWxsLDEsMTAsMThedV9hcHBfY29tcG9uZW50LnVfbmFtZVNUQVJUU1dJVEhQ"
    "Q08tT1BTXk9SREVSQllERVNDc3lzX3VwZGF0ZWRfb25eR1JPVVBCWXVfYXBwX2NvbXBvbmVudCIsImN1cn"
    "JlbnRQYWdlIjowfX19fQ%3D%3D"
)

SNOW_INC_URL = (
    "https://itsm.services.sap/now/sow/list/params/list-id/ab4a9da22fe9b69cb9d4fb2cbfa4e38f"
    "/__state__/b64-tiny~hAUQmgzDrcthpjEosMCtWer9NpUy8qgS"
)


def setup_session():
    """Open browser so you can log in manually, then saves the session for future runs."""
    print("Opening browser — log in to ServiceNow, then wait until the list is fully loaded.")
    print("Once the page is loaded and you can see the records, press ENTER here to save the session.")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel=BROWSER_CHANNEL)
        context = browser.new_context()
        page = context.new_page()
        page.goto(SNOW_ALERT_URL)

        input("\n>>> Press ENTER after you have logged in and the list is visible: ")

        context.storage_state(path=str(SESSION_FILE))
        print(f"Session saved to {SESSION_FILE}")
        browser.close()


def export_report(page, output_file: Path, debug: bool = False):
    """Clicks Export → Export in the popup and saves the downloaded file."""
    print(f"Clicking Export for {output_file.name} …")
    try:
        page.locator("button:has-text('Export')").first.click(timeout=15_000)
    except PlaywrightTimeout:
        print("Could not find Export button. Run: playwright codegen to record the selector.")
        return False

    # Wait for the Cancel button — it only exists inside the popup dialog
    print("Waiting for export popup …")
    try:
        page.wait_for_selector("button:has-text('Cancel')", timeout=15_000)
    except PlaywrightTimeout:
        screenshot = SCRIPT_DIR / f"snow_debug_{output_file.stem}.png"
        page.screenshot(path=str(screenshot))
        print(f"Export popup did not appear. Debug screenshot saved → {screenshot}")
        return False

    if debug:
        # Dump DOM info around the dialog to identify the Export button selector
        info = page.evaluate("""
            () => {
                const results = [];
                function collect(root, depth) {
                    for (const el of root.querySelectorAll('*')) {
                        const text  = (el.textContent || '').trim().substring(0, 40);
                        const label = el.getAttribute('label') || '';
                        const role  = el.getAttribute('role')  || '';
                        if (text === 'Export' || text === 'Cancel' || label === 'Export') {
                            results.push({
                                tag: el.tagName, text, label, role,
                                id: el.id,
                                cls: (el.className || '').substring(0, 60),
                                shadowDepth: depth
                            });
                        }
                    }
                    for (const el of root.querySelectorAll('*'))
                        if (el.shadowRoot) collect(el.shadowRoot, depth + 1);
                }
                collect(document, 0);
                return results;
            }
        """)
        print("=== Dialog DOM dump ===")
        for item in info:
            print(f"  depth={item['shadowDepth']} <{item['tag']}> "
                  f"text='{item['text']}' label='{item['label']}' "
                  f"role='{item['role']}' id='{item['id']}' class='{item['cls']}'")
        print("=== end dump ===")
        print("Pausing — use the Playwright Inspector to hover over the Export button "
              "and copy its locator, then press Resume.")
        page.pause()
        return False

    try:
        with page.expect_download(timeout=60_000) as dl_info:
            # Dialog Export button has class 'now-button -primary'; toolbar has '-secondary'
            page.locator("button[class*='-primary']:has-text('Export')").click(timeout=10_000)
        download = dl_info.value
    except PlaywrightTimeout:
        screenshot = SCRIPT_DIR / f"snow_debug_{output_file.stem}.png"
        page.screenshot(path=str(screenshot))
        print(f"Download did not start. Debug screenshot saved → {screenshot}")
        return False

    download.save_as(str(output_file))
    print(f"Saved → {output_file}")
    return True


def download_reports(do_alert: bool, do_case: bool, do_inc: bool = False, debug: bool = False):
    if not SESSION_FILE.exists():
        print("No session found. Run with --setup first.")
        sys.exit(1)

    if do_case and not SNOW_CASE_URL:
        print("SNOW_CASE_URL is not set in snow_download.py — edit the file and add it.")
        sys.exit(1)

    if do_inc and not SNOW_INC_URL:
        print("SNOW_INC_URL is not set in snow_download.py — edit the file and add it.")
        sys.exit(1)

    print("Launching browser with saved session …")
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, channel=BROWSER_CHANNEL)
        context = browser.new_context(storage_state=str(SESSION_FILE))
        page = context.new_page()

        if do_alert:
            print("Navigating to SnowAlert list …")
            page.goto(SNOW_ALERT_URL)
            try:
                page.wait_for_selector("button", timeout=60_000)
                page.wait_for_load_state("networkidle", timeout=60_000)
            except PlaywrightTimeout:
                print("Page took too long to load. Session may have expired — run --setup again.")
                browser.close()
                sys.exit(1)
            export_report(page, SCRIPT_DIR / "SnowAlert.xlsx", debug)

        if do_case:
            print("Navigating to SnowCase list …")
            page.goto(SNOW_CASE_URL)
            try:
                page.wait_for_selector("button", timeout=60_000)
                page.wait_for_load_state("networkidle", timeout=60_000)
            except PlaywrightTimeout:
                print("Page took too long to load. Session may have expired — run --setup again.")
                browser.close()
                sys.exit(1)
            export_report(page, SCRIPT_DIR / "SnowCase.xlsx", debug)

        if do_inc:
            print("Navigating to Snow Internal INCs list …")
            page.goto(SNOW_INC_URL)
            try:
                page.wait_for_selector("button", timeout=60_000)
                page.wait_for_load_state("networkidle", timeout=60_000)
            except PlaywrightTimeout:
                print("Page took too long to load. Session may have expired — run --setup again.")
                browser.close()
                sys.exit(1)
            export_report(page, SCRIPT_DIR / "SnowInc.xlsx", debug)

        context.storage_state(path=str(SESSION_FILE))
        browser.close()


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--setup", action="store_true",
                        help="First-time login: opens browser for manual login and saves session")
    parser.add_argument("--alert", action="store_true", help="Download SnowAlert.xlsx")
    parser.add_argument("--case",  action="store_true", help="Download SnowCase.xlsx")
    parser.add_argument("--inc",   action="store_true", help="Download SnowInc.xlsx")
    parser.add_argument("--debug", action="store_true",
                        help="Dump dialog DOM and pause for inspection instead of downloading")
    args = parser.parse_args()

    if args.setup:
        setup_session()
    elif args.alert or args.case or args.inc:
        download_reports(args.alert, args.case, args.inc, args.debug)
    else:
        parser.print_help()
