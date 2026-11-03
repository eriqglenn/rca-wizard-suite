"""
HCAD Property Record Downloader
================================
Planner-friendly HCAD downloader for the MLS/RCA wizard.

This tool intentionally uses a visible Chrome window. It does not bypass HCAD's
human verification. In v13.9 it can connect to a normal Chrome session that the
wizard starts with a remote debugging port, after the planner verifies manually.
A legacy persistent Playwright profile path is kept only as a fallback.
"""

from __future__ import annotations

import argparse
import base64
import csv
import os
import random
import re
import subprocess
import sys
import time
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from urllib.parse import urljoin


# ═══════════════════════════════════════════════════════
#  AUTO-INSTALL DEPENDENCIES
# ═══════════════════════════════════════════════════════
def install_dependencies() -> bool:
    """Install playwright and browser on first run."""
    try:
        from playwright.sync_api import sync_playwright  # noqa: F401
        return True
    except ImportError:
        pass

    print()
    print("  ╔═══════════════════════════════════════════╗")
    print("  ║     First-time setup (one time only)      ║")
    print("  ╚═══════════════════════════════════════════╝")
    print()
    print("   Installing Playwright...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "pip", "install", "playwright"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print("   [OK] Playwright installed")
    except Exception as e:
        print(f"   [ERROR] Could not install Playwright: {e}")
        print("   Try running: pip install playwright")
        input("\n   Press Enter to exit...")
        sys.exit(1)

    print("   Downloading Chrome browser (this may take a minute)...")
    try:
        subprocess.check_call(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )
        print("   [OK] Chrome browser ready")
    except Exception as e:
        print(f"   [ERROR] Could not install browser: {e}")
        print("   Try running: playwright install chromium")
        input("\n   Press Enter to exit...")
        sys.exit(1)

    print()
    print("   Setup complete! Starting program...")
    print()
    return True


install_dependencies()
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError  # noqa: E402
from playwright.sync_api import sync_playwright  # noqa: E402


# ═══════════════════════════════════════════════════════
#  SETTINGS
# ═══════════════════════════════════════════════════════
DEFAULT_SAVE_FOLDER = Path.home() / "Desktop" / "HCAD_Records"
SAVE_FOLDER = DEFAULT_SAVE_FOLDER
SAVE_FOLDER.mkdir(parents=True, exist_ok=True)

PROPERTY_DETAIL_MARKERS = [
    "Back to Search Results",
    "Valuation History",
    "Property Details",
    "Jurisdictions/Exemptions",
    "Mailing Address:",
]

HCAD_HOME = "https://search.hcad.org/"


@dataclass
class SearchItem:
    address: str
    account: str = ""

    @property
    def display(self) -> str:
        if self.account:
            return f"{self.address} [{self.account}]" if self.address else self.account
        return self.address


@dataclass
class DownloadResult:
    item: SearchItem
    status: str
    pdf_path: str = ""
    notes: str = ""


def clean_address(address: str) -> str:
    search_addr = (address or "").strip().upper()
    for city in ["HOUSTON", "KATY", "SPRING", "HUMBLE", "PASADENA", "CYPRESS", "PEARLAND", "SUGAR LAND", "TOMBALL", "BELLAIRE"]:
        search_addr = re.sub(r",?\s*" + city + r"\b", "", search_addr)
    search_addr = re.sub(r",?\s*\bTX\b\.?,?", "", search_addr)
    search_addr = re.sub(r",?\s*\bTEXAS\b", "", search_addr)
    search_addr = re.sub(r"\s*\d{5}(-\d{4})?\s*$", "", search_addr)
    return re.sub(r"\s+", " ", search_addr).strip().strip(",").strip()


from core.hcad_identity import clean_account


def safe_filename_part(value: str, fallback: str = "PROPERTY") -> str:
    part = re.sub(r"[^A-Z0-9]+", "_", (value or "").upper()).strip("_")
    return part[:120] if part else fallback


def parse_search_line(line: str) -> SearchItem | None:
    raw = (line or "").strip()
    if not raw or raw.startswith("#"):
        return None
    if raw.lower().startswith("address") and ("tax" in raw.lower() or "account" in raw.lower()):
        return None

    # Supported formats:
    #   1601 BANKS ST
    #   1601 BANKS ST<TAB>0123456789012
    #   1601 BANKS ST | 0123456789012
    #   0123456789012 | 1601 BANKS ST
    parts = [p.strip() for p in re.split(r"\t|\|", raw) if p.strip()]
    if len(parts) >= 2:
        p1, p2 = parts[0], parts[1]
        a1, a2 = clean_account(p1), clean_account(p2)
        if a1 and not a2:
            return SearchItem(address=clean_address(p2), account=a1)
        if a2:
            return SearchItem(address=clean_address(p1), account=a2)
        return SearchItem(address=clean_address(p1), account="")

    acct = clean_account(raw)
    if acct and len(re.sub(r"\D+", "", raw)) == len(raw.replace(" ", "")):
        return SearchItem(address="", account=acct)
    return SearchItem(address=clean_address(raw), account="")


def unique_search_items(items):
    result=[];seen=set()
    for item in items:
        item=SearchItem(clean_address(item.address),clean_account(item.account))
        key=('account',item.account) if item.account else ('address',item.address.upper())
        if not (item.account or item.address) or key in seen:continue
        seen.add(key);result.append(item)
    return result


def load_search_items(path: str | Path) -> list[SearchItem]:
    with open(path,'r',encoding='utf-8-sig') as stream:
        return unique_search_items(item for line in stream if (item:=parse_search_line(line)))



def page_text(page, timeout: int = 2000) -> str:
    try:
        return page.inner_text("body", timeout=timeout)
    except Exception:
        return ""


def page_needs_verification(page) -> bool:
    text = page_text(page, timeout=1000).lower()
    verify_terms = ["verify", "human", "checking your browser", "captcha", "cf-turnstile", "challenge"]
    return any(term in text for term in verify_terms)


def page_has_property_record(page) -> bool:
    text = page_text(page, timeout=2000)
    normalized_text = re.sub(r"\s+", " ", text)
    if "Account:" not in normalized_text:
        return False
    return any(marker in normalized_text for marker in PROPERTY_DETAIL_MARKERS)


def page_has_search_ready(page) -> bool:
    text = page_text(page, timeout=2000).lower()
    if "find property records" in text:
        return True
    # fallback: search inputs visible
    for selector in ["#PROPERTYADDRESS", "input.searchTerm", "button.btn-primary[aria-label='Search']"]:
        try:
            if page.locator(selector).count() > 0:
                return True
        except Exception:
            pass
    return False


def wait_for_hcad_ready(page, timeout: int = 180, verify_first: bool = False) -> bool:
    """Wait for HCAD search page or property page. Planner may complete verification manually."""
    start = time.time()
    announced = False
    while time.time() - start < timeout:
        if page_has_search_ready(page) or page_has_property_record(page):
            return True
        if page_needs_verification(page) and not announced:
            print()
            print("   >>> HCAD is asking for human verification.")
            print("   >>> Complete it in the Chrome window. The downloader will wait here.")
            if verify_first:
                print("   >>> After the search page appears, the batch will continue automatically.")
            print()
            announced = True
        time.sleep(2)
    return False


def wait_for_property_record(page, timeout: int = 90, announce_captcha: bool = True) -> bool:
    start = time.time()
    announced = False
    while time.time() - start < timeout:
        if page_has_property_record(page):
            return True
        if announce_captcha and page_needs_verification(page) and not announced:
            print("   >>> HCAD verification detected. Complete it in the Chrome window.")
            announced = True
        time.sleep(1)
    return False


def open_result_candidate(page, locator) -> bool:
    try:
        locator.scroll_into_view_if_needed(timeout=3000)
    except Exception:
        pass
    try:
        href = locator.get_attribute("href", timeout=1000)
    except Exception:
        href = None
    try:
        if href and not href.lower().startswith(("javascript:", "#")):
            page.goto(urljoin(page.url, href), timeout=60000)
        else:
            locator.click(timeout=5000)
    except Exception:
        try:
            locator.click(timeout=5000, force=True)
        except Exception:
            return False
    try:
        page.wait_for_load_state("domcontentloaded", timeout=10000)
    except Exception:
        pass
    return wait_for_property_record(page, timeout=15)


def click_first_search_result(page) -> bool:
    try:
        page.wait_for_selector("tr.resulttr, table tbody tr", timeout=15000)
    except Exception:
        pass

    body_text = page_text(page, timeout=2000)
    account_match = re.search(r"\b\d{10,13}\b", body_text)
    account_number = account_match.group(0) if account_match else None

    candidate_selectors = ["tr.resulttr", "tr:has(td.resulttd)", "table tbody tr:has(td)"]
    if account_number:
        candidate_selectors.insert(0, f"tr:has-text('{account_number}')")
        candidate_selectors.insert(1, f"text={account_number}")

    for row_selector in candidate_selectors:
        rows = page.locator(row_selector)
        try:
            row_count = min(rows.count(), 10)
        except Exception:
            row_count = 0
        for row_index in range(row_count):
            row = rows.nth(row_index)
            try:
                row_text = row.inner_text(timeout=1000).strip().lower()
            except Exception:
                row_text = ""
            if not row_text:
                continue
            if len(row_text) < 120 and all(word in row_text for word in ["account", "address", "owner"]):
                continue
            if open_result_candidate(page, row):
                return True
    return False


def click_search_mode(page, mode: str) -> None:
    """Select HCAD search mode. mode is 'address' or 'account'."""
    selector_candidates = []
    if mode == "account":
        selector_candidates = [
            "#ACCOUNT",
            "input#ACCOUNT",
            "label:has-text('Account')",
            "text=Account",
        ]
    else:
        selector_candidates = [
            "#PROPERTYADDRESS",
            "input#PROPERTYADDRESS",
            "label:has-text('Property Address')",
            "text=Property Address",
        ]
    for selector in selector_candidates:
        try:
            loc = page.locator(selector)
            if loc.count() > 0:
                loc.first.click(timeout=3000)
                time.sleep(0.3)
                return
        except Exception:
            continue


def perform_search(page, search_value: str, mode: str) -> None:
    if "search.hcad.org" not in page.url.lower():
        page.goto(HCAD_HOME, timeout=60000)
        wait_for_hcad_ready(page, timeout=180)

    click_search_mode(page, mode)
    time.sleep(random.uniform(0.4, 1.0))
    # HCAD currently uses input.searchTerm for the active search method.
    search_box = page.locator("input.searchTerm").first
    search_box.fill("")
    time.sleep(random.uniform(0.2, 0.6))
    search_box.fill(search_value)
    time.sleep(random.uniform(0.5, 1.2))
    page.click("button.btn-primary[aria-label='Search']")


def find_existing_pdf(item: SearchItem) -> Path | None:
    account=clean_account(item.account)
    candidates=[]
    if account:candidates.extend(sorted(SAVE_FOLDER.glob(f"HCAD_*{account}*.pdf")))
    if item.address:candidates.extend(sorted(SAVE_FOLDER.glob(f"HCAD_{safe_filename_part(clean_address(item.address))}*.pdf")))
    seen=set()
    for path in candidates:
        if path in seen:continue
        seen.add(path)
        try:
            if not path.is_file() or not path.stat().st_size:continue
            if account:
                from core.application_import import record_pdf_facts
                if clean_account(record_pdf_facts(path)['account'])!=account:continue
            return path
        except Exception:
            # An unreadable/wrong record must not hide a later matching PDF.
            continue
    return None



def save_property_pdf(page, item: SearchItem) -> str | None:
    print("   Saving PDF...")
    safe_addr = safe_filename_part(clean_address(item.address)) if item.address else "PROPERTY"
    suffix = f"_{item.account}" if item.account else ""
    filename = f"HCAD_{safe_addr}{suffix}.pdf"
    pdf_path = SAVE_FOLDER / filename
    try:
        cdp = page.context.new_cdp_session(page)
        result = cdp.send("Page.printToPDF", {
            "landscape": False,
            "printBackground": True,
            "paperWidth": 8.5,
            "paperHeight": 11,
        })
        cdp.detach()
        with open(pdf_path, "wb") as f:
            f.write(base64.b64decode(result["data"]))
        print(f"   SUCCESS! Saved: {filename}")
        return str(pdf_path)
    except Exception as e:
        print(f"   PDF save error: {e}")
        return None


def download_hcad_record(page, item: SearchItem, resume: bool = True) -> DownloadResult:
    """Search for a property and save the HCAD page as PDF.

    v14.1 search order:
    1. Search by property address first when an address is available.
    2. Only if address search fails, fallback to a cleaned 13-digit HCAD account
       number. This avoids failures caused by workbook TAX_ID values missing a
       leading zero or otherwise not matching HCAD's account search exactly.
    """
    if resume:
        existing = find_existing_pdf(item)
        if existing and existing.exists() and existing.stat().st_size > 0:
            print(f"   Already downloaded, skipping: {existing.name}")
            return DownloadResult(item=item, status="Skipped", pdf_path=str(existing), notes="Existing PDF found")

    search_addr = clean_address(item.address)
    search_account = clean_account(item.account)

    attempts: list[tuple[str, str]] = []
    if search_addr:
        attempts.append(("address", search_addr))
    if search_account:
        attempts.append(("account", search_account))

    if not attempts:
        return DownloadResult(item=item, status="Failed", notes="Missing usable address or valid 13-digit account number")

    failure_notes: list[str] = []

    for attempt_index, (mode, search_value) in enumerate(attempts, start=1):
        if mode == "account" and search_addr:
            print(f"   Address search failed. Trying 13-digit account fallback: {search_value}")
        else:
            print(f"   Searching by {mode}: {search_value}")

        try:
            perform_search(page, search_value, mode=mode)
        except Exception as e:
            note = f"{mode.title()} search error: {e}"
            failure_notes.append(note)
            continue

        try:
            page.wait_for_selector("table tbody tr, [role='row'], .search-results, .results, .no-results, .alert", timeout=30000)
        except PlaywrightTimeoutError:
            pass
        except Exception:
            pass

        opened_record = False
        if page_has_property_record(page):
            print("   Auto-redirected to property page.")
            opened_record = True
        else:
            text = page_text(page, timeout=3000).lower()
            if "no results" in text or "no records" in text:
                failure_notes.append(f"{mode.title()} search returned no results for {search_value}")
                continue
            print("   Opening first result...")
            if click_first_search_result(page):
                opened_record = True
            else:
                failure_notes.append(f"{mode.title()} search found results but could not open the first result")
                continue

        if opened_record:
            print("   Waiting for property page...")
            if not wait_for_property_record(page, timeout=120, announce_captcha=True):
                failure_notes.append(f"{mode.title()} search did not load a property page; check browser/verification")
                continue

            if search_account:
                record_text = page_text(page, timeout=3000)
                match = re.search(r'Account(?:\s+(?:Number|No\.?))?\s*:?\s*(\d{10,13})\b', record_text, re.I)
                actual = clean_account(match.group(1)) if match else ''
                if actual != search_account:
                    failure_notes.append(f'{mode.title()} search opened account {actual or "unreadable"}, expected {search_account}; record was not saved')
                    continue

            time.sleep(random.uniform(3.5, 6.5))
            pdf = save_property_pdf(page, item)
            if pdf:
                note = "Downloaded by address search" if mode == "address" else "Downloaded by 13-digit account fallback"
                return DownloadResult(item=item, status="Success", pdf_path=pdf, notes=note)
            failure_notes.append(f"{mode.title()} search loaded a property page but PDF save failed")

    status = "Needs Review" if any("could not open" in n.lower() or "did not load" in n.lower() for n in failure_notes) else "Failed"
    return DownloadResult(item=item, status=status, notes="; ".join(failure_notes) or "No usable search result")

def write_report(report_path: Path, results: list[DownloadResult]) -> None:
    report_path.parent.mkdir(parents=True, exist_ok=True)
    with open(report_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.writer(f)
        writer.writerow(["Address", "Account", "Status", "PDF File", "Notes", "Downloaded At"])
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        for r in results:
            writer.writerow([r.item.address, r.item.account, r.status, r.pdf_path, r.notes, now])


def run_batch(page, items: list[SearchItem], delay_min: float, delay_max: float, resume: bool) -> list[DownloadResult]:
    items=unique_search_items(items)
    results: list[DownloadResult] = []
    for i, item in enumerate(items, 1):
        print(f"\n   [{i}/{len(items)}] {item.display}")
        try:
            result = download_hcad_record(page, item, resume=resume)
        except Exception as e:
            result = DownloadResult(item=item, status="Failed", notes=f"Unexpected error: {e}")
        results.append(result)

        if result.pdf_path:
            print(f"   Result: {result.status} - {Path(result.pdf_path).name}")
        else:
            print(f"   Result: {result.status} - {result.notes}")

        try:
            page.goto(HCAD_HOME, timeout=30000)
            wait_for_hcad_ready(page, timeout=90)
        except Exception:
            pass
        if i < len(items):
            delay = random.uniform(delay_min, delay_max)
            print(f"   Pausing {delay:.1f} seconds before next record...")
            time.sleep(delay)
    return results


def launch_context(p, profile_dir: Path):
    profile_dir.mkdir(parents=True, exist_ok=True)
    # Persistent context keeps HCAD cookies/session after the planner completes verification.
    launch_kwargs = dict(
        headless=False,
        viewport={"width": 1400, "height": 900},
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/122.0.0.0 Safari/537.36",
        args=["--disable-blink-features=AutomationControlled"],
    )
    try:
        return p.chromium.launch_persistent_context(str(profile_dir), channel="chrome", **launch_kwargs)
    except Exception:
        # Fallback if Chrome channel is unavailable; Playwright's bundled Chromium still uses the same persistent profile folder.
        return p.chromium.launch_persistent_context(str(profile_dir), **launch_kwargs)


def main() -> None:
    parser = argparse.ArgumentParser(add_help=True)
    parser.add_argument("--file", default=None, help="Text file with one address/account per line")
    parser.add_argument("--output", default=None, help="Folder where HCAD PDFs should be saved")
    parser.add_argument("--profile", default=None, help="Persistent Chrome profile folder")
    parser.add_argument("--verify-first", action="store_true", help="Open HCAD and wait for planner verification before batch starts")
    parser.add_argument("--delay-min", type=float, default=8.0, help="Minimum pause between records")
    parser.add_argument("--delay-max", type=float, default=15.0, help="Maximum pause between records")
    parser.add_argument("--no-resume", action="store_true", help="Do not skip existing PDFs")
    parser.add_argument("--cdp-url", default=None, help="Connect to an already-running Chrome remote-debugging session, e.g. http://127.0.0.1:9222")
    args = parser.parse_args()

    global SAVE_FOLDER
    if args.output:
        SAVE_FOLDER = Path(args.output).expanduser().resolve()
    SAVE_FOLDER.mkdir(parents=True, exist_ok=True)

    profile_dir = Path(args.profile).expanduser().resolve() if args.profile else SAVE_FOLDER / "hcad_chrome_profile"
    report_path = SAVE_FOLDER / "HCAD_Download_Report.csv"

    print()
    print("  ╔═══════════════════════════════════════════╗")
    print("  ║     HCAD Property Record Downloader       ║")
    print("  ╚═══════════════════════════════════════════╝")
    print()
    print(f"   PDFs save to: {SAVE_FOLDER}")
    if args.cdp_url:
        print(f"   Chrome connection: {args.cdp_url}")
    else:
        print(f"   Chrome profile: {profile_dir}")
    print(f"   Report: {report_path}")
    print()

    try:
        with sync_playwright() as p:
            connected_to_existing = False
            browser = None
            context = None

            if args.cdp_url:
                print("   Connecting to existing Chrome session...")
                print("   Make sure HCAD verification is already complete in that Chrome window.")
                browser = p.chromium.connect_over_cdp(args.cdp_url)
                connected_to_existing = True
                context = browser.contexts[0] if browser.contexts else browser.new_context()
                existing_pages = []
                for ctx in browser.contexts:
                    existing_pages.extend(ctx.pages)
                page = next((pg for pg in existing_pages if "search.hcad.org" in (pg.url or "").lower()), None)
                if page is None:
                    page = context.new_page()
            else:
                print("   Starting legacy persistent Chrome browser...")
                context = launch_context(p, profile_dir)
                page = context.pages[0] if context.pages else context.new_page()

            print("   Opening HCAD...")
            if "search.hcad.org" not in (page.url or "").lower():
                page.goto(HCAD_HOME, timeout=60000)
            else:
                try:
                    page.bring_to_front()
                except Exception:
                    pass
            if not wait_for_hcad_ready(page, timeout=240, verify_first=args.verify_first):
                print("   HCAD did not become ready. Complete verification and run again.")
                if context is not None and not connected_to_existing:
                    context.close()
                return

            if args.verify_first:
                print("   HCAD is ready. Starting batch.")

            if args.file:
                items = load_search_items(args.file)
                print(f"\n   Auto-batch: {len(items)} records from {args.file}\n")
                results = run_batch(page, items, args.delay_min, args.delay_max, resume=not args.no_resume)
                write_report(report_path, results)
                success = sum(1 for r in results if r.status in {"Success", "Skipped"})
                print(f"\n   Completed! {success}/{len(items)} saved or already present.")
                print(f"   Report saved: {report_path}")
                if context is not None and not connected_to_existing:
                    print("\n   Closing browser...")
                    context.close()
                else:
                    print("\n   Leaving your Chrome session open.")
                return

            print()
            print("   ╔═══════════════════════════════════════════╗")
            print("   ║  Ready! Choose an option:                 ║")
            print("   ║                                           ║")
            print("   ║  1 = Single address mode                  ║")
            print("   ║  2 = Batch mode (multiple addresses)      ║")
            print("   ║  Q = Quit                                 ║")
            print("   ╚═══════════════════════════════════════════╝")
            print()
            mode = input("   Choose (1/2/Q): ").strip()
            results: list[DownloadResult] = []

            if mode.upper() == "Q":
                pass
            elif mode == "2":
                print()
                print("   Paste addresses/accounts, one per line. Type DONE when finished.")
                raw_lines = []
                while True:
                    line = input("   ").strip()
                    if line.upper() == "DONE":
                        break
                    if line:
                        raw_lines.append(line)
                items = [i for i in (parse_search_line(x) for x in raw_lines) if i]
                if items:
                    results = run_batch(page, items, args.delay_min, args.delay_max, resume=not args.no_resume)
            else:
                print("\n   Type address/account and press Enter. Type Q to quit.\n")
                while True:
                    line = input("   Address/account: ").strip()
                    if not line or line.upper() == "Q":
                        break
                    item = parse_search_line(line)
                    if not item:
                        continue
                    results.extend(run_batch(page, [item], args.delay_min, args.delay_max, resume=not args.no_resume))

            if results:
                write_report(report_path, results)
                print(f"\n   Report saved: {report_path}")
            if context is not None and not connected_to_existing:
                print("\n   Closing browser...")
                context.close()
            else:
                print("\n   Leaving your Chrome session open.")
    except Exception as e:
        print(f"\n   Error: {e}")
        print("\n   Make sure Chrome browser is installed, or run install_requirements.bat first.")

    print("\n   Done!")
    input("\n   Press Enter to close...")


if __name__ == "__main__":
    main()
