"""
Recipe scraper (GitHub Actions ready):
  1. Collects links from the main page (follows "Next page" until MAX_LINKS is reached)
  2. Opens every link and extracts: image, title (from CSS ::before), caption
  3. Downloads images (original file names)
  4. Saves recipes.xlsx  (image_filename | caption | fullname)
  5. ONLY AFTER everything is scraped: uploads images + recipes.xlsx to MEGA ONCE

Settings come from environment variables (GitHub workflow inputs / secrets):
  MAIN_URL, MAX_LINKS, MEGA_USER, MEGA_PASS (obscured password)
"""
import asyncio
import os
import re
import shutil
import subprocess
import sys
import threading
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import unquote, urljoin, urlparse

import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

# ====== SETTINGS ======
MAIN_URL = os.environ.get("MAIN_URL") or "https://foodiesposts.com/page1/"
MAX_LINKS = int(os.environ.get("MAX_LINKS") or 10)
LINK_SELECTOR = "#main-grid .grid_item a[href]"
NEXT_SELECTOR = ".paginator a.next[href]"
MAX_PAGES = 50

IMG_SELECTOR = "#recipe-card .card-image img"
TITLE_SELECTOR = "#recipe-card .card-title"
DESC_SELECTOR = "#recipe-card .desc p"

OUT_DIR = "images"
XLSX_FILE = "recipes.xlsx"

MEGA_REMOTE = "mega"
MEGA_TYPE = "mega"
MEGA_USER = os.environ.get("MEGA_USER", "your@email.com")
MEGA_PASS = os.environ.get("MEGA_PASS", "<OBSCURED_PASSWORD>")
MEGA_FOLDER = "recipe_images"
DELAY = 0.3                                   # pause between pagination requests
WORKERS = int(os.environ.get("WORKERS") or 8)  # parallel recipe pages / image downloads
BROWSER_WORKERS = 3                            # parallel headless browsers (slow fallback)
# ======================

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                         "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}

TITLE_TAIL = TITLE_SELECTOR.split()[-1]


# ---------- fetching ----------
def get_html_requests(url):
    r = requests.get(url, headers=HEADERS, timeout=30)
    r.raise_for_status()
    return r.text


async def _get_html_browser(url, wait_selector):
    from playwright.async_api import async_playwright
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True)
        page = await browser.new_page(user_agent=HEADERS["User-Agent"])
        await page.goto(url, wait_until="networkidle", timeout=60000)
        await page.wait_for_selector(wait_selector, state="attached", timeout=30000)
        for _ in range(3):
            await page.mouse.wheel(0, 3000)
            await page.wait_for_timeout(1000)
        html = await page.content()
        css_title = await page.evaluate(
            """(sel) => {
                const el = document.querySelector(sel);
                return el ? getComputedStyle(el, '::before').content : '';
            }""",
            TITLE_SELECTOR,
        )
        await browser.close()
    return html, clean_css_content(css_title)


def get_page_browser(url, wait_selector):
    """Returns (html, title_from_css_computed_style)."""
    loop = asyncio.new_event_loop()
    try:
        return loop.run_until_complete(_get_html_browser(url, wait_selector))
    finally:
        loop.close()


def get_html_browser(url, wait_selector):
    return get_page_browser(url, wait_selector)[0]


# ---------- CSS title ----------
_css_cache = {}


def fetch_css(url):
    if url not in _css_cache:
        try:
            _css_cache[url] = get_html_requests(url)
        except Exception as e:
            print(f"  could not load stylesheet {url}: {e}")
            _css_cache[url] = ""
    return _css_cache[url]


def collect_css(soup, page_url):
    parts = [s.get_text() for s in soup.find_all("style")]
    for link in soup.find_all("link", href=True):
        rel = link.get("rel") or []
        rel = rel if isinstance(rel, str) else " ".join(rel)
        if "stylesheet" in rel.lower():
            parts.append(fetch_css(urljoin(page_url, link["href"])))
    return "\n".join(parts)


def unescape_css(s):
    s = re.sub(r"\\([0-9a-fA-F]{1,6})\s?", lambda m: chr(int(m.group(1), 16)), s)
    return re.sub(r"\\(.)", r"\1", s)


def clean_css_content(value):
    value = (value or "").strip()
    if value in ("", "none", "normal"):
        return ""
    m = re.match(r"""^(["'])(.*)\1$""", value, re.S)
    if m:
        value = m.group(2)
    return unescape_css(value).strip()


def title_from_css(css_text):
    css_text = re.sub(r"/\*.*?\*/", "", css_text, flags=re.S)
    title = ""
    for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", css_text):
        selectors = [s.strip() for s in m.group(1).split(",")]
        if not any(s.endswith((TITLE_TAIL + "::before", TITLE_TAIL + ":before")) for s in selectors):
            continue
        c = re.search(r"""content\s*:\s*(["'])(.*?)(?<!\\)\1""", m.group(2), re.S)
        if c:
            title = unescape_css(c.group(2)).strip()
    return title


# ---------- parsing ----------
def extract_links(html, base_url, limit, links=None, seen=None):
    soup = BeautifulSoup(html, "html.parser")
    links = [] if links is None else links
    seen = set(links) if seen is None else seen
    for a in soup.select(LINK_SELECTOR):
        if len(links) >= limit:
            break
        href = a.get("href", "").strip()
        if not href or href.startswith(("#", "javascript:")):
            continue
        full = urljoin(base_url, href)
        if full not in seen:
            seen.add(full)
            links.append(full)
            print(f"  link {len(links)}/{limit}: {full}", flush=True)
    return links


def find_next_page(html, base_url):
    soup = BeautifulSoup(html, "html.parser")
    a = soup.select_one(NEXT_SELECTOR)
    if not a:
        return ""
    href = a.get("href", "").strip()
    if not href or href.startswith(("#", "javascript:")):
        return ""
    return urljoin(base_url, href)


def collect_links(start_url, limit):
    links, seen = [], set()
    visited = set()
    page_url = start_url
    page_no = 0
    while page_url and page_url not in visited and len(links) < limit and page_no < MAX_PAGES:
        visited.add(page_url)
        page_no += 1
        before = len(links)

        html = ""
        try:
            html = get_html_requests(page_url)
            extract_links(html, page_url, limit, links, seen)
        except Exception as e:
            print(f"requests failed on page {page_no}: {e}")
        if len(links) == before:
            print(f"Page {page_no}: trying browser rendering...")
            try:
                html = get_html_browser(page_url, LINK_SELECTOR)
                extract_links(html, page_url, limit, links, seen)
            except Exception as e:
                print(f"browser failed on page {page_no}: {e}")

        print(f"Page {page_no}: {len(links) - before} new links ({len(links)}/{limit}) - {page_url}")
        if len(links) >= limit:
            break

        next_url = find_next_page(html, page_url) if html else ""
        if not next_url:
            print("No next page - stopping.")
            break
        page_url = next_url
        time.sleep(DELAY)
    return links


def parse_recipe(html, page_url, css_title=""):
    soup = BeautifulSoup(html, "html.parser")
    img = soup.select_one(IMG_SELECTOR)
    desc = soup.select_one(DESC_SELECTOR)
    img_url = ""
    if img:
        src = img.get("src") or img.get("data-src") or ""
        img_url = urljoin(page_url, src.strip()) if src else ""

    fullname = css_title or title_from_css(collect_css(soup, page_url))
    if not fullname:
        el = soup.select_one(TITLE_SELECTOR)
        fullname = el.get_text(strip=True) if el else ""

    return {
        "img_url": img_url,
        "fullname": fullname,
        "caption": desc.get_text(" ", strip=True) if desc else "",
    }


def safe_filename(name, fallback):
    name = re.sub(r'[\\/:*?"<>|]+', "_", name).strip(" .")
    return name or fallback


_used_names = set()
_names_lock = threading.Lock()


def download_image(img_url, fullname, index):
    original = unquote(os.path.basename(urlparse(img_url).path))
    original = safe_filename(original, f"recipe_{index}.jpg")
    stem, ext = os.path.splitext(original)
    if not ext:
        ext = ".jpg"
    filename = stem + ext
    n = 2
    with _names_lock:
        while filename.lower() in _used_names:
            filename = f"{stem}_{n}{ext}"
            n += 1
        _used_names.add(filename.lower())
    path = os.path.join(OUT_DIR, filename)
    r = requests.get(img_url, headers=HEADERS, timeout=60)
    r.raise_for_status()
    with open(path, "wb") as f:
        f.write(r.content)
    return filename


# ---------- MEGA upload (called ONCE, at the very end) ----------
def upload_to_mega():
    if "your@email.com" in MEGA_USER or "OBSCURED_PASSWORD" in MEGA_PASS:
        print("MEGA_USER / MEGA_PASS not set - skipping upload.")
        return False
    if not shutil.which("rclone"):
        print("rclone not found, installing...")
        subprocess.run("curl -fsSL https://rclone.org/install.sh | sudo bash", shell=True, check=True)

    name = MEGA_REMOTE.upper()
    env = os.environ.copy()
    env[f"RCLONE_CONFIG_{name}_TYPE"] = MEGA_TYPE
    env[f"RCLONE_CONFIG_{name}_USER"] = MEGA_USER
    env[f"RCLONE_CONFIG_{name}_PASS"] = MEGA_PASS
    dest = f"{MEGA_REMOTE}:{MEGA_FOLDER}"

    print(f"Uploading everything to {dest} ...")
    r1 = subprocess.run(["rclone", "copy", OUT_DIR, dest], env=env, capture_output=True, text=True)
    r2 = subprocess.run(["rclone", "copyto", XLSX_FILE, f"{dest}/{XLSX_FILE}"],
                        env=env, capture_output=True, text=True)
    ok = r1.returncode == 0 and r2.returncode == 0
    if not ok:
        print("Upload FAILED:")
        print((r1.stderr or "") + (r2.stderr or ""))
    else:
        print("Upload finished.")
    return ok


# ---------- main ----------
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    print(f"Start URL: {MAIN_URL} | max links: {MAX_LINKS}")

    # 1) collect links
    links = collect_links(MAIN_URL, MAX_LINKS)
    print(f"Collected {len(links)} links")

    # 2) scrape + download in PARALLEL (NO uploading here)
    total = len(links)

    def finish(i, link, data):
        """Download the image and return the xlsx row (or None)."""
        if not data["img_url"]:
            print(f"[{i}/{total}] no image on {link}", flush=True)
            return None
        filename = download_image(data["img_url"], data["fullname"], i)
        print(f"[{i}/{total}] OK: {data['fullname']} -> {filename}", flush=True)
        return (filename, data["caption"], data["fullname"])

    def fast_job(i, link):
        """Fast path: plain HTTP. Returns ('row', row) / ('skip', None) / ('browser', None)."""
        try:
            data = parse_recipe(get_html_requests(link), link)
        except Exception:
            return "browser", None
        if not data["fullname"]:
            return "browser", None
        try:
            row = finish(i, link, data)
        except Exception as e:
            print(f"[{i}/{total}] failed {link}: {e}", flush=True)
            return "skip", None
        return ("row", row) if row else ("skip", None)

    def browser_job(i, link):
        """Slow path: headless browser, only for pages the fast path could not read."""
        try:
            html, css_title = get_page_browser(link, TITLE_SELECTOR)
            return finish(i, link, parse_recipe(html, link, css_title))
        except Exception as e:
            print(f"[{i}/{total}] failed {link}: {e}", flush=True)
            return None

    results = {}
    need_browser = []
    print(f"Scraping {total} pages with {WORKERS} parallel workers...", flush=True)
    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futs = {ex.submit(fast_job, i, link): (i, link) for i, link in enumerate(links, 1)}
        for f in as_completed(futs):
            i, link = futs[f]
            kind, row = f.result()
            if kind == "row":
                results[i] = row
            elif kind == "browser":
                need_browser.append((i, link))

    if need_browser:
        print(f"{len(need_browser)} pages need the browser ({BROWSER_WORKERS} at a time)...", flush=True)
        with ThreadPoolExecutor(max_workers=BROWSER_WORKERS) as ex:
            futs = {ex.submit(browser_job, i, link): i for i, link in need_browser}
            for f in as_completed(futs):
                row = f.result()
                if row:
                    results[futs[f]] = row

    rows = [results[i] for i in sorted(results)]   # keep the original page order

    # 3) write xlsx
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipes"
    ws.append(["image_filename", "caption", "fullname"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for row in rows:
        ws.append(row)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 35
    wb.save(XLSX_FILE)
    print(f"Saved {XLSX_FILE} with {len(rows)} rows")

    # 4) scraping is fully done -> upload ONCE to MEGA
    if rows:
        if not upload_to_mega():
            sys.exit(1)


if __name__ == "__main__":
    main()
