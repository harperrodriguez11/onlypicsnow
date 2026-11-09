"""
Recipe scraper (disconnect-proof version):
  1. Collects links from the main grid page; if fewer than MAX_LINKS are found,
     follows the paginator "Next page" button until MAX_LINKS is reached or no pages are left
  2. Opens every link and extracts: image, title (fullname), desc (caption)
     NOTE: the title is NOT in the HTML text. It lives in CSS:
         #recipe-card .card-title::before { content: "..." }
     so it is read from the page's CSS (inline <style> + linked stylesheets),
     with a browser fallback that reads the computed ::before content.
  3. Downloads images (ORIGINAL file names kept)
  4. Saves recipes.xlsx  (image_filename | caption | fullname)
  5. Uploads images + recipes.xlsx to MEGA with rclone

DISCONNECT PROTECTION (Colab etc.):
  - Every UPLOAD_EVERY recipes, the images, recipes.xlsx and progress.json are uploaded to MEGA.
  - If the runtime dies, just run the script again: it restores progress.json from MEGA,
    skips the recipes already done, and continues where it stopped.
  - Ctrl-C / stopping the cell also triggers a final upload.

Setup:
  pip install requests beautifulsoup4 openpyxl playwright nest_asyncio
  playwright install chromium            # only needed for the browser fallback
  install rclone: https://rclone.org/install/   (the script installs it automatically if missing)

MEGA login is written directly in the SETTINGS block below (MEGA_USER, MEGA_PASS).
No rclone.conf file is needed. MEGA_PASS must be the OBSCURED password:
  rclone obscure "yourpassword"     -> paste the output into MEGA_PASS
"""
import asyncio
import json
import os
import re
import shutil
import subprocess
import time
from urllib.parse import unquote, urljoin, urlparse

import nest_asyncio
import requests
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font

nest_asyncio.apply()

# ====== SETTINGS ======
MAIN_URL = "https://www.onlypics.com/teen/"        # <-- FIX: put your real main URL here
MAX_LINKS = 100
LINK_SELECTOR = "#main-grid .grid_item a[href]"
NEXT_SELECTOR = ".paginator a.next[href]"   # "Next page" button
MAX_PAGES = 50                              # safety limit so it can never loop forever

IMG_SELECTOR = "#model-card .card-image img"
TITLE_SELECTOR = "#model-card .card-title"
DESC_SELECTOR = "#model-card .desc p"

OUT_DIR = "images"
XLSX_FILE = "recipes.xlsx"
STATE_FILE = "progress.json"             # remembers finished recipes (also uploaded to MEGA)
UPLOAD_EVERY = 5                         # upload to MEGA after every N recipes
# --- MEGA login (same as the [mega] block, but written here, no rclone.conf) ---

MEGA_REMOTE = "mega"                     # [mega]  (name of the remote)
MEGA_TYPE = "mega"                       # type = mega
MEGA_USER = "tanweerahmedjw8@gmail.com"             # user = your@email.com
MEGA_PASS = "tGCuikkmr3crlky87xBplmvL02yCHSCWAVxPH5U"        # pass = <OBSCURED_PASSWORD>  (get it with: rclone obscure "yourpassword")
MEGA_FOLDER = "recipe_images"   
        # folder inside MEGA
DELAY = 1.0                     # seconds between requests (be polite)
# ======================

HEADERS = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                         "AppleWebKit/537.36 Chrome/124.0 Safari/537.36"}

# Last part of TITLE_SELECTOR, e.g. ".card-title" -> used to match CSS rules
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
        # computed ::before content of the title element (the real rendered title)
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
    return asyncio.get_event_loop().run_until_complete(_get_html_browser(url, wait_selector))


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
    """Inline <style> blocks + every linked stylesheet (cached between pages)."""
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
    """'"Ashley ..."' (computed style) -> 'Ashley ...'; 'none'/'normal' -> ''."""
    value = (value or "").strip()
    if value in ("", "none", "normal"):
        return ""
    m = re.match(r"""^(["'])(.*)\1$""", value, re.S)
    if m:
        value = m.group(2)
    return unescape_css(value).strip()


def title_from_css(css_text):
    """Find `<...>.card-title::before { content: "..." }` and return its text."""
    css_text = re.sub(r"/\*.*?\*/", "", css_text, flags=re.S)
    title = ""
    for m in re.finditer(r"([^{}]+)\{([^{}]*)\}", css_text):
        selectors = [s.strip() for s in m.group(1).split(",")]
        if not any(s.endswith((TITLE_TAIL + "::before", TITLE_TAIL + ":before")) for s in selectors):
            continue
        c = re.search(r"""content\s*:\s*(["'])(.*?)(?<!\\)\1""", m.group(2), re.S)
        if c:
            title = unescape_css(c.group(2)).strip()   # last matching rule wins (cascade)
    return title


# ---------- parsing ----------
def extract_links(html, base_url, limit, links=None, seen=None):
    """Adds new recipe links from one page to `links` (shared across pages) until `limit` is reached."""
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
    return links


def find_next_page(html, base_url):
    """Returns the absolute URL of the 'Next page' button, or '' if there isn't one."""
    soup = BeautifulSoup(html, "html.parser")
    a = soup.select_one(NEXT_SELECTOR)
    if not a:
        return ""
    href = a.get("href", "").strip()
    if not href or href.startswith(("#", "javascript:")):
        return ""
    return urljoin(base_url, href)


def collect_links(start_url, limit):
    """Walks the paginator (Next page) until `limit` links are found or there are no more pages."""
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
        if len(links) == before:  # nothing found -> maybe JS-rendered, try the browser
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

    # Title always comes from CSS (::before content)
    fullname = css_title or title_from_css(collect_css(soup, page_url))
    if not fullname:  # last resort: normal text inside the element, if any
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


def download_image(img_url, fullname, index):
    # Keep the ORIGINAL file name from the image URL
    original = unquote(os.path.basename(urlparse(img_url).path))
    original = safe_filename(original, f"recipe_{index}.jpg")
    stem, ext = os.path.splitext(original)
    if not ext:
        ext = ".jpg"
    filename = stem + ext
    # Only if two recipes use the exact same original name, add a suffix so one doesn't overwrite the other
    n = 2
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


# ---------- MEGA ----------
def mega_ready():
    """Checks the login values and installs rclone if needed. Done at START so problems show up early."""
    if "your@email.com" in MEGA_USER or "OBSCURED_PASSWORD" in MEGA_PASS:
        print("Set MEGA_USER and MEGA_PASS in the SETTINGS block - uploads are disabled.")
        return False
    if not shutil.which("rclone"):
        print("rclone not found, installing...")
        try:
            subprocess.run("curl -fsSL https://rclone.org/install.sh | bash", shell=True, check=True)
        except Exception as e:
            print("rclone install failed:", e)
            return False
    return True


def _mega_env():
    # Builds the [mega] remote on the fly from the values in SETTINGS (no rclone.conf needed)
    name = MEGA_REMOTE.upper()
    env = os.environ.copy()
    env[f"RCLONE_CONFIG_{name}_TYPE"] = MEGA_TYPE
    env[f"RCLONE_CONFIG_{name}_USER"] = MEGA_USER
    env[f"RCLONE_CONFIG_{name}_PASS"] = MEGA_PASS
    return env


def _rclone(args, quiet=False):
    """Runs rclone. Never raises, so a failed upload can't kill the scraping."""
    try:
        r = subprocess.run(["rclone", *args], env=_mega_env(),
                           capture_output=quiet, text=True)
        if r.returncode != 0 and not quiet:
            print(f"rclone exited with code {r.returncode}")
        return r.returncode == 0
    except Exception as e:
        print("rclone failed:", e)
        return False


def upload_to_mega():
    """Uploads images + recipes.xlsx + progress.json. Files already on MEGA are skipped."""
    dest = f"{MEGA_REMOTE}:{MEGA_FOLDER}"
    print(f"Uploading to {dest} ...")
    ok = _rclone(["copy", OUT_DIR, dest, "--stats-one-line", "-v"])
    for f in (XLSX_FILE, STATE_FILE):
        if os.path.exists(f):
            ok = _rclone(["copyto", f, f"{dest}/{f}"]) and ok
    print("Upload finished." if ok else "Upload had errors (will retry at the next batch).")
    return ok


# ---------- progress / resume ----------
def load_state(mega_ok):
    """Restores progress.json from MEGA (after a disconnect) and returns (rows, done_links)."""
    if not os.path.exists(STATE_FILE) and mega_ok:
        print("Looking for earlier progress on MEGA...")
        _rclone(["copyto", f"{MEGA_REMOTE}:{MEGA_FOLDER}/{STATE_FILE}", STATE_FILE], quiet=True)
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, encoding="utf-8") as f:
                state = json.load(f)
            rows = [tuple(r) for r in state.get("rows", [])]
            done = set(state.get("done", []))
            print(f"Resuming: {len(done)} recipes already done.")
            return rows, done
        except Exception as e:
            print("Could not read progress file, starting fresh:", e)
    return [], set()


def save_state(rows, done):
    tmp = STATE_FILE + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"rows": rows, "done": sorted(done)}, f, ensure_ascii=False)
    os.replace(tmp, STATE_FILE)


def write_xlsx(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Recipes"
    ws.append(["image_filename", "caption", "fullname"])
    for c in ws[1]:
        c.font = Font(bold=True)
    for row in rows:
        ws.append(list(row))
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 70
    ws.column_dimensions["C"].width = 35
    wb.save(XLSX_FILE)


# ---------- main ----------
def main():
    os.makedirs(OUT_DIR, exist_ok=True)
    mega_ok = mega_ready()

    rows, done = load_state(mega_ok)
    for r in rows:
        _used_names.add(r[0].lower())

    # 1) collect links
    links = collect_links(MAIN_URL, MAX_LINKS)
    print(f"Collected {len(links)} links")

    # 2) open each link, scrape, download - saving + uploading as we go
    since_upload = 0
    try:
        for i, link in enumerate(links, 1):
            if link in done:
                print(f"[{i}/{len(links)}] already done, skipping")
                continue
            try:
                try:
                    html = get_html_requests(link)
                    data = parse_recipe(html, link)
                    if not data["fullname"]:
                        raise ValueError("no title found in CSS via requests")
                except Exception:
                    html, css_title = get_page_browser(link, TITLE_SELECTOR)
                    data = parse_recipe(html, link, css_title)

                if not data["img_url"]:
                    print(f"[{i}] no image on {link}")
                else:
                    filename = download_image(data["img_url"], data["fullname"], i)
                    rows.append((filename, data["caption"], data["fullname"]))
                    print(f"[{i}/{len(links)}] {data['fullname']} -> {filename}")
                done.add(link)

                # save after EVERY recipe so nothing is lost
                write_xlsx(rows)
                save_state(rows, done)
                since_upload += 1
            except Exception as e:
                print(f"[{i}] failed {link}: {e}")

            if mega_ok and since_upload >= UPLOAD_EVERY:
                upload_to_mega()
                since_upload = 0
            time.sleep(DELAY)
    finally:
        # runs at the end AND if the cell is stopped / interrupted
        write_xlsx(rows)
        save_state(rows, done)
        print(f"Saved {XLSX_FILE} with {len(rows)} rows")
        if mega_ok and (since_upload or rows):
            upload_to_mega()


if __name__ == "__main__":
    main()