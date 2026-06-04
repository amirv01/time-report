"""
Build script: inlines every dependency into a single self-contained HTML file
suitable for offline use. Run from project root:
    python scripts/build.py            -> produces dist/timereport-v1.0.html
    python scripts/build.py 1.1        -> produces dist/timereport-v1.1.html

On first run, downloads CDN libraries and Google Fonts (incl. .woff2 files)
into .build-cache/. Subsequent builds work offline and are fast.
"""

import os
import re
import sys
import hashlib
import base64
import datetime
import urllib.request
import urllib.error

DEFAULT_VERSION = "1.0"
VERSION = sys.argv[1] if len(sys.argv) > 1 else DEFAULT_VERSION

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
OUT_DIR = os.path.join(ROOT, "dist")
CACHE_DIR = os.path.join(ROOT, ".build-cache")
OUT_FILE = os.path.join(OUT_DIR, f"timereport-v{VERSION}.html")

# CDN scripts in the order they appear in index.html (preserves load order semantics)
CDN_SCRIPTS = [
    "https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js",
    "https://cdn.jsdelivr.net/npm/chart.js@4.4.7/dist/chart.umd.min.js",
    "https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js",
    "https://cdn.jsdelivr.net/npm/jspdf@2.5.2/dist/jspdf.umd.min.js",
    "https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js",
]

FONT_CSS_URL = "https://fonts.googleapis.com/css2?family=Assistant:wght@300;400;600;700&display=swap"
# Google Fonts returns .ttf for older UAs and .woff2 for modern ones — we want woff2.
MODERN_UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
             "(KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")

EXT_TO_MIME = {"woff2": "font/woff2", "woff": "font/woff",
               "ttf": "font/ttf", "otf": "font/otf"}


def cache_path_for(url):
    last_seg = url.split("/")[-1].split("?")[0] or "file"
    h = hashlib.sha1(url.encode("utf-8")).hexdigest()[:8]
    return os.path.join(CACHE_DIR, f"{h}-{last_seg}")


def download(url, headers=None):
    # Some CDNs (sheetjs.com in particular) 403 the default Python User-Agent,
    # so always send a modern browser UA unless the caller overrides it.
    merged = {"User-Agent": MODERN_UA}
    if headers:
        merged.update(headers)
    req = urllib.request.Request(url, headers=merged)
    with urllib.request.urlopen(req, timeout=60) as resp:
        return resp.read()


def fetch_cached(url, headers=None):
    p = cache_path_for(url)
    if os.path.exists(p):
        with open(p, "rb") as f:
            return f.read()
    print(f"  fetching: {url}")
    data = download(url, headers)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "wb") as f:
        f.write(data)
    return data


def replace_script_tag(html, url, replacement):
    """
    Replace <script src="URL" ...></script> with `replacement`.
    Attributes may span multiple lines (integrity/crossorigin), so we locate the
    URL first, then scan outward to the surrounding <script and </script>.
    """
    url_idx = html.find(url)
    if url_idx == -1:
        raise RuntimeError(f'<script src="{url}"> not found in HTML')
    start_idx = html.rfind("<script", 0, url_idx)
    end_marker = "</script>"
    end_idx = html.find(end_marker, url_idx)
    if start_idx == -1 or end_idx == -1:
        raise RuntimeError(f"Could not find script bounds for {url}")
    return html[:start_idx] + replacement + html[end_idx + len(end_marker):]


def inline_fonts():
    css = fetch_cached(FONT_CSS_URL, {"User-Agent": MODERN_UA}).decode("utf-8")
    font_urls = sorted(set(re.findall(r"url\((https?://[^)]+)\)", css)))
    for font_url in font_urls:
        font_data = fetch_cached(font_url)
        m = re.search(r"\.(woff2|woff|ttf|otf)(\?|$)", font_url, re.IGNORECASE)
        ext = (m.group(1) if m else "woff2").lower()
        mime = EXT_TO_MIME.get(ext, "font/woff2")
        b64 = base64.b64encode(font_data).decode("ascii")
        css = css.replace(font_url, f"data:{mime};base64,{b64}")
    return css


def main():
    print(f"Building timereport v{VERSION}\n")

    os.makedirs(OUT_DIR, exist_ok=True)
    os.makedirs(CACHE_DIR, exist_ok=True)

    print("Fetching CDN libraries...")
    libs = {url: fetch_cached(url).decode("utf-8") for url in CDN_SCRIPTS}

    print("Fetching fonts (CSS + .woff2 files)...")
    fonts_css = inline_fonts()

    print("Reading project files...")
    with open(os.path.join(ROOT, "index.html"), encoding="utf-8") as f:
        html = f.read()
    with open(os.path.join(ROOT, "style.css"), encoding="utf-8") as f:
        app_css = f.read()
    with open(os.path.join(ROOT, "app.js"), encoding="utf-8") as f:
        app_js = f.read()

    print("Inlining...")

    # 1. Remove the <link rel="preconnect"> for fonts (irrelevant when fonts are inline)
    html = re.sub(
        r'\s*<link rel="preconnect" href="https://fonts\.googleapis\.com">\s*',
        "\n    ", html, count=1
    )

    # 2. Replace Google Fonts <link> with inline <style>
    html = re.sub(
        r'<link href="https://fonts\.googleapis\.com/[^"]+" rel="stylesheet">',
        f"<style>\n{fonts_css}\n</style>",
        html, count=1
    )

    # 3. Replace local style.css link with inline <style>
    html = html.replace(
        '<link rel="stylesheet" href="style.css">',
        f"<style>\n{app_css}\n</style>"
    )

    # 4. Inline each CDN <script>
    for url in CDN_SCRIPTS:
        html = replace_script_tag(html, url, f"<script>\n{libs[url]}\n</script>")

    # 5. Inline app.js
    html = html.replace(
        '<script src="app.js"></script>',
        f"<script>\n{app_js}\n</script>"
    )

    # 6. Build banner
    built = datetime.datetime.now(datetime.UTC).strftime("%Y-%m-%dT%H:%M:%SZ")
    html = html.replace(
        "<!DOCTYPE html>",
        f"<!DOCTYPE html>\n<!-- timereport v{VERSION} — built {built} -->",
        1
    )

    with open(OUT_FILE, "w", encoding="utf-8") as f:
        f.write(html)

    size_mb = os.path.getsize(OUT_FILE) / (1024 * 1024)
    rel = os.path.relpath(OUT_FILE, ROOT)
    print(f"\n[ok] {rel}  ({size_mb:.2f} MB)")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nBuild failed: {e}", file=sys.stderr)
        sys.exit(1)
