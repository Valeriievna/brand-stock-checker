#!/usr/bin/env python3
"""
Brand Stock Scraper
Discovers all products of a brand on Epicenter and Eva.ua,
exports results to a multi-sheet Excel workbook.
"""

import requests
import subprocess
import time
import json
import re
import sys
from datetime import datetime
from pathlib import Path

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("ERROR: Missing package. Please run START_BRANDS.bat to install dependencies.")
    input("Press Enter to close...")
    sys.exit(1)

try:
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
except ImportError:
    print("ERROR: Missing package. Please run START_BRANDS.bat to install dependencies.")
    input("Press Enter to close...")
    sys.exit(1)

# ─────────────────────────────────────────────────────────
#  SETTINGS
# ─────────────────────────────────────────────────────────
DEFAULT_BRAND = "Paclan"
DELAY_SECONDS = 1.5
MAX_PAGES     = 30
# ─────────────────────────────────────────────────────────

REQUEST_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/122.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "uk-UA,uk;q=0.9,ru;q=0.8,en;q=0.6",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
}

XLSX_COLS = [
    ("#",               5),
    ("Product Name",   55),
    ("SKU / Article",  18),
    ("Regular Price",  14),
    ("On Discount",    12),
    ("Discount Price", 14),
    ("In Stock",       11),
    ("Product URL",    65),
]

C_HEADER    = "2E75B6"
C_IN_STOCK  = "E2EFDA"
C_OUT_STOCK = "FCE4D6"
C_UNKNOWN   = "FFF2CC"
C_TITLE     = "1F4E79"
C_DISCOUNT  = "FCE4C4"  # orange tint for discounted rows


# ─────────────────────────────────────────────────────────
#  HELPERS
# ─────────────────────────────────────────────────────────

def fetch(url, session):
    try:
        r = session.get(url, headers=REQUEST_HEADERS, timeout=25, allow_redirects=True)
        r.raise_for_status()
        r.encoding = 'utf-8'
        return BeautifulSoup(r.text, "html.parser"), r.text
    except requests.RequestException as e:
        print(f"  WARN: {e}")
        return None, None


def node_available():
    try:
        result = subprocess.run(["node", "--version"], capture_output=True, timeout=5)
        return result.returncode == 0
    except (FileNotFoundError, subprocess.TimeoutExpired, OSError):
        return False


def decode_nuxt(script_text):
    """Run the __NUXT__ script in Node.js and return the parsed state dict."""
    js = "const window={};\n" + script_text + "\nprocess.stdout.write(JSON.stringify(window.__NUXT__||null));"
    try:
        result = subprocess.run(
            ["node", "-e", js],
            capture_output=True, text=True, timeout=30
        )
        if result.returncode == 0 and result.stdout:
            return json.loads(result.stdout)
    except Exception:
        pass
    return None


# ─────────────────────────────────────────────────────────
#  EPICENTER SCRAPER
# ─────────────────────────────────────────────────────────

def scrape_epicenter(brand, session, has_node, log_fn=print):
    all_products = []

    if not has_node:
        log_fn("Epicenter: SKIPPED — Node.js is not installed. Install from https://nodejs.org/")
        return []

    brand_enc = requests.utils.quote(brand)
    base_url  = f"https://epicentrk.ua/ua/search/?q={brand_enc}&per-page=60"
    log_fn(f"Epicenter: searching for '{brand}'...")
    total_pages = None

    for page_num in range(1, MAX_PAGES + 1):
        url = f"{base_url}&page={page_num}" if page_num > 1 else base_url

        soup, raw = fetch(url, session)
        if not soup:
            log_fn(f"  Page {page_num}: failed to load")
            break

        nuxt_data = None
        for script in soup.find_all("script"):
            txt = script.string or ""
            if "window.__NUXT__" in txt:
                nuxt_data = decode_nuxt(txt)
                break

        if not nuxt_data:
            log_fn(f"  Page {page_num}: could not decode page data")
            break

        if total_pages is None:
            try:
                total_pages = nuxt_data["data"][0]["params"]["pagination"]["pages"]
            except (KeyError, IndexError, TypeError):
                total_pages = 1

        try:
            raw_products = nuxt_data["state"]["products"]["products"]
        except (KeyError, TypeError):
            log_fn(f"  Page {page_num}: no products in state")
            break

        if not raw_products:
            log_fn(f"  Page {page_num}: no products found")
            break

        page_products = []
        for p in raw_products:
            name      = p.get("name_ua") or p.get("name_ru") or ""
            sku       = str(p.get("id") or "")
            url_p     = p.get("url") or ""

            price_now = p.get("price") or 0
            price_old = p.get("price_old") or 0
            on_discount    = bool(price_old and price_old > price_now)
            regular_price  = str(price_old if on_discount else price_now)
            discount_price = str(price_now) if on_discount else ""

            avail = p.get("avail")
            avail_str = str(p.get("available") or "")
            if avail is not None:
                in_stock = avail > 0
            elif "наявності" in avail_str:
                in_stock = "Немає" not in avail_str
            else:
                in_stock = None

            if name:
                page_products.append({
                    "name": name, "sku": sku,
                    "price": regular_price, "on_discount": on_discount,
                    "discount_price": discount_price,
                    "url": url_p, "in_stock": in_stock,
                })

        log_fn(f"  Page {page_num}: {len(page_products)} products")
        all_products.extend(page_products)

        if page_num >= (total_pages or 1):
            break

        time.sleep(DELAY_SECONDS)

    return all_products


# ─────────────────────────────────────────────────────────
#  EVA SCRAPER
# ─────────────────────────────────────────────────────────

def find_eva_brand_id(brand_name, session):
    """Look up Eva brand ID by name. Returns (id, title) or (None, None)."""
    try:
        r = session.get(
            "https://api.eva.ua/v1/ua/api/brands",
            headers={**REQUEST_HEADERS, "Accept": "application/json"},
            timeout=15,
        )
        data = r.json()
        for group_brands in data.get("data", {}).get("groups", {}).values():
            for b in group_brands:
                if b.get("title", "").lower() == brand_name.lower():
                    m = re.search(r"brnd-(\d+)", b.get("url", ""))
                    if m:
                        return m.group(1), b["title"]
    except Exception as e:
        print(f"  WARN: could not fetch Eva brand list: {e}")
    return None, None


def parse_eva_initial_state(html_text):
    """Extract and parse window.__INITIAL_STATE__ from Eva page HTML."""
    for script in BeautifulSoup(html_text, "html.parser").find_all("script"):
        txt = script.string or ""
        if "__INITIAL_STATE__" in txt:
            try:
                idx = txt.index("{")
                decoder = json.JSONDecoder()
                data, _ = decoder.raw_decode(txt, idx)
                return data
            except (ValueError, json.JSONDecodeError):
                pass
    return None


def scrape_eva(brand, session, log_fn=print):
    all_products = []

    log_fn(f"Eva: searching for '{brand}'...")
    log_fn("  Looking up brand in Eva catalogue...")
    brand_id, found_title = find_eva_brand_id(brand, session)
    if not brand_id:
        log_fn(f"  Brand '{brand}' was not found in Eva's brand list.")
        return []
    log_fn(f"  Found: '{found_title}' (ID {brand_id})")
    time.sleep(DELAY_SECONDS)

    base_url    = f"https://eva.ua/ua/brnd-{brand_id}/"
    total_pages = None

    for page_num in range(1, MAX_PAGES + 1):
        url = f"{base_url}?p={page_num}" if page_num > 1 else base_url

        soup, raw = fetch(url, session)
        if not soup:
            log_fn(f"  Page {page_num}: failed to load")
            break

        state = parse_eva_initial_state(raw)
        if not state:
            log_fn(f"  Page {page_num}: could not parse page state")
            break

        brand_state = state.get("brand", {})

        if total_pages is None:
            total_pages = brand_state.get("totalPages", 1)

        raw_products = brand_state.get("products", [])
        if not raw_products:
            log_fn(f"  Page {page_num}: no products")
            break

        page_products = []
        for p in raw_products:
            name    = p.get("name") or ""
            sku     = str(p.get("sku") or "")
            url_key = p.get("url_key") or ""
            url_p   = f"https://eva.ua/ua/{url_key}/" if url_key else ""
            stock   = p.get("stock") or {}
            in_stock = stock.get("is_in_stock")

            # Eva: price = discounted price, original_price = full price before discount
            original = p.get("original_price") or 0
            final    = p.get("final_price") or p.get("price") or 0
            on_discount    = bool(original and original > final)
            regular_price  = str(original if on_discount else final)
            discount_price = str(final) if on_discount else ""

            if name:
                page_products.append({
                    "name": name, "sku": sku,
                    "price": regular_price, "on_discount": on_discount,
                    "discount_price": discount_price,
                    "url": url_p, "in_stock": in_stock,
                })

        log_fn(f"  Page {page_num}: {len(page_products)} products")
        all_products.extend(page_products)

        if page_num >= (total_pages or 1):
            break

        time.sleep(DELAY_SECONDS)

    return all_products


# ─────────────────────────────────────────────────────────
#  EXCEL WRITER
# ─────────────────────────────────────────────────────────

def _fmt_price(raw):
    digits = re.sub(r"[^\d.]", "", str(raw).replace(",", "."))
    try:
        return "{:,.2f}".format(float(digits)).replace(",", " ")
    except ValueError:
        return str(raw)


def write_store_sheet(ws, products, store_name, brand, checked_at):
    hdr_font = Font(bold=True, color="FFFFFF", size=11)
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    in_fill  = PatternFill("solid", fgColor=C_IN_STOCK)
    out_fill = PatternFill("solid", fgColor=C_OUT_STOCK)
    unk_fill = PatternFill("solid", fgColor=C_UNKNOWN)
    center   = Alignment(horizontal="center", vertical="center")
    left     = Alignment(horizontal="left",   vertical="center")
    wrap     = Alignment(horizontal="left",   vertical="center", wrap_text=True)

    last_col = get_column_letter(len(XLSX_COLS))

    ws.merge_cells(f"A1:{last_col}1")
    t = ws["A1"]
    t.value     = f"{store_name}  ·  Brand: {brand}  ·  Checked: {checked_at}"
    t.font      = Font(bold=True, size=13, color=C_TITLE)
    t.alignment = center
    ws.row_dimensions[1].height = 26

    for ci, (col_name, col_width) in enumerate(XLSX_COLS, 1):
        c = ws.cell(row=2, column=ci, value=col_name)
        c.font      = hdr_font
        c.fill      = hdr_fill
        c.alignment = center
        ws.column_dimensions[get_column_letter(ci)].width = col_width
    ws.row_dimensions[2].height = 20
    ws.freeze_panes = "A3"

    disc_fill = PatternFill("solid", fgColor=C_DISCOUNT)
    aligns = [center, wrap, center, center, center, center, center, left]
    for i, p in enumerate(products, 1):
        row = i + 2
        if p["in_stock"] is True:
            status, stock_fill = "Yes", in_fill
        elif p["in_stock"] is False:
            status, stock_fill = "No", out_fill
        else:
            status, stock_fill = "?", unk_fill

        on_disc = p.get("on_discount", False)
        vals = [
            i,
            p.get("name", ""),
            p.get("sku", ""),
            _fmt_price(p.get("price", "")),
            "Yes" if on_disc else "No",
            _fmt_price(p.get("discount_price", "")) if on_disc else "",
            status,
            p.get("url", ""),
        ]
        for ci, (val, aln) in enumerate(zip(vals, aligns), 1):
            c = ws.cell(row=row, column=ci, value=val)
            c.alignment = aln
            # Colour: discount columns get orange tint; stock column gets green/red/yellow
            if ci in (5, 6) and on_disc:
                c.fill = disc_fill
            elif ci == 7:
                c.fill = stock_fill

    fr    = len(products) + 4
    in_n  = sum(1 for p in products if p["in_stock"] is True)
    out_n = sum(1 for p in products if p["in_stock"] is False)
    unk_n = sum(1 for p in products if p["in_stock"] is None)
    disc_n = sum(1 for p in products if p.get("on_discount"))
    bf    = Font(bold=True)
    ws.cell(row=fr, column=1, value="TOTAL").font = bf
    ws.cell(row=fr, column=2, value=f"{len(products)} products").font = bf
    ws.cell(row=fr, column=5,
            value=f"On discount: {disc_n}").font = bf
    ws.cell(row=fr, column=7,
            value=f"In stock: {in_n}  |  Out of stock: {out_n}  |  Unknown: {unk_n}").font = bf


def write_summary_sheet(ws, results, brand, checked_at):
    hdr_font = Font(bold=True, color="FFFFFF")
    hdr_fill = PatternFill("solid", fgColor=C_HEADER)
    in_fill  = PatternFill("solid", fgColor=C_IN_STOCK)
    out_fill = PatternFill("solid", fgColor=C_OUT_STOCK)
    center   = Alignment(horizontal="center", vertical="center")

    ws.merge_cells("A1:E1")
    t = ws["A1"]
    t.value     = f"Summary  ·  Brand: {brand}  ·  {checked_at}"
    t.font      = Font(bold=True, size=13, color=C_TITLE)
    t.alignment = center
    ws.row_dimensions[1].height = 26

    disc_fill = PatternFill("solid", fgColor=C_DISCOUNT)
    headers = ["Store", "Total Products", "In Stock", "Out of Stock", "On Discount", "Unknown"]
    widths  = [20, 16, 12, 14, 13, 12]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        c = ws.cell(row=2, column=ci, value=h)
        c.font = hdr_font
        c.fill = hdr_fill
        c.alignment = center
        ws.column_dimensions[get_column_letter(ci)].width = w

    for ri, (store, products) in enumerate(results.items(), 3):
        in_n   = sum(1 for p in products if p["in_stock"] is True)
        out_n  = sum(1 for p in products if p["in_stock"] is False)
        unk_n  = sum(1 for p in products if p["in_stock"] is None)
        disc_n = sum(1 for p in products if p.get("on_discount"))
        for ci, val in enumerate([store, len(products), in_n, out_n, disc_n, unk_n], 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.alignment = center
            if ci == 3 and in_n > 0:
                c.fill = in_fill
            if ci == 4 and out_n > 0:
                c.fill = out_fill
            if ci == 5 and disc_n > 0:
                c.fill = disc_fill


# ─────────────────────────────────────────────────────────
#  MAIN
# ─────────────────────────────────────────────────────────

def main():
    print("=" * 56)
    print("  BRAND STOCK SCRAPER")
    print("=" * 56)

    has_node = node_available()
    if not has_node:
        print("\n  WARNING: Node.js not found.")
        print("  Epicenter scraping requires Node.js.")
        print("  Install from https://nodejs.org/ then re-run.\n")

    entered = input(f"\nEnter brand name (press Enter for '{DEFAULT_BRAND}'): ").strip()
    brand = entered if entered else DEFAULT_BRAND

    checked_at  = datetime.now().strftime("%d.%m.%Y %H:%M")
    ts          = datetime.now().strftime("%Y%m%d_%H%M")
    output_file = f"{brand.replace(' ', '_')}_{ts}.xlsx"
    output_path = Path(__file__).parent / output_file

    print(f"\n  Brand:  {brand}")
    print(f"  Output: {output_file}")
    print("-" * 56)

    session = requests.Session()
    results = {}

    results["Epicenter"] = scrape_epicenter(brand, session, has_node, log_fn=print)
    results["Eva"]       = scrape_eva(brand, session, log_fn=print)

    for store, products in results.items():
        in_n = sum(1 for p in products if p["in_stock"] is True)
        print(f"  {store}: {len(products)} products, {in_n} in stock")

    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    ws_sum = wb.create_sheet("Summary")
    write_summary_sheet(ws_sum, results, brand, checked_at)

    for store_name, products in results.items():
        ws = wb.create_sheet(store_name)
        if products:
            write_store_sheet(ws, products, store_name, brand, checked_at)
        else:
            ws["A1"] = f'No products found for "{brand}" on {store_name}.'
            ws["A1"].font = Font(bold=True, color="FF0000")
            if store_name == "Epicenter" and not has_node:
                ws["A2"] = "Node.js is required — install from https://nodejs.org/"
            else:
                ws["A2"] = "Check that the brand name matches exactly how it appears on the store."

    wb.save(output_path)

    total    = sum(len(v) for v in results.values())
    total_in = sum(sum(1 for p in v if p["in_stock"] is True) for v in results.values())

    print(f"\n{'=' * 56}")
    print(f"  DONE  —  {total} products total, {total_in} in stock")
    print(f"  File: {output_file}")
    print(f"{'=' * 56}")
    input("\nPress Enter to close...")


if __name__ == "__main__":
    main()
