#!/usr/bin/env python3
"""
=============================================================
COMPETITOR PRICE SCRAPER  —  Run this on YOUR LOCAL machine
=============================================================

SETUP (one time):
    pip install requests beautifulsoup4 xlrd openpyxl

USAGE:
    1. Copy your XLS file path into FILE1 below
    2. Run:  python scraper_local.py
    3. When done, run: python build_excel.py   (to rebuild the Excel)

The script:
- Searches all 4 competitor sites for each of the 584 products
- Saves results to confirmed_matches.json
- Resumes from where it left off if you stop/restart it
=============================================================
"""

import xlrd
import json
import time
import re
import os
import random
import urllib.parse

import requests
from bs4 import BeautifulSoup

# ──────────────────────────────────────────────────────────
# UPDATE THIS PATH to wherever the .xls file is on your PC
# ──────────────────────────────────────────────────────────
FILE1 = r'5415aa07-034272.xls'          # <-- change this

OUTPUT_JSON   = 'confirmed_matches.json'
DELAY_MIN     = 2.0    # seconds between requests (same site)
DELAY_MAX     = 4.0
REQUEST_TIMEOUT = 20   # seconds

COMPETITORS = [
    {'key': 'dar',     'name': 'دار الاميرات',  'domain': 'daralamirat.com.sa', 'platform': 'salla'},
    {'key': 'mybrand', 'name': 'ماي براند',      'domain': 'mybrand.com.sa',     'platform': 'salla'},
    {'key': 'knooz',   'name': 'كنوز العناية',  'domain': 'knooz-store.com',    'platform': 'shopify'},
    {'key': 'makh',    'name': 'مخازن العناية', 'domain': 'makhazenalenaya.sa', 'platform': 'auto'},
]

HEADERS = {
    'User-Agent': (
        'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
        'AppleWebKit/537.36 (KHTML, like Gecko) '
        'Chrome/125.0.0.0 Safari/537.36'
    ),
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
    'Accept-Language': 'ar-SA,ar;q=0.9,en;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Connection': 'keep-alive',
}

JSON_HEADERS = {
    **HEADERS,
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'X-Requested-With': 'XMLHttpRequest',
}


# ── helpers ────────────────────────────────────────────────

def clean_barcode(raw):
    bc = str(raw).strip()
    return bc[:-2] if bc.endswith('.0') else bc


def get_search_query(desc):
    """Return 3-4 key Arabic words from product description."""
    words = desc.strip().split()
    skip = re.compile(r'^\d+(?:[.,]\d+)?$|^(?:مل|جم|غرام|ml|g|gr)$', re.I)
    key_words = [w for w in words[:8] if not skip.match(w)][:4]
    return ' '.join(key_words) if key_words else desc[:30]


def extract_price(text):
    """Pull a SAR price out of a string."""
    text = text.replace('\xa0', ' ').replace(',', '').strip()
    patterns = [
        r'(\d{1,5}(?:\.\d{1,2})?)\s*(?:ريال|SAR|SR|﷼)',
        r'(?:ريال|SAR|SR|﷼)\s*(\d{1,5}(?:\.\d{1,2})?)',
        r'(?:السعر|Price)[:\s]+(\d{1,5}(?:\.\d{1,2})?)',
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            try:
                val = float(m.group(1))
                if 1 <= val <= 9999:
                    return val
            except ValueError:
                pass
    return None


# ── Shopify scraper ────────────────────────────────────────

def scrape_shopify(domain, query, session):
    encoded = urllib.parse.quote(query)

    # 1. Storefront suggest API (works on most Shopify stores)
    api = f"https://{domain}/search/suggest.json?q={encoded}&resources[type]=product&resources[limit]=3"
    try:
        r = session.get(api, headers=JSON_HEADERS, timeout=REQUEST_TIMEOUT)
        if r.status_code == 200:
            data = r.json()
            products = data.get('resources', {}).get('results', {}).get('products', [])
            if products:
                p = products[0]
                raw_price = p.get('price', '0')
                # Shopify returns price in minor currency units (e.g. 1995 = 19.95)
                try:
                    price_val = float(str(raw_price).replace(',', ''))
                    price_val = price_val / 100 if price_val > 500 else price_val
                except Exception:
                    price_val = None
                handle = p.get('handle', '')
                link = f"https://{domain}/products/{handle}" if handle else ''
                return {'price': price_val, 'link': link}
    except Exception:
        pass

    # 2. HTML search fallback
    search_url = f"https://{domain}/search?q={encoded}&type=product"
    try:
        r = session.get(search_url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            price = None
            for sel in ('[class*="price"]', '.price', '.money', 'span[data-price]'):
                el = soup.select_one(sel)
                if el:
                    price = extract_price(el.get_text())
                    if price:
                        break
            link_el = soup.select_one('a[href*="/products/"]')
            link = f"https://{domain}{link_el['href']}" if link_el else ''
            return {'price': price, 'link': link}
    except Exception:
        pass

    return None


# ── Salla scraper ──────────────────────────────────────────

def scrape_salla(domain, query, session):
    encoded = urllib.parse.quote(query)

    # 1. Try Salla JSON search endpoints
    for endpoint in [
        f"https://{domain}/api/search?q={encoded}",
        f"https://{domain}/megamenu/search?q={encoded}",
        f"https://{domain}/search?q={encoded}&format=json",
    ]:
        try:
            r = session.get(endpoint, headers=JSON_HEADERS, timeout=REQUEST_TIMEOUT)
            if r.status_code == 200 and 'application/json' in r.headers.get('Content-Type', ''):
                data = r.json()
                products = (
                    data.get('data', {}).get('products', [])
                    or data.get('products', {}).get('data', [])
                    or data.get('products', [])
                )
                if products:
                    p = products[0]
                    price_data = p.get('price', {})
                    if isinstance(price_data, dict):
                        price_val = float(price_data.get('amount', 0) or 0)
                    else:
                        try:
                            price_val = float(price_data or 0)
                        except Exception:
                            price_val = 0.0
                    link = p.get('url', '') or p.get('link', '')
                    if link and not link.startswith('http'):
                        link = f"https://{domain}{link}"
                    return {
                        'price': price_val if price_val > 0 else None,
                        'link': link
                    }
        except Exception:
            pass

    # 2. HTML search fallback
    search_url = f"https://{domain}/search?q={encoded}"
    try:
        r = session.get(search_url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
        if r.status_code == 200:
            soup = BeautifulSoup(r.text, 'html.parser')
            price = None
            for sel in (
                '[class*="price"]', '.s-product-card__price',
                'span[itemprop="price"]', '[data-price-amount]', '.product-price'
            ):
                el = soup.select_one(sel)
                if el:
                    price = extract_price(el.get_text())
                    if price:
                        break
            link_el = soup.select_one('a[href*="/p"], a[href*="/products/"]')
            link = ''
            if link_el:
                href = link_el['href']
                link = href if href.startswith('http') else f"https://{domain}{href}"
            return {'price': price, 'link': link}
    except Exception:
        pass

    return None


# ── Generic scraper ────────────────────────────────────────

def scrape_generic(domain, query, session):
    encoded = urllib.parse.quote(query)
    for url in [
        f"https://{domain}/search?q={encoded}",
        f"https://{domain}/search?search={encoded}",
    ]:
        try:
            r = session.get(url, headers=HEADERS, timeout=REQUEST_TIMEOUT)
            if r.status_code == 200:
                soup = BeautifulSoup(r.text, 'html.parser')
                price = None
                for sel in ('[class*="price"]', '.price', 'span[itemprop="price"]'):
                    el = soup.select_one(sel)
                    if el:
                        price = extract_price(el.get_text())
                        if price:
                            break
                link_el = soup.select_one('a[href*="product"]') or soup.select_one('.product-card a, .product a')
                link = ''
                if link_el:
                    href = link_el.get('href', '')
                    link = href if href.startswith('http') else f"https://{domain}{href}"
                return {'price': price, 'link': link}
        except Exception:
            pass
    return None


# ── main per-competitor dispatch ───────────────────────────

def search_competitor(comp, product, session):
    query = get_search_query(product['description'])
    domain = comp['domain']
    platform = comp['platform']

    if platform == 'shopify':
        result = scrape_shopify(domain, query, session)
    elif platform == 'salla':
        result = scrape_salla(domain, query, session)
    else:
        result = scrape_generic(domain, query, session)

    time.sleep(random.uniform(DELAY_MIN, DELAY_MAX))
    return result or {'price': None, 'link': ''}


# ── product loader ─────────────────────────────────────────

def load_products():
    wb = xlrd.open_workbook(FILE1)
    ws = wb.sheet_by_index(0)
    products = []
    for i in range(1, ws.nrows):
        bc_raw = ws.cell_value(i, 1)
        if not bc_raw:
            continue
        bc = clean_barcode(bc_raw)
        desc = str(ws.cell_value(i, 2)).strip()
        try:
            price = float(ws.cell_value(i, 4)) if ws.cell_value(i, 4) else 0.0
        except (ValueError, TypeError):
            price = 0.0
        if bc and bc != '0' and price > 0:
            products.append({'barcode': bc, 'description': desc, 'our_price': price})
    return products


# ── entry point ────────────────────────────────────────────

def main():
    print("=" * 62)
    print("  COMPETITOR PRICE SCRAPER")
    print("=" * 62)

    # Load products
    print(f"\nLoading products from: {FILE1}")
    try:
        products = load_products()
        print(f"✓  Loaded {len(products)} products")
    except FileNotFoundError:
        print(f"✗  File not found: {FILE1}")
        print("   Update the FILE1 variable at the top of this script.")
        return
    except Exception as e:
        print(f"✗  Error loading products: {e}")
        return

    # Load existing results (allows resume)
    results = {}
    if os.path.exists(OUTPUT_JSON):
        with open(OUTPUT_JSON, 'r', encoding='utf-8') as f:
            results = json.load(f)
        already_done = sum(1 for v in results.values() if v)
        print(f"✓  Resuming — {already_done} barcodes already in {OUTPUT_JSON}")

    session = requests.Session()
    session.headers.update(HEADERS)

    total = len(products)
    newly_done = 0
    prices_found = 0

    print(f"\nSearching {len(COMPETITORS)} competitor sites for {total} products…")
    print("-" * 62)

    for idx, product in enumerate(products):
        bc = product['barcode']

        if bc in results and results[bc]:
            continue  # already processed

        print(f"\n[{idx+1}/{total}]  {product['description'][:55]}")
        print(f"        barcode={bc}  our_price={product['our_price']} SAR")

        results[bc] = {}
        for comp in COMPETITORS:
            result = search_competitor(comp, product, session)
            results[bc][comp['key']] = result

            if result.get('price'):
                prices_found += 1
                print(f"  ✓  {comp['name']}: {result['price']:.2f} SAR  {result.get('link','')[:50]}")
            else:
                print(f"  ✗  {comp['name']}: not found")

        newly_done += 1

        # Save every 10 products so nothing is lost on interrupt
        if newly_done % 10 == 0:
            _save(results, OUTPUT_JSON)
            pct = (idx + 1) / total * 100
            print(f"\n  [Saved — {idx+1}/{total} products ({pct:.0f}%), {prices_found} prices found so far]")

    _save(results, OUTPUT_JSON)

    print("\n" + "=" * 62)
    print("  DONE!")
    print(f"  Products processed : {newly_done}")
    print(f"  Competitor prices  : {prices_found}")
    print(f"  Output             : {OUTPUT_JSON}")
    print("\n  Next step: run  python build_excel.py")
    print("  to generate the price comparison Excel file.")
    print("=" * 62)


def _save(data, path):
    with open(path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


if __name__ == '__main__':
    main()
