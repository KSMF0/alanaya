#!/usr/bin/env python3
"""
Build price comparison Excel for 584 products.
Includes our prices, pre-built competitor search links, and auto-calc formulas.
"""

import xlrd
import json
import urllib.parse
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.hyperlink import Hyperlink

FILE1 = '/root/.claude/uploads/fdd8558c-9708-5264-982e-bcf811630994/5415aa07-034272________________________________________.xls'
CONFIRMED_JSON = '/home/user/alanaya/confirmed_matches.json'
OUTPUT_FILE = '/home/user/alanaya/price_comparison_output.xlsx'

COMPETITORS = [
    {'key': 'dar',    'name': 'دار الاميرات',   'domain': 'daralamirat.com.sa',  'search_base': 'https://daralamirat.com.sa/search?q='},
    {'key': 'mybrand','name': 'ماي براند',       'domain': 'mybrand.com.sa',       'search_base': 'https://mybrand.com.sa/search?q='},
    {'key': 'knooz',  'name': 'كنوز العناية',   'domain': 'knooz-store.com',      'search_base': 'https://knooz-store.com/search?q='},
    {'key': 'makh',   'name': 'مخازن العناية',  'domain': 'makhazenalenaya.sa',   'search_base': 'https://makhazenalenaya.sa/search?q='},
]


def clean_barcode(raw):
    bc = str(raw).strip()
    if bc.endswith('.0'):
        bc = bc[:-2]
    return bc


def shorten_desc(desc):
    """Get key product search terms (first 4-5 meaningful words)."""
    words = desc.strip().split()
    # Take first 4 words, exclude quantity/number words that are too generic
    key_words = []
    skip_patterns = re.compile(r'^\d+$|^مل$|^جم$|^غرام$')
    for w in words[:6]:
        if not skip_patterns.match(w):
            key_words.append(w)
        if len(key_words) >= 4:
            break
    return ' '.join(key_words)


def build_search_url(comp, desc):
    """Build a search URL for this product on this competitor's site."""
    query = urllib.parse.quote(desc)
    return comp['search_base'] + query


def load_our_products():
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


def load_confirmed():
    if os.path.exists(CONFIRMED_JSON):
        with open(CONFIRMED_JSON, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def generate_excel(products, confirmed):
    wb = Workbook()
    ws = wb.active
    ws.title = "مقارنة الأسعار"
    ws.sheet_view.rightToLeft = True

    # Fills
    hdr_fill  = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    our_fill  = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")
    alt_fill  = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    wht_fill  = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    lnk_fill  = PatternFill(start_color="E8F0FE", end_color="E8F0FE", fill_type="solid")
    chk_fill  = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    res_fill  = PatternFill(start_color="FCE5CD", end_color="FCE5CD", fill_type="solid")
    cheap_fill = PatternFill(start_color="D9EAD3", end_color="D9EAD3", fill_type="solid")

    # Fonts
    hdr_font   = Font(name='Arial', bold=True, color="FFFFFF", size=10)
    data_font  = Font(name='Arial', size=9)
    bold_font  = Font(name='Arial', bold=True, size=9)
    link_font  = Font(name='Arial', size=9, color="1155CC", underline='single')
    price_font = Font(name='Arial', bold=True, size=10, color="0B5394")

    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right  = Alignment(horizontal='right',  vertical='center', wrap_text=True)
    left   = Alignment(horizontal='left',   vertical='center', wrap_text=True)

    border = Border(
        left=Side(style='thin', color='D0D0D0'),
        right=Side(style='thin', color='D0D0D0'),
        top=Side(style='thin', color='D0D0D0'),
        bottom=Side(style='thin', color='D0D0D0'),
    )

    # ── Column layout ──
    # Col 1: #
    # Col 2: Barcode
    # Col 3: Product Description
    # Col 4: Our Price (SAR)
    # For each competitor (×4):
    #   Col [5+4n]:   Competitor Price (SAR) — manual entry
    #   Col [6+4n]:   Competitor Product Link
    # Col last-1: Cheapest Competitor Price (formula)
    # Col last:   Result (formula / "Our product is the cheapest")

    # Column indices (1-based)
    COL_NUM   = 1
    COL_BC    = 2
    COL_DESC  = 3
    COL_OURS  = 4
    comp_price_cols = {}  # comp key → column index
    comp_link_cols  = {}

    col = 5
    for comp in COMPETITORS:
        comp_price_cols[comp['key']] = col
        comp_link_cols[comp['key']]  = col + 1
        col += 2

    COL_MIN_PRICE = col
    COL_RESULT    = col + 1
    TOTAL_COLS    = col + 1

    # ── Headers row 1 ──
    ws.row_dimensions[1].height = 45
    header_cells = {
        COL_NUM:   '#',
        COL_BC:    'الباركود',
        COL_DESC:  'وصف المنتج',
        COL_OURS:  'سعرنا\n(ريال)',
    }
    for comp in COMPETITORS:
        header_cells[comp_price_cols[comp['key']]] = f"سعر\n{comp['name']}\n(ريال)"
        header_cells[comp_link_cols[comp['key']]]  = f"رابط\n{comp['name']}"
    header_cells[COL_MIN_PRICE] = 'أقل سعر\nمنافس'
    header_cells[COL_RESULT]    = 'النتيجة / الرابط الأرخص'

    for c, title in header_cells.items():
        cell = ws.cell(row=1, column=c, value=title)
        cell.fill   = hdr_fill
        cell.font   = hdr_font
        cell.alignment = center
        cell.border = border

    # ── Column widths ──
    ws.column_dimensions[get_column_letter(COL_NUM)].width   = 5
    ws.column_dimensions[get_column_letter(COL_BC)].width    = 18
    ws.column_dimensions[get_column_letter(COL_DESC)].width  = 42
    ws.column_dimensions[get_column_letter(COL_OURS)].width  = 12
    for comp in COMPETITORS:
        ws.column_dimensions[get_column_letter(comp_price_cols[comp['key']])].width = 12
        ws.column_dimensions[get_column_letter(comp_link_cols[comp['key']])].width  = 35
    ws.column_dimensions[get_column_letter(COL_MIN_PRICE)].width = 12
    ws.column_dimensions[get_column_letter(COL_RESULT)].width    = 45

    # ── Data rows ──
    for row_idx, product in enumerate(products, start=1):
        r = row_idx + 1
        ws.row_dimensions[r].height = 22
        bc       = product['barcode']
        desc     = product['description']
        our_price = product['our_price']
        base_fill = alt_fill if row_idx % 2 == 0 else wht_fill

        # Confirmed data for this barcode
        c_data = confirmed.get(bc, {})

        # Row #
        cell = ws.cell(row=r, column=COL_NUM, value=row_idx)
        cell.fill = base_fill; cell.font = data_font; cell.border = border; cell.alignment = center

        # Barcode
        cell = ws.cell(row=r, column=COL_BC, value=bc)
        cell.fill = base_fill; cell.font = data_font; cell.border = border; cell.alignment = center

        # Description
        cell = ws.cell(row=r, column=COL_DESC, value=desc)
        cell.fill = base_fill; cell.font = data_font; cell.border = border; cell.alignment = right

        # Our price
        cell = ws.cell(row=r, column=COL_OURS, value=our_price)
        cell.fill = our_fill; cell.font = price_font; cell.border = border; cell.alignment = center
        cell.number_format = '#,##0.00'

        # Competitor columns
        price_col_letters = []
        for comp in COMPETITORS:
            pc = comp_price_cols[comp['key']]
            lc = comp_link_cols[comp['key']]
            pc_letter = get_column_letter(pc)
            price_col_letters.append(pc_letter)

            # Price cell — confirmed data or empty for manual entry
            comp_info = c_data.get(comp['key'], {})
            price_val = comp_info.get('price')  # float or None
            link_val  = comp_info.get('link', '')

            pcell = ws.cell(row=r, column=pc, value=price_val)
            pcell.fill = chk_fill if price_val is None else wht_fill
            pcell.font = data_font
            pcell.border = border
            pcell.alignment = center
            if price_val is not None:
                pcell.number_format = '#,##0.00'

            # Link cell — confirmed link or auto-search URL
            if not link_val:
                search_terms = shorten_desc(desc)
                link_val = build_search_url(comp, search_terms)

            lcell = ws.cell(row=r, column=lc, value='بحث مباشر' if not comp_info.get('link') else 'رابط المنتج')
            lcell.fill = lnk_fill
            lcell.font = link_font
            lcell.border = border
            lcell.alignment = center
            lcell.hyperlink = link_val

        # Minimum competitor price formula
        price_refs = ','.join([f'{l}{r}' for l in price_col_letters])
        min_formula = f'=IFERROR(MIN({price_refs}),"")'
        min_cell = ws.cell(row=r, column=COL_MIN_PRICE, value=min_formula)
        min_cell.fill = res_fill; min_cell.font = bold_font; min_cell.border = border
        min_cell.alignment = center; min_cell.number_format = '#,##0.00'

        # Result formula: cheapest comparison
        min_col_letter = get_column_letter(COL_MIN_PRICE)
        our_col_letter = get_column_letter(COL_OURS)
        result_formula = (
            f'=IFERROR(IF({min_col_letter}{r}="","لا توجد أسعار منافسة - يرجى البحث يدوياً",'
            f'IF({our_col_letter}{r}<={min_col_letter}{r},'
            f'"منتجنا هو الأرخص",'
            f'CONCATENATE("أرخص منافس: ",'
            f'IF(MIN({",".join([f"{l}{r}" for l in price_col_letters])})={price_col_letters[0]}{r},"دار الاميرات",'
            f'IF(MIN({",".join([f"{l}{r}" for l in price_col_letters])})={price_col_letters[1]}{r},"ماي براند",'
            f'IF(MIN({",".join([f"{l}{r}" for l in price_col_letters])})={price_col_letters[2]}{r},"كنوز العناية",'
            f'"مخازن العناية"))),'
            f'" - ",'
            f'TEXT({min_col_letter}{r},"#,##0.00"),'
            f'" ريال"))),"")'
        )

        rcell = ws.cell(row=r, column=COL_RESULT, value=result_formula)
        rcell.fill = base_fill; rcell.font = data_font; rcell.border = border
        rcell.alignment = right

    # ── Summary row ──
    sr = len(products) + 2
    ws.cell(row=sr, column=COL_NUM,    value='إجمالي').font = bold_font
    ws.cell(row=sr, column=COL_BC,     value=f'{len(products)} منتج').font = data_font
    ws.cell(row=sr, column=COL_DESC,   value='الخلايا الصفراء = يرجى إدخال السعر يدوياً بعد زيارة الرابط').font = data_font
    ws.cell(row=sr, column=COL_OURS,   value=f'=SUM({get_column_letter(COL_OURS)}2:{get_column_letter(COL_OURS)}{sr-1})').font = bold_font

    # ── Instructions sheet ──
    ws2 = wb.create_sheet(title="تعليمات")
    ws2.sheet_view.rightToLeft = True
    instructions = [
        ("كيفية استخدام هذا الملف", True),
        ("", False),
        ("1. في عمود 'سعر دار الاميرات' (وبقية المنافسين): أدخل سعر المنتج بعد زيارة الرابط المحدد في عمود الرابط.", False),
        ("2. عند النقر على زر 'بحث مباشر' في عمود الرابط، سيتم فتح صفحة بحث مباشرة في موقع المنافس.", False),
        ("3. بعد إدخال الأسعار، سيتم حساب عمود 'النتيجة' تلقائياً.", False),
        ("4. عمود 'النتيجة' يوضح ما إذا كان منتجنا هو الأرخص أم لا، مع ذكر المنافس الأرخص.", False),
        ("", False),
        ("ملاحظة: تعذّر الوصول التلقائي لأسعار المنافسين بسبب سياسة أمنية تمنع الروابط الآلية.", False),
        ("جميع روابط البحث مُهيَّأة مسبقاً لكل منتج ومنافس لتسريع عملية التحقق اليدوي.", False),
    ]
    for row_i, (text, bold) in enumerate(instructions, 1):
        cell = ws2.cell(row=row_i, column=1, value=text)
        if bold:
            cell.font = Font(name='Arial', bold=True, size=12)
        else:
            cell.font = Font(name='Arial', size=10)
    ws2.column_dimensions['A'].width = 90

    # ── Freeze panes ──
    ws.freeze_panes = 'E2'

    wb.save(OUTPUT_FILE)
    print(f"✓ Saved: {OUTPUT_FILE}")
    print(f"  Products: {len(products)}")
    print(f"  Columns: {TOTAL_COLS} per row")
    print(f"  Competitors: {[c['name'] for c in COMPETITORS]}")


if __name__ == '__main__':
    print("Loading products...")
    products = load_our_products()
    print(f"Loaded {len(products)} products")

    print("Loading confirmed competitor matches...")
    confirmed = load_confirmed()
    print(f"Confirmed matches: {len(confirmed)} barcodes")

    print("Building Excel...")
    generate_excel(products, confirmed)
