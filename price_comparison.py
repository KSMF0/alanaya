#!/usr/bin/env python3
"""
Price comparison script: reads our products from File 1,
loads competitor data from search results JSON,
generates output Excel file.
"""

import xlrd
import json
import os
import re
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

FILE1 = '/root/.claude/uploads/fdd8558c-9708-5264-982e-bcf811630994/5415aa07-034272________________________________________.xls'
RESULTS_JSON = '/home/user/alanaya/competitor_results.json'
OUTPUT_FILE = '/home/user/alanaya/price_comparison_output.xlsx'

COMPETITORS = [
    {'name': 'دار الاميرات', 'name_en': 'Dar Al Amirates', 'domain': 'daralamirat.com.sa'},
    {'name': 'ماي براند', 'name_en': 'My Brand', 'domain': 'mybrand.com.sa'},
    {'name': 'كنوز العناية', 'name_en': 'Knooz Al Inaya', 'domain': 'knooz-store.com'},
    {'name': 'مخازن العناية', 'name_en': 'Makhazen Alenayah', 'domain': 'makhazenalenaya.sa'},
]


def load_our_products():
    wb = xlrd.open_workbook(FILE1)
    ws = wb.sheet_by_index(0)
    products = []
    for i in range(1, ws.nrows):
        barcode_raw = ws.cell_value(i, 1)
        if not barcode_raw:
            continue
        barcode = str(barcode_raw).strip()
        if barcode.endswith('.0'):
            barcode = barcode[:-2]
        desc = str(ws.cell_value(i, 2)).strip()
        price_raw = ws.cell_value(i, 4)
        try:
            price = float(price_raw) if price_raw else 0.0
        except (ValueError, TypeError):
            price = 0.0
        if not barcode or barcode == '0':
            continue
        products.append({
            'barcode': barcode,
            'description': desc,
            'our_price': price,
        })
    return products


def load_competitor_results():
    if os.path.exists(RESULTS_JSON):
        with open(RESULTS_JSON, 'r', encoding='utf-8') as f:
            return json.load(f)
    return {}


def generate_excel(products, competitor_results):
    wb = Workbook()
    ws = wb.active
    ws.title = "مقارنة الأسعار"
    ws.sheet_view.rightToLeft = True

    # Color definitions
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    our_price_fill = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    cheapest_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    not_found_fill = PatternFill(start_color="F4CCCC", end_color="F4CCCC", fill_type="solid")
    alt_row_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    white_fill = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
    our_cheapest_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")

    header_font = Font(name='Arial', bold=True, color="FFFFFF", size=11)
    data_font = Font(name='Arial', size=10)
    bold_font = Font(name='Arial', bold=True, size=10)
    center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    right_align = Alignment(horizontal='right', vertical='center', wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC'),
    )

    # Build headers
    headers = [
        '#',
        'الباركود',
        'وصف المنتج',
        'سعرنا (ريال)',
    ]
    for comp in COMPETITORS:
        headers.append(f"سعر {comp['name']} (ريال)")
        headers.append(f"رابط {comp['name']}")
    headers.append('أرخص منافس (ريال)')
    headers.append('رابط / الحكم')

    # Write header row
    ws.row_dimensions[1].height = 40
    for col_idx, header in enumerate(headers, start=1):
        cell = ws.cell(row=1, column=col_idx, value=header)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = thin_border

    # Column widths
    col_widths = [5, 18, 45, 14]
    for _ in COMPETITORS:
        col_widths.extend([14, 40])
    col_widths.extend([14, 45])

    for col_idx, width in enumerate(col_widths, start=1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width

    # Write data rows
    NOT_FOUND = 'غير متوفر'
    OUR_CHEAPEST = 'منتجنا هو الأرخص'

    total_cheapest_ours = 0
    total_found = 0

    for row_idx, product in enumerate(products, start=1):
        excel_row = row_idx + 1
        is_alt_row = row_idx % 2 == 0
        base_fill = alt_row_fill if is_alt_row else white_fill
        ws.row_dimensions[excel_row].height = 25

        barcode = product['barcode']
        our_price = product['our_price']

        # Gather competitor data for this barcode
        comp_data = competitor_results.get(barcode, {})

        row_values = [
            row_idx,
            barcode,
            product['description'],
            our_price,
        ]

        competitor_prices = []
        competitor_links = []
        for comp in COMPETITORS:
            domain = comp['domain']
            info = comp_data.get(domain, {})
            price = info.get('price')
            link = info.get('link', '')
            if price is not None and price > 0:
                competitor_prices.append((price, link, comp['name']))
            else:
                competitor_prices.append(None)
            competitor_links.append(link)
            row_values.append(price if price else NOT_FOUND)
            row_values.append(link if link else '')

        # Determine cheapest
        valid_prices = [(p[0], p[1], p[2]) for p in competitor_prices if p is not None]

        if valid_prices:
            min_price, min_link, min_comp_name = min(valid_prices, key=lambda x: x[0])
            if our_price > 0 and our_price <= min_price:
                cheapest_price = our_price
                cheapest_label = OUR_CHEAPEST
            else:
                cheapest_price = min_price
                cheapest_label = min_link if min_link else f"({min_comp_name}) {min_price} ريال"
        else:
            cheapest_price = ''
            cheapest_label = 'لم يُعثر على بيانات المنافسين'

        row_values.append(cheapest_price if cheapest_price else '')
        row_values.append(cheapest_label)

        # Write cells
        for col_idx, value in enumerate(row_values, start=1):
            cell = ws.cell(row=excel_row, column=col_idx, value=value)
            cell.font = data_font
            cell.border = thin_border
            cell.alignment = center_align if col_idx <= 4 else right_align

            # Apply fills
            if col_idx == 4:  # Our price
                cell.fill = our_price_fill
                cell.font = bold_font
            elif cheapest_label == OUR_CHEAPEST and col_idx == len(headers) - 1:
                cell.fill = our_cheapest_fill
            elif cheapest_label == OUR_CHEAPEST and col_idx == len(headers):
                cell.fill = our_cheapest_fill
                cell.font = bold_font
            elif value == NOT_FOUND:
                cell.fill = not_found_fill
            else:
                cell.fill = base_fill

            # Highlight minimum competitor price
            if valid_prices and col_idx >= 5 and col_idx <= len(headers) - 2:
                # Check if this is a price column (odd offset from col 5)
                offset = col_idx - 5
                if offset % 2 == 0:  # price column
                    comp_index = offset // 2
                    cp = competitor_prices[comp_index]
                    if cp is not None and cp[0] == min_price and cheapest_label != OUR_CHEAPEST:
                        cell.fill = cheapest_fill

        if cheapest_label == OUR_CHEAPEST:
            total_cheapest_ours += 1
        if valid_prices:
            total_found += 1

    # Summary row
    summary_row = len(products) + 2
    ws.cell(row=summary_row, column=1, value='الإجمالي').font = bold_font
    ws.cell(row=summary_row, column=2, value=f'المنتجات: {len(products)}').font = data_font
    ws.cell(row=summary_row, column=3, value=f'منتجاتنا الأرخص: {total_cheapest_ours}').font = data_font
    ws.cell(row=summary_row, column=4, value=f'منتجات وُجدت عند المنافسين: {total_found}').font = data_font

    # Freeze header row
    ws.freeze_panes = 'A2'

    wb.save(OUTPUT_FILE)
    print(f"Output saved to: {OUTPUT_FILE}")
    print(f"Total products: {len(products)}")
    print(f"Products found at competitors: {total_found}")
    print(f"Our products cheapest: {total_cheapest_ours}")


if __name__ == '__main__':
    print("Loading our products from File 1...")
    products = load_our_products()
    print(f"Loaded {len(products)} products")

    print("Loading competitor results...")
    competitor_results = load_competitor_results()
    print(f"Competitor data loaded for {len(competitor_results)} barcodes")

    print("Generating Excel output...")
    generate_excel(products, competitor_results)
