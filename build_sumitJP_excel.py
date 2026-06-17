#!/usr/bin/env python3
"""Build an Excel workbook mapping every analysed sumitclub.jp URL to its
category, migration mode and DOM template family.

Sheet 1 (URL Mapping): one row per URL.
Sheet 2 (Category Summary): counts per category.
Sheet 3 (Mode Summary): counts per migration mode.
"""
import json
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

PER_URL = '/workspace/current/.migration/sumitJP/per-url.json'
OUT = '/workspace/current/SumitclubJP_URL_Category_Mapping.xlsx'

rows = json.load(open(PER_URL))

# Family mapping (same grouping as the Word report).
FAMILY = {
    'Standard AEM Content Shell (/ja/)': [
        'Usage / How-to Page', 'Travel Service Page', 'Insurance Product/Info',
        'Point Program Page', 'Legal / Policy Text', 'Contact / Support',
        'Campaign Landing Page', 'Entertainment/Lifestyle Page', 'Gourmet Service Page',
        'Club Online Promo (unique)', 'Section Landing Page', 'Section Index',
        'Commercial Card Page', 'Other Content Page', 'Notice / News Detail',
        'Notice / News Index', 'Sitemap',
    ],
    'AEM Card-Detail / Listing Shell': [
        'Card Product Detail', 'Card Lineup Listing/Index',
    ],
    'AEM Corporate Shell (/ja/corporate_site)': [
        'Corporate Info (AEM)', 'Corporate News List', 'Corporate Special/Anniversary',
    ],
    'Legacy Corporate Microsite (/corporate)': [
        'Corporate Info (legacy shell)', 'Corporate Kaizen Notice (archive)',
        'Corporate Recruit Page',
    ],
    'Legacy English Shell (/en)': [
        'Homepage (locale)', 'Service Page (legacy en)', 'Announcement Page',
        'FAQ Page', 'Info Index', 'HTTP Error Page',
    ],
    'Standalone Microsites (non-AEM)': [
        'Card Application Landing Page (LP)', 'Identity Verification Microsite',
        'Club Online Bumper Page', 'Cancellation/Termination Flow',
    ],
    'Auth-gated / System (excluded)': [
        'Member Dining-Selection Detail (R####)', 'Member Dining-Selection Index/Search',
        'Login/Member System Page', 'Sign-on / Login System', 'System Error Stub',
    ],
    'Non-Page Assets (excluded)': [
        'SSI Include Fragment', 'LP Partial Fragment', 'Bumper/Redirect Stub',
        'Blank/Helper Stub', 'Tracking/System Stub', 'Modal/Popup Fragment',
        'Auth Widget Fragment', 'Redirect / Router Stub', 'App-bridge Stub',
    ],
    'E-statement Email Templates (separate track)': [
        'E-statement Email Template',
    ],
}
cat_to_family = {c: f for f, cs in FAMILY.items() for c in cs}

# Styling helpers
HEADER_FILL = PatternFill('solid', fgColor='1F4E79')
HEADER_FONT = Font(bold=True, color='FFFFFF', size=11)
THIN = Side(style='thin', color='D9D9D9')
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP = Alignment(vertical='center', wrap_text=True)

MODE_FILL = {
    'Automated': 'C6EFCE',
    'Assisted': 'FFEB9C',
    'Manual': 'FFC7CE',
    'Manual / Replace': 'FFC7CE',
    'Manual / Auth-gated': 'F4CCCC',
    'Exclude (not a page)': 'D9D9D9',
    'Exclude / Re-point': 'D9D9D9',
    'Exclude / Separate track': 'D9D9D9',
}


def style_header(ws, ncols):
    for c in range(1, ncols + 1):
        cell = ws.cell(row=1, column=c)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(vertical='center', horizontal='center', wrap_text=True)
        cell.border = BORDER
    ws.freeze_panes = 'A2'
    ws.row_dimensions[1].height = 28


wb = openpyxl.Workbook()

# --- Sheet 1: URL Mapping ---
ws = wb.active
ws.title = 'URL Mapping'
ws.append(['#', 'URL', 'Category', 'Migration Mode', 'Template Family'])
for i, r in enumerate(sorted(rows, key=lambda x: x['url']), start=1):
    fam = cat_to_family.get(r['category'], 'Other')
    ws.append([i, r['url'], r['category'], r['mode'], fam])
    mode_cell = ws.cell(row=i + 1, column=4)
    fill = MODE_FILL.get(r['mode'])
    if fill:
        mode_cell.fill = PatternFill('solid', fgColor=fill)
    for c in range(1, 6):
        ws.cell(row=i + 1, column=c).border = BORDER
        ws.cell(row=i + 1, column=c).alignment = WRAP
style_header(ws, 5)
ws.auto_filter.ref = f'A1:E{ws.max_row}'
widths = [6, 78, 34, 24, 38]
for i, w in enumerate(widths, start=1):
    ws.column_dimensions[get_column_letter(i)].width = w

# --- Sheet 2: Category Summary ---
ws2 = wb.create_sheet('Category Summary')
ws2.append(['Category', 'Migration Mode', 'Template Family', 'URL Count'])
cat_counts = {}
cat_mode = {}
for r in rows:
    cat_counts[r['category']] = cat_counts.get(r['category'], 0) + 1
    cat_mode[r['category']] = r['mode']
for cat, n in sorted(cat_counts.items(), key=lambda x: -x[1]):
    ws2.append([cat, cat_mode[cat], cat_to_family.get(cat, 'Other'), n])
    rr = ws2.max_row
    fill = MODE_FILL.get(cat_mode[cat])
    if fill:
        ws2.cell(row=rr, column=2).fill = PatternFill('solid', fgColor=fill)
    for c in range(1, 5):
        ws2.cell(row=rr, column=c).border = BORDER
        ws2.cell(row=rr, column=c).alignment = WRAP
ws2.append(['TOTAL (distinct URLs)', '', '', sum(cat_counts.values())])
tot_row = ws2.max_row
for c in range(1, 5):
    ws2.cell(row=tot_row, column=c).font = Font(bold=True)
style_header(ws2, 4)
for i, w in enumerate([34, 24, 38, 12], start=1):
    ws2.column_dimensions[get_column_letter(i)].width = w

# --- Sheet 3: Mode Summary ---
ws3 = wb.create_sheet('Mode Summary')
ws3.append(['Migration Mode', 'URL Count', '% of analysed'])
mode_counts = {}
for r in rows:
    mode_counts[r['mode']] = mode_counts.get(r['mode'], 0) + 1
total = sum(mode_counts.values())
for m, n in sorted(mode_counts.items(), key=lambda x: -x[1]):
    ws3.append([m, n, f'{100*n/total:.1f}%'])
    rr = ws3.max_row
    fill = MODE_FILL.get(m)
    if fill:
        ws3.cell(row=rr, column=1).fill = PatternFill('solid', fgColor=fill)
    for c in range(1, 4):
        ws3.cell(row=rr, column=c).border = BORDER
ws3.append(['TOTAL', total, '100%'])
for c in range(1, 4):
    ws3.cell(row=ws3.max_row, column=c).font = Font(bold=True)
style_header(ws3, 3)
for i, w in enumerate([28, 12, 14], start=1):
    ws3.column_dimensions[get_column_letter(i)].width = w

# Note row about R#### set
ws3.append([])
ws3.append(['Note: The supplied list also contains ~1,004 member dining-selection detail '
            'pages (…/search/R####.html), all mapping to the single "Member Dining-Selection '
            'Detail (R####)" category (Manual / Auth-gated). They are represented here by the '
            'index/search/stocklist entries, not as 1,004 separate rows.'])
ws3.merge_cells(start_row=ws3.max_row, start_column=1, end_row=ws3.max_row, end_column=3)
ws3.cell(row=ws3.max_row, column=1).alignment = Alignment(wrap_text=True, vertical='top')
ws3.row_dimensions[ws3.max_row].height = 60

wb.save(OUT)
print('Saved:', OUT)
print('URL Mapping rows:', len(rows), '| categories:', len(cat_counts), '| modes:', len(mode_counts))
