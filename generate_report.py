#!/usr/bin/env python3
"""Generate comprehensive site analysis Word document for Kyowa Kirin Medical Site."""

from docx import Document
from docx.shared import Inches, Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
import os

doc = Document()

# ============================================================
# STYLES
# ============================================================
style = doc.styles['Normal']
font = style.font
font.name = 'Calibri'
font.size = Pt(10)

for level in range(1, 4):
    heading_style = doc.styles[f'Heading {level}']
    heading_style.font.color.rgb = RGBColor(0xEA, 0x55, 0x04)  # Kyowa Kirin orange
    heading_style.font.name = 'Calibri'

def add_table_with_style(doc, headers, rows, col_widths=None):
    """Add a formatted table."""
    table = doc.add_table(rows=1 + len(rows), cols=len(headers))
    table.style = 'Light Grid Accent 1'
    table.alignment = WD_TABLE_ALIGNMENT.LEFT

    # Header row
    hdr = table.rows[0]
    for i, header in enumerate(headers):
        cell = hdr.cells[i]
        cell.text = header
        for paragraph in cell.paragraphs:
            paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
            for run in paragraph.runs:
                run.font.bold = True
                run.font.size = Pt(9)
                run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        # Orange background
        shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="EA5504"/>')
        cell._tc.get_or_add_tcPr().append(shading)

    # Data rows
    for r_idx, row_data in enumerate(rows):
        row = table.rows[r_idx + 1]
        for c_idx, cell_text in enumerate(row_data):
            cell = row.cells[c_idx]
            cell.text = str(cell_text)
            for paragraph in cell.paragraphs:
                for run in paragraph.runs:
                    run.font.size = Pt(9)

    if col_widths:
        for i, width in enumerate(col_widths):
            for row in table.rows:
                row.cells[i].width = Cm(width)

    return table

def add_screenshot(doc, filename, caption, width=5.5):
    """Add a screenshot with caption if file exists."""
    filepath = f'/workspace/{filename}'
    if os.path.exists(filepath):
        try:
            doc.add_picture(filepath, width=Inches(width))
            last_paragraph = doc.paragraphs[-1]
            last_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap = doc.add_paragraph(f'Figure: {caption}')
            cap.alignment = WD_ALIGN_PARAGRAPH.CENTER
            cap.runs[0].font.size = Pt(8)
            cap.runs[0].font.italic = True
            cap.runs[0].font.color.rgb = RGBColor(0x66, 0x66, 0x66)
            return True
        except Exception as e:
            doc.add_paragraph(f'[Screenshot: {caption} - {filename}]')
            return False
    else:
        doc.add_paragraph(f'[Screenshot not available: {filename}]')
        return False

# ============================================================
# TITLE PAGE
# ============================================================
doc.add_paragraph('')
doc.add_paragraph('')
doc.add_paragraph('')
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('Site Analysis Report')
run.font.size = Pt(28)
run.font.color.rgb = RGBColor(0xEA, 0x55, 0x04)
run.font.bold = True

subtitle = doc.add_paragraph()
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = subtitle.add_run('Kyowa Kirin Medical Site')
run.font.size = Pt(20)
run.font.color.rgb = RGBColor(0x33, 0x33, 0x33)

url_para = doc.add_paragraph()
url_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = url_para.add_run('https://medical.kyowakirin.co.jp/')
run.font.size = Pt(12)
run.font.color.rgb = RGBColor(0x00, 0x56, 0xB3)

doc.add_paragraph('')
purpose = doc.add_paragraph()
purpose.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = purpose.add_run('AEM Edge Delivery Services Migration Assessment')
run.font.size = Pt(14)
run.font.color.rgb = RGBColor(0x66, 0x66, 0x66)

doc.add_paragraph('')
date_para = doc.add_paragraph()
date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = date_para.add_run('Date: March 27, 2026')
run.font.size = Pt(11)

doc.add_paragraph('')
doc.add_paragraph('')

# Confidential notice
conf = doc.add_paragraph()
conf.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = conf.add_run('CONFIDENTIAL')
run.font.size = Pt(10)
run.font.bold = True
run.font.color.rgb = RGBColor(0xCC, 0x00, 0x00)

doc.add_page_break()

# ============================================================
# TABLE OF CONTENTS
# ============================================================
doc.add_heading('Table of Contents', level=1)
toc_items = [
    '1. Executive Summary',
    '2. Templates Inventory',
    '3. Blocks / Components Catalog',
    '4. Page Counts by Template',
    '5. Integrations Analysis',
    '6. Complex Use Cases & Observations',
    '7. Migration Estimates',
    '8. Appendix: Screenshots'
]
for item in toc_items:
    p = doc.add_paragraph(item)
    p.paragraph_format.space_after = Pt(4)

doc.add_page_break()

# ============================================================
# 1. EXECUTIVE SUMMARY
# ============================================================
doc.add_heading('1. Executive Summary', level=1)
doc.add_paragraph(
    'This report provides a comprehensive analysis of the Kyowa Kirin Medical Site '
    '(https://medical.kyowakirin.co.jp/), a healthcare professional (HCP) portal '
    'operated by Kyowa Kirin Co., Ltd. The site serves as the primary digital channel '
    'for medical professionals in Japan to access drug information, clinical resources, '
    'web seminars, medical literature, and practice support tools.'
)
doc.add_paragraph(
    'The site is built on a custom technology stack using Vue.js 3 (development build), '
    'jQuery, and a custom backend API layer. It features role-based access gating '
    '(requiring HCP verification), member authentication with SSO (medPass, DLink), '
    'and personalized content recommendations via Marketo RTP.'
)

doc.add_heading('Key Findings', level=2)
findings = [
    '11 distinct page templates identified, ranging from Simple to Complex',
    '20+ reusable blocks/components cataloged across the site',
    'Estimated 350-450+ total pages (including deep content articles, PDFs, and product pages)',
    '15+ third-party integrations including analytics, marketing automation, chatbot, and SSO',
    'Significant complexity in authentication gating, recommendation engine, and interactive tools',
    'Vue.js-based dynamic rendering creates challenges for static content migration',
    'Estimated total migration effort: 55-80 person-days'
]
for f in findings:
    doc.add_paragraph(f, style='List Bullet')

doc.add_paragraph('')
add_screenshot(doc, 'homepage-clean.png', 'Kyowa Kirin Medical Site - Homepage', 5.0)

doc.add_page_break()

# ============================================================
# 2. TEMPLATES INVENTORY
# ============================================================
doc.add_heading('2. Templates Inventory', level=1)
doc.add_paragraph(
    'The following table lists all unique page templates identified across the site. '
    'Each template represents a distinct layout pattern with specific component composition.'
)

template_data = [
    ['T01', 'Homepage', 'Complex',
     'Rich landing page with hero carousel (6 slides), product search widget, seminar listing with category filters, news feed with tabs, recommendation carousel, contact section, and promotional cards. Uses Vue.js for dynamic content loading.',
     'https://medical.kyowakirin.co.jp/'],
    ['T02', 'Medical Area Landing', 'Complex',
     'Category landing for clinical specialties (Kidney, Hematonco, Allergy, Neuro, Rare Disease). Features expandable category accordion, seminar CTA, content carousel, product logo grid, and optional feature cards. Single-column layout.',
     'https://medical.kyowakirin.co.jp/kidney/\nhttps://medical.kyowakirin.co.jp/allergy/\nhttps://medical.kyowakirin.co.jp/hematonco/\nhttps://medical.kyowakirin.co.jp/neuro/\nhttps://medical.kyowakirin.co.jp/raredisease/'],
    ['T03', 'Product List / Catalog', 'Complex',
     'Data-driven product listing with Japanese syllabary (Aiueo) filter tabs, search functionality, and product data table with document availability indicators. Vue.js handles dynamic filtering via /api/v1.0/product.',
     'https://medical.kyowakirin.co.jp/druginfo/detail/'],
    ['T04', 'Product Detail', 'Medium',
     'Individual product page with document links table (PI, IF, Guide), product photos, filtered notice list, and related products section. 2-column layout with sidebar.',
     'https://medical.kyowakirin.co.jp/druginfo/detail/activacin-for-injection-600/'],
    ['T05', 'Article / Column', 'Medium',
     'Long-form content page with table of contents, structured sections (H2/H3), author attribution, content gating for member-only sections, and related articles. 2-column with right sidebar.',
     'https://medical.kyowakirin.co.jp/support/work-style/generative-ai.html'],
    ['T06', 'Facility Case Study', 'Medium',
     'Deep article template with facility overview (photo, key facts), 8+ numbered sections, staff photos with captions, printable PDF, related facilities, and product promotions. 2-column with sidebar.',
     'https://medical.kyowakirin.co.jp/kidney/disease/toseki_s/051/index.html'],
    ['T07', 'Content Listing / Library', 'Medium',
     'Grid or list-based content collections with category filtering. Used for seminar listings, news, video library, materials, booklets. Includes tab-based filters, year selectors, and pagination. 2-column with sidebar.',
     'https://medical.kyowakirin.co.jp/webseminar/index.html\nhttps://medical.kyowakirin.co.jp/newslist.html\nhttps://medical.kyowakirin.co.jp/support/movie_library/\nhttps://medical.kyowakirin.co.jp/leaf/index.html\nhttps://medical.kyowakirin.co.jp/druginfo/newslist/'],
    ['T08', 'Data Table Page', 'Simple',
     'Tabular data display for product information (code tables, discontinued products, RMP materials). Static tables with optional PDF download links. 2-column with sidebar.',
     'https://medical.kyowakirin.co.jp/druginfo/code_table/\nhttps://medical.kyowakirin.co.jp/druginfo/drgdiscon/\nhttps://medical.kyowakirin.co.jp/druginfo/rmp.html'],
    ['T09', 'Interactive Tool', 'Complex',
     'Medical calculator page with multiple assessment tools (BSA, PASI, DLQI, PDI). Features tab navigation, dropdown selectors, interactive calculation interfaces, and visual diagrams. Vue.js-powered.',
     'https://medical.kyowakirin.co.jp/allergy/disease/tool-index.html'],
    ['T10', 'Authentication / Form', 'Complex',
     'Login and registration pages with multi-field forms, SSO integrations (medPass, DLink), postal code API lookup, multi-step wizard flow, and client-side validation. Single-column centered layout.',
     'https://medical.kyowakirin.co.jp/login.html\nhttps://medical.kyowakirin.co.jp/menu/entry_input.html'],
    ['T11', 'Utility / Static', 'Simple',
     'Informational pages including sitemap, FAQ (accordion), terms of use, contact info, and external link confirmation. Minimal dynamic features. Variable layouts (single or 2-column).',
     'https://medical.kyowakirin.co.jp/sitemap.html\nhttps://medical.kyowakirin.co.jp/faq.html\nhttps://medical.kyowakirin.co.jp/terms.html\nhttps://medical.kyowakirin.co.jp/contact.html\nhttps://medical.kyowakirin.co.jp/external/confirmation.html'],
]

add_table_with_style(doc,
    ['ID', 'Template Name', 'Complexity', 'Description & Reasoning', 'Reference URL(s)'],
    template_data,
    [1.2, 2.8, 1.5, 6.5, 5.0]
)

doc.add_paragraph('')
doc.add_heading('Template Complexity Legend', level=3)
doc.add_paragraph('Simple: Static content, minimal JS, straightforward layout (<2 days to implement)', style='List Bullet')
doc.add_paragraph('Medium: Some dynamic features, structured content model, moderate interactivity (2-4 days)', style='List Bullet')
doc.add_paragraph('Complex: Heavy JS/Vue.js, API integrations, dynamic filtering, interactive tools (4-8 days)', style='List Bullet')

doc.add_paragraph('')
doc.add_heading('Template Screenshots', level=2)

template_screenshots = [
    ('homepage-clean.png', 'T01 - Homepage'),
    ('kidney-landing.png', 'T02 - Medical Area Landing (Kidney)'),
    ('drug-list.png', 'T03 - Product List / Catalog'),
    ('drug-detail.png', 'T04 - Product Detail'),
    ('article-column.png', 'T05 - Article / Column'),
    ('facility-case-detail.png', 'T06 - Facility Case Study'),
    ('seminar-list.png', 'T07 - Content Listing (Seminar)'),
    ('discontinued-products.png', 'T08 - Data Table Page'),
    ('tool-calculator.png', 'T09 - Interactive Tool'),
    ('login-page.png', 'T10 - Authentication / Form (Login)'),
    ('faq-page.png', 'T11 - Utility / Static (FAQ)'),
]

for filename, caption in template_screenshots:
    add_screenshot(doc, filename, caption, 4.5)
    doc.add_paragraph('')

doc.add_page_break()

# ============================================================
# 3. BLOCKS / COMPONENTS CATALOG
# ============================================================
doc.add_heading('3. Blocks / Components Catalog', level=1)
doc.add_paragraph(
    'The following catalog identifies all reusable blocks and components observed across the site. '
    'Design variations of the same content model are grouped as variants rather than separate blocks.'
)

blocks_data = [
    ['B01', 'Global Header', 'High',
     'Persistent orange header with Kyowa Kirin logo, utility links (contact, search, login, register), and 5-item icon-based main navigation with mega-menu dropdowns. Includes mobile responsive hamburger menu.',
     'All pages'],
    ['B02', 'Global Footer', 'Medium',
     '5-column footer with comprehensive sitemap links organized by section (Drug Info, Regional Info, Seminars, Materials, Support). Includes legal links row (privacy, terms, sitemap, FAQ, contact) and copyright.',
     'All pages'],
    ['B03', 'HCP Gate Dialog', 'High',
     'Full-screen modal overlay requiring healthcare professional role selection (Doctor, Pharmacist, Nurse, Other). Includes login form with email/password, medPass SSO, DLink SSO. Uses Lity.js lightbox library. Sets cookies for persistent verification.',
     'All pages (first visit)'],
    ['B04', 'Hero Carousel', 'High',
     'Auto-rotating image slider with 6 slides, prev/next arrows, dot indicators. Full-width banner images linking to promotional content. Uses Slick slider library with touch/swipe support.',
     'https://medical.kyowakirin.co.jp/'],
    ['B05', 'Product Search Widget', 'High',
     'Inline search box with product name input, search button, and Japanese syllabary (Aiueo) filter links. Includes chatbot launcher button. Vue.js handles API calls to /api/v1.0/product.',
     'https://medical.kyowakirin.co.jp/'],
    ['B06', 'Seminar Listing', 'High',
     'Dynamic seminar list with category filter tabs (All, Kidney, Cancer, Allergy, Neuro, Rare Disease, Other). Each item shows category badge, title, and date/time. "Show all" CTA button. Vue.js-rendered.',
     'https://medical.kyowakirin.co.jp/\nhttps://medical.kyowakirin.co.jp/webseminar/index.html'],
    ['B07', 'News/Content Feed', 'Medium',
     'Tabbed content feed showing latest articles with thumbnail images, titles, descriptions, and publish dates. Category filter tabs. Horizontal card layout (4 cards per row). Link to full listing.',
     'https://medical.kyowakirin.co.jp/'],
    ['B08', 'Recommendation Carousel', 'High',
     'Personalized content carousel powered by /api/v1.0/recommend API. Shows 4-5 cards with thumbnails and titles. Appears on homepage and in sidebar. Vue.js component with Slick slider.',
     'Homepage + sidebar on most pages'],
    ['B09', 'Right Sidebar', 'Medium',
     'Standard sidebar composition: opt-in notification banner (top), recommendation carousel cards, and e-learning promotional banner (bottom). Consistent across 2-column pages.',
     'Most content pages'],
    ['B10', 'Breadcrumb Navigation', 'Low',
     'Horizontal breadcrumb trail showing page hierarchy (HOME > Section > Sub-section > Page). Standard HTML links with ">" separators.',
     'All pages except homepage'],
    ['B11', 'Category Accordion', 'Medium',
     'Expandable category sections with thumbnail images, 2 recent items preview, and "View All" link. Used on medical area landing pages. Accordion expand/collapse behavior.',
     'https://medical.kyowakirin.co.jp/kidney/\nhttps://medical.kyowakirin.co.jp/allergy/ (etc.)'],
    ['B12', 'Product Logo Grid', 'Low',
     'Grid of clickable product name/logo images. Responsive layout adapting from 4 to 6 columns. Links to individual product detail pages.',
     'Medical area landing pages'],
    ['B13', 'Content Card Grid', 'Medium',
     'Grid of cards with image thumbnails, titles, and descriptions. Used for facility case studies, audio lectures, video library, article series. Variants: 3-column, 4-column, horizontal.',
     'Multiple templates (T05, T06, T07)'],
    ['B14', 'Data Table', 'Medium',
     'Structured HTML tables for product data display. Variants include: document links table (product detail), discontinued products table, code/price table, and RMP materials table. Optional PDF download links.',
     'Drug info pages (T03, T04, T08)'],
    ['B15', 'Contact / CTA Block', 'Low',
     'Contact information section with phone number table (free call 0120-850-150), business hours, usage notes, and "Contact Us" CTA button. Used on homepage.',
     'https://medical.kyowakirin.co.jp/'],
    ['B16', 'Promotional Banner Cards', 'Low',
     '3-column cards linking to external/internal promotional content (Clinical Research, Stories, e-Learning). Each card has image, title, and description text.',
     'https://medical.kyowakirin.co.jp/'],
    ['B17', 'FAQ Accordion', 'Medium',
     'Tabbed FAQ sections with numbered Q&A accordion items. Horizontal tabs for category filtering, clickable questions expanding to show answers. Vue.js toggle behavior.',
     'https://medical.kyowakirin.co.jp/faq.html'],
    ['B18', 'Floating Seminar CTA', 'Low',
     'Fixed-position floating button at bottom-right of screen. Green calendar icon with "Seminar Schedule" text. Links to /webseminar/index.html. Persistent across all pages.',
     'All pages'],
    ['B19', 'Cookie Consent Banner', 'Medium',
     'Ensighten Privacy-powered consent banner with Accept All/Settings options. Full-width bottom overlay. Category-based consent toggles in settings modal.',
     'All pages (first visit)'],
    ['B20', 'Chatbot Widget', 'High',
     'Fujitsu CHORDSHIP AI chatbot for product Q&A. Floating chat launcher button. Full chat window with message input, decorations, and option lists. Loaded from external Fujitsu cloud.',
     'All pages (Drug Info section)'],
    ['B21', 'Supply Status Alert', 'Low',
     'Orange/red alert banner at top of Drug Info pages notifying about product shipment suspension or discontinuation. Links to relevant product pages.',
     'Drug Info pages'],
    ['B22', 'Keyword Search Bar', 'Low',
     'Full-width search input with magnifying glass button. Submits to /search.html. Appears below hero on homepage.',
     'https://medical.kyowakirin.co.jp/'],
    ['B23', 'Content Area Carousel', 'Medium',
     'Rotating banner slider within medical area landing pages. Shows seminar/content promotions specific to the clinical area. Variant of Hero Carousel with smaller size.',
     'Medical area landing pages'],
    ['B24', 'External Link Confirmation', 'Low',
     'Centered confirmation dialog for external site navigation. Shows disclaimer text with Yes/No buttons. Vue.js reads outLink query parameter.',
     'https://medical.kyowakirin.co.jp/external/confirmation.html'],
]

add_table_with_style(doc,
    ['ID', 'Block Name', 'Complexity', 'Description & Functionality', 'Reference URL(s)'],
    blocks_data,
    [1.0, 2.5, 1.3, 7.5, 4.5]
)

doc.add_paragraph('')
doc.add_heading('Block Complexity Legend', level=3)
doc.add_paragraph('Low: Static HTML/CSS, no JavaScript interactivity required', style='List Bullet')
doc.add_paragraph('Medium: Moderate interactivity (accordion, tabs, sliders), limited API calls', style='List Bullet')
doc.add_paragraph('High: Complex JS/Vue.js, external API integrations, real-time data, SSO', style='List Bullet')

doc.add_paragraph('')
doc.add_heading('Key Component Screenshots', level=2)

block_screenshots = [
    ('homepage-gate.png', 'B03 - HCP Gate Dialog with Login/SSO'),
    ('homepage-clean.png', 'B04/B05/B06 - Homepage: Hero Carousel, Product Search, Seminar Listing'),
    ('kidney-landing.png', 'B11/B12/B23 - Area Landing: Category Accordion, Product Grid, Area Carousel'),
    ('leaf-materials.png', 'B07/B13 - Content Card Grid / Materials Library'),
    ('movie-library.png', 'B13 - Content Card Grid Variant (Video Library)'),
    ('illust-collection.png', 'B13 - Content Card Grid Variant (Illustration Collection with Left Nav)'),
]
for filename, caption in block_screenshots:
    add_screenshot(doc, filename, caption, 4.5)
    doc.add_paragraph('')

doc.add_page_break()

# ============================================================
# 4. PAGE COUNTS BY TEMPLATE
# ============================================================
doc.add_heading('4. Page Counts by Template', level=1)
doc.add_paragraph(
    'The following table estimates the number of pages using each template type, '
    'based on sitemap analysis, navigation exploration, and content structure patterns. '
    'Note: Many pages require member login for access, so counts are estimated from '
    'visible navigation and sitemap data.'
)

page_count_data = [
    ['T01', 'Homepage', '1', 'Manual',
     'Highly customized; requires custom block development for hero carousel, product search, recommendation engine, and dynamic seminar listing.'],
    ['T02', 'Medical Area Landing', '5', 'Semi-Auto',
     'Consistent template across 5 clinical areas. Content varies (number of categories, products, carousel items) but structure is identical. Template can be automated; content needs manual mapping.'],
    ['T03', 'Product List / Catalog', '1-2', 'Manual',
     'API-driven dynamic page with syllabary filtering. Requires reimplementation of product search API and Vue.js filtering logic.'],
    ['T04', 'Product Detail', '30-40', 'Semi-Auto',
     'One page per pharmaceutical product. Structured data (document links, photos, notices). Could be automated with proper data extraction from API.'],
    ['T05', 'Article / Column', '50-80', 'Semi-Auto',
     'Includes work-style columns (~10), audio lectures (~10), disease info articles (~30-40), and other editorial content. Structured but member-gated content requires manual verification.'],
    ['T06', 'Facility Case Study', '20-25', 'Semi-Auto',
     'Deep articles across 5 series: CKD LIAISON (4), Dialysis Frontline (6+), Chemo Reports (6), PD med.front (1), Rare Disease (6). Rich media and structured sections.'],
    ['T07', 'Content Listing / Library', '10-15', 'Manual',
     'Dynamic listing pages with category filters, year selectors, pagination. Includes seminar list, news list, video library, materials, product news, Q&A list. Vue.js rendering.'],
    ['T08', 'Data Table Page', '5-8', 'Automated',
     'Relatively static tabular content. Code tables, discontinued products, RMP materials, price lists. Low complexity, easy to migrate.'],
    ['T09', 'Interactive Tool', '3-5', 'Manual',
     'Medical calculators (BSA, PASI, DLQI, PDI) and other interactive assessment tools. Require custom JavaScript reimplementation.'],
    ['T10', 'Authentication / Form', '5-8', 'Manual',
     'Login, registration (multi-step), password reset, opt-in preferences. Require SSO integration (medPass, DLink) and form backend.'],
    ['T11', 'Utility / Static', '8-10', 'Automated',
     'Sitemap, FAQ, terms, contact, external confirmation. Static content easily migrated.'],
    ['', 'PDF Documents', '100-150+', 'N/A (Asset)',
     'Product information PDFs, guideline documents, booklet PDFs. These are assets to be migrated as-is to DAM, not HTML pages.'],
]

add_table_with_style(doc,
    ['ID', 'Template', 'Est. Page Count', 'Migration Type', 'Notes'],
    page_count_data,
    [1.0, 2.8, 2.0, 1.8, 9.0]
)

doc.add_paragraph('')

# Summary box
summary = doc.add_paragraph()
run = summary.add_run('Total Estimated HTML Pages: 138-199')
run.font.bold = True
run.font.size = Pt(11)
doc.add_paragraph('')
run2 = doc.add_paragraph().add_run('Total Including PDF Assets: 238-349+')
run2.font.bold = True
run2.font.size = Pt(11)

doc.add_paragraph('')
doc.add_heading('Migration Type Summary', level=2)

migration_type_data = [
    ['Automated', '13-18 pages', 'Data tables, static utility pages. Standard HTML extraction.', 'T08, T11'],
    ['Semi-Automated', '105-150 pages', 'Structured content with consistent templates. Requires template setup then automated content extraction.', 'T02, T04, T05, T06'],
    ['Manual', '20-31 pages', 'Dynamic/interactive pages requiring custom reimplementation (API-driven, Vue.js, SSO, calculators).', 'T01, T03, T07, T09, T10'],
]

add_table_with_style(doc,
    ['Migration Type', 'Page Count', 'Description', 'Templates'],
    migration_type_data,
    [2.5, 2.5, 7.5, 4.0]
)

doc.add_page_break()

# ============================================================
# 5. INTEGRATIONS ANALYSIS
# ============================================================
doc.add_heading('5. Integrations Analysis', level=1)
doc.add_paragraph(
    'The site integrates with numerous third-party services and internal APIs. '
    'This section catalogs all identified integrations with their complexity assessment.'
)

doc.add_heading('5.1 Third-Party Integrations', level=2)

integrations_data = [
    ['Google Tag Manager', 'Tag Management', 'Embed', 'Medium',
     'Container GTM-MRQPPQ8. Manages GA4, UA, and other marketing tags.',
     'All pages'],
    ['Google Analytics 4', 'Analytics', 'Embed', 'Medium',
     'Measurement ID G-K97ZXD4E0W. Tracks page views, scroll, engagement. Custom dimensions: uid, userjob.',
     'All pages'],
    ['Google Universal Analytics', 'Analytics (Legacy)', 'Embed', 'Low',
     'Tracking ID UA-139042179-1. Legacy implementation running alongside GA4.',
     'All pages'],
    ['Ensighten Privacy', 'Consent Management', 'Embed/Plugin', 'High',
     'Full cookie consent management with category-based toggles. Controls tag firing. Client: kirin, publish path: prod_medical-kyowakirin.',
     'All pages'],
    ['Marketo Munchkin', 'Marketing Automation', 'Embed', 'Medium',
     'Account 801-WGL-416. Tracks page visits for lead scoring and email automation.',
     'All pages'],
    ['Marketo RTP', 'Web Personalization', 'Embed', 'High',
     'Real-time personalization engine providing campaign overlays and segment-based content targeting.',
     'All pages'],
    ['Nakanohito (User Insight)', 'B2B Analytics', 'Embed', 'Low',
     'Account 55517. Japanese B2B analytics identifying visiting companies.',
     'All pages'],
    ['Fujitsu CHORDSHIP', 'AI Chatbot', 'Embed/SaaS', 'High',
     'Instance bctrl123-standard. Full AI chatbot for product Q&A. Cloud-hosted on Fujitsu infrastructure.',
     'All pages (Drug Info)'],
    ['medPass SSO', 'Authentication', 'API/Redirect', 'High',
     'Japanese medical professional SSO platform. Redirect-based authentication flow.',
     'Login, Gate Dialog'],
    ['DLink SSO', 'Authentication', 'API/Plugin', 'High',
     'Alternative SSO provider for healthcare professionals. Button-based login.',
     'Login, Gate Dialog'],
    ['Google Fonts', 'CDN/Typography', 'Embed', 'Low',
     'External font loading from fonts.googleapis.com.',
     'All pages'],
    ['Kirin FAQ Platform', 'External Forms', 'Redirect', 'Medium',
     'External form hosting at faq.kirin.co.jp for inquiries and feedback.',
     'Contact page'],
]

add_table_with_style(doc,
    ['Integration', 'Type', 'Method', 'Complexity', 'Description', 'Pages Used'],
    integrations_data,
    [2.5, 2.0, 1.5, 1.3, 5.5, 2.5]
)

doc.add_paragraph('')
doc.add_heading('5.2 Internal API Endpoints', level=2)
doc.add_paragraph(
    'The site exposes several internal REST API endpoints used by the Vue.js frontend:'
)

api_data = [
    ['/api/v1.0/medical_confirm', 'GET', 'Verify HCP status and set medical_check cookie', 'Gate Dialog'],
    ['/api/v1.0/product', 'GET', 'Product search/listing (params: category, word, action)', 'Product List/Search'],
    ['/api/v1.0/product_uselimit', 'GET', 'Product expiry date search (params: name, no)', 'Lot Search'],
    ['/api/v1.0/recommend', 'GET', 'Personalized content recommendations', 'Homepage, Sidebar'],
    ['/api/v1.0/zip_search', 'GET', 'Postal code to address lookup', 'Registration Form'],
    ['/api/v1.0/Reg_Keii_Log', 'POST', 'Audit logging for external journal link access', 'Journal Links'],
    ['/api/v1.0/mypage/set_favorite', 'POST', 'Bookmark/favorite content items', 'Member Pages'],
]

add_table_with_style(doc,
    ['Endpoint', 'Method', 'Purpose', 'Used By'],
    api_data,
    [4.0, 1.5, 6.5, 3.5]
)

doc.add_page_break()

# ============================================================
# 6. COMPLEX USE CASES & OBSERVATIONS
# ============================================================
doc.add_heading('6. Complex Use Cases & Observations', level=1)
doc.add_paragraph(
    'The following complex behaviors, edge cases, and functionality require special '
    'attention during migration planning.'
)

complex_data = [
    ['CU-01', 'HCP Gate Authentication',
     'Every page requires Healthcare Professional (HCP) verification via a modal dialog. Users must select their role (Doctor, Pharmacist, Nurse, etc.) before accessing any content. Session managed via cookies (medical_check). Includes full login form with email/password and two SSO providers (medPass, DLink).',
     'All pages', '1 (global)',
     'Requires reimplementation of role-based access control, cookie management, and SSO integration with medPass and DLink. Forms the foundation of the entire site access model.'],
    ['CU-02', 'Member-Only Content Gating',
     'Certain content sections and full articles are only accessible to logged-in members. Non-members see truncated content with "Continue reading" CTAs that redirect to login. Content visibility is controlled server-side based on session state.',
     'Article pages, area sub-pages', '50-80 pages',
     'Content gating logic needs to be reimplemented. Affects content migration since full content may not be extractable without authentication.'],
    ['CU-03', 'Personalized Recommendations',
     'The /api/v1.0/recommend API provides personalized content suggestions based on browsing history. Vue.js components render recommendation carousels on homepage and in sidebar. Marketo RTP provides additional personalization layer.',
     'Homepage, sidebar on all pages', '2 implementations',
     'Requires alternative recommendation/personalization strategy in EDS. Could use Edge Delivery personalization or simplified static recommendations.'],
    ['CU-04', 'Dynamic Product Search & Filtering',
     'Product listing uses Vue.js with real-time API calls to /api/v1.0/product for filtering by Japanese syllabary categories and keyword search. Product expiry search uses separate API endpoint.',
     'Product list, Lot search', '2-3 pages',
     'Requires custom JavaScript block with API backend or alternative data source. Japanese syllabary filtering is a unique requirement.'],
    ['CU-05', 'Interactive Medical Calculators',
     'BSA, PASI, DLQI, PDI medical assessment calculators with complex scoring logic, body region diagrams, and multi-field input forms. Used for psoriasis severity and QOL evaluation.',
     'Allergy tool page', '3-5 tools',
     'Each calculator requires custom JavaScript reimplementation. Medical accuracy is critical - requires clinical validation after migration.'],
    ['CU-06', 'AI Chatbot (Fujitsu CHORDSHIP)',
     'Full AI-powered chatbot for pharmaceutical Q&A. Loads from external Fujitsu cloud infrastructure. Complex initialization with multiple JS files, CSS, and configuration.',
     'All pages (Drug Info section)', '1 (global)',
     'External SaaS dependency. Embed code needs to be preserved or chatbot needs to be re-integrated in EDS template.'],
    ['CU-07', 'Webinar Platform Integration',
     'Seminar pages link to member-only webinar streaming at /member/webseminar?seminar_id=XXX. Requires authenticated access. Seminar listing is dynamically rendered with category filtering.',
     'Seminar listing, individual seminars', '10-20 seminars',
     'Webinar platform integration needs investigation. May require custom embed or redirect solution.'],
    ['CU-08', 'External Link Audit Trail',
     'All external links (especially journal/literature links) route through confirmation pages (/external/confirmation.html, /external/journal.html) with audit logging via /api/v1.0/Reg_Keii_Log. Required for pharmaceutical regulatory compliance.',
     'All external links', '50+ instances',
     'Regulatory requirement. External link interstitial pattern must be preserved for compliance.'],
    ['CU-09', 'Vue.js Development Build in Production',
     'The site runs Vue.js 3 development build (not production minified). Console shows warning about this. Suggests the site may be using features that depend on development-mode behavior.',
     'All pages', 'Global',
     'Risk factor: behavior may change if upgraded to production build. Should be flagged for testing.'],
    ['CU-10', 'Aggressive Client-Side Redirects',
     'common.js contains logic that automatically redirects users to previously-viewed content based on cookies/browsing history. This causes unexpected navigation during crawling and content extraction.',
     'All pages', 'Global',
     'Complicates automated content migration. Crawlers/importers need to handle or bypass redirect logic.'],
    ['CU-11', 'Disaster Recovery Mode',
     'disasterRecovery.js controls visibility of login/registration UI elements during emergencies. Ensures basic drug information remains accessible without authentication during disaster scenarios.',
     'All pages', '1 mode',
     'Edge case but important for pharmaceutical compliance. Needs equivalent mechanism in EDS.'],
    ['CU-12', 'PDF Document Ecosystem',
     'Extensive PDF document library including product information (PI/IF), guidelines, booklets, and patient materials. Custom PDF viewer (/pdfview/web/viewer.html) using PDF.js. Some PDFs are gated behind authentication.',
     'Throughout site', '100-150+ PDFs',
     'PDFs need to be migrated to DAM. PDF viewer functionality needs reimplementation or simplification. Gated PDFs need access control.'],
]

add_table_with_style(doc,
    ['ID', 'Use Case', 'Description', 'Where Found', 'Instances', 'Why Complex'],
    complex_data,
    [1.0, 2.5, 5.5, 2.5, 1.5, 4.0]
)

doc.add_page_break()

# ============================================================
# 7. MIGRATION ESTIMATES
# ============================================================
doc.add_heading('7. Migration Estimates', level=1)
doc.add_paragraph(
    'The following estimates are based on the identified templates, components, integrations, '
    'and complex use cases. Estimates assume a team familiar with AEM Edge Delivery Services '
    'and the current technology stack.'
)

doc.add_heading('7.1 Effort Breakdown by Template', level=2)

effort_data = [
    ['T01', 'Homepage', '1', '8-10', 'Manual',
     'Custom blocks: hero carousel, product search widget, seminar listing, recommendation engine, contact section.'],
    ['T02', 'Medical Area Landing', '5', '6-8', 'Semi-Auto',
     'Single template with 5 content variations. Category accordion, product grid, area carousel.'],
    ['T03', 'Product List / Catalog', '1-2', '5-7', 'Manual',
     'API-driven product search with syllabary filtering. Requires backend API or data indexing.'],
    ['T04', 'Product Detail', '30-40', '4-6', 'Semi-Auto',
     'Template development + automated content extraction via product API. ~0.5 hr per page for QA.'],
    ['T05', 'Article / Column', '50-80', '4-6', 'Semi-Auto',
     'Template development + bulk import. Content gating logic adds complexity.'],
    ['T06', 'Facility Case Study', '20-25', '3-5', 'Semi-Auto',
     'Structured long-form template. Rich media handling. ~1 hr per page for QA.'],
    ['T07', 'Content Listing / Library', '10-15', '6-8', 'Manual',
     'Multiple listing variants (seminars, news, videos, materials). Dynamic filtering.'],
    ['T08', 'Data Table Page', '5-8', '2-3', 'Automated',
     'Simple table extraction. Minimal customization needed.'],
    ['T09', 'Interactive Tool', '3-5', '8-12', 'Manual',
     'Medical calculators require custom JS. Clinical accuracy validation critical.'],
    ['T10', 'Authentication / Form', '5-8', '6-8', 'Manual',
     'SSO integration (medPass, DLink), multi-step registration, validation.'],
    ['T11', 'Utility / Static', '8-10', '2-3', 'Automated',
     'Simple content extraction. FAQ accordion needs minor JS.'],
]

add_table_with_style(doc,
    ['ID', 'Template', 'Pages', 'Dev Effort (days)', 'Migration Type', 'Notes'],
    effort_data,
    [1.0, 2.8, 1.5, 2.0, 1.5, 7.5]
)

doc.add_paragraph('')
doc.add_heading('7.2 Cross-Cutting Effort', level=2)

crosscut_data = [
    ['Design System / CSS Migration', '5-7', 'Extract and map design tokens, typography, color palette (orange #EA5504 brand), spacing system.'],
    ['Global Header & Navigation', '3-4', 'Implement responsive header with mega-menu, mobile hamburger, utility links.'],
    ['Global Footer', '2-3', '5-column footer with sitemap, legal links, copyright.'],
    ['HCP Gate / Authentication', '5-8', 'Role-based gating, login form, medPass SSO, DLink SSO integration.'],
    ['Recommendation Engine', '3-5', 'Replace or reimplement personalized content recommendations.'],
    ['Chatbot Integration', '2-3', 'Re-embed Fujitsu CHORDSHIP chatbot in EDS templates.'],
    ['Search Functionality', '3-4', 'Implement site search and product search in EDS.'],
    ['Analytics & Tag Management', '2-3', 'Migrate GTM, GA4, Ensighten, Marketo, Nakanohito tags.'],
    ['PDF Migration & Viewer', '2-3', 'Migrate 100-150+ PDFs to DAM. Implement or simplify PDF viewer.'],
    ['External Link Compliance', '1-2', 'Implement external link confirmation interstitial for regulatory compliance.'],
    ['Content Import Scripts', '5-7', 'Develop import infrastructure: parsers, transformers, bulk import scripts.'],
    ['Cookie Consent (Ensighten)', '2-3', 'Reintegrate Ensighten Privacy or migrate to alternative consent platform.'],
]

add_table_with_style(doc,
    ['Work Item', 'Effort (days)', 'Description'],
    crosscut_data,
    [3.5, 2.0, 11.0]
)

doc.add_paragraph('')
doc.add_heading('7.3 QA & Testing Effort', level=2)

qa_data = [
    ['Template Validation', '3-5', 'Visual comparison of all 11 templates against originals.'],
    ['Content Accuracy Check', '5-8', 'Verify all migrated content matches source. Japanese text, product names, dosages.'],
    ['Interactive Features Testing', '3-5', 'Test calculators, search, filtering, carousels, chatbot, SSO.'],
    ['Responsive/Mobile Testing', '2-3', 'Verify all pages on mobile, tablet, desktop breakpoints.'],
    ['Accessibility Compliance', '2-3', 'WCAG 2.1 AA compliance check (required for medical sites in Japan).'],
    ['Cross-Browser Testing', '1-2', 'Edge, Chrome, Safari (as specified in site terms).'],
    ['Performance Testing', '1-2', 'Core Web Vitals, Lighthouse scores, page load times.'],
    ['Security / Auth Testing', '2-3', 'Verify HCP gating, login flow, SSO, content access control.'],
    ['PDF & Document Testing', '1-2', 'Verify all PDF links, viewer functionality, download access.'],
    ['Regulatory Compliance', '2-3', 'External link interstitials, content gating, audit trails.'],
]

add_table_with_style(doc,
    ['QA Activity', 'Effort (days)', 'Description'],
    qa_data,
    [3.5, 2.0, 11.0]
)

doc.add_paragraph('')
doc.add_heading('7.4 Total Migration Estimate Summary', level=2)

summary_data = [
    ['Template & Block Development', '54-76', '11 templates, 24 blocks, custom components'],
    ['Cross-Cutting Infrastructure', '35-52', 'Auth, search, recommendations, analytics, PDF, compliance'],
    ['Content Migration & Import', '10-15', 'Import scripts, bulk migration, content QA'],
    ['QA & Testing', '22-36', 'Comprehensive testing across all dimensions'],
    ['Project Management & Buffer', '10-15', 'Coordination, documentation, contingency (15-20%)'],
    ['', '', ''],
    ['TOTAL ESTIMATE', '131-194 person-days', '(approx. 6.5-9.5 months with 1 developer, or 3-5 months with 2 developers)'],
]

add_table_with_style(doc,
    ['Phase', 'Effort (person-days)', 'Scope'],
    summary_data,
    [3.5, 3.0, 10.0]
)

doc.add_paragraph('')
doc.add_heading('7.5 Recommended Phasing', level=2)

phase_data = [
    ['Phase 1: Foundation\n(Weeks 1-4)', '20-30',
     'Design system migration, global header/footer, HCP gate, authentication framework, EDS project setup.'],
    ['Phase 2: Core Templates\n(Weeks 5-10)', '30-40',
     'Homepage, medical area landing, product list/detail, article/column templates. Import infrastructure.'],
    ['Phase 3: Content Migration\n(Weeks 11-14)', '15-25',
     'Bulk content import for articles, product pages, facility case studies. PDF migration.'],
    ['Phase 4: Advanced Features\n(Weeks 15-18)', '20-30',
     'Interactive tools, search, recommendations, chatbot, seminar integration, webinar platform.'],
    ['Phase 5: Integration & Testing\n(Weeks 19-22)', '20-30',
     'SSO integration, analytics, consent management. Full QA cycle across all templates.'],
    ['Phase 6: Launch Prep\n(Weeks 23-24)', '5-10',
     'Final QA, performance optimization, redirect mapping, go-live preparation.'],
]

add_table_with_style(doc,
    ['Phase', 'Effort (person-days)', 'Description'],
    phase_data,
    [3.5, 2.5, 10.5]
)

doc.add_page_break()

# ============================================================
# 8. APPENDIX: SCREENSHOTS
# ============================================================
doc.add_heading('8. Appendix: Additional Screenshots', level=1)

appendix_screenshots = [
    ('hematonco-landing.png', 'Medical Area Landing - Hematonco (Cancer/Blood)'),
    ('allergy-landing.png', 'Medical Area Landing - Allergy/Immunology'),
    ('neuro-landing.png', 'Medical Area Landing - Central Nervous System'),
    ('raredisease-landing.png', 'Medical Area Landing - Rare Diseases'),
    ('support-landing.png', 'Support Landing Page'),
    ('drug-qa.png', 'Drug Q&A Page'),
    ('drug-lot.png', 'Product Lot/Expiry Search'),
    ('drug-code.png', 'Drug Price/Code Tables'),
    ('drug-rmp.png', 'RMP Information Materials'),
    ('news-list.png', 'News Listing Page'),
    ('audio-lectures.png', 'Audio Lectures Page'),
    ('medical-facilities.png', 'Medical Facility Case Studies'),
    ('contact-page.png', 'Contact Page'),
    ('sitemap-page.png', 'Sitemap Page'),
    ('terms-page.png', 'Terms of Use Page'),
    ('registration-page.png', 'Member Registration Form'),
    ('external-confirm.png', 'External Link Confirmation'),
]

for filename, caption in appendix_screenshots:
    filepath = f'/workspace/{filename}'
    if not os.path.exists(filepath):
        filepath_alt = f'/tmp/playwright/{filename}'
        if os.path.exists(filepath_alt):
            import shutil
            shutil.copy2(filepath_alt, filepath)
    add_screenshot(doc, filename, caption, 4.5)
    doc.add_paragraph('')

# ============================================================
# SAVE DOCUMENT
# ============================================================
output_path = '/workspace/Kyowa_Kirin_Medical_Site_Analysis_Report.docx'
doc.save(output_path)
print(f'Report saved to: {output_path}')
print(f'File size: {os.path.getsize(output_path) / 1024:.1f} KB')
