import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, AlignmentType, HeadingLevel, BorderStyle, ImageRun,
  ShadingType, TableBorders, PageBreak, Header, Footer, TabStopType,
  TabStopPosition, convertInchesToTwip
} from 'docx';
import fs from 'fs';

// Helper to create a styled heading
function heading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, spacing: { before: 300, after: 200 }, children: [new TextRun({ text, bold: true, color: '003478' })] });
}

function subheading(text) {
  return heading(text, HeadingLevel.HEADING_2);
}

function subsubheading(text) {
  return heading(text, HeadingLevel.HEADING_3);
}

function para(text, opts = {}) {
  return new Paragraph({ spacing: { after: 120 }, children: [new TextRun({ text, size: 22, ...opts })] });
}

function boldPara(text) {
  return para(text, { bold: true });
}

function bulletPoint(text, level = 0) {
  return new Paragraph({
    bullet: { level },
    spacing: { after: 60 },
    children: [new TextRun({ text, size: 22 })]
  });
}

// Table styling
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  insideVertical: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
};

function headerCell(text) {
  return new TableCell({
    shading: { type: ShadingType.SOLID, color: '003478' },
    children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20 })] })],
    verticalAlign: 'center',
  });
}

function cell(text, opts = {}) {
  return new TableCell({
    shading: opts.shading ? { type: ShadingType.SOLID, color: opts.shading } : undefined,
    children: [new Paragraph({ spacing: { before: 40, after: 40 }, children: [new TextRun({ text: text || '', size: 20, ...opts })] })],
    verticalAlign: 'center',
    width: opts.width ? { size: opts.width, type: WidthType.PERCENTAGE } : undefined,
  });
}

function createTable(headers, rows) {
  return new Table({
    borders: tableBorders,
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: [
      new TableRow({ children: headers.map(h => headerCell(h)), tableHeader: true }),
      ...rows.map((r, i) => new TableRow({
        children: r.map(c => cell(c, i % 2 === 1 ? { shading: 'F5F8FC' } : {})),
      })),
    ],
  });
}

function loadImage(filename) {
  try {
    const data = fs.readFileSync(`/tmp/playwright/${filename}`);
    return new ImageRun({ data, transformation: { width: 600, height: 400 }, type: 'png' });
  } catch { return null; }
}

function screenshotParagraph(filename, caption) {
  const img = loadImage(filename);
  const children = [];
  if (img) {
    children.push(new Paragraph({ spacing: { before: 200, after: 60 }, children: [img] }));
  }
  children.push(new Paragraph({
    spacing: { after: 200 },
    alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: caption, italics: true, size: 18, color: '666666' })]
  }));
  return children;
}

// ============ BUILD DOCUMENT ============

const sections = [];

// ---- TITLE PAGE ----
sections.push({
  properties: {},
  children: [
    new Paragraph({ spacing: { before: 3000 } }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: 'SACOMBANK WEBSITE', size: 56, bold: true, color: '003478' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 200 },
      children: [new TextRun({ text: 'MIGRATION ANALYSIS REPORT', size: 48, bold: true, color: '003478' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━', size: 28, color: '003478' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 400 },
      children: [new TextRun({ text: 'www.sacombank.com.vn', size: 28, color: '0066CC' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { after: 100 },
      children: [new TextRun({ text: 'Comprehensive Site Analysis for AEM Edge Delivery Services Migration', size: 24, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 600, after: 100 },
      children: [new TextRun({ text: `Date: April 16, 2026`, size: 24, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: 'Status: Draft for Review', size: 24, color: '555555' })]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 400 },
      children: [new TextRun({ text: 'CONFIDENTIAL', size: 20, bold: true, color: 'CC0000' })]
    }),
  ]
});

// ---- TABLE OF CONTENTS ----
const tocItems = [
  '1. Executive Summary',
  '2. Templates Inventory',
  '3. Blocks / Components Catalog',
  '4. Page Counts by Template',
  '5. Integrations Analysis',
  '6. Complex Use Cases & Observations',
  '7. Migration Estimates',
  'Appendix: Screenshots'
];

sections.push({
  properties: {},
  children: [
    heading('Table of Contents'),
    new Paragraph({ spacing: { after: 200 } }),
    ...tocItems.map((item, i) => new Paragraph({
      spacing: { after: 100 },
      children: [new TextRun({ text: item, size: 24, color: '003478' })]
    })),
    new Paragraph({ children: [new PageBreak()] }),

    // ---- 1. EXECUTIVE SUMMARY ----
    heading('1. Executive Summary'),
    para('This report presents a comprehensive analysis of the Sacombank website (www.sacombank.com.vn) in preparation for migration to Adobe Edge Delivery Services (AEM EDS). The analysis covers template inventory, component catalog, page counts, third-party integrations, complex use cases, and migration effort estimates.'),
    new Paragraph({ spacing: { after: 100 } }),
    boldPara('Key Findings:'),
    bulletPoint('The site is currently built on Adobe Experience Manager (AEM) with extensive custom components'),
    bulletPoint('Approximately 2,000 total pages identified across 14 distinct template types'),
    bulletPoint('News articles constitute the largest content volume (~1,271 pages spanning 2014-2026)'),
    bulletPoint('The site serves two primary customer segments: Personal Banking (Cá Nhân) and Enterprise Banking (Doanh Nghiệp)'),
    bulletPoint('16+ third-party integrations identified including dual analytics stacks (Google + Adobe), FPT.AI chatbot, FiinGroup stock data, and social plugins'),
    bulletPoint('Several complex interactive features require special migration attention: real-time exchange rates, financial calculators, product comparison tools, and stock data dashboards'),
    bulletPoint('Estimated total migration effort: 85-115 working days (approximately 4-6 months)'),
    new Paragraph({ children: [new PageBreak()] }),

    // ---- 2. TEMPLATES INVENTORY ----
    heading('2. Templates Inventory'),
    para('The following table lists all unique page templates identified across the Sacombank website. Each template is classified by complexity and includes reference URLs.'),
    new Paragraph({ spacing: { after: 200 } }),

    createTable(
      ['#', 'Template Name', 'Complexity', 'Description', 'Reference URL(s)'],
      [
        ['1', 'Segment Landing Page (Homepage)', 'High', 'Rich multi-section homepage for customer segments with hero carousel, navigation bar, product recommendations, app download section, news feed, and awards carousel. Two variants: Personal & Enterprise.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/doanh-nghiep.html'],
        ['2', 'Product Category Landing', 'Medium', 'Category-level product page with hero banner, sub-category navigation, featured product cards grid, value proposition section, FAQ accordion, and lead capture form.', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/ca-nhan/tiet-kiem/lai-suat-cao.html\nsacombank.com.vn/ca-nhan/ngan-hang-so.html'],
        ['3', 'Filterable Product Listing', 'High', 'Product listing with filters, category tabs, checkbox filters, product cards with specs, comparison feature ("So sánh"), and pagination. Includes card comparison bar and lead form.', 'sacombank.com.vn/ca-nhan/the/the-tin-dung.html\nsacombank.com.vn/ca-nhan/vay/vay-tieu-dung.html\nsacombank.com.vn/ca-nhan/bao-hiem/bao-hiem-nhan-tho.html'],
        ['4', 'Product Detail Page', 'Medium', 'Single product deep-dive with anchor navigation bar, feature highlights, specifications, conditions/requirements, related products, and lead form.', 'sacombank.com.vn/ca-nhan/tai-khoan/tai-khoan-thanh-toan.html'],
        ['5', 'News/Article Listing', 'Medium', 'News listing page with tab navigation (Tin Sacombank / Thông báo), year filter dropdown, text search, 2-column article card grid, and "Load more" pagination.', 'sacombank.com.vn/trang-chu/tin-tuc/tin-sacombank.html'],
        ['6', 'News Article Detail', 'Low', 'Article page with breadcrumb, social sharing (Facebook, Zalo), like counter, rich text body with inline images, and "Latest news" sidebar.', 'sacombank.com.vn/trang-chu/tin-tuc/tin-sacombank/2026/...html'],
        ['7', 'Contact Page', 'Medium', 'Contact page with tabbed interface (Hotline, Form, Live chat), phone number cards, info cards, headquarters details, and embedded Google Maps.', 'sacombank.com.vn/trang-chu/lien-he.html'],
        ['8', 'Utility/Tools Page', 'High', 'Data-centric pages with tables, interactive calculators (deposit, loan, FX, installment, insurance), real-time exchange rates, date pickers, and PDF export. Multiple sub-page variants.', 'sacombank.com.vn/cong-cu/ty-gia.html\nsacombank.com.vn/cong-cu/lai-suat.html\nsacombank.com.vn/cong-cu/cong-cu-khac.html'],
        ['9', 'Document Repository', 'Medium', 'Structured document listing with category tabs, search, date filters, accordion sections (fee schedules), and PDF download links. Used for fees, reports, disclosures.', 'sacombank.com.vn/cong-cu/bieu-phi.html\nsacombank.com.vn/trang-chu/nha-dau-tu/bao-cao.html\nsacombank.com.vn/trang-chu/nha-dau-tu/cong-bo-thong-tin.html'],
        ['10', 'Promotions Hub', 'Medium', 'Promotion aggregation page with audience selector dropdown, search bar, categorized offer sections, partner deals, and newsletter signup with reCAPTCHA.', 'sacombank.com.vn/trang-chu/khuyen-mai/khcn.html\nsacombank.com.vn/trang-chu/khuyen-mai/khdn.html'],
        ['11', 'Audience Segment Page', 'Low', 'Cross-product recommendation by life stage (Student, Career, Family, Retirement). Hero, segment tabs, category filters, product cards, and lead form.', 'sacombank.com.vn/ca-nhan/nhom-san-pham-khcn/sinh-vien-hoc-sinh.html'],
        ['12', 'Investor Relations Dashboard', 'High', 'Composite dashboard with stock ticker (FiinGroup iframes), financial data charts, shareholder meeting docs, disclosure tables, reports, and governance docs.', 'sacombank.com.vn/trang-chu/nha-dau-tu.html'],
        ['13', 'Product Suggestion Wizard', 'High', 'Interactive JS-driven questionnaire/wizard for product recommendations. Minimal server-rendered HTML; logic runs client-side.', 'sacombank.com.vn/trang-chu/Goi-y-san-pham/goi-y-the-ca-nhan.html'],
        ['14', '404 Error Page', 'Low', 'Error page with 404 display, error message, home button, and suggestion cards.', 'N/A (triggered by invalid URLs)'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    ...screenshotParagraph('sacombank-homepage.png', 'Figure 1: Segment Landing Page - Personal Banking Homepage'),
    ...screenshotParagraph('sacombank-credit-cards.png', 'Figure 2: Filterable Product Listing - Credit Cards with comparison feature'),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- 3. BLOCKS / COMPONENTS CATALOG ----
    heading('3. Blocks / Components Catalog'),
    para('The following table catalogs all reusable blocks and components identified across the site. Design variations of the same content model are grouped together rather than listed as separate blocks.'),
    new Paragraph({ spacing: { after: 200 } }),

    createTable(
      ['#', 'Block Name', 'Complexity', 'Description & Behavior', 'Reference URL(s)'],
      [
        ['1', 'Header / Mega Menu', 'High', 'Sticky header with hamburger menu, customer segment switcher (Cá Nhân / Doanh Nghiệp), Sacombank logo, and Internet Banking CTA. Mega menu reveals full navigation tree on click. Responsive behavior.', 'All pages (global)'],
        ['2', 'Footer', 'Medium', 'Multi-column footer with logo, quick links (Liên hệ, Tuyển dụng, Công cụ, Biểu phí, An toàn bảo mật), expandable sitemap columns (6 sections), app download QR codes (mBanking + Pay), social links (Facebook, YouTube), and copyright.', 'All pages (global)'],
        ['3', 'Hero Banner', 'Medium', 'Full-width hero with breadcrumb navigation, H1 title, subtitle text, CTA button(s), and hero image. Two design variants: (a) Standard with single CTA, (b) Dual CTA variant (e.g., "Gợi ý Thẻ" + "Nhận tư vấn"). Consistent content model.', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/ca-nhan/the/the-tin-dung.html'],
        ['4', 'Hero Carousel/Slider', 'High', 'Full-width rotating carousel on homepage with 3-4 slides. Each slide: background image, heading, description, CTA link. Auto-rotate with manual navigation arrows. Used on segment landing pages only.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/doanh-nghiep.html'],
        ['5', 'Category Tab Navigation', 'Medium', 'Horizontal scrollable tab bar for sub-category navigation. Active tab highlighted. Used consistently across product, tools, and listing pages. Chevron arrow for overflow scroll.', 'sacombank.com.vn/ca-nhan/the/the-tin-dung.html\nsacombank.com.vn/cong-cu/ty-gia.html'],
        ['6', 'Product Card Grid', 'Medium', 'Grid of product cards (2-4 columns). Each card: product image, title, description/features, dual CTA buttons ("Nhận tư vấn" + "Xem chi tiết"). Two variants: (a) Simple with image+title+CTA, (b) Detailed with spec columns (Ưu đãi, Hạn mức, Phí).', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/ca-nhan/the/the-tin-dung.html'],
        ['7', 'Filter Chips / Checkbox Filter', 'Medium', 'Inline filter chips (pill-shaped tags) with checkbox toggles and "Lọc" (Filter) button. Used for product attribute filtering (Hoàn tiền, Đặc quyền, Tích dặm, Tích điểm).', 'sacombank.com.vn/ca-nhan/the/the-tin-dung.html\nsacombank.com.vn/ca-nhan/bao-hiem/bao-hiem-nhan-tho.html'],
        ['8', 'Comparison Bar', 'Medium', 'Sticky bottom bar showing selected products count ("Bạn đã chọn X thẻ") with "Bắt đầu so sánh" (Start comparing) CTA link. Appears on filterable product listing pages.', 'sacombank.com.vn/ca-nhan/the/the-tin-dung.html'],
        ['9', 'Lead Capture Form', 'Medium', 'Contact/consultation form with fields: Họ và tên (Name), Email, Số điện thoại (Phone), CCCD/Mã số thuế (ID/Tax ID), and "Đăng ký" button. Two variants: (a) Personal (uses CCCD), (b) Enterprise (uses Tax ID).', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/doanh-nghiep/tin-dung.html'],
        ['10', 'Value Proposition Grid', 'Low', '3-4 column grid of benefit cards. Each card: icon, stat/heading, description. Used for "Vì sao nên chọn SACOMBANK" sections.', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/ca-nhan/ngan-hang-so.html'],
        ['11', 'FAQ Accordion', 'Low', 'Expandable/collapsible Q&A section with numbered items. Click to toggle answer visibility. Standard accordion behavior.', 'sacombank.com.vn/ca-nhan/tai-khoan.html\nsacombank.com.vn/doanh-nghiep/tai-khoan-dich-vu.html'],
        ['12', 'Feature Promo Block', 'Medium', 'Split layout: text content (heading, description, key stats) on one side, product image on the other. Stats shown as large numbers (e.g., "0% lãi suất", "55 ngày"). Includes CTA link.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/ca-nhan/ngan-hang-so.html'],
        ['13', 'News Card Grid', 'Low', 'Grid of news article cards (2-4 columns). Each card: thumbnail image, article title (linked), date, excerpt text. Horizontal scroll variant on homepage.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/trang-chu/tin-tuc/tin-sacombank.html'],
        ['14', 'Awards Carousel', 'Low', 'Horizontal scrollable carousel of award/certification logos with captions. Navigation arrows (left/right). Used on segment landing pages.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/doanh-nghiep.html'],
        ['15', 'App Download Section', 'Low', 'Promotional block for mobile app download. App icon, heading, description, app store badges (iOS + Android), QR code for direct download. Two app variants (mBanking + Pay).', 'sacombank.com.vn/ca-nhan.html (in-page)\nAll pages (footer)'],
        ['16', '"I Want" Product Selector', 'Medium', 'Interactive dropdown-based product finder. Dropdown with suggestions (Gợi ý, Thẻ tín dụng, etc.) + "Bắt đầu" CTA. Routes to product suggestion wizard.', 'sacombank.com.vn/ca-nhan.html'],
        ['17', 'Personalized Recommendations ("Dành riêng cho bạn")', 'Medium', 'Horizontal scrollable card carousel with product/promo cards. Each card: image, title, description, CTA link. Used for personalized content sections.', 'sacombank.com.vn/ca-nhan.html\nsacombank.com.vn/ca-nhan/ngan-hang-so.html'],
        ['18', 'Product Segment Grid', 'Low', 'Grid of life-stage category cards with icons (Student, Career, Family, Retirement). Each card links to audience-specific product page.', 'sacombank.com.vn/ca-nhan.html'],
        ['19', 'Anchor Navigation Bar', 'Medium', 'Sticky in-page navigation with anchor links to sections (Đặc tính, Tiện ích, Điều kiện, Ưu đãi). Highlights active section on scroll. Used on product detail pages.', 'sacombank.com.vn/ca-nhan/tai-khoan/tai-khoan-thanh-toan.html'],
        ['20', 'Social Sharing Bar', 'Low', 'Floating left sidebar with like/reaction button, Facebook share, and Zalo share buttons. Used on news article detail pages.', 'sacombank.com.vn/trang-chu/tin-tuc/.../[article].html'],
        ['21', 'Latest News Sidebar', 'Low', 'Right sidebar widget showing 4 recent news article titles with dates. Links to individual articles.', 'sacombank.com.vn/trang-chu/tin-tuc/.../[article].html'],
        ['22', 'Exchange Rate Table', 'High', 'Data table with real-time exchange rates (columns: Currency, Cash Buy, Transfer Buy, Cash Sell, Transfer Sell). Date/time picker, PDF download, "View all" expandable. API-driven real-time data.', 'sacombank.com.vn/cong-cu/ty-gia.html'],
        ['23', 'Currency Converter', 'High', 'Interactive calculator with transaction type selector (4 options), currency dropdown, amount slider + text input, and computed VND output. Live rate integration.', 'sacombank.com.vn/cong-cu/ty-gia.html'],
        ['24', 'Financial Calculator', 'High', 'Multi-tab calculator widget (Deposit, Loan, Transaction Fee, Installment, Insurance). Sliders, text inputs, dropdown selectors, computed output fields. Reusable across multiple pages.', 'sacombank.com.vn/cong-cu/lai-suat.html\nsacombank.com.vn/cong-cu/cong-cu-khac.html'],
        ['25', 'Data Table with Filters', 'Medium', 'Sortable/filterable data table for documents, forms, fee schedules. Columns vary by context. Search input, category filter dropdown, date picker.', 'sacombank.com.vn/cong-cu/bieu-mau.html\nsacombank.com.vn/trang-chu/nha-dau-tu/cong-bo-thong-tin.html'],
        ['26', 'Fee Schedule Accordion', 'Medium', 'Expandable accordion organized by fee category (A through G + special sections). Each category expands to show fee details. Customer segment toggle (KHCN/KHDN).', 'sacombank.com.vn/cong-cu/bieu-phi.html'],
        ['27', 'Stock Ticker / Financial Data Widget', 'High', 'Embedded FiinGroup iframes showing STB stock price, trading volume, market cap, interactive stock chart with time period selectors. Third-party dependency.', 'sacombank.com.vn/trang-chu/nha-dau-tu.html'],
        ['28', 'Contact Info Cards', 'Low', 'Grid of 3 info cards with icon, title, description, and CTA link. Used on contact page for Hotline, FAQ, and Branch locator links.', 'sacombank.com.vn/trang-chu/lien-he.html'],
        ['29', 'Google Maps Embed', 'Low', 'Embedded Google Maps iframe showing headquarters location. Includes address, phone, email, SWIFT code, working hours.', 'sacombank.com.vn/trang-chu/lien-he.html'],
        ['30', 'Newsletter Signup', 'Low', 'Email subscription form with email text input, Google reCAPTCHA Enterprise checkbox, and submit button.', 'sacombank.com.vn/trang-chu/khuyen-mai/khcn.html'],
        ['31', 'Cookie Consent Banner', 'Low', 'Bottom-of-page cookie consent bar with message, "Từ chối" (Reject) and "Đồng ý" (Accept) buttons. Global across all pages.', 'All pages (global)'],
        ['32', 'FPT.AI Live Chat Widget', 'Medium', 'Floating chat button (bottom-right) that opens FPT.AI chatbot iframe. Persistent across all pages. Third-party integration.', 'All pages (global)'],
        ['33', 'Related Products Carousel', 'Low', 'Horizontal card carousel showing 3-4 related product cards with image, title, and "Xem chi tiết" link. Appears at bottom of product and tools pages.', 'sacombank.com.vn/ca-nhan/tai-khoan/tai-khoan-thanh-toan.html\nsacombank.com.vn/cong-cu/ty-gia.html'],
        ['34', 'CTA Banner ("Bạn chưa tìm được thông tin?")', 'Low', 'Full-width banner with heading, subtitle, and "Liên hệ ngay" CTA button. Used as fallback call-to-action on various pages.', 'sacombank.com.vn/cong-cu/ty-gia.html'],
        ['35', 'Request Lookup Form', 'Medium', 'Form with request type dropdown, request code input, ID number input, reCAPTCHA, and submit button. Specific to request tracking utility.', 'sacombank.com.vn/tien-ich/tra-cuu-yeu-cau.html'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    ...screenshotParagraph('sacombank-exchange-rate.png', 'Figure 3: Utility/Tools Page - Exchange Rate Table with Currency Converter'),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- 4. PAGE COUNTS BY TEMPLATE ----
    heading('4. Page Counts by Template'),
    para('Page counts are derived from the XML sitemap (1,029 URLs) supplemented by API endpoint analysis revealing additional content not indexed in the sitemap. The total estimated page count is approximately 2,000 pages.'),
    new Paragraph({ spacing: { after: 200 } }),

    createTable(
      ['#', 'Template Type', 'Est. Page Count', 'Migration Approach', 'Notes'],
      [
        ['1', 'Segment Landing Page', '2-4', 'Manual', 'Complex layout with many interactive components; 2 primary (Personal, Enterprise) + possible sub-variants'],
        ['2', 'Product Category Landing', '~25', 'Semi-automated', 'Consistent structure across categories; hero + tabs + cards + form pattern repeats'],
        ['3', 'Filterable Product Listing', '~15', 'Manual', 'Dynamic filtering, comparison logic, and AJAX-loaded product cards require custom JS'],
        ['4', 'Product Detail Page', '~420', 'Automated', 'High volume (~291 personal + ~130 enterprise); standardized structure with anchor nav + specs + form'],
        ['5', 'News/Article Listing', '~5', 'Manual', 'Small number of listing pages but dynamic search/filter behavior needs custom implementation'],
        ['6', 'News Article Detail', '~1,271', 'Automated', '1,193 bank news + 78 announcements; consistent article template; 13 years of content (2014-2026)'],
        ['7', 'Contact Page', '1', 'Manual', 'Unique page with Google Maps embed and tabbed interface'],
        ['8', 'Utility/Tools Page', '~13', 'Manual', 'Interactive calculators, real-time data feeds, and complex widgets require custom development'],
        ['9', 'Document Repository', '~10', 'Semi-automated', 'Structured document listings; content is standardized but filters need implementation'],
        ['10', 'Promotions Hub', '~5', 'Semi-automated', 'Hub pages are few; promotion detail pages (~150) can be automated'],
        ['11', 'Promotion Detail Pages', '~150', 'Automated', 'Not in sitemap but discovered via API; consistent card/detail structure'],
        ['12', 'Audience Segment Page', '~4', 'Semi-automated', '4 life-stage segments; consistent template with filter + product cards'],
        ['13', 'Investor Relations Dashboard', '1', 'Manual', 'Unique composite page with third-party FiinGroup stock data iframes'],
        ['14', 'Product Suggestion Wizard', '~2', 'Manual', 'JS-driven interactive wizard; minimal server HTML; requires full rebuild'],
        ['15', '404 Error Page', '1', 'Manual', 'Simple static page'],
        ['', 'Miscellaneous / Other', '~75', 'Mixed', 'Root pages, search, standalone pages, support/FAQ'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),

    subsubheading('Summary of Migration Approach'),
    createTable(
      ['Migration Approach', 'Page Count', 'Percentage'],
      [
        ['Automated (standardized, repeatable templates)', '~1,841', '~92%'],
        ['Semi-automated (requires template setup + some manual review)', '~54', '~3%'],
        ['Manual (unique/complex pages requiring custom development)', '~105', '~5%'],
        ['Total Estimated Pages', '~2,000', '100%'],
      ]
    ),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- 5. INTEGRATIONS ANALYSIS ----
    heading('5. Integrations Analysis'),
    para('The Sacombank website integrates with numerous third-party services across analytics, marketing, customer service, financial data, and social media. The following table provides a complete inventory.'),
    new Paragraph({ spacing: { after: 200 } }),

    subsubheading('5.1 Analytics & Tracking'),
    createTable(
      ['Integration', 'Type', 'Complexity', 'Details', 'Pages Used'],
      [
        ['Google Tag Manager', 'Tag Management', 'Medium', 'Container ID: GTM-M5276CD. Central tag management system orchestrating all Google tags.', 'Global (all pages)'],
        ['Google Analytics 4', 'Analytics', 'Medium', '5 separate GA4 property IDs: G-NXMYL66EL2, G-D61HXHHQ9P, G-2KYEH9105E, G-V8SGVWN98N, G-JVYMGCZZS9. Multiple properties suggest departmental tracking.', 'Global (all pages)'],
        ['Google Ads Conversion', 'Advertising', 'Medium', '8+ conversion tracking IDs (AW-11327170602, AW-11339929524, etc.) including DoubleClick (dc.js) integration.', 'Global (all pages)'],
        ['Adobe Analytics / Launch', 'Analytics', 'High', 'Adobe DTM/Launch (launch-44025ca85cd2.min.js), AppMeasurement.js. Full Adobe Analytics stack running in parallel with Google Analytics.', 'Global (all pages)'],
        ['Adobe Helix RUM', 'Monitoring', 'Low', 'Real User Monitoring for performance tracking.', 'Global (all pages)'],
        ['Microsoft Clarity', 'Session Replay', 'Low', 'Project ID: dpwtej66sj. Session recording and heatmap analytics (clarity.ms).', 'Global (all pages)'],
        ['Facebook Pixel', 'Marketing', 'Medium', 'Two pixel IDs: 1243382036810144 and 2083683551969168. Domain verified (rr7yui3zls2ivix0bchkb44ofwrjgd).', 'Global (all pages)'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('5.2 Customer Service & Social'),
    createTable(
      ['Integration', 'Type', 'Complexity', 'Details', 'Pages Used'],
      [
        ['FPT.AI Live Chat', 'Chatbot / Embed', 'High', 'Full chatbot widget (stb-livechat.fpt.ai/v36) loaded as floating iframe. Persistent across site. Custom JS + CSS bundle.', 'Global (all pages)'],
        ['Facebook SDK / Share', 'Social Plugin', 'Low', 'Facebook sharer.php integration for article sharing. SDK loaded from connect.facebook.net.', 'News article pages'],
        ['Zalo Social Plugin', 'Social Plugin', 'Low', 'Zalo share button SDK (sp.zalo.me/plugins/sdk.js). Vietnam-specific social sharing.', 'News article pages'],
        ['Google reCAPTCHA Enterprise', 'Security / Embed', 'Medium', 'reCAPTCHA v2 Enterprise checkbox. Note: "exceeding free tier" warning observed.', 'Promotions newsletter, Request lookup page'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('5.3 Financial Data & Business'),
    createTable(
      ['Integration', 'Type', 'Complexity', 'Details', 'Pages Used'],
      [
        ['FiinGroup Stock Data', 'Embed / iframe', 'High', 'STB stock ticker, trading charts, financial data via iframes from stb-embed.fiingroup.vn. Interactive time-period selectors.', 'sacombank.com.vn/trang-chu/nha-dau-tu.html'],
        ['Exchange Rate API', 'API / Custom', 'High', 'Real-time exchange rate data feed (likely internal API). Updated every few minutes. Powers both rate table and currency converter.', 'sacombank.com.vn/cong-cu/ty-gia.html'],
        ['Interest Rate Data', 'API / Custom', 'Medium', 'Deposit and lending rate data. Powers rate tables and calculator widgets.', 'sacombank.com.vn/cong-cu/lai-suat.html'],
        ['Internet Banking (iSacombank)', 'External Link', 'Low', 'Links to isacombank.com.vn for online banking. Separate application, not embedded.', 'Header (all pages)'],
        ['MISA Lending', 'Partner API', 'Medium', 'Fintech lending partnership mentioned in enterprise credit section. Integration details unclear.', 'sacombank.com.vn/doanh-nghiep/tin-dung.html'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('5.4 Infrastructure & Content'),
    createTable(
      ['Integration', 'Type', 'Complexity', 'Details', 'Pages Used'],
      [
        ['Adobe Experience Manager', 'CMS Platform', 'High', 'Current CMS. AEM components (data-cmp-is attributes), clientlibs, Core Components, Adaptive Forms, CSRF protection.', 'All pages (infrastructure)'],
        ['Google Maps', 'Embed / iframe', 'Low', 'Embedded map showing headquarters location.', 'sacombank.com.vn/trang-chu/lien-he.html'],
        ['Mozilla PDF.js', 'Library / Plugin', 'Low', 'Client-side PDF rendering library (mozilla.github.io/pdf.js) for inline document viewing.', 'Document pages'],
        ['Sacombank Career Portal', 'External Link', 'Low', 'Links to sacombankcareer.com for recruitment. Separate application.', 'Footer (all pages)'],
        ['Loyalty Program Portal', 'External Link', 'Low', 'Links to khachhangthanthiet.sacombank.com for loyalty program.', 'Footer (all pages)'],
        ['E-Invoice Portal', 'External Link', 'Low', 'Links to hoadondientu.sacombank.com for electronic invoices.', 'Footer (all pages)'],
        ['Quick Registration Portal', 'External Link', 'Low', 'Links to khachhang.sacombank.com/dangkynhanh for customer onboarding.', 'Footer (all pages)'],
      ]
    ),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- 6. COMPLEX USE CASES ----
    heading('6. Complex Use Cases & Observations'),
    para('The following section identifies complex behaviors, edge cases, and functionality requiring special attention during migration.'),
    new Paragraph({ spacing: { after: 200 } }),

    createTable(
      ['#', 'Use Case', 'Instances', 'Where Found', 'Complexity', 'Why It Is Complex'],
      [
        ['1', 'Real-Time Exchange Rate Display', '1 page + footer widget', 'cong-cu/ty-gia.html', 'High', 'Requires live API integration for rate data updated every few minutes. Includes date/time picker for historical rates, PDF export, and currency converter calculator with live computation. Cannot be static content.'],
        ['2', 'Financial Calculators (5 types)', '2 pages (shared widget)', 'cong-cu/lai-suat.html, cong-cu/cong-cu-khac.html', 'High', 'Five interactive calculator types (deposit, loan, transaction fee, installment, insurance) with sliders, real-time computation, and dynamic output. Requires custom JavaScript implementation in EDS.'],
        ['3', 'Product Comparison Tool', '~4 listing pages', 'ca-nhan/the/the-tin-dung.html, bao-hiem/', 'High', 'Users select multiple products via checkboxes, sticky comparison bar tracks selections, "Start comparing" navigates to comparison page. Requires client-side state management and comparison page generation.'],
        ['4', 'Dual Analytics Stack (Google + Adobe)', 'All pages (global)', 'Global', 'Medium', 'Both Google Analytics (GTM + GA4 with 5 properties) and Adobe Analytics run simultaneously. Migration must preserve both tracking implementations or consolidate. 8+ Google Ads conversion tags add complexity.'],
        ['5', 'FiinGroup Stock Data Dashboard', '1 page', 'trang-chu/nha-dau-tu.html', 'High', 'Third-party FiinGroup iframes for stock price, trading volume, interactive charts with time-period selectors. Cannot be replicated — must maintain iframe embed approach or negotiate API access.'],
        ['6', 'AEM Adaptive Forms', '~15+ pages', 'Product pages, contact page', 'Medium', 'Lead capture forms use AEM Adaptive Forms infrastructure. Must be rebuilt as EDS form blocks or integrated with a form service. Two variants: Personal (CCCD field) and Enterprise (Tax ID field).'],
        ['7', 'Dynamic Product Filtering (AJAX)', '~4 listing pages', 'the/the-tin-dung.html, vay/', 'High', 'Product cards loaded dynamically via AJAX API calls. Filter checkboxes, category tabs, and "Load more" pagination all trigger API requests. Requires custom fetch logic in EDS.'],
        ['8', 'Product Suggestion Wizard', '~2 pages', 'Goi-y-san-pham/', 'High', 'Interactive multi-step questionnaire driven entirely by client-side JavaScript. Minimal server-rendered HTML. Requires full rebuild of wizard logic, question flow, and recommendation engine.'],
        ['9', 'FPT.AI Chatbot Integration', 'All pages (global)', 'Global', 'Medium', 'Third-party chatbot widget loads custom JS/CSS bundle and creates persistent iframe. Must ensure script loading order and positioning work in EDS. May conflict with EDS optimization strategies.'],
        ['10', 'Bilingual / Segment-Based Routing', 'All pages', 'Global', 'Medium', 'Site routes differ by customer segment (ca-nhan vs. doanh-nghiep). Header switcher changes context. URL structure and navigation trees differ per segment. Must maintain dual content hierarchies.'],
        ['11', 'PDF Generation / Export', '~3 pages', 'cong-cu/ty-gia.html, bieu-phi.html', 'Medium', 'Exchange rate and fee schedule pages offer PDF export/download functionality. PDF.js library for inline viewing. May need server-side PDF generation or client-side library in EDS.'],
        ['12', 'Cookie Consent Management', 'All pages (global)', 'Global', 'Low', 'Custom cookie consent implementation (not a standard CMP). Accept/Reject buttons control cookie behavior. Must reimplement or integrate standard CMP solution.'],
        ['13', 'News Archive (13 years)', '~1,271 pages', 'trang-chu/tin-tuc/', 'Medium', 'Large content volume spanning 2014-2026. Year-based URL structure. Bulk migration feasible but requires careful URL mapping and redirect strategy for SEO preservation.'],
        ['14', 'reCAPTCHA Enterprise (Over Quota)', '~2 pages', 'khuyen-mai, tra-cuu-yeu-cau', 'Low', 'reCAPTCHA Enterprise shows "exceeding free tier" warning. Current implementation may have billing/quota issues that need resolution during migration.'],
      ]
    ),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- 7. MIGRATION ESTIMATES ----
    heading('7. Migration Estimates'),
    para('The following estimates assume a team of 2-3 AEM EDS developers, 1 content migration specialist, and 1 QA engineer. Estimates are in working days (8 hours/day).'),
    new Paragraph({ spacing: { after: 200 } }),

    subsubheading('7.1 Phase Breakdown'),
    createTable(
      ['Phase', 'Scope', 'Effort (Days)', 'Resources', 'Approach'],
      [
        ['1. Discovery & Setup', 'Environment setup, design token extraction, global CSS, header/footer blocks, EDS project scaffolding', '8-10', '2 developers', 'Manual'],
        ['2. Template Development', 'Build 14 page templates in EDS (blocks, JS, CSS for each template type)', '25-35', '2-3 developers', 'Manual development'],
        ['3. Block Development', 'Implement 35 reusable blocks/components including interactive widgets (calculators, rate tables, comparison tool)', '20-30', '2-3 developers', 'Manual development'],
        ['4. Integration Migration', 'Migrate 16+ third-party integrations (analytics, chat, social, stock data, reCAPTCHA, forms)', '8-12', '1-2 developers', 'Manual + configuration'],
        ['5. Content Migration - Automated', 'Bulk migrate ~1,841 standardized pages (product details, news articles, promotions)', '8-10', '1 migration specialist + 1 developer', 'Automated (import scripts)'],
        ['6. Content Migration - Manual', 'Migrate ~105 complex/unique pages (homepages, tools, IR dashboard, wizards)', '10-15', '1-2 developers + 1 content specialist', 'Manual'],
        ['7. URL Redirect Mapping', 'Create 301 redirect rules for ~2,000 pages. SEO audit and canonical URL setup.', '3-5', '1 developer + 1 SEO specialist', 'Semi-automated'],
        ['8. QA & Testing', 'Cross-browser testing, responsive testing, accessibility audit, performance testing, visual regression, integration verification', '15-20', '1-2 QA engineers', 'Manual + automated tests'],
        ['9. UAT & Stakeholder Review', 'User acceptance testing, content review, stakeholder sign-off', '5-8', 'Client team + support', 'Manual review'],
        ['10. Go-Live & Hypercare', 'DNS cutover, monitoring, bug fixes, performance tuning', '3-5', 'Full team', 'Managed rollout'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('7.2 Effort Summary'),
    createTable(
      ['Category', 'Effort Range (Days)', 'Effort Range (Hours)'],
      [
        ['Automated Migration (scripted content import)', '8-10 days', '64-80 hours'],
        ['Manual / Custom Migration (templates, blocks, integrations, complex pages)', '71-102 days', '568-816 hours'],
        ['QA & Testing', '15-20 days', '120-160 hours'],
        ['UAT & Go-Live', '8-13 days', '64-104 hours'],
        ['Total Estimated Effort', '102-145 days', '816-1,160 hours'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('7.3 Timeline Estimate'),
    para('With a team of 3-4 dedicated resources working in parallel:'),
    new Paragraph({ spacing: { after: 100 } }),
    createTable(
      ['Scenario', 'Calendar Duration', 'Team Size', 'Notes'],
      [
        ['Aggressive', '3.5-4 months', '4-5 resources', 'Parallel development tracks; assumes minimal client review cycles; higher risk'],
        ['Recommended', '4.5-5.5 months', '3-4 resources', 'Balanced pace with adequate QA and review cycles; buffer for unknowns'],
        ['Conservative', '6-7 months', '2-3 resources', 'Sequential development; thorough testing; ample stakeholder review time'],
      ]
    ),

    new Paragraph({ spacing: { after: 200 } }),
    subsubheading('7.4 Risk Factors & Assumptions'),
    boldPara('Key Risks:'),
    bulletPoint('Real-time data integrations (exchange rates, stock data) may require backend API development if current endpoints are not accessible from EDS'),
    bulletPoint('FPT.AI chatbot compatibility with EDS script loading and performance optimization'),
    bulletPoint('Dual analytics stack (Google + Adobe) migration may require coordination with separate teams managing each platform'),
    bulletPoint('Product comparison tool and suggestion wizard require significant custom JavaScript development'),
    bulletPoint('13 years of news content (~1,271 articles) needs careful URL redirect mapping for SEO preservation'),
    bulletPoint('AEM Adaptive Forms must be replaced with EDS-compatible form solution'),
    new Paragraph({ spacing: { after: 100 } }),
    boldPara('Assumptions:'),
    bulletPoint('Client provides access to all API endpoints and data sources currently used by the site'),
    bulletPoint('Third-party vendors (FiinGroup, FPT.AI) will cooperate with integration migration'),
    bulletPoint('Content freeze period will be agreed upon for final migration cutover'),
    bulletPoint('Existing AEM DAM assets (images, PDFs) will be migrated to new hosting'),
    bulletPoint('Vietnamese language content only (no multilingual requirement beyond current scope)'),

    new Paragraph({ children: [new PageBreak()] }),

    // ---- APPENDIX ----
    heading('Appendix: Visual References'),
    para('The following screenshots provide visual reference for the key page templates and components identified in this analysis.'),
    new Paragraph({ spacing: { after: 200 } }),

    ...screenshotParagraph('sacombank-homepage.png', 'A1: Personal Banking Homepage (Segment Landing Page template)'),
    ...screenshotParagraph('page01-enterprise.png', 'A2: Enterprise Banking Homepage (Segment Landing Page template)'),
    ...screenshotParagraph('page02-personal-account.png', 'A3: Product Category Landing - Personal Accounts'),
    ...screenshotParagraph('sacombank-credit-cards.png', 'A4: Filterable Product Listing - Credit Cards with comparison'),
    ...screenshotParagraph('page05-savings.png', 'A5: Product Category Landing - Savings (with tabs)'),
    ...screenshotParagraph('page14-product-detail.png', 'A6: Product Detail Page - Payment Account'),
    ...screenshotParagraph('page06-news-listing.png', 'A7: News/Article Listing'),
    ...screenshotParagraph('sacombank-news-article.png', 'A8: News Article Detail with social sharing sidebar'),
    ...screenshotParagraph('page08-contact.png', 'A9: Contact Page with tabs and Google Maps'),
    ...screenshotParagraph('sacombank-exchange-rate.png', 'A10: Utility/Tools - Exchange Rate Table with Currency Converter'),
    ...screenshotParagraph('page09-tools.png', 'A11: Utility/Tools - Forms/Documents listing'),
    ...screenshotParagraph('page10-investor-reports.png', 'A12: Document Repository - Investor Reports'),
  ]
});

// ---- BUILD DOCUMENT ----
const doc = new Document({
  creator: 'Site Analysis Tool',
  title: 'Sacombank Website Migration Analysis Report',
  description: 'Comprehensive analysis of www.sacombank.com.vn for AEM Edge Delivery Services migration',
  styles: {
    paragraphStyles: [
      {
        id: 'Heading1',
        name: 'Heading 1',
        run: { size: 36, bold: true, color: '003478', font: 'Calibri' },
        paragraph: { spacing: { before: 360, after: 200 } },
      },
      {
        id: 'Heading2',
        name: 'Heading 2',
        run: { size: 28, bold: true, color: '003478', font: 'Calibri' },
        paragraph: { spacing: { before: 300, after: 160 } },
      },
      {
        id: 'Heading3',
        name: 'Heading 3',
        run: { size: 24, bold: true, color: '003478', font: 'Calibri' },
        paragraph: { spacing: { before: 240, after: 120 } },
      },
    ],
    default: {
      document: {
        run: { size: 22, font: 'Calibri' },
        paragraph: { spacing: { line: 276 } },
      },
    },
  },
  sections,
});

const buffer = await Packer.toBuffer(doc);
fs.writeFileSync('/workspace/Sacombank_Migration_Analysis_Report.docx', buffer);
console.log('Report generated: /workspace/Sacombank_Migration_Analysis_Report.docx');
console.log(`File size: ${(buffer.length / 1024).toFixed(1)} KB`);
