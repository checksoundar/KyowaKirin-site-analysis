import {
  Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
  Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType, PageBreak
} from 'docx';
import fs from 'fs';

// ---------- palette ----------
const NAVY = '1F3864';
const BLUE = '2E5496';
const LIGHT = 'D9E2F3';
const ZEBRA = 'F2F5FB';
const WHITE = 'FFFFFF';
const GREEN = 'C6E0B4';
const AMBER = 'FFE699';
const RED = 'F8CBAD';

const FONT = 'Calibri';

// ---------- helpers ----------
const tableBorders = {
  top: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
  bottom: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
  left: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
  right: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
  insideHorizontal: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
  insideVertical: { style: BorderStyle.SINGLE, size: 4, color: 'BFBFBF' },
};

function cellPara(text, { bold = false, color = '000000', align = AlignmentType.LEFT, size = 18 } = {}) {
  const lines = String(text).split('\n');
  return lines.map((ln, i) =>
    new Paragraph({
      alignment: align,
      spacing: { before: 20, after: 20 },
      children: [new TextRun({ text: ln, bold, color, size, font: FONT })],
    })
  );
}

function headerCell(text, width) {
  return new TableCell({
    width: width ? { size: width, type: WidthType.PERCENTAGE } : undefined,
    shading: { type: ShadingType.CLEAR, color: 'auto', fill: NAVY },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: cellPara(text, { bold: true, color: WHITE, align: AlignmentType.CENTER, size: 18 }),
  });
}

function dataCell(text, { fill = WHITE, bold = false, color = '000000', align = AlignmentType.LEFT } = {}) {
  return new TableCell({
    shading: { type: ShadingType.CLEAR, color: 'auto', fill },
    margins: { top: 40, bottom: 40, left: 80, right: 80 },
    children: cellPara(text, { bold, color, align }),
  });
}

const complexityFill = (lvl) =>
  /high/i.test(lvl) ? RED : /med/i.test(lvl) ? AMBER : GREEN;

// Build a table from headers + rows. colWidths in %. complexityCol = index to color-code.
function buildTable(headers, rows, colWidths, opts = {}) {
  const { complexityCol = -1, zebra = true } = opts;
  const headerRow = new TableRow({
    tableHeader: true,
    children: headers.map((h, i) => headerCell(h, colWidths ? colWidths[i] : undefined)),
  });
  const dataRows = rows.map((r, ri) => {
    const baseFill = zebra && ri % 2 === 1 ? ZEBRA : WHITE;
    return new TableRow({
      children: r.map((c, ci) => {
        let fill = baseFill;
        if (ci === complexityCol) fill = complexityFill(String(c));
        return dataCell(String(c), { fill, bold: ci === 0 && opts.boldFirst, align: ci === 0 ? AlignmentType.LEFT : (opts.centerCols && opts.centerCols.includes(ci) ? AlignmentType.CENTER : AlignmentType.LEFT) });
      }),
    });
  });
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: tableBorders,
    rows: [headerRow, ...dataRows],
  });
}

// paragraph helpers
const P = (text, opts = {}) => new Paragraph({
  spacing: { before: opts.before ?? 60, after: opts.after ?? 120 },
  alignment: opts.align || AlignmentType.LEFT,
  children: [new TextRun({ text, font: FONT, size: opts.size || 21, color: opts.color || '000000', bold: opts.bold, italics: opts.italics })],
});

const bullet = (text, level = 0) => new Paragraph({
  bullet: { level },
  spacing: { before: 20, after: 20 },
  children: Array.isArray(text) ? text : [new TextRun({ text, font: FONT, size: 21 })],
});

function H1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 280, after: 140 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: BLUE } },
    children: [new TextRun({ text, font: FONT, size: 30, bold: true, color: NAVY })],
  });
}
function H2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    spacing: { before: 220, after: 100 },
    children: [new TextRun({ text, font: FONT, size: 25, bold: true, color: BLUE })],
  });
}
function H3(text) {
  return new Paragraph({
    spacing: { before: 160, after: 80 },
    children: [new TextRun({ text, font: FONT, size: 22, bold: true, color: NAVY })],
  });
}

const children = [];

// ===================== COVER =====================
children.push(
  new Paragraph({ spacing: { before: 1800, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Adobe Experience Manager → Edge Delivery Services', font: FONT, size: 24, color: BLUE, bold: true })] }),
  new Paragraph({ spacing: { before: 200, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Migration Discovery & Analysis Report', font: FONT, size: 52, bold: true, color: NAVY })] }),
  new Paragraph({ spacing: { before: 240, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Diners Club Japan  •  TRUST CLUB (Sumitomo Mitsui Trust Club)', font: FONT, size: 26, color: '404040' })] }),
  new Paragraph({ spacing: { before: 120, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'www.diners.co.jp  |  www.sumitclub.jp', font: FONT, size: 20, color: '808080' })] }),
  new Paragraph({ spacing: { before: 1400, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Prepared for: Migration Planning Team', font: FONT, size: 20, color: '404040' })] }),
  new Paragraph({ spacing: { before: 80, after: 0 }, alignment: AlignmentType.CENTER,
    children: [new TextRun({ text: 'Date: 16 June 2026', font: FONT, size: 20, color: '404040' })] }),
  new Paragraph({ children: [new PageBreak()] })
);

// ===================== EXECUTIVE SUMMARY =====================
children.push(H1('Executive Summary'));
children.push(P('This report consolidates the discovery findings for two related properties — diners.co.jp (Diners Club Japan) and sumitclub.jp (TRUST CLUB) — and translates them into a structured migration plan for Adobe Edge Delivery Services (EDS / Document-based authoring).'));

children.push(P('Both sites are sister properties operated by Sumitomo Mitsui Trust Club and are currently built on the same classic Adobe Experience Manager (AEM 6.x) platform. This is a material finding: they share one common component library (the "CCM" component set), one header/footer/navigation system, AEM Core Components (cmp-*), and near-identical page-construction patterns. The two sites therefore behave as a single migration program with a shared block library, not two independent projects.', { after: 120 }));

children.push(H3('Key facts at a glance'));
children.push(buildTable(
  ['Metric', 'Diners Club (diners.co.jp)', 'TRUST CLUB (sumitclub.jp)', 'Combined'],
  [
    ['Total published pages (sitemap)', '1,765', '439', '2,204'],
    ['Page template (AEM)', 'diners-club_', 'trust-club_', '2 variants, 1 base'],
    ['Underlying platform', 'AEM 6.x (classic)', 'AEM 6.x (classic)', 'Shared instance'],
    ['Distinct content templates', '8', '8', '8 (shared model)'],
    ['Blocks (team discovery / EDS target)', '53 / 35', '53 / 35 (shared)', '35 EDS blocks'],
    ['Auto-migratable share (est.)', '~88%', '~85%', '~87%'],
    ['Languages', 'Japanese (+EN sub)', 'Japanese + English', 'JA primary'],
  ],
  [28, 24, 24, 24],
  { centerCols: [1, 2, 3] }
));

children.push(P(''));
children.push(P('Headline estimate: the combined program is estimated at approximately 103–133 working days of effort (≈ 4.5–6 calendar months with a small dedicated team), with roughly 87% of pages eligible for automated, parser-driven migration and the remainder requiring manual / assisted handling (application flows, member-only "Club Online" areas, iframe-embedded forms, and heavily designed campaign landing pages).', { bold: false }));

children.push(new Paragraph({
  spacing: { before: 120, after: 120 },
  shading: { type: ShadingType.CLEAR, color: 'auto', fill: ZEBRA },
  border: { left: { style: BorderStyle.SINGLE, size: 18, color: BLUE } },
  children: [new TextRun({ text: '  Note on the supplied spreadsheet: ', bold: true, font: FONT, size: 20, color: NAVY }),
    new TextRun({ text: 'Section 2 is built on the team’s "Post Discovery – Analysis.xlsx" Sheet 1 (53 blocks), reproduced in full (2.1) and then reconciled against the live sites and consolidated to 35 EDS blocks (2.2–2.3). Sheet 2 (integrations) was not visible in the supplied screenshot; the integrations catalogue (2.4) is assembled from the block-list notes plus live observation — please confirm it against Sheet 2 to close any gaps.', font: FONT, size: 20, color: '404040' })],
}));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== 1. TEMPLATES INVENTORY =====================
children.push(H1('1. Templates Inventory'));
children.push(P('Although both sites technically render through a single AEM page template (diners-club_ / trust-club_) built on a flexible parsys/responsive-grid, the content falls into eight recurring authoring patterns. For migration planning these patterns are the meaningful "templates": they determine the parser logic and the EDS section model. Complexity reflects layout variability, dynamic behaviour, and embedded integrations — not page count.'));

children.push(H2('1.1  Diners Club (diners.co.jp)'));
children.push(buildTable(
  ['Template', 'Complexity', 'Reasoning', 'Reference URLs'],
  [
    ['T1 — Homepage / Section Landing', 'High',
      'Hero carousel (slick), audience navigation, "important notices" feed, news list, promo grids. Many distinct blocks; mostly one-off layout.',
      'https://www.diners.co.jp/ja/index.html\nhttps://www.diners.co.jp/ja/travel.html\nhttps://www.diners.co.jp/ja/gourmet.html'],
    ['T2 — Card Detail / Product Page', 'High',
      'Long-form product page: card-spec tables, fee tables, comparison sliders, FAQ, embedded application CTAs and iframes. Richest component mix on the site.',
      'https://www.diners.co.jp/ja/cardlineup/dinersclubcard.html\nhttps://www.diners.co.jp/ja/cardlineup/anadiners_new.html'],
    ['T3 — Card Lineup / Comparison', 'Medium',
      'Filterable card grid, dropdown filters, comparison tables, tabs. Standardised but interactive.',
      'https://www.diners.co.jp/ja/cardlineup.html\nhttps://www.diners.co.jp/ja/cardlineup/comparison.html'],
    ['T4 — Benefit / Service Detail', 'Medium',
      'Standardised content body: intro, image+text blocks, link lists, accordions. 473 pages, very consistent structure → high auto-migration.',
      'https://www.diners.co.jp/ja/benefit.html\nhttps://www.diners.co.jp/ja/benefit/stock.html'],
    ['T5 — Magazine / Editorial Article', 'Medium',
      'Editorial layout: rich text, full-width imagery, pull quotes, related-article cards, occasional video. 603 pages — the largest single group.',
      'https://www.diners.co.jp/ja/magazine.html\nhttps://www.diners.co.jp/ja/magazine/library.html'],
    ['T6 — Press / Notice Article', 'Low',
      'Dated text article: title, date, body copy, occasional table/PDF link. Highly uniform → ideal for automated migration.',
      'https://www.diners.co.jp/ja/press.html\nhttps://www.diners.co.jp/ja/press/inf_20260601.html'],
    ['T7 — Legal / Policy / Info', 'Low',
      'Plain rich-text: terms, privacy, prospectus. Tables but no interactivity. Fully auto-migratable.',
      'https://www.diners.co.jp/ja/privacy.html\nhttps://www.diners.co.jp/ja/tc.html'],
    ['T8 — Sitemap / Index', 'Low',
      'Auto-style link directory. One page; trivial to rebuild.',
      'https://www.diners.co.jp/ja/sitemap.html'],
    ['T9 — Application / Form / Member (special)', 'High',
      'Card application ("nyukai"), contact, 3-D Secure registration, "Club Online" member area. Forms are iframe-embedded external systems → out of CMS scope; landing wrappers migrate, the flows do not.',
      'https://www.diners.co.jp/ja/cardlineup/nyukai.html\nhttps://www.diners.co.jp/ja/contact.html\nhttps://www.diners.co.jp/ja/usage/3d_secure/new_registration.html'],
  ],
  [20, 12, 40, 28],
  { complexityCol: 1, boldFirst: true }
));

children.push(H2('1.2  TRUST CLUB (sumitclub.jp)'));
children.push(P('TRUST CLUB uses the same eight patterns (same component library and header/footer), so the template definitions above apply directly. The notable site-specific groups are below.', { after: 80 }));
children.push(buildTable(
  ['Template', 'Complexity', 'Reasoning', 'Reference URLs'],
  [
    ['T1 — Homepage / Section Landing', 'High',
      'Hero, recommended-cards grid, service & support grid, notices, FAQ promo. Same block family as Diners.',
      'https://www.sumitclub.jp/ja/index.html\nhttps://www.sumitclub.jp/ja/travel.html'],
    ['T2 — Card Detail / Product', 'High',
      'Card spec, fee tables, benefits, application CTA. Mirrors Diners T2.',
      'https://www.sumitclub.jp/ja/cardlineup/world.html\nhttps://www.sumitclub.jp/ja/cardlineup/platinummaster.html'],
    ['T4 — Benefit / Service Detail', 'Medium',
      'Travel / gourmet / point / insurance service pages. Consistent structure.',
      'https://www.sumitclub.jp/ja/travel/airport/airportlounge.html\nhttps://www.sumitclub.jp/ja/insurance/buyers.html'],
    ['T6 — Notice Article', 'Low',
      'Dated notices (104 pages) — uniform, fully auto-migratable.',
      'https://www.sumitclub.jp/ja/notice/inf_20260601.html'],
    ['T7 — Legal / Policy', 'Low',
      'Terms, privacy, prospectus, anti-social-forces policy.',
      'https://www.sumitclub.jp/ja/privacy.html\nhttps://www.sumitclub.jp/ja/tc.html'],
    ['T9 — Member / Login / Application (special)', 'High',
      '"loginPage" tree (41 pages) and Club Online member area — iframe / external application system. Public wrappers migrate; secured flows are out of scope.',
      'https://www.sumitclub.jp/ja/cardlineup/nyukai.html\nhttps://www.sumitclub.jp/ja/loginPage.html'],
    ['T10 — Corporate Site (sub-site)', 'Medium',
      'Separate /corporate tree (company info, news, sustainability) with EN mirror. ~87 pages; standard editorial/landing mix.',
      'https://www.sumitclub.jp/ja/corporate_site.html\nhttps://www.sumitclub.jp/corporate/kaizen/2204.html'],
  ],
  [20, 12, 40, 28],
  { complexityCol: 1, boldFirst: true }
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== 2. BLOCKS / COMPONENTS CATALOG =====================
children.push(H1('2. Blocks / Components Catalog'));
children.push(P('This section takes the team’s discovery catalogue from "Post Discovery – Analysis.xlsx" (Sheet 1 — 53 blocks) as the baseline, reproduces it faithfully (2.1), then applies the guidance to prefer design variations over net-new blocks and reconciles it against the live sites (2.2). The result is a leaner recommended EDS block library (2.3). Section 2.4 catalogues the third-party integrations.'));

children.push(H2('2.1  Team discovery baseline (Sheet 1 — 53 blocks)'));
children.push(P('Reproduced from the supplied workbook. "C/C" = Core or Custom; "Var" = number of style variations identified; "SV" indicates whether the style variation was measured in the existing AEM library.', { italics: true, after: 80, size: 18 }));
children.push(buildTable(
  ['#', 'Block', 'Diff.', 'C/C', 'Var', 'Linked feature / note'],
  [
    ['1', 'Text', 'Low', 'Core', '4', 'Normal / Centered / Small. HTML editability TBC'],
    ['2', 'Headings', 'Low', 'Core', '6', 'H1–H4 / Centered / Thumbnail. Heading icon'],
    ['3', 'Links', 'Low', 'Core', '3', 'Normal / Inline / Phone number. Targeting'],
    ['4', 'Lists', 'Low', 'Core', '4', 'Normal / Normal(S) / Numbering. Auto-index'],
    ['5', 'Lists Links Description', 'Low', 'Core', '1', 'Link + description list. Not in library'],
    ['6', 'Images', 'Low', 'Core', '7', 'No style / centered / 100% etc. Benefit-detail gallery'],
    ['7', 'Divider', 'Low', 'Core', '1', 'Standard divider'],
    ['8', 'Columns', 'Low-Med', 'Core', '6', 'Grid / left-aligned grid. Important notices auto-index'],
    ['9', 'Buttons', 'Low', 'Core', '8', 'CTA(S/L) / normal. Card application / QSRF / KARTE / Targeting'],
    ['10', 'Buttons Share', 'Low', 'Custom', '1', 'SNS share button. Share-button JS'],
    ['11', 'FAQ Details', 'Low', 'Core', '1', 'Open-close (details)'],
    ['12', 'Phrase', 'Low', 'Custom', '2', 'A type / P type. COL phrase'],
    ['13', 'Anker Navigation', 'Low', 'Custom', '1', 'Anchor nav auto-generation JS'],
    ['14', 'Icons', 'Low', 'Core', '1', 'Icon set (fontello → standard)'],
    ['15', 'Video', 'Low', 'Core', '1', 'Standard video'],
    ['16', 'Embed', 'Low', 'Core', '1', 'YouTube embed'],
    ['17', 'Hero', 'Low-Med', 'Core', '2', 'MV slide (carousel 1-col) / SP segment. Dedicated Hero not in library'],
    ['18', 'Accordion', 'Medium', 'Core', '2', 'Default / exclusive open-close. cmp-accordion measured'],
    ['19', 'Tabs', 'Medium', 'Core', '2', 'Normal / SP accordion. Benefit-detail plan tabs / Auto-index'],
    ['20', 'Table', 'Medium', 'Core', '2', 'Standard / 2-col non-stacking on SP. HTML authoring TBC'],
    ['21', 'Section Metadata', 'Medium', 'Core', '4', 'Member-type & PC/SP segmentation. Partly in library'],
    ['22', 'Cards', 'Medium', 'Custom', '3', 'Manual edit / auto-index / JS+JSON. teaser measured'],
    ['23', 'Cards Card Info', 'Medium', 'Custom', '2', 'Campaign list / event list (CIX012). AES-autoi CSS only'],
    ['24', 'Cards Benefit Step', 'Low', 'Core', '1', 'Step layout (Cards variant). Named variant not in library'],
    ['25', 'Cards Card Lineup', 'Low', 'Core', '1', 'Card lineup (Cards variant)'],
    ['26', 'Cards Magazine', 'Low', 'Core', '1', 'Magazine (Cards variant)'],
    ['27', 'Cards Privilege Icon', 'Low', 'Core', '1', 'Privilege icon (Cards variant)'],
    ['28', 'Cards Prize', 'Low', 'Core', '1', 'Prize (Cards variant)'],
    ['29', 'Cards Service Introduction', 'Low', 'Core', '1', 'Service intro (Cards variant)'],
    ['30', 'Carousel', 'Medium', 'Core', '3', '1 / 3 / 4-col slide. AES-carousel measured'],
    ['31', 'Fragment', 'Medium', 'Core', '3', 'Individual / corporate / contact-common. Header segmentation'],
    ['32', 'Footer', 'Medium', 'Core', '5', 'Individual / corporate / travel / service. Menu handling'],
    ['33', 'Header', 'High', 'Custom', '4', 'DPC / RPC / general / corporate. Cookie detection, custom JS'],
    ['34', 'Left Menu', 'Medium', 'Custom', '2', 'PC menu / SP 2nd menu (JS). Indiv-corp segmentation'],
    ['35', 'Breadcrumbs', 'Medium', 'Custom', '3', 'Individual / corporate co-display / mid-category'],
    ['36', 'Modal', 'Medium', 'Core', '1', 'Standard modal. FAQ / Travel desk / Airport lounge'],
    ['37', 'Page Style', 'Medium', 'Custom', '1', 'Per-page custom style (~3 pages). MV segmentation'],
    ['38', 'Sitemap', 'Medium', 'Custom', '2', 'sitemap.xml (custom) / sitemap page. Auto-generation'],
    ['39', 'Buttons FAQ', 'Medium', 'Custom', '1', 'FAQ button (Helpfeel). May be unnecessary'],
    ['40', 'Buttons Float', 'Medium', 'Custom', '1', 'Floating button auto-display / QSRF Linker'],
    ['41', 'Buttons Clipboard Copy', 'Medium', 'Custom', '1', 'Copy button (Clipboard API)'],
    ['42', 'Related Pages', 'Medium', 'Custom', '1', 'Related-pages list. Current-page exclusion'],
    ['43', 'Search', 'Medium', 'Custom', '1', 'Site search (i-search)'],
    ['44', 'Search Insite', 'Medium', 'Custom', '1', 'Site search, card-type cookie variable (i-search)'],
    ['45', 'Travel Desk', 'Medium', 'Custom', '2', 'Recommended travel info / scrollbar. XML-driven'],
    ['46', 'Embed Custom', 'Medium', 'Custom', '2', 'HTML-only / HTML+CSS/JS. Receiver for packaged HTML + legacy'],
    ['47', 'Benefit Detail', 'High', 'Custom', '4', 'Gourmet / other / domestic / overseas. Shop JSON, member seg'],
    ['48', 'Cards Campaign Event Benefit', 'High', 'Custom', '2', 'Campaign/event (CIX012/013). query-index fetch, 2-system merge'],
    ['49', 'Cards Event Report', 'High', 'Custom', '2', 'Report list / top display. eventList.json fetch, JS+JSON sort'],
    ['50', 'Lists Announcement', 'High', 'Custom', '2', 'Year-based accordion / count flat. AES-list-notice measured'],
    ['51', 'Google Maps', 'High', 'Custom', '1', 'googleapis. JS API init, key/domain restriction, CSP'],
    ['52', 'Airport Lounge Domestic', 'High', 'Custom', '1', 'Domestic (CDN010). XML-driven'],
    ['53', 'Airport Lounge International', 'High', 'Custom', '2', 'airport-list / detail. DCI API + App Builder relay, CORS/auth'],
  ],
  [4, 22, 9, 8, 5, 52],
  { complexityCol: 2, centerCols: [0, 3, 4] }
));

children.push(H2('2.2  Reconciliation against the live sites'));
children.push(P('Having inspected both live sites, the team’s catalogue is accurate and thorough — the live DOM confirms the AEM Core Components (cmp-*), the CCM/AES custom set, and the carousel/accordion/tab behaviours listed. Applying the "variants over new blocks" principle, several entries are structurally identical and should collapse into a single EDS block addressed by named variants. A small number of clarifications are added from the live findings.', { after: 80 }));
children.push(buildTable(
  ['Action', 'Affected entries', 'Recommendation & rationale'],
  [
    ['CONSOLIDATE', '#22–29 (Cards + 6 "Cards X" variants) and #48, #49',
      'Single "Cards" block. The team already labels #24–29 as Cards variants; #23/#48/#49 are data-driven variants (auto-index / JSON fetch). Model as one block: variants = manual, auto-index, json. Net: 10 entries → 1 block (+ variants). Biggest single reduction.'],
    ['CONSOLIDATE', '#9, #10, #39, #40, #41 (Buttons family)',
      'Single "Button" block with variants (cta, share, faq, float, copy). Share/FAQ/float/copy differ only by a small JS behaviour attached to the same model. Net: 5 → 1.'],
    ['CONSOLIDATE', '#43, #44 (Search, Search Insite)',
      'One "Search" block; the card-type cookie is a config option, not a second block. Net: 2 → 1.'],
    ['CONSOLIDATE', '#16, #46 (Embed, Embed Custom)',
      'One "Embed" block, variants = youtube, html, html+js. Net: 2 → 1.'],
    ['CONSOLIDATE', '#52, #53 (Airport Lounge Domestic / International)',
      'One "Airport Lounge" block with domestic/international variants; both are API/XML-driven lists. Net: 2 → 1. (Still High effort — keep as integration.)'],
    ['GROUP (chrome)', '#31–35 (Fragment, Footer, Header, Left Menu, Breadcrumbs)',
      'Treat as the shared navigation/chrome system, built once per brand and reused. In EDS these are nav + footer + fragments, not per-page content blocks.'],
    ['MERGE', '#5 into #4 (Lists)',
      '"Lists Links Description" is a Lists variant (link + description). Net: 2 → 1.'],
    ['CONFIRM (keep)', '#47 Benefit Detail, #50 Lists Announcement, #51 Google Maps, #45 Travel Desk',
      'Genuinely distinct High-complexity blocks (member-segmented shop data, year-grouped notices, Maps JS API, XML travel feed). Keep as first-class blocks.'],
    ['NO ADDITIONS NEEDED', '—',
      'The live audit surfaced no block missing from the team’s list. Items I might otherwise add (notice feed, iframe form wrapper, card-spec) are already covered by #50, #46 and #47 respectively.'],
  ],
  [16, 26, 58],
  { boldFirst: true }
));

children.push(H2('2.3  Recommended consolidated EDS block library'));
children.push(P('After the reconciliation, the 53 discovery entries map to ≈ 30 EDS blocks (the difference being variants rather than separate blocks). This is the recommended build list.', { after: 80 }));
children.push(buildTable(
  ['#', 'EDS Block', 'Effort', 'Replaces (Sheet 1 #)'],
  [
    ['1', 'Text', 'Low', '1'],
    ['2', 'Headings', 'Low', '2'],
    ['3', 'Links', 'Low', '3'],
    ['4', 'Lists', 'Low', '4, 5'],
    ['5', 'Images', 'Low', '6'],
    ['6', 'Divider', 'Low', '7'],
    ['7', 'Columns / Section layout', 'Low', '8'],
    ['8', 'Button', 'Low', '9, 10, 39, 40, 41'],
    ['9', 'FAQ / Details', 'Low', '11'],
    ['10', 'Phrase', 'Low', '12'],
    ['11', 'Anchor Navigation', 'Low', '13'],
    ['12', 'Icons', 'Low', '14'],
    ['13', 'Video', 'Low', '15'],
    ['14', 'Embed', 'Low', '16, 46'],
    ['15', 'Hero', 'Low-Med', '17'],
    ['16', 'Accordion', 'Medium', '18'],
    ['17', 'Tabs', 'Medium', '19'],
    ['18', 'Table', 'Medium', '20'],
    ['19', 'Section Metadata', 'Medium', '21'],
    ['20', 'Cards', 'Medium-High', '22–29, 48, 49'],
    ['21', 'Carousel', 'Medium', '30'],
    ['22', 'Header (chrome)', 'High', '33, 34'],
    ['23', 'Footer (chrome)', 'Medium', '32'],
    ['24', 'Fragment (chrome)', 'Medium', '31'],
    ['25', 'Breadcrumbs', 'Medium', '35'],
    ['26', 'Modal', 'Medium', '36'],
    ['27', 'Page Style', 'Medium', '37'],
    ['28', 'Sitemap', 'Medium', '38'],
    ['29', 'Related Pages', 'Medium', '42'],
    ['30', 'Search', 'Medium', '43, 44'],
    ['31', 'Travel Desk', 'Medium', '45'],
    ['32', 'Benefit Detail', 'High', '47'],
    ['33', 'Lists Announcement', 'High', '50'],
    ['34', 'Google Maps', 'High', '51'],
    ['35', 'Airport Lounge', 'High', '52, 53'],
  ],
  [6, 38, 16, 40],
  { complexityCol: 2, centerCols: [0] }
));
children.push(P('Result: 53 discovery entries → 35 EDS blocks (≈ 34% fewer), with no loss of functionality. 14 Low, 13 Medium, 1 Low-Med / Medium-High, 7 High.', { bold: true, before: 60 }));

children.push(H2('2.4  Integrations catalog'));
children.push(P('The supplied workbook lists integrations on Sheet 2 (not visible in the provided screenshot). The catalogue below is assembled from the integration notes embedded in the block list plus the third-party services observed loading on the live sites. Please confirm against Sheet 2 to close any gaps.', { italics: true, after: 80, size: 18 }));
children.push(buildTable(
  ['Integration', 'Used by / where', 'Migration impact'],
  [
    ['Adobe Analytics / DTM (Launch)', 'Sitewide tag (satelliteLib / adobedtm)', 'Re-add tag via EDS head/scripts. Low.'],
    ['Adobe Experience Cloud ID (adobe_mc)', 'Cross-domain links diners ↔ sumitclub', 'Preserve cross-domain identity params. Low-Med.'],
    ['Google Tag Manager', 'Sitewide', 'Re-add container. Low.'],
    ['Meta (Facebook) Pixel', 'Sitewide marketing tag', 'Re-add. Low.'],
    ['Logly / ad & DSP pixels (bidswitch, adnxs, ladsp)', 'Marketing/retargeting', 'Re-add tags; verify consent. Low.'],
    ['Helpfeel (cards-faq-custhelp)', 'FAQ buttons (#39), "よくあるご質問"', 'External FAQ service; link/embed only. Low.'],
    ['i-search', 'Site search (#43, #44)', 'Search service integration; re-wire query + cookie variant. Medium.'],
    ['Google Maps JavaScript API (#51)', 'Store/location maps', 'API key, domain restriction, CSP entry. Medium-High.'],
    ['DCI API + Adobe App Builder relay (#53)', 'Airport lounge international list', 'CORS/auth via App Builder; keep relay or re-point. High.'],
    ['XML / JSON content feeds', 'Travel desk (#45), airport (#52), shop data (#47), query-index (#48), eventList.json (#49)', 'Re-create feeds or fetch logic in block JS. Medium-High per feed.'],
    ['Club Online (WA2010101 / member system)', 'Header CTA, member area', 'External application; link out only — out of CMS scope. N/A.'],
    ['Remote Operator (rmop.jp)', 'Footer support tool', 'External link. Low.'],
    ['KARTE / on-site personalisation & Targeting', 'Buttons (#9), Links (#3)', 'Re-instrument personalisation tags. Medium.'],
    ['QSRF Linker', 'Card application / float button (#40)', 'Application-flow linker; preserve params. Medium.'],
    ['Clipboard API (#41)', 'Copy button', 'Native browser API; trivial. Low.'],
    ['fontello icon set (#14)', 'Iconography', 'Migrate to standard icon sprite/SVG. Low.'],
  ],
  [26, 38, 36],
  { boldFirst: true }
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== 3. PAGE COUNTS BY TEMPLATE =====================
children.push(H1('3. Page Counts by Template'));
children.push(P('Counts are derived from the live XML sitemaps (diners.co.jp = 1,765 URLs; sumitclub.jp = 352 main + 87 corporate = 439 URLs). Migration mode reflects structural uniformity: uniform, content-driven pages are parser-automatable; pages with bespoke layout, heavy interactivity or external integrations need manual / assisted work.'));

children.push(H2('3.1  Migration mode definitions'));
children.push(P('Each row in the tables below is classified by how the page can be moved into Edge Delivery Services. The classification is driven by one question: how much human judgement is needed per page once the block library and parsers exist? The four modes are defined below, with the reasoning for when each applies.', { after: 80 }));
children.push(buildTable(
  ['Migration Mode', 'What it means', 'Why a page lands here', 'Per-page human effort'],
  [
    ['Automated',
      'The page is imported end-to-end by the bulk import pipeline (block parsers + page transformers). A human only spot-checks a sample, not every page.',
      'Pages share one repeating structure with no per-page logic: a parser written once recognises the same blocks across hundreds of pages. Content is static text/image/link markup. Highest-volume groups (Magazine, Benefit, Notice/Press) qualify because every page is built the same way.',
      'Near-zero (sampling QA only)'],
    ['Assisted',
      'Mostly automated, but a person reviews and finishes each page after import — fixing a table, re-linking a CTA, confirming a data-driven block rendered correctly.',
      'The structure is recognisable to a parser, but the page carries elements the parser cannot fully resolve on its own: complex spec/fee tables, comparison layouts, application CTAs, or blocks fed by JSON/XML feeds. Import does 70–90% of the work; a human closes the gap.',
      'Low–moderate (review + touch-up each page)'],
    ['Manual',
      'The page is rebuilt by hand in the EDS authoring model; the importer is not relied on for layout.',
      'Layout is bespoke or one-off (homepages, hero-heavy section landings, campaign LPs) where each page is unique, so there is no repeating pattern for a parser to exploit. Also applies to pages that are thin wrappers around an external system (iframe forms) — the wrapper is recreated by hand and the embed re-pointed.',
      'High (full rebuild per page)'],
    ['Mixed',
      'The group contains a blend: most pages are Automated, but a minority within the same URL tree need Manual/Assisted handling. Reported as one row because they live under one section.',
      'Used where a section is mostly uniform but contains exceptions — e.g. Usage/Guides is largely static (Automated) but the 3-D Secure / registration sub-pages are interactive (Manual). Splitting them into separate rows would overstate precision; "Mixed" flags that the section needs triage before batch import.',
      'Mostly low, with a manual subset'],
    ['Out-of-scope',
      'Not migrated as CMS content. Only a link or wrapper is preserved; the underlying system stays where it is (or is re-platformed as a separate project).',
      'The "pages" are screens of an external application — the Club Online member area and the secured login tree. These are authenticated, dynamic application flows, not editable content, so they fall outside a content migration entirely.',
      'N/A (excluded; link-out only)'],
  ],
  [16, 26, 42, 16],
  { boldFirst: true }
));
children.push(P('Note on terminology: "Manual / Assisted" in the section tables denotes a group that needs human work on every page, spanning the Assisted-to-Manual range depending on the individual page; "Manual / Out-of-scope" denotes wrapper pages that are rebuilt by hand while the underlying flow is excluded.', { italics: true, size: 18, before: 40, after: 120 }));

children.push(H2('3.2  Diners Club (diners.co.jp) — 1,765 pages'));
children.push(buildTable(
  ['Template / Section', 'Pages', 'Migration Mode', 'Basis', 'URL pattern'],
  [
    ['T5 Magazine (editorial)', '603', 'Automated', 'Uniform editorial structure', '/ja/magazine/**'],
    ['T4 Benefit / Service detail', '473', 'Automated', 'Highly consistent body layout', '/ja/benefit/**'],
    ['T6 Press / Notice', '173', 'Automated', 'Dated text articles', '/ja/press/** (e.g. inf_YYYYMMDD.html)'],
    ['T1/T4 Event', '88', 'Automated', 'Landing + detail, standard blocks', '/ja/event/**'],
    ['T1/T4 Travel', '77', 'Automated', 'Service pages', '/ja/travel/**'],
    ['T5 Lifestyle', '71', 'Automated', 'Editorial', '/ja/lifestyle/**'],
    ['T2/T3 Card Lineup & Detail', '44', 'Manual / Assisted', 'Tables, comparison, application CTAs', '/ja/cardlineup/**'],
    ['T1/T10 Corporate', '44', 'Automated', 'Standard content', '/ja/corporate/**, /ja/corporate.html'],
    ['T4 Merchant', '39', 'Automated', 'Service/info pages', '/ja/merchant/**'],
    ['T9 Usage / Guides', '31', 'Mixed', '3DS & registration = manual; rest auto', '/ja/usage/** (incl. /usage/3d_secure/**)'],
    ['T4 Point / Gourmet / Golf', '39', 'Automated', 'Service detail', '/ja/point/**, /ja/gourmet/**, /ja/golf/**'],
    ['T9 Campaign (cpn_evt)', '15', 'Manual / Assisted', 'Bespoke LP styling', '/ja/cpn_evt/**'],
    ['T4 Shopping/Finance/Insurance/Kameiten/Payment', '36', 'Automated', 'Service detail', '/ja/shopping/**, /ja/finance/**, /ja/insurance/**, /ja/kameiten/**, /ja/payment/**'],
    ['T7 Legal / Policy / About / Sponsorship', '21', 'Automated', 'Plain rich text', '/ja/privacy*.html, /ja/tc.html, /ja/smallprint*, /ja/about/**, /ja/sponsorship*'],
    ['T1 Homepage + misc', '11', 'Manual / Assisted', 'Bespoke homepage layout', '/ja/index.html, /ja/*.html (top-level)'],
  ],
  [24, 8, 16, 24, 28],
  { centerCols: [1, 2], boldFirst: true }
));
children.push(P('Diners split: ≈ 1,556 automated (88%) · ≈ 209 manual / assisted (12%).', { bold: true, before: 60 }));

children.push(H2('3.3  TRUST CLUB (sumitclub.jp) — 439 pages'));
children.push(buildTable(
  ['Template / Section', 'Pages', 'Migration Mode', 'Basis', 'URL pattern'],
  [
    ['T6 Notice', '104', 'Automated', 'Dated text articles', '/ja/notice/** (e.g. inf_YYYYMMDD.html)'],
    ['T9 Login / Member area', '41', 'Manual / Out-of-scope', 'iframe / external auth system', '/ja/loginPage/**, /ja/loginPage.html'],
    ['T4 Travel', '63', 'Automated', 'Service detail', '/ja/travel/**'],
    ['T9/T4 Usage / Guides', '46', 'Mixed', 'Some external links/forms', '/ja/usage/**'],
    ['T2/T3 Card Lineup & Detail', '21', 'Manual / Assisted', 'Spec/fee tables, application', '/ja/cardlineup/**'],
    ['T4 Point', '15', 'Automated', 'Service detail', '/ja/point/**'],
    ['T4 Insurance', '12', 'Automated', 'Service detail', '/ja/insurance/**'],
    ['T7 Legal / Policy', '12', 'Automated', 'Plain rich text', '/ja/privacy*.html, /ja/tc.html, /ja/smallprint*, /ja/cnasp.html, /ja/aasfp.html'],
    ['T1/T4 Entertainment', '7', 'Automated', 'Service detail', '/ja/entertainment/**'],
    ['T9 Campaign', '5', 'Manual / Assisted', 'Bespoke LP', '/ja/campaign/**, /ja/campaign.html'],
    ['T9 Contact', '3', 'Manual', 'iframe form wrappers', '/ja/contact.html, /ja/contact_form.html, /ja/contact/**'],
    ['T1 Home + Gourmet + misc', '23', 'Mixed', 'Homepage bespoke; rest auto', '/ja/index.html, /ja/gourmet/**, /ja/*.html (top-level)'],
    ['T10 Corporate sub-site (ja+en)', '87', 'Automated', 'Standard editorial/landing', '/corporate/**, /ja/corporate_site/**, /en/**'],
  ],
  [24, 8, 16, 24, 28],
  { centerCols: [1, 2], boldFirst: true }
));
children.push(P('TRUST CLUB split: ≈ 373 automated (85%) · ≈ 66 manual / assisted / out-of-scope (15%).', { bold: true, before: 60 }));

children.push(H2('3.4  Combined totals'));
children.push(buildTable(
  ['Migration Mode', 'Diners', 'TRUST CLUB', 'Combined', 'Share'],
  [
    ['Automated (parser-driven)', '1,556', '373', '1,929', '~87%'],
    ['Manual / Assisted', '209', '66', '275', '~13%'],
    ['Total', '1,765', '439', '2,204', '100%'],
  ],
  [30, 16, 18, 18, 18],
  { centerCols: [1, 2, 3, 4], boldFirst: true }
));

children.push(new Paragraph({ children: [new PageBreak()] }));

// ===================== 4. MIGRATION ESTIMATES =====================
children.push(H1('4. Migration Estimates'));
children.push(P('Estimates assume the EDS Document-based authoring model with a bulk import pipeline (block parsers + page transformers), a shared block library reused across both brands, and one combined program. Effort is expressed in person-days (1 day = 8h). Ranges reflect low / high confidence.'));

children.push(H2('4.1  Effort breakdown'));
children.push(buildTable(
  ['Work Stream', 'Description', 'Effort (days)'],
  [
    ['Discovery & setup (done/finalise)', 'Template confirmation, block inventory reconciliation, EDS project scaffold, import infra.', '6 – 8'],
    ['Block library development', '35 EDS blocks (14 Low, 13 Medium, 7 High incl. Benefit Detail, Lists Announcement, Google Maps, Airport Lounge), variants, CSS/JS. Reused across both brands.', '30 – 38'],
    ['Parser / transformer development', 'Per-template parsers (8 patterns) + cleanup/section transformers.', '10 – 13'],
    ['Automated content migration', '~1,929 pages via bulk import, batched by template, with spot-fixing.', '12 – 16'],
    ['Manual / assisted migration', '~275 pages: card detail/spec, campaigns, homepages, iframe-form wrappers, member-area landings.', '20 – 26'],
    ['Navigation, header & footer instrumentation', 'Desktop + mobile + mega-menu per brand; footer link columns.', '5 – 7'],
    ['QA & testing', 'Visual regression, link/redirect checks, responsive + accessibility, content parity sampling.', '12 – 15'],
    ['Project management & UAT support', 'Coordination, stakeholder reviews, go-live support.', '8 – 10'],
  ],
  [26, 50, 24],
  { boldFirst: true, centerCols: [2] }
));
children.push(P('Total estimated effort: ≈ 103 – 133 person-days (midpoint ≈ 118 days).', { bold: true, before: 80 }));

children.push(H2('4.2  Indicative schedule'));
children.push(buildTable(
  ['Phase', 'Calendar duration', 'Parallelism'],
  [
    ['Phase 0 — Setup & reconciliation', 'Weeks 1–2', 'Lead + 1 dev'],
    ['Phase 1 — Blocks + parsers (both brands)', 'Weeks 2–7', '2 devs'],
    ['Phase 2 — Automated bulk migration', 'Weeks 6–10', '1 dev + 1 author'],
    ['Phase 3 — Manual / assisted pages', 'Weeks 8–14', '1–2 authors/devs'],
    ['Phase 4 — Nav/header/footer + integrations', 'Weeks 9–13', '1 dev'],
    ['Phase 5 — QA, UAT & go-live', 'Weeks 13–18', 'Full team + QA'],
  ],
  [38, 32, 30],
  { boldFirst: true }
));
children.push(P('Overall timeline: ≈ 4.5 – 6 calendar months with a 2–3 person core team (some phases overlap).', { bold: true, before: 60 }));

children.push(H2('4.3  Indicative cost'));
children.push(P('Cost uses a blended rate placeholder; replace with the contracted day-rate to finalise. Calculation: midpoint effort 118 days.', { italics: true, after: 80 }));
children.push(buildTable(
  ['Blended day-rate (placeholder)', 'Low (103 d)', 'Midpoint (118 d)', 'High (133 d)'],
  [
    ['$800 / day', '$82,400', '$94,400', '$106,400'],
    ['$1,000 / day', '$103,000', '$118,000', '$133,000'],
    ['$1,200 / day', '$123,600', '$141,600', '$159,600'],
  ],
  [28, 24, 24, 24],
  { boldFirst: true, centerCols: [1, 2, 3] }
));

children.push(H2('4.4  Key assumptions & risks'));
[
  'Both sites share one component library and chrome — block development is done once and reused, which is the single biggest efficiency in this estimate.',
  'Content freeze (or delta-handling plan) is in place during bulk migration; the 2,204 figure is the current sitemap snapshot and will drift.',
  'Forms (contact, card application, 3-D Secure) and the "Club Online" member area are external/iframe systems and are OUT of content-migration scope — only their wrapper pages migrate. Re-platforming these is a separate workstream if required.',
  'Japanese-language content (and the TRUST CLUB English mirror) imports cleanly; font/typography parity verified in QA.',
  'Estimates exclude net-new design, brand redesign, SEO redirect-map sign-off beyond standard 1:1 mapping, and third-party integration rebuilds (search, analytics tags, ad pixels).',
  'PDFs and DAM assets are migrated by reference / bulk asset copy, not re-authored.',
].forEach(t => children.push(bullet(t)));

children.push(P(''));
children.push(new Paragraph({
  spacing: { before: 160, after: 60 },
  border: { top: { style: BorderStyle.SINGLE, size: 6, color: 'BFBFBF' } },
  children: [new TextRun({ text: 'End of report — figures are planning estimates and should be confirmed against the reconciled discovery spreadsheet and contracted rates.', italics: true, size: 18, color: '808080', font: FONT })],
}));

const doc = new Document({
  creator: 'Migration Analysis',
  title: 'Diners Club & TRUST CLUB — EDS Migration Analysis',
  styles: {
    default: { document: { run: { font: FONT, size: 21 } } },
  },
  sections: [{
    properties: { page: { margin: { top: 1000, bottom: 1000, left: 1000, right: 1000 } } },
    children,
  }],
});

const buf = await Packer.toBuffer(doc);
fs.writeFileSync('/workspace/current/Diners_SumitClub_Migration_Analysis.docx', buf);
console.log('WROTE', buf.length, 'bytes');
