import {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  WidthType, BorderStyle, HeadingLevel, AlignmentType, ImageRun,
  TableBorders, ShadingType, VerticalAlign, PageBreak
} from 'docx';
import fs from 'fs';
import path from 'path';

const screenshotsDir = '/tmp/playwright';

function loadImage(filename) {
  const filePath = path.join(screenshotsDir, filename);
  if (fs.existsSync(filePath)) {
    return fs.readFileSync(filePath);
  }
  return null;
}

function createHeading(text, level = HeadingLevel.HEADING_1) {
  return new Paragraph({ heading: level, children: [new TextRun({ text, bold: true })] });
}

function createParagraph(text, options = {}) {
  return new Paragraph({
    spacing: { after: 120 },
    ...options,
    children: [new TextRun({ text, size: 22, ...options.textOptions })]
  });
}

function createBoldParagraph(text) {
  return new Paragraph({
    spacing: { after: 120 },
    children: [new TextRun({ text, bold: true, size: 22 })]
  });
}

const noBorders = {
  top: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  bottom: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  left: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
  right: { style: BorderStyle.SINGLE, size: 1, color: 'CCCCCC' },
};

function headerCell(text) {
  return new TableCell({
    borders: noBorders,
    shading: { type: ShadingType.SOLID, color: '1B3A5C' },
    verticalAlign: VerticalAlign.CENTER,
    children: [new Paragraph({
      spacing: { before: 60, after: 60 },
      children: [new TextRun({ text, bold: true, color: 'FFFFFF', size: 20 })]
    })]
  });
}

function dataCell(text, options = {}) {
  return new TableCell({
    borders: noBorders,
    verticalAlign: VerticalAlign.TOP,
    ...options,
    children: [new Paragraph({
      spacing: { before: 40, after: 40 },
      children: [new TextRun({ text, size: 20 })]
    })]
  });
}

function multiLineDataCell(lines) {
  return new TableCell({
    borders: noBorders,
    verticalAlign: VerticalAlign.TOP,
    children: lines.map(line => new Paragraph({
      spacing: { before: 20, after: 20 },
      children: [new TextRun({ text: line, size: 20 })]
    }))
  });
}

// Templates Data
const templates = [
  {
    name: 'Homepage',
    complexity: 'High',
    reasoning: 'Multiple interactive sections (hero carousel, tabbed industry navigation, accordion capabilities, card carousels), dual contact forms, animated statistics, complex layout with multiple content zones',
    urls: ['https://www.broadridge.com/']
  },
  {
    name: 'Industry Vertical Landing Page',
    complexity: 'High',
    reasoning: 'Feature carousels, accordion solution grid (8+ expandable items), statistics carousel, awards section with mixed card layouts, insights carousel with pagination, audience segments carousel, dual forms',
    urls: ['https://www.broadridge.com/who-we-serve/asset-management', 'https://www.broadridge.com/who-we-serve/capital-markets', 'https://www.broadridge.com/who-we-serve/wealth-management', 'https://www.broadridge.com/who-we-serve/issuers', 'https://www.broadridge.com/who-we-serve/consumer-industries']
  },
  {
    name: 'Capability Landing Page',
    complexity: 'High',
    reasoning: 'Value proposition grid with 7 feature blocks, statistics section (5 metrics), featured solution cards, awards section, dual contact forms, similar structure to industry pages but focused on capabilities',
    urls: ['https://www.broadridge.com/capability/front-office-solutions/', 'https://www.broadridge.com/capability/middle-and-back-office-solutions/', 'https://www.broadridge.com/capability/governance-and-regulatory-compliance/', 'https://www.broadridge.com/capability/customer-communications/', 'https://www.broadridge.com/capability/data-analytics-and-intelligence/']
  },
  {
    name: 'Product/Solution Detail Page',
    complexity: 'Medium',
    reasoning: 'Hero with value proposition, three-pillar feature blocks, product variant cards (4 items), client testimonial quote, related solutions cross-links, resource hub, portal access buttons, dual forms',
    urls: ['https://www.broadridge.com/capability/governance-and-regulatory-compliance/proxy-services/', 'https://www.broadridge.com/capability/front-office-solutions/portfolio-management/', 'https://www.broadridge.com/capability/middle-and-back-office-solutions/asset-servicing/']
  },
  {
    name: 'Topic/Campaign Hub Page',
    complexity: 'High',
    reasoning: 'Comprehensive single-topic page with multiple content zones: insights cards, value propositions, 6 featured services, 4 industry verticals, news section (8 items), executive quote, expandable FAQ (11 items), dual forms',
    urls: ['https://www.broadridge.com/hub/tokenization']
  },
  {
    name: 'Insights/Resource Hub Page',
    complexity: 'Medium',
    reasoning: 'Featured insights carousel (8 cards with pagination), search interface with keyword input, dynamic article listing with "Load More" pagination, promotional CTA banner, dual forms',
    urls: ['https://www.broadridge.com/insight-pages/broadridge-insights', 'https://www.broadridge.com/insight-pages/artificial-intelligence', 'https://www.broadridge.com/insight-pages/transformation-innovation', 'https://www.broadridge.com/insight-pages/investor-trends']
  },
  {
    name: 'Insight/Whitepaper Detail (Gated Content)',
    complexity: 'Medium',
    reasoning: 'Hero with study preview image, key findings infographics (4 data visualizations), lead capture form with multiple fields, secondary form variant, FAQ section (10 items), progressive disclosure pattern',
    urls: ['https://www.broadridge.com/insights/2026-digital-transformation-study', 'https://www.broadridge.com/insights/finadium-bank-balance-sheet-benefits-from-intraday-distributed-ledger-repo']
  },
  {
    name: 'About/Corporate Page',
    complexity: 'Medium',
    reasoning: 'Video hero, narrative content blocks, key metrics (4 data points), subsection card navigation (6 cards), career CTA, investor relations block, partner logo grid, dual forms',
    urls: ['https://www.broadridge.com/about/', 'https://www.broadridge.com/about/the-broadridge-culture', 'https://www.broadridge.com/about/sustainability/']
  },
  {
    name: 'Leadership/Team Page',
    complexity: 'Medium',
    reasoning: 'Tab navigation (Executive Leadership / Board of Directors), headshot photo grid with uniform card format, expandable bio sections with "Read bio" functionality, In Memoriam section',
    urls: ['https://www.broadridge.com/our-leadership-team']
  },
  {
    name: 'News/Press Hub Page',
    complexity: 'Low',
    reasoning: 'Simple layout with page title, featured story highlight, four-card news grid, media contact information. No filters, pagination, or search functionality visible',
    urls: ['https://www.broadridge.com/news-room']
  },
  {
    name: 'Contact/Office Directory Page',
    complexity: 'Medium',
    reasoning: 'Hero with intro text, regional headquarters cards (3 locations), quick links section (6 items), expandable regional tabs (North America/EMEA/APAC), office cards with flags and phone links, dual forms',
    urls: ['https://www.broadridge.com/contact-us']
  },
  {
    name: 'Awards & Recognition Page',
    complexity: 'Low',
    reasoning: 'Featured award logo cards (4 prominent), solution awards linked list, comprehensive bulleted archive list, "Load More" functionality. Mostly static content with minimal interactivity',
    urls: ['https://www.broadridge.com/about/awards-and-recognition']
  },
  {
    name: 'Legal/Policy Page',
    complexity: 'Low',
    reasoning: 'Long-form document template with hierarchical sections, numbered navigation, data tables, bulleted lists. Primarily text content with no interactive components beyond navigation',
    urls: ['https://www.broadridge.com/legal/privacy-statement-english', 'https://www.broadridge.com/legal/terms-of-use', 'https://www.broadridge.com/legal/accessibility']
  }
];

// Blocks Data
const blocks = [
  {
    name: 'Global Header/Navigation',
    complexity: 'High',
    description: 'Sticky header with logo, search trigger, "Contact us" CTA button, and hamburger menu. Expands into full-screen mega-menu overlay with multi-level navigation organized by: Who We Serve, Capabilities, Insights, About Us. Includes language selector and utility links (Client access, Careers).',
    urls: ['https://www.broadridge.com/ (all pages)'],
    behavior: 'Hamburger menu toggles full-screen overlay; search opens expandable search panel; mega-menu supports hover/click navigation with nested sub-items; responsive collapse on mobile.'
  },
  {
    name: 'Global Footer',
    complexity: 'Medium',
    description: 'Multi-column footer with Broadridge logo, company description, NYSE stock ticker widget with real-time pricing, social media links (LinkedIn, Instagram, YouTube), three-column link lists (Company info, Who we serve, Quick links), legal links bar, language selector (FR/DE/JP), copyright.',
    urls: ['https://www.broadridge.com/ (all pages)'],
    behavior: 'Stock ticker updates dynamically; social links open in new tabs; language links redirect to localized sites; responsive stacking on mobile.'
  },
  {
    name: 'Hero Banner - Carousel',
    complexity: 'High',
    description: 'Full-width animated hero with auto-rotating content slides featuring headline, description text, and animated graphic/illustration. Includes play/pause button and pagination dots. Used exclusively on homepage.',
    urls: ['https://www.broadridge.com/'],
    behavior: 'Auto-advances with configurable timing; pause/play toggle; supports animated SVG/canvas illustrations; mobile-optimized with stacked layout.'
  },
  {
    name: 'Hero Banner - Static',
    complexity: 'Low',
    description: 'Full-width hero with dark gradient background, page title (H1), and descriptive subtitle paragraph. Used on industry/capability landing pages and about pages.',
    urls: ['https://www.broadridge.com/who-we-serve/asset-management', 'https://www.broadridge.com/capability/front-office-solutions/'],
    behavior: 'Static display; responsive text sizing; dark-to-transparent gradient overlay.'
  },
  {
    name: 'Hero Banner - Image Background',
    complexity: 'Low',
    description: 'Full-width hero with background image, overlaid page title and description text. Used on insights hub and contact pages.',
    urls: ['https://www.broadridge.com/insight-pages/broadridge-insights', 'https://www.broadridge.com/contact-us'],
    behavior: 'Parallax-style background image; text overlay with semi-transparent backdrop for readability.'
  },
  {
    name: 'Contact Form (Lead Capture)',
    complexity: 'High',
    description: 'Full-width form section with dark blue background, heading "What\'s next for your business?", 8 form fields (first name, last name, email, telephone, job title, company, country dropdown, message textarea), submit button, reCAPTCHA, and regional phone numbers sidebar.',
    urls: ['https://www.broadridge.com/ (appears on nearly all pages, often duplicated)'],
    behavior: 'Form validation with required field indicators; custom country dropdown with search; reCAPTCHA verification; UTM tracking via hidden fields; campaign ID tracking; auto-population from cookies; phone number links (tel: protocol).'
  },
  {
    name: 'Announcement Cards',
    complexity: 'Medium',
    description: 'Two-column layout with announcement cards featuring "ANNOUNCEMENT" tag label, heading text, and directional arrow icon. Links to press releases or external content.',
    urls: ['https://www.broadridge.com/'],
    behavior: 'Card links navigate to announcement detail; "Opens in new tab" indicator for external links; hover state with subtle animation.'
  },
  {
    name: 'Featured Solutions Grid',
    complexity: 'Medium',
    description: 'Section with heading and 4-column grid of solution cards. Each card has an icon/illustration, solution name, brief description, and directional arrow. Full card is clickable.',
    urls: ['https://www.broadridge.com/'],
    behavior: 'Cards link to solution detail pages; hover effect with card elevation; responsive grid (4-col desktop → 2-col tablet → 1-col mobile).'
  },
  {
    name: 'Insights Carousel',
    complexity: 'High',
    description: 'Horizontal carousel of content cards (4-8 items) with content type tag (WHITEPAPER, REPORT, ARTICLE, WEBINAR, CASE STUDY, RESOURCE), heading, and thumbnail image. Includes dot pagination indicators and prev/next arrow buttons. "Explore all insights" CTA link below.',
    urls: ['https://www.broadridge.com/', 'https://www.broadridge.com/who-we-serve/asset-management', 'https://www.broadridge.com/insight-pages/broadridge-insights'],
    behavior: 'Swipe/drag on touch devices; dot pagination with slide-to; prev/next arrow navigation; responsive card count (4 desktop → 2 tablet → 1 mobile); cards link to content detail pages.'
  },
  {
    name: 'Industry Tabs (Tabbed Content)',
    complexity: 'High',
    description: 'Tabbed interface with 5 industry vertical buttons (Asset Management, Capital Markets, Issuers, Wealth Management, Consumer Industries). Each tab reveals a content panel with large background image, industry name, description paragraph, and "Explore" CTA link with arrow.',
    urls: ['https://www.broadridge.com/'],
    behavior: 'Tab switching with smooth content transition; active tab indicator; image preloading for smooth transitions; responsive (tabs become scrollable on mobile).'
  },
  {
    name: 'Capabilities Accordion',
    complexity: 'High',
    description: 'Vertically stacked accordion with 7 capability categories. Each item shows a category name and chevron icon; clicking expands to reveal heading, description paragraph, and directional arrow link. Only one item expanded at a time.',
    urls: ['https://www.broadridge.com/', 'https://www.broadridge.com/who-we-serve/asset-management'],
    behavior: 'Click toggles expand/collapse with smooth animation; auto-closes other items when one opens; links to capability detail pages; mobile-optimized with full-width items.'
  },
  {
    name: 'Awards/Recognition Section',
    complexity: 'Medium',
    description: 'Section with heading, featured large award card (image + text), and 3-column grid of smaller award cards. Each card shows year/category tag, award title, and arrow icon. "Explore all" CTA link at bottom.',
    urls: ['https://www.broadridge.com/', 'https://www.broadridge.com/who-we-serve/asset-management'],
    behavior: 'Cards link to awards page; mixed large/small card layout (1 featured + 3 standard); responsive grid adjustment.'
  },
  {
    name: 'Statistics/Metrics Carousel',
    complexity: 'Medium',
    description: 'Horizontal carousel displaying 3 key statistics with large numeric values (e.g., "$100T", "2B", "$1.5T") and supporting description text. Includes dot pagination and prev/next arrows.',
    urls: ['https://www.broadridge.com/who-we-serve/asset-management', 'https://www.broadridge.com/capability/front-office-solutions/'],
    behavior: 'Auto-rotation with pause on hover; animated number counting on first view; dot pagination; responsive (single stat per slide on mobile).'
  },
  {
    name: 'Key Metrics Block (Static)',
    complexity: 'Low',
    description: 'Horizontal row of 4 key data points with large numbers and brief descriptions. Used on About page to display company scale (15K employees, 10K companies, $15T settlements, $100T AUM).',
    urls: ['https://www.broadridge.com/about/'],
    behavior: 'Static display; animated count-up on scroll into view; responsive (2x2 grid on tablet, stacked on mobile).'
  },
  {
    name: 'Content Feature Carousel (Text-only)',
    complexity: 'Medium',
    description: 'Text-based carousel showing rotating value propositions. Each slide has a heading and paragraph description. Used to cycle through key messaging points (e.g., 3 value pillars for an industry vertical or 4 audience segments).',
    urls: ['https://www.broadridge.com/who-we-serve/asset-management'],
    behavior: 'Auto-advance with dot indicators; prev/next navigation; smooth text transitions; responsive.'
  },
  {
    name: 'Search/Filter Interface',
    complexity: 'Medium',
    description: 'Search input field with placeholder text and search button, positioned above a dynamic article listing. Results display as simple linked headings in a list format with "Load More" button for pagination.',
    urls: ['https://www.broadridge.com/insight-pages/broadridge-insights'],
    behavior: 'Keyword search with instant filtering; "Load More" loads additional results (AJAX); results are clickable links; clear search functionality.'
  },
  {
    name: 'Promotional CTA Banner',
    complexity: 'Low',
    description: 'Full-width banner with gradient/image background featuring study/campaign title, brief description, and action link button. Used to promote key research or initiatives.',
    urls: ['https://www.broadridge.com/insight-pages/broadridge-insights'],
    behavior: 'Static display with link; responsive text scaling; stands out visually from surrounding content.'
  },
  {
    name: 'Office Location Cards',
    complexity: 'Medium',
    description: 'Regional office directory with country headings, individual office cards showing country flag icon, city name, full address, and clickable phone number (tel: link). Organized by geographic region with expandable sections.',
    urls: ['https://www.broadridge.com/contact-us'],
    behavior: 'Regional tab switching (North America/EMEA/APAC); expandable/collapsible country sections; phone numbers as clickable tel: links; "Back to top" floating button; flag icons for country identification.'
  },
  {
    name: 'Quick Links Section',
    complexity: 'Low',
    description: 'Vertical list of linked items with text label and directional arrow icon. Used for navigation shortcuts to related pages (e.g., Shareholder info, Careers, Media relations).',
    urls: ['https://www.broadridge.com/contact-us'],
    behavior: 'Simple clickable links; "Opens in new tab" indicator for external links; hover highlight effect.'
  },
  {
    name: 'Team Member Grid',
    complexity: 'Medium',
    description: 'Grid of team member cards with professional headshot photo, name, and job title. Tab navigation to switch between Executive Leadership and Board of Directors. Cards have expandable "Read bio" functionality.',
    urls: ['https://www.broadridge.com/our-leadership-team'],
    behavior: 'Tab switching between groups; "Read bio" expands inline biography; uniform card sizing; responsive grid (4-col → 3-col → 2-col → 1-col).'
  },
  {
    name: 'FAQ Accordion',
    complexity: 'Medium',
    description: 'Expandable FAQ section with question headings that toggle answer content. Used on topic/campaign pages to address common questions.',
    urls: ['https://www.broadridge.com/hub/tokenization', 'https://www.broadridge.com/insights/2026-digital-transformation-study'],
    behavior: 'Click to expand/collapse individual Q&A items; smooth animation; can have multiple items open simultaneously; anchor link support for direct linking to specific questions.'
  },
  {
    name: 'Partner Logo Grid',
    complexity: 'Low',
    description: 'Horizontal row of partner/client logos displayed in grayscale or brand colors. Shows ecosystem partnerships (Salesforce, Morningstar, SS&C, etc.).',
    urls: ['https://www.broadridge.com/about/'],
    behavior: 'Static logo display; responsive wrapping; optional link to partner page; hover effect may show color version.'
  },
  {
    name: 'Executive Quote/Testimonial',
    complexity: 'Low',
    description: 'Highlighted quote block with quotation marks, quote text, speaker name, and title/company attribution. Used for client testimonials or executive statements.',
    urls: ['https://www.broadridge.com/hub/tokenization', 'https://www.broadridge.com/capability/governance-and-regulatory-compliance/proxy-services/'],
    behavior: 'Static display; styled with large quotation marks; may include speaker photo; responsive text sizing.'
  },
  {
    name: 'Cookie Consent Banner',
    complexity: 'Low',
    description: 'Bottom-anchored privacy banner with cookie usage message, "Cookie Notice" link, and three action buttons: "Customize Settings", "Reject all", "Accept all cookies". Powered by OneTrust.',
    urls: ['https://www.broadridge.com/ (all pages, first visit)'],
    behavior: 'Appears on first visit; dismisses on any button click; "Customize Settings" opens preference center dialog; persists choice in cookie; does not reappear after consent.'
  },
  {
    name: 'Video Embed Block',
    complexity: 'Medium',
    description: 'Embedded video player with play button overlay, video thumbnail, and optional transcript link. Used on About/Culture pages for leadership messages.',
    urls: ['https://www.broadridge.com/about/', 'https://www.broadridge.com/about/sustainability/'],
    behavior: 'Click to play; video controls; optional full-screen; transcript toggle; responsive sizing maintaining aspect ratio.'
  },
  {
    name: 'Three-Pillar Cards',
    complexity: 'Low',
    description: 'Three-column card layout presenting ESG/value pillars with icon, heading, and description text. Each card has consistent visual treatment. Used for sustainability (Environmental, Social, Governance) or company values.',
    urls: ['https://www.broadridge.com/about/sustainability/', 'https://www.broadridge.com/about/the-broadridge-culture'],
    behavior: 'Static card display; optional "Learn More" links; responsive (3-col → stacked); consistent card heights.'
  },
  {
    name: 'News Card Grid',
    complexity: 'Low',
    description: 'Four-column grid of news/press release cards with title text and link. Used in newsroom for displaying recent press releases.',
    urls: ['https://www.broadridge.com/news-room', 'https://www.broadridge.com/hub/tokenization'],
    behavior: 'Cards link to press release detail; responsive grid; simple text-based cards without images.'
  }
];

// Build document
const doc = new Document({
  styles: {
    default: {
      document: {
        run: { font: 'Calibri', size: 22 }
      }
    }
  },
  sections: [{
    properties: {},
    children: [
      // Title Page
      new Paragraph({ spacing: { before: 2000 } }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: 'BROADRIDGE FINANCIAL SOLUTIONS', bold: true, size: 48, color: '1B3A5C' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: 'Website Analysis Report', bold: true, size: 36, color: '1B3A5C' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 400 },
        children: [new TextRun({ text: 'Templates Inventory & Blocks/Components Catalog', size: 28, color: '555555' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 200 },
        children: [new TextRun({ text: 'https://www.broadridge.com/', size: 24, color: '0066CC' })]
      }),
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { after: 800 },
        children: [new TextRun({ text: `Date: ${new Date().toLocaleDateString('en-US', { year: 'numeric', month: 'long', day: 'numeric' })}`, size: 22, color: '666666' })]
      }),

      // Homepage screenshot
      ...(loadImage('broadridge-homepage-full.png') ? [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new ImageRun({
            data: loadImage('broadridge-homepage-full.png'),
            transformation: { width: 300, height: 900 },
            type: 'png'
          })]
        })
      ] : []),

      // Page break
      new Paragraph({ children: [new PageBreak()] }),

      // Table of Contents
      createHeading('Table of Contents', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createParagraph('1. Executive Summary'),
      createParagraph('2. Templates Inventory'),
      createParagraph('3. Blocks / Components Catalog'),
      createParagraph('4. Visual References'),
      createParagraph('5. Recommendations'),

      new Paragraph({ children: [new PageBreak()] }),

      // Executive Summary
      createHeading('1. Executive Summary', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createParagraph('This document provides a comprehensive analysis of the Broadridge Financial Solutions website (www.broadridge.com), identifying all unique page templates and reusable blocks/components used across the site.'),
      new Paragraph({ spacing: { after: 100 } }),
      createBoldParagraph('Key Findings:'),
      createParagraph('• 13 unique page templates identified across the site'),
      createParagraph('• 26 reusable blocks/components cataloged'),
      createParagraph('• Site uses a modular, component-based architecture'),
      createParagraph('• Consistent design system with dark navy (#1B3A5C) primary color'),
      createParagraph('• Dual contact form placement pattern (top and bottom of most pages)'),
      createParagraph('• Heavy use of carousels and accordion patterns for content density'),
      createParagraph('• Responsive design with mobile-optimized layouts'),
      new Paragraph({ spacing: { after: 200 } }),
      createBoldParagraph('Site Architecture Overview:'),
      createParagraph('The site is organized around four primary content hierarchies: Who We Serve (industry verticals), Capabilities (solutions), Insights (thought leadership), and About (corporate information). Each hierarchy uses specialized templates while sharing a common block library.'),

      new Paragraph({ children: [new PageBreak()] }),

      // Templates Inventory
      createHeading('2. Templates Inventory', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createParagraph('The following table catalogs all unique page templates identified across the Broadridge website. Templates are classified by complexity level based on the number of interactive components, content zones, and technical implementation requirements.'),
      new Paragraph({ spacing: { after: 300 } }),

      // Templates Table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              headerCell('Template Name'),
              headerCell('Complexity'),
              headerCell('Reasoning'),
              headerCell('Reference URL(s)')
            ]
          }),
          ...templates.map(t => new TableRow({
            children: [
              dataCell(t.name),
              dataCell(t.complexity),
              dataCell(t.reasoning),
              multiLineDataCell(t.urls)
            ]
          }))
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Detailed Template Descriptions
      createHeading('2.1 Template Details', HeadingLevel.HEADING_2),
      new Paragraph({ spacing: { after: 200 } }),

      ...templates.flatMap(t => [
        createHeading(t.name, HeadingLevel.HEADING_3),
        createParagraph(`Complexity: ${t.complexity}`),
        createParagraph(`Reasoning: ${t.reasoning}`),
        createBoldParagraph('Reference URLs:'),
        ...t.urls.map(url => createParagraph(`  • ${url}`)),
        new Paragraph({ spacing: { after: 200 } }),
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // Blocks Catalog
      createHeading('3. Blocks / Components Catalog', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createParagraph('The following catalog identifies all reusable blocks and components present across the Broadridge website. Design variations of the same content model are documented as variants of a single block rather than separate components.'),
      new Paragraph({ spacing: { after: 300 } }),

      // Blocks Summary Table
      new Table({
        width: { size: 100, type: WidthType.PERCENTAGE },
        rows: [
          new TableRow({
            children: [
              headerCell('Block Name'),
              headerCell('Complexity'),
              headerCell('Description'),
              headerCell('Reference URL(s)')
            ]
          }),
          ...blocks.map(b => new TableRow({
            children: [
              dataCell(b.name),
              dataCell(b.complexity),
              dataCell(b.description.substring(0, 150) + (b.description.length > 150 ? '...' : '')),
              multiLineDataCell(b.urls.slice(0, 2))
            ]
          }))
        ]
      }),

      new Paragraph({ children: [new PageBreak()] }),

      // Detailed Block Descriptions
      createHeading('3.1 Block Details', HeadingLevel.HEADING_2),
      new Paragraph({ spacing: { after: 200 } }),

      ...blocks.flatMap(b => [
        createHeading(b.name, HeadingLevel.HEADING_3),
        createParagraph(`Complexity: ${b.complexity}`),
        createBoldParagraph('Description:'),
        createParagraph(b.description),
        createBoldParagraph('Behaviour & Functionality:'),
        createParagraph(b.behavior),
        createBoldParagraph('Reference URLs:'),
        ...b.urls.map(url => createParagraph(`  • ${url}`)),
        new Paragraph({ spacing: { after: 300 } }),
      ]),

      new Paragraph({ children: [new PageBreak()] }),

      // Visual References
      createHeading('4. Visual References', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createParagraph('The following screenshots provide visual reference for key page templates and block implementations.'),
      new Paragraph({ spacing: { after: 300 } }),

      // Homepage screenshot
      createHeading('4.1 Homepage', HeadingLevel.HEADING_2),
      createParagraph('URL: https://www.broadridge.com/'),
      ...(loadImage('broadridge-homepage-hero.png') ? [
        new Paragraph({
          spacing: { after: 200 },
          children: [new ImageRun({
            data: loadImage('broadridge-homepage-hero.png'),
            transformation: { width: 550, height: 340 },
            type: 'png'
          })]
        })
      ] : []),
      new Paragraph({ spacing: { after: 100 } }),

      // Industry page
      createHeading('4.2 Industry Vertical Landing Page (Asset Management)', HeadingLevel.HEADING_2),
      createParagraph('URL: https://www.broadridge.com/who-we-serve/asset-management'),
      ...(loadImage('broadridge-industry-page.png') ? [
        new Paragraph({
          spacing: { after: 200 },
          children: [new ImageRun({
            data: loadImage('broadridge-industry-page.png'),
            transformation: { width: 300, height: 900 },
            type: 'png'
          })]
        })
      ] : []),

      new Paragraph({ children: [new PageBreak()] }),

      // Insights hub
      createHeading('4.3 Insights/Resource Hub Page', HeadingLevel.HEADING_2),
      createParagraph('URL: https://www.broadridge.com/insight-pages/broadridge-insights'),
      ...(loadImage('broadridge-insights-hub.png') ? [
        new Paragraph({
          spacing: { after: 200 },
          children: [new ImageRun({
            data: loadImage('broadridge-insights-hub.png'),
            transformation: { width: 300, height: 900 },
            type: 'png'
          })]
        })
      ] : []),

      // Contact page
      createHeading('4.4 Contact / Office Directory Page', HeadingLevel.HEADING_2),
      createParagraph('URL: https://www.broadridge.com/contact-us'),
      ...(loadImage('broadridge-contact-page.png') ? [
        new Paragraph({
          spacing: { after: 200 },
          children: [new ImageRun({
            data: loadImage('broadridge-contact-page.png'),
            transformation: { width: 550, height: 340 },
            type: 'png'
          })]
        })
      ] : []),

      new Paragraph({ children: [new PageBreak()] }),

      // Recommendations
      createHeading('5. Recommendations', HeadingLevel.HEADING_1),
      new Paragraph({ spacing: { after: 200 } }),
      createBoldParagraph('Migration Considerations:'),
      new Paragraph({ spacing: { after: 100 } }),
      createParagraph('1. Shared Components: The Contact Form and Header/Footer are used site-wide and should be implemented as global components first.'),
      createParagraph('2. Carousel Pattern: Multiple blocks use carousel functionality (Insights, Statistics, Features). A shared carousel utility should be developed once and reused.'),
      createParagraph('3. Accordion Pattern: The Capabilities Accordion and FAQ Accordion share similar expand/collapse behavior and can leverage the same base implementation.'),
      createParagraph('4. Card Variations: Multiple card types exist (Solution, Insight, Award, Announcement, Office) that share a common card model but differ in visual layout — these should be design variants of a single Card block.'),
      createParagraph('5. Form Complexity: The lead capture form has significant hidden functionality (UTM tracking, campaign IDs, cookie management, geolocation) that must be replicated.'),
      createParagraph('6. Dynamic Content: The Insights hub search/filter and "Load More" pagination require dynamic data loading capabilities.'),
      new Paragraph({ spacing: { after: 200 } }),
      createBoldParagraph('Complexity Distribution:'),
      createParagraph(`• High complexity: ${templates.filter(t => t.complexity === 'High').length} templates, ${blocks.filter(b => b.complexity === 'High').length} blocks`),
      createParagraph(`• Medium complexity: ${templates.filter(t => t.complexity === 'Medium').length} templates, ${blocks.filter(b => b.complexity === 'Medium').length} blocks`),
      createParagraph(`• Low complexity: ${templates.filter(t => t.complexity === 'Low').length} templates, ${blocks.filter(b => b.complexity === 'Low').length} blocks`),
    ]
  }]
});

const buffer = await Packer.toBuffer(doc);
const outputPath = '/backups/checksoundar/KyowaKirin-site-analysis/repo/Broadridge_Site_Analysis_Report.docx';
fs.writeFileSync(outputPath, buffer);
console.log(`Report generated: ${outputPath}`);
console.log(`File size: ${(buffer.length / 1024).toFixed(1)} KB`);
