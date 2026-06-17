const fs = require('fs');
const lines = fs.readFileSync(process.argv[2], 'utf8').trim().split('\n').filter(Boolean);

// Category definitions with migration tag. Order matters (first match wins).
const rules = [
  // Non-page assets / fragments
  [/\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/include\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/common\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/member\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/_assets\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/(_offer|_header|_footer|_contents|_header_main|_project\d+|parts|qs_parts)\b/, 'LP Partial Fragment', 'Exclude (not a page)'],
  [/\/bumper\/link\.html$/, 'Bumper/Redirect Stub', 'Exclude (not a page)'],
  [/\/iframe\/blank/, 'Blank/Helper Stub', 'Exclude (not a page)'],
  [/\/popup\/javascriptoff/, 'Blank/Helper Stub', 'Exclude (not a page)'],
  [/\/pushcode\.|\/body_code\.|\/log\.html$|\/activecheck\.html$|google[0-9a-f]+\.html$/, 'Tracking/System Stub', 'Exclude (not a page)'],
  [/\/popup\/error\.html$|modal-window|modal_0/, 'Modal/Popup Fragment', 'Exclude (not a page)'],
  [/virtual-numpad|btn-online-sp/, 'Auth Widget Fragment', 'Exclude (not a page)'],

  // Auth-gated / system
  [/\/platinum\/member\/campaign\/dining-selection\/search\/R\d+\.html$/, 'Member Dining-Selection Detail (R####)', 'Manual / Auth-gated'],
  [/\/platinum\/member\/campaign\/dining-selection(\/(index|search\/(index|stocklist)))?\.html$/, 'Member Dining-Selection Index/Search', 'Manual / Auth-gated'],
  [/\/signon\//, 'Sign-on / Login System', 'Exclude / Re-point'],
  [/\/loginPage\//, 'Login/Member System Page', 'Exclude / Re-point'],
  [/\/rd\/|\/rd_|\/WD1100112|\/redirect\//, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/pointmall\/.*sysfail/, 'System Error Stub', 'Exclude / Re-point'],
  [/\/(400|401|403|404|500|501|503)\.html$/, 'HTTP Error Page', 'Manual / Replace'],

  // Email templates
  [/\/email\/estatement\//, 'E-statement Email Template', 'Exclude / Separate track'],

  // Application landing pages
  [/\/entry_form\/.*\/index\.html$/, 'Card Application Landing Page (LP)', 'Manual'],
  [/\/entry_form\//, 'Card Application Landing Page (LP)', 'Manual'],

  // Verification microsites
  [/\/honninkakunin/, 'Identity Verification Microsite', 'Manual'],
  [/\/(corp-verify|how-verify|before-verify|id-doc)\.html$/, 'Identity Verification Microsite', 'Manual'],
  [/\/ccol\//, 'Club Online Bumper Page', 'Manual'],
  [/\/taikai\//, 'Cancellation/Termination Flow', 'Manual'],
  [/\/bp\/apple/, 'App-bridge Stub', 'Exclude / Re-point'],

  // Corporate recruit (separate shell)
  [/\/corporate\/recruit\//, 'Corporate Recruit Page', 'Assisted'],
  [/\/corporate\/kaizen\//, 'Corporate Kaizen Notice (archive)', 'Automated'],
  [/\/corporate\/news\//, 'Corporate News List', 'Automated'],
  [/\/corporate\/(index|greeting|summary)\//, 'Corporate Info (legacy shell)', 'Assisted'],
  [/\/corporate\//, 'Corporate Info (legacy shell)', 'Assisted'],

  // /ja/ corporate_site (AEM corporate)
  [/\/corporate_site\/news\.html$/, 'Corporate News List', 'Automated'],
  [/\/corporate_site\//, 'Corporate Info (AEM)', 'Assisted'],
  [/\/corporate\/d60th/, 'Corporate Special/Anniversary', 'Assisted'],
  [/\/ja\/corporate\//, 'Corporate Info (AEM)', 'Assisted'],

  // Content families
  [/\/notice\/(index|inf_|20\d{6}|2019|maintenance|phishing|201907)/, 'Notice / News Detail', 'Automated'],
  [/\/notice(\/index)?\.html$/, 'Notice / News Index', 'Automated'],
  [/\/notice\//, 'Notice / News Detail', 'Automated'],
  [/\/cardlineup\/(list|application|nyukai)/, 'Card Lineup Listing/Index', 'Assisted'],
  [/\/cardlineup(\.html|\/index\.html)$/, 'Card Lineup Listing/Index', 'Assisted'],
  [/\/cardlineup\/card_services_list/, 'Card Lineup Listing/Index', 'Assisted'],
  [/\/cardlineup\//, 'Card Product Detail', 'Manual'],
  [/\/campaign\//, 'Campaign Landing Page', 'Assisted'],
  [/\/insurance\//, 'Insurance Product/Info', 'Assisted'],
  [/\/point\//, 'Point Program Page', 'Assisted'],
  [/\/travel\//, 'Travel Service Page', 'Assisted'],
  [/\/usage\//, 'Usage / How-to Page', 'Assisted'],
  [/\/unique\//, 'Club Online Promo (unique)', 'Assisted'],
  [/\/entertainment\//, 'Entertainment/Lifestyle Page', 'Assisted'],
  [/\/gourmet\//, 'Gourmet Service Page', 'Assisted'],
  [/\/commercialcard\//, 'Commercial Card Page', 'Assisted'],
  [/\/(privacy|smallprint|tc|policy_customerresponse|cnasp|localbank|shinkinbank|email)\b/, 'Legal / Policy Text', 'Automated'],
  [/\/(privacy_kzk|privacy\/)/, 'Legal / Policy Text', 'Automated'],
  [/\/contact\//, 'Contact / Support', 'Manual'],
  [/\/contact\.html$/, 'Contact / Support', 'Manual'],
  [/\/sitemap\//, 'Sitemap', 'Automated'],
  [/\/faq\//, 'FAQ Page', 'Assisted'],
  [/\/info\/index/, 'Info Index', 'Automated'],
  [/\/service\//, 'Service Page (legacy en)', 'Assisted'],
  [/\/announcement\//, 'Announcement Page', 'Assisted'],
  [/\/(en|ja)\/index\.html$/, 'Homepage (locale)', 'Manual'],
  [/\/index\.html$/, 'Section Index', 'Assisted'],
  [/\/(corporate_site|entertainment|gourmet|insurance|point|travel|usage|cardlineup|campaign)\.html$/, 'Section Landing Page', 'Assisted'],
  [/\.html$/, 'Other Content Page', 'Assisted'],
];

const cats = {};
const tagByCat = {};
const examples = {};
const unmatched = [];
for (const url of lines) {
  let matched = false;
  for (const [re, cat, tag] of rules) {
    if (re.test(url)) {
      cats[cat] = (cats[cat] || 0) + 1;
      tagByCat[cat] = tag;
      (examples[cat] = examples[cat] || []).push(url);
      matched = true;
      break;
    }
  }
  if (!matched) { unmatched.push(url); }
}

const sorted = Object.entries(cats).sort((a, b) => b[1] - a[1]);
console.log('TOTAL classified:', lines.length, '| categories:', sorted.length, '| unmatched:', unmatched.length);
console.log('');
for (const [cat, n] of sorted) {
  console.log(String(n).padStart(4), '|', tagByCat[cat].padEnd(22), '|', cat);
}
if (unmatched.length) {
  console.log('\n--- UNMATCHED ---');
  unmatched.slice(0, 40).forEach(u => console.log('   ', u));
}
fs.writeFileSync(process.argv[3] || 'classification.json',
  JSON.stringify({ total: lines.length, cats, tagByCat, examples, unmatched }, null, 2));
