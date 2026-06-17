const fs = require('fs');
const lines = fs.readFileSync(process.argv[2], 'utf8').trim().split('\n').filter(Boolean);

// Order matters (first match wins). Diners Club Japan (AEM, same SMTC platform as sumitclub).
const rules = [
  // --- Non-page assets / fragments ---
  [/\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/honninkakunin\/include\//, 'Verification Partial Fragment', 'Exclude (not a page)'],
  [/\/travel-guides\/assets\/ssi\//, 'SSI Include Fragment', 'Exclude (not a page)'],
  [/\/(_offer\d*|_header(_main)?|_footer|_contents\d*|_mainvisual|_service\d+|parts|qs_parts|point_parts)\b/, 'LP Partial Fragment', 'Exclude (not a page)'],
  [/\/qs\/(notice|code|index|index2|biz_|dpc_|mail_|a_diners|business|regular|dinersclub|point_parts)/, 'LP Partial Fragment', 'Exclude (not a page)'],
  [/\/qs_parts\//, 'LP Partial Fragment', 'Exclude (not a page)'],
  [/\/js\/.*\/(test|hash-file)\.html$/, 'Test/Dev Stub', 'Exclude (not a page)'],
  [/\/bumper\/link\.html$/, 'Bumper/Redirect Stub', 'Exclude (not a page)'],
  [/\/popup\/error\.html$/, 'Modal/Popup Fragment', 'Exclude (not a page)'],
  [/\/include\/(modal-window|virtual-numpad|btn-online)/, 'Verification Partial Fragment', 'Exclude (not a page)'],
  [/\/iframe\/blank/, 'Blank/Helper Stub', 'Exclude (not a page)'],
  [/google[0-9a-f]{10,}\.html$/, 'Site-verification Stub', 'Exclude (not a page)'],
  [/\/(0vyahwt15y|4jl3puk3fx|4z7czfm9uh|XnG7gL1vXs|bmfqqrplq6|vmbqes3cu5|yi6kdzujjz|hbapn60hafb69yi7ir473jyrgo889y)\.html$/, 'Random-hash Stub', 'Exclude (not a page)'],

  // --- Redirect / system ---
  [/\/(rp_rd|rp\d+_rd|rp\d+|premium\d+|premium\d+_rd|premium_app_rd[a-z_]*|premium_rd|premium_serviceguide_dpc)\.html$/, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/rd\//, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/(rd_jal|rd_jal_lineup|rd_reg_signup|rd_signon)\.html$/, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/to\/redirect_/, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/thankyou_mail\//, 'Thank-you / Mail Stub', 'Exclude / Re-point'],
  [/\/signon\//, 'Sign-on / Login System', 'Exclude / Re-point'],
  [/\/golf\/gdo\/login|gdo_reserve/, 'External Booking Redirect', 'Exclude / Re-point'],
  [/\/WD1100112\.html$/, 'Redirect / Router Stub', 'Exclude / Re-point'],
  [/\/(400|401|403|404|500|501|502|503)\.html$/, 'HTTP Error Page', 'Manual / Replace'],

  // --- Auth-gated member area (premium/member) ---
  [/\/premium\/member\//, 'Premium Member Page (auth-gated)', 'Manual / Auth-gated'],
  [/\/premium\/(member_rp|loginErr|card)\.html$/, 'Premium Member Page (auth-gated)', 'Manual / Auth-gated'],
  [/\/premium\.html$/, 'Premium Member Landing', 'Manual / Auth-gated'],

  // --- Application landing pages (entry_form) ---
  [/\/entry_form\/corporate\//, 'Corporate Card SEO/Oyakudachi Article', 'Assisted'],
  [/\/entry_form\/(lp|campaign|dlpo\d+|smtb|entrance)\b/, 'Card Application Landing Page (LP)', 'Manual'],
  [/\/entry_form\//, 'Card Application Landing Page (LP)', 'Manual'],

  // --- Verification microsites ---
  [/\/honninkakunin/, 'Identity Verification Microsite', 'Manual'],
  [/\/(corp-verify|how-verify|before-verify|id-doc)\.html$/, 'Identity Verification Microsite', 'Manual'],
  [/\/taikai\//, 'Cancellation/Termination Flow', 'Manual'],

  // --- Lounge display microsites ---
  [/\/(ginzalounge|osakalounge)\/display\//, 'Lounge Display Screen', 'Manual'],

  // --- Content families ---
  // Magazine
  [/\/magazine\/(all|login)\.html$/, 'Magazine Index/Listing', 'Assisted'],
  [/\/magazine\.html$/, 'Magazine Index/Listing', 'Assisted'],
  [/\/magazine\/article\/[a-z_]+\.html$/, 'Magazine Category Index', 'Assisted'],
  [/\/magazine\/article\//, 'Magazine Article', 'Automated'],
  // Press / merchant press
  [/\/merchant\/press\//, 'Press / News Detail', 'Automated'],
  [/\/press\/(maintenance|phishing|email_address)\.html$/, 'Press / News Detail', 'Automated'],
  [/\/press\.html$/, 'Press / News Index', 'Automated'],
  [/\/press\//, 'Press / News Detail', 'Automated'],
  // Events
  [/\/event\/report\//, 'Event Report Detail', 'Automated'],
  [/\/event\/(report|nav|nav_banner)\.html$/, 'Event Index/Nav', 'Assisted'],
  [/\/event\.html$/, 'Event Index/Nav', 'Assisted'],
  [/\/event\//, 'Event Detail', 'Assisted'],
  [/\/corporate\/privilege\/business_service\/event\/report\//, 'Event Report Detail', 'Automated'],
  // Ginza restaurant microsite
  [/\/ginzarestaurant\//, 'Ginza Restaurant Shop Page', 'Assisted'],
  // Travel guides microsite
  [/\/travel-guides\//, 'Travel Guide Article', 'Assisted'],
  // Cards
  [/\/cardlineup\/(comparison|status|nyukai|etc|family|loancard|revo)\b/, 'Card Lineup Listing/Index', 'Assisted'],
  [/\/cardlineup(\.html)?$/, 'Card Lineup Listing/Index', 'Assisted'],
  [/\/cardlineup\//, 'Card Product Detail', 'Manual'],
  [/\/corporate\/cardlineup\//, 'Corporate Card Detail', 'Manual'],
  [/\/corporate\/cardlineup(\.html)?$/, 'Corporate Card Listing', 'Assisted'],
  // Corporate (business site sections)
  [/\/corporate\/ccol\//, 'Corporate Club Online How-to', 'Assisted'],
  [/\/corporate\/privilege\//, 'Corporate Privilege Page', 'Assisted'],
  [/\/corporate\/sitemap\.html$/, 'Sitemap', 'Automated'],
  [/\/corporate\/(businesspoint|finance|insurance|honninkakunin)\.html$/, 'Corporate Service Page', 'Assisted'],
  [/\/corporate\/privilege\.html$/, 'Corporate Privilege Page', 'Assisted'],
  [/\/corporate\.html$/, 'Corporate Top/Landing', 'Assisted'],
  [/\/corporate\//, 'Corporate Service Page', 'Assisted'],
  // Service / how-to families (standard AEM content shell)
  [/\/cpn_evt\//, 'Campaign/Event Landing', 'Assisted'],
  [/\/cpn_evt\.html$/, 'Campaign/Event Index', 'Assisted'],
  [/\/(travel|gourmet|golf|lifestyle|kameiten|shopping|point|payment|finance|insurance|usage|sponsorship|privilege)\//, 'Service / How-to / Category Page', 'Assisted'],
  [/\/(travel|gourmet|golf|lifestyle|kameiten|kameiten_overseas|shopping|point|payment|finance|insurance|usage|sponsorship|privilege|community|kyoto|cm)\.html$/, 'Section Landing Page', 'Assisted'],
  [/\/about\//, 'About / History Page', 'Assisted'],
  [/\/about\.html$/, 'Section Landing Page', 'Assisted'],
  [/\/benefit\/(detail|expired|stock)\.html$/, 'Benefit Detail/Listing', 'Assisted'],
  [/\/benefit\.html$/, 'Benefit Index', 'Assisted'],
  [/\/merchant\/(faq)\.html$/, 'FAQ Page', 'Assisted'],
  [/\/merchant\.html$/, 'Merchant Top/Landing', 'Assisted'],
  [/\/merchant\//, 'Merchant Service Page', 'Assisted'],
  [/\/unique\//, 'Club Online Promo (unique)', 'Assisted'],
  // Legal / policy
  [/\/(privacy|privacy_cic|privacy_kzk|privacy_law|smallprint|tc|policy_customerresponse|cnasp|aasfp|signature_essay|pvt|email)\b/, 'Legal / Policy Text', 'Automated'],
  // Contact
  [/\/contact\b/, 'Contact / Support', 'Manual'],
  [/\/faq\//, 'FAQ Page', 'Assisted'],
  [/\/topic\//, 'Topic Page', 'Assisted'],
  // Sitemap / search / home
  [/\/Sitemap\.html$/i, 'Sitemap', 'Automated'],
  [/\/search\.html$/, 'Search Page', 'Manual'],
  [/\/(ja\/)?index\.html$/, 'Homepage', 'Manual'],
  [/\/(company|biz|BMW|jc)\.html$/, 'Short-name Landing/Redirect', 'Assisted'],
  // generic catch-alls
  [/\/index\.html$/, 'Section Index', 'Assisted'],
  [/\.html$/, 'Other Content Page', 'Assisted'],
];

const cats = {}, tagByCat = {}, examples = {}, unmatched = [];
const perUrl = [];
for (const url of lines) {
  let matched = false;
  for (const [re, cat, tag] of rules) {
    if (re.test(url)) {
      cats[cat] = (cats[cat] || 0) + 1;
      tagByCat[cat] = tag;
      (examples[cat] = examples[cat] || []).push(url);
      perUrl.push({ url, category: cat, mode: tag });
      matched = true; break;
    }
  }
  if (!matched) { unmatched.push(url); perUrl.push({ url, category: 'UNMATCHED', mode: 'Review' }); }
}
fs.writeFileSync('per-url.json', JSON.stringify(perUrl, null, 2));
fs.writeFileSync('classification.json', JSON.stringify({ total: lines.length, cats, tagByCat, examples, unmatched }, null, 2));

const sorted = Object.entries(cats).sort((a, b) => b[1] - a[1]);
console.log('TOTAL:', lines.length, '| categories:', sorted.length, '| unmatched:', unmatched.length);
for (const [c, n] of sorted) console.log(String(n).padStart(5), '|', tagByCat[c].padEnd(22), '|', c);
if (unmatched.length) { console.log('\n--- UNMATCHED (first 30) ---'); unmatched.slice(0, 30).forEach(u => console.log('  ', u)); }
