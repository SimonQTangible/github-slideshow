// B-Catcher Group investor deck - every number is read from the recalculated model (numbers.json).
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");
const { applyTheme } = require("/root/.claude/skills/synced/e3a10683-0f6f-4df7-9d97-1f3de7e0ea41_a12ffef0-b97c-4fce-adde-69bad4852925/pptx/scripts/apply_theme.js");

const D = JSON.parse(fs.readFileSync(process.argv[2], "utf8"));
const OUT = process.argv[3];

const THEME = {
  name: "B-Catcher Hive",
  headFontFace: "Calibri",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1B1F2A", lt1: "FFFFFF", dk2: "3A4152", lt2: "F3F4F7",
    accent1: "E8A317", accent2: "1F3A5F", accent3: "2A7F8A", accent4: "C0504D",
    accent5: "8C93A3", accent6: "F6D58E", hlink: "2A7F8A", folHlink: "1F3A5F",
  },
};
const HEX = { ink: "1B1F2A", slate: "3A4152", mist: "F3F4F7", honey: "E8A317", navy: "1F3A5F", teal: "2A7F8A", red: "C0504D", grey: "8C93A3", pale: "F6D58E", white: "FFFFFF" };

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.33 x 7.5
pres.title = "B-Catcher Group - Investor Presentation";
pres.company = "B-Catcher Group";
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
const C = pres.SchemeColor;

// ---------- helpers
const m = (v, d = 1) => (v / 1e6).toFixed(d);
const $m = (v, d = 1) => (v < 0 ? "-$" : "$") + Math.abs(v / 1e6).toFixed(d) + "m";
const $k = (v) => "$" + Math.round(v / 1000).toLocaleString("en-AU") + "k";
const pct = (v, d = 0) => (v * 100).toFixed(d) + "%";
const YRS = ["2027", "2028", "2029", "2030", "2031"];
const FOOT = "B-Catcher Group  |  Confidential planning document  |  Not an offer or financial advice";

// ---------- layouts
pres.defineSlideMaster({
  title: "DARK",
  background: { color: HEX.ink },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.8, y: 1.9, w: 11.7, h: 1.6, fontSize: 44, bold: true, color: C.background1, valign: "bottom", align: "left", margin: 0 }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: 0.8, y: 3.75, w: 11.0, h: 1.0, fontSize: 20, color: C.accent6, valign: "top", align: "left", margin: 0 }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: HEX.white },
  margin: [0.5, 0.6, 0.6, 0.6],
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: 0.6, y: 0.35, w: 12.1, h: 0.75, fontSize: 30, bold: true, color: C.text1, valign: "top", align: "left", margin: 0 }, text: "" } },
    { text: { text: FOOT, options: { x: 0.6, y: 7.05, w: 10, h: 0.3, fontSize: 9, color: C.accent5, margin: 0 } } },
  ],
  slideNumber: { x: 12.3, y: 7.05, w: 0.5, h: 0.3, fontSize: 9, color: C.accent5, align: "right" },
});

function content(section, title, lead) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: section });
  s.addText(title, { placeholder: "title" });
  if (lead) s.addText(lead, { x: 0.6, y: 1.1, w: 12.1, h: 0.5, fontSize: 15, color: C.text2, margin: 0, isTextBox: true, objectName: "lead" });
  return s;
}
function card(s, x, y, w, h, fill, name) {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, rectRadius: 0.08, fill: { color: fill }, line: { color: fill }, objectName: name });
}
function hexBadge(s, x, y, size, fill, txt, txtColor, name) {
  s.addShape(pres.shapes.HEXAGON, { x, y, w: size, h: size * 0.88, fill: { color: fill }, line: { color: fill }, objectName: name + "-hex" });
  s.addText(txt, { x, y, w: size, h: size * 0.88, align: "center", valign: "middle", fontSize: Math.max(12, size * 22), bold: true, color: txtColor, margin: 0, isTextBox: true, objectName: name + "-txt" });
}
function stat(s, x, y, w, big, small, colBig, colSmall, name) {
  s.addText(big, { x, y, w, h: 0.85, fontSize: 40, bold: true, color: colBig, margin: 0, isTextBox: true, objectName: name + "-big" });
  s.addText(small, { x, y: y + 0.85, w, h: 0.6, fontSize: 13, color: colSmall, margin: 0, isTextBox: true, valign: "top", objectName: name + "-small" });
}
const chartBase = {
  catAxisLabelColor: "3A4152", valAxisLabelColor: "3A4152", catAxisLabelFontSize: 11, valAxisLabelFontSize: 10,
  catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt", dataLabelFontFace: "+mn-lt", legendFontFace: "+mn-lt", titleFontFace: "+mn-lt",
  valGridLine: { color: "E3E5EA", size: 0.5 }, catGridLine: { style: "none" }, dataLabelFontSize: 10, dataLabelColor: "3A4152",
  legendFontSize: 11, legendColor: "3A4152", titleFontSize: 13, titleColor: "1B1F2A",
};

// =====================================================================================
// 1 COVER
pres.addSection({ title: "Opening" });
let s = pres.addSlide({ masterName: "DARK", sectionTitle: "Opening" });
s.addText("B-Catcher Group", { placeholder: "title" });
s.addText("One controlled journey from broker introduction to investor ownership", { placeholder: "body" });
[[0.8, HEX.navy, "B"], [1.55, HEX.honey, "H"], [2.3, HEX.teal, "B"]].forEach(([x, col, t], i) =>
  hexBadge(s, x, 0.9, 0.62, col, t, i === 1 ? HEX.ink : HEX.white, "cover-hex" + i));
s.addText("B-Catcher  ·  Honey Pot  ·  BHive", { x: 3.15, y: 0.95, w: 6, h: 0.5, fontSize: 16, color: C.accent5, margin: 0, isTextBox: true, objectName: "cover-brands" });
s.addText([
  { text: "Investor presentation  |  December 2026  |  Seed round", options: { breakLine: true } },
  { text: "Confidential planning document. Not an offer, PDS or financial advice. Regulated services require licences and partner agreements before launch." },
], { x: 0.8, y: 6.2, w: 11.5, h: 0.8, fontSize: 12, color: C.accent5, margin: 0, isTextBox: true, objectName: "cover-foot" });
s.addNotes("Open with the one-line thesis: three companies, one controlled lifecycle. Investors subscribe directly through BHive; Honey Pot runs brokers and distribution behind it; B-Catcher originates and manages the investments.");

// 2 PROBLEM
pres.addSection({ title: "Opportunity" });
s = content("Opportunity", "Private markets are hard to reach, hard to sell and hard to exit",
  "Wholesale investors want property, private credit and pre-IPO exposure; the path to it is fragmented and manual");
const pains = [
  ["Access", "Quality private deals reach investors through informal networks, paper applications and inconsistent eligibility checks."],
  ["Distribution", "Brokers and advisers lack one compliant channel: no shared offer library, attribution or commission ledger."],
  ["Liquidity", "Once invested, holders have no visibility of their portfolio and no structured way to transfer interests."],
];
pains.forEach(([h, t], i) => {
  const x = 0.6 + i * 4.1;
  card(s, x, 1.9, 3.8, 3.9, HEX.mist, "pain-card" + i);
  hexBadge(s, x + 0.35, 2.2, 0.75, HEX.honey, String(i + 1), HEX.ink, "pain" + i);
  s.addText(h, { x: x + 0.35, y: 3.05, w: 3.1, h: 0.55, fontSize: 22, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "pain-h" + i });
  s.addText(t, { x: x + 0.35, y: 3.65, w: 3.1, h: 1.9, fontSize: 15, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "pain-t" + i });
});
s.addText("Market sizing and competitor evidence to be independently verified before external circulation (Deck Content Brief).", { x: 0.6, y: 6.15, w: 12.1, h: 0.4, fontSize: 11, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "pain-note" });

// 3 SOLUTION - three companies
pres.addSection({ title: "Business model" });
s = content("Business model", "Three companies, one lifecycle",
  "Each entity owns one job and earns its own revenue - so each can be licensed, staffed and valued on its own");
const ents = [
  [HEX.navy, HEX.white, "B-Catcher", "Parent - originate & manage", ["Sources and structures property trusts, private credit, PE / pre-IPO", "Investment committee, asset management, group finance and risk", "Earns entry and management fees (BHive fee schedule)"]],
  [HEX.honey, HEX.ink, "Honey Pot", "Distribution & broker network", ["Recruits, accredits and supervises brokers", "Approves offers for distribution; owns lead attribution", "Earns placement fees; pays broker commissions"]],
  [HEX.teal, HEX.white, "BHive", "Investor platform & marketplace", ["Onboarding, KYC, wholesale verification, e-sign", "Subscriptions, wallet, reporting, secondary market", "Earns technology, onboarding, exchange and buy-back resale fees"]],
];
ents.forEach(([fill, txt, name, role, pts], i) => {
  const x = 0.6 + i * 4.1;
  card(s, x, 1.85, 3.8, 1.25, fill, "ent-head" + i);
  s.addText(name, { x: x + 0.3, y: 1.95, w: 3.3, h: 0.55, fontSize: 24, bold: true, color: txt, margin: 0, isTextBox: true, objectName: "ent-name" + i });
  s.addText(role, { x: x + 0.3, y: 2.5, w: 3.3, h: 0.45, fontSize: 14, color: txt, margin: 0, isTextBox: true, objectName: "ent-role" + i });
  card(s, x, 3.2, 3.8, 2.75, HEX.mist, "ent-body" + i);
  s.addText(pts.map((p, j) => ({ text: p, options: { bullet: true, breakLine: j < pts.length - 1 } })),
    { x: x + 0.25, y: 3.35, w: 3.35, h: 2.5, fontSize: 14, color: C.text1, paraSpaceAfter: 8, valign: "top", margin: 0, isTextBox: true, objectName: "ent-pts" + i });
});
s.addText("Investors subscribe directly through BHive. Honey Pot manages brokers and distribution behind that process. B-Catcher originates and manages the investments.",
  { x: 0.6, y: 6.15, w: 12.1, h: 0.5, fontSize: 14, bold: true, color: C.accent2, margin: 0, isTextBox: true, objectName: "ent-foot" });

// 4 STRUCTURE
s = content("Business model", "Client money and assets sit outside the operating group",
  "Each investment is a separate trust or SPV; licensed third parties hold assets and run regulated roles until the group holds its own permissions");
const box = (x, y, w, h, fill, t1, t2, tc, name) => {
  card(s, x, y, w, h, fill, name);
  s.addText([{ text: t1, options: { bold: true, fontSize: 16, breakLine: !!t2 } }].concat(t2 ? [{ text: t2, options: { fontSize: 12 } }] : []),
    { x, y, w, h, align: "center", valign: "middle", color: tc, margin: 4, isTextBox: true, objectName: name + "-t" });
};
box(4.9, 1.85, 3.5, 0.95, HEX.navy, "B-Catcher Group (parent)", "Strategy · origination · asset management", HEX.white, "st-parent");
box(1.4, 3.35, 3.5, 0.95, HEX.honey, "Honey Pot Pty Ltd", "CAR -> own AFSL (yr 2) · brokers", HEX.ink, "st-hp");
box(8.4, 3.35, 3.5, 0.95, HEX.teal, "BHive Pty Ltd", "Platform · wallet · marketplace tech", HEX.white, "st-bh");
box(4.9, 3.35, 3.5, 0.95, HEX.mist, "Investment trusts / SPVs", "Hold assets · issue units · trust debt", HEX.ink, "st-spv");
[[6.65, 2.8, 0, 0.55], [3.15, 2.8, 0, 0.55], [10.15, 2.8, 0, 0.55]].forEach(([x, y, w, h], i) =>
  s.addShape(pres.shapes.LINE, { x, y, w, h, line: { color: HEX.grey, width: 1.5, endArrowType: "triangle" }, objectName: "st-arrow" + i }));
s.addShape(pres.shapes.LINE, { x: 3.15, y: 2.8, w: 7.0, h: 0, line: { color: HEX.grey, width: 1.5 }, objectName: "st-bus" });
const ext = [["Principal licensee / RE", "regulated roles until own AFSL"], ["Trustee & custodian", "holds legal title and client money"], ["Fund admin & registry", "authoritative unit register"], ["Licensed market operator", "executes secondary trades"]];
ext.forEach(([a, b], i) => box(0.6 + i * 3.1, 4.9, 2.9, 1.05, HEX.white, a, b, HEX.ink, "st-ext" + i));
ext.forEach((_, i) => s.addShape(pres.shapes.RECTANGLE, { x: 0.6 + i * 3.1, y: 4.9, w: 2.9, h: 1.05, fill: { type: "none" }, line: { color: HEX.grey, width: 1, dashType: "dash" }, objectName: "st-extline" + i }));
s.addText("External regulated counterparties (contracted)", { x: 0.6, y: 4.5, w: 8, h: 0.35, fontSize: 12, bold: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "st-extlabel" });
s.addText("Subscriptions and trust assets are never group revenue or group cash.", { x: 0.6, y: 6.2, w: 12, h: 0.4, fontSize: 13, italic: true, color: C.text2, margin: 0, isTextBox: true, objectName: "st-note" });

// 5 JOURNEY
s = content("Business model", "The investor journey happens entirely inside BHive",
  "Broker or direct campaign  ->  verified, documented subscription  ->  reconciled holding  ->  optional secondary sale");
const steps = ["Referral via broker link or campaign", "Identity, AML/KYC, tax, consents", "Wholesale classification approved", "Review approved offer documents", "Apply and e-sign in BHive",
  "Funds to trust / custody account", "Issuer allots; register reconciled", "Holdings, distributions, tax statements", "List on secondary market (when licensed)"];
steps.forEach((t, i) => {
  const col = i % 3, rw = Math.floor(i / 3);
  const x = 0.6 + col * 4.1, y = 1.85 + rw * 1.5;
  const fill = rw === 0 ? HEX.honey : rw === 1 ? HEX.navy : HEX.teal;
  hexBadge(s, x, y + 0.1, 0.8, fill, String(i + 1), rw === 0 ? HEX.ink : HEX.white, "j" + i);
  s.addText(t, { x: x + 0.95, y: y + 0.05, w: 2.9, h: 0.9, fontSize: 15, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "jt" + i });
});
s.addText([{ text: "Who does what:  ", options: { bold: true } }, { text: "Honey Pot approves eligibility and offer access (steps 2-4); the issuer/RE accepts and allots (7); BHive displays and reconciles (8) and runs the order workflow with a licensed operator (9)." }],
  { x: 0.6, y: 6.25, w: 12.1, h: 0.55, fontSize: 13, color: C.text2, margin: 0, isTextBox: true, objectName: "j-note" });

// 6 HONEY POT ENGINE
s = content("Business model", "Honey Pot turns brokers into a scalable, supervised sales channel",
  `${pct(D.inputs["Share of capital sourced via brokers"])} of capital is modelled through brokers, who receive ${pct(D.inputs["Broker share of placement fee (commission)"])} of the placement fee on cleared subscriptions`);
s.addChart(pres.charts.BAR, [{ name: "Active accredited brokers", labels: YRS, values: D.brk.map(Math.round) }], {
  ...chartBase, x: 0.6, y: 1.8, w: 6.3, h: 4.6, barDir: "col", chartColors: [HEX.honey], showValue: true, dataLabelPosition: "outEnd",
  showTitle: true, title: "Active accredited brokers (year end)", showLegend: false, valAxisHidden: true, objectName: "hp-chart",
});
const life = ["Recruit & verify AFSL / ACL status", "Classify: AR, licensed distributor or referrer", "Accredit, train, grant BHive broker portal", "Attribute leads with unique referral links", "Pay commissions only on cleared, accepted money", "Supervise, sample files, annual review"];
card(s, 7.3, 1.8, 5.4, 4.6, HEX.mist, "hp-life-card");
s.addText("Broker lifecycle", { x: 7.6, y: 1.95, w: 4.8, h: 0.45, fontSize: 17, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "hp-life-h" });
s.addText(life.map((t, i) => ({ text: t, options: { bullet: { type: "number" }, breakLine: i < life.length - 1 } })),
  { x: 7.6, y: 2.45, w: 4.9, h: 3.8, fontSize: 15, color: C.text1, paraSpaceAfter: 9, valign: "top", margin: 0, isTextBox: true, objectName: "hp-life" });

// 7 BHIVE MARKETPLACE
s = content("Business model", "BHive's secondary market gives holders a structured route to exit",
  "Built and tested before launch; switched on only with a licensed market-operator partner and settlement in place");
const mk = [["Verified holders list", "Eligible units, transfer restrictions checked automatically"], ["Disclosed pricing", "Bids, offers and dated valuations shown to verified buyers"],
  ["Licensed execution", "Matching by authorised operator; partner takes " + pct(D.inputs["Licensed market-operator partner share"]) + " of fees"], ["Register + wallet update together", "Settlement reconciled to the legal register; full audit trail"]];
mk.forEach(([h, t], i) => {
  const x = 0.6 + i * 3.1;
  card(s, x, 1.9, 2.9, 2.6, i % 2 ? HEX.mist : "E8F2F3", "mk-card" + i);
  s.addText(String(i + 1), { x: x + 0.25, y: 2.0, w: 0.6, h: 0.6, fontSize: 28, bold: true, color: C.accent3, margin: 0, isTextBox: true, objectName: "mk-n" + i });
  s.addText(h, { x: x + 0.25, y: 2.6, w: 2.45, h: 0.65, fontSize: 16, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "mk-h" + i });
  s.addText(t, { x: x + 0.25, y: 3.3, w: 2.45, h: 1.1, fontSize: 13, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "mk-t" + i });
});
stat(s, 0.6, 4.9, 3.8, $m(D.trade[4]), "value traded in 2031 at " + pct(D.inputs["Secondary market turnover"]) + " annual turnover", C.accent3, C.text2, "mk-s1");
stat(s, 4.7, 4.9, 3.8, pct(D.inputs["Exchange fee per trade (BHive schedule)"], 2), "exchange fee per trade (BHive schedule)", C.accent3, C.text2, "mk-s2");
stat(s, 8.8, 4.9, 3.9, "Buy-back ladder", "Tokens: 25% / 15% / 10% below valuation (yrs 3-5). Trust units: 15% / 10% / 8%. Funded reserve held in cash", C.accent2, C.text2, "mk-s3");

// 8 PRODUCTS & FEES
pres.addSection({ title: "Economics" });
s = content("Economics", "Four product lines, priced at market and paid by the issuing vehicle",
  "BHive's investor fee schedule is kept; Honey Pot's placement fee pays broker introduction commissions");
const I = D.inputs;
const prows = [
  ["Product", "Mix", "Honey Pot placement fee", "B-Catcher upfront", "B-Catcher ongoing", "BHive"],
  ["Property trusts / tokens", pct(I["Residential & commercial property trusts"]), pct(I["Placement fee - property trusts"], 1), pct(I["Property entry fee (BHive schedule)"], 1) + " entry fee", pct(I["Management fee - property (BHive schedule)"], 2) + " p.a.", "rowspan"],
  ["Private credit", pct(I["Private credit"]), pct(I["Placement fee - private credit"], 1), pct(I["Entry fee - credit & PE products (BHive schedule)"], 1) + " entry fee", pct(I["Net management fee - private credit"], 2) + " p.a.", ""],
  ["Private equity / pre-IPO", pct(I["Private company equity / pre-IPO (incl. UTEGBO-style unit trusts)"]), pct(I["Placement fee - private equity / pre-IPO"], 1), pct(I["Entry fee - credit & PE products (BHive schedule)"], 1) + " entry fee", pct(I["Net management fee - private equity"], 2) + " p.a.", ""],
  ["ASX placements / IPO", pct(I["ASX placements / IPO allocations (third-party issuers)"]), pct(I["Placement fee - ASX placements / IPO"], 1), "-", "-", ""],
];
const bhiveCell = pct(I["Technology / platform fee (BHive schedule)"], 2) + " technology fee p.a. + $" + Math.round(I["Offer onboarding fee (data room, workflows, registry set-up)"] / 1000) + "k per offer + " + pct(I["Exchange fee per trade (BHive schedule)"], 2) + " exchange fee";
const tbl = prows.map((r, i) => r.filter(c => c !== "").map((c, j) => {
  if (c === "rowspan") return { text: bhiveCell, options: { rowspan: 4, valign: "middle", fontSize: 13, color: HEX.ink, fill: { color: "E8F2F3" } } };
  return { text: c, options: i === 0 ? { bold: true, color: HEX.white, fill: { color: HEX.ink }, fontSize: 13 } : { fontSize: 13, color: HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist }, bold: j === 0 } };
}));
s.addTable(tbl, { x: 0.6, y: 1.85, w: 12.1, colW: [3.0, 0.9, 2.0, 2.2, 2.0, 2.0], rowH: 0.62, border: { type: "solid", color: "D9DCE3", pt: 0.75 }, valign: "middle", objectName: "fee-table" });
s.addText(`Broker commissions: ${pct(I["Broker share of placement fee (commission)"])} of placement fee on broker-sourced capital. Average investor ticket $${Math.round(I["Average investor subscription"] / 1000)}k. Buy-back ladder applies to property tokens and trust units. Performance fees excluded.`,
  { x: 0.6, y: 5.35, w: 12.1, h: 0.7, fontSize: 13, color: C.text2, margin: 0, isTextBox: true, objectName: "fee-note" });

// 9 REVENUE BY ENTITY
s = content("Economics", "Revenue shifts from one-off placement fees to recurring AUM fees",
  `Group revenue grows to ${$m(D.rev[4])} in 2031 as managed AUM reaches ${$m(D.aum[4], 0)}; Honey Pot falls from ${pct(D.rhp[1] / D.rev[1])} to ${pct(D.rhp[4] / D.rev[4])} of revenue`);
s.addChart(pres.charts.BAR, [
  { name: "Honey Pot - distribution", labels: YRS, values: D.rhp.map(v => +m(v, 2)) },
  { name: "B-Catcher - investment & asset mgmt", labels: YRS, values: D.rbc.map(v => +m(v, 2)) },
  { name: "BHive - platform & marketplace", labels: YRS, values: D.rbh.map(v => +m(v, 2)) },
], { ...chartBase, x: 0.6, y: 1.75, w: 8.0, h: 4.95, barDir: "col", barGrouping: "stacked", chartColors: [HEX.honey, HEX.navy, HEX.teal],
  showLegend: true, legendPos: "b", showTitle: true, title: "External revenue by entity ($m)", valAxisLabelFormatCode: "$0", showValue: false, objectName: "rev-chart" });
card(s, 9.0, 1.75, 3.7, 4.95, HEX.mist, "rev-card");
stat(s, 9.3, 1.95, 3.2, $m(D.rev[4]), "group net revenue 2031", C.text1, C.text2, "rev-s1");
stat(s, 9.3, 3.5, 3.2, $m(D.aum[4], 0), "managed AUM end 2031 (held in trusts, not group assets)", C.accent2, C.text2, "rev-s2");
stat(s, 9.3, 5.05, 3.2, Math.round(D.inv[4]).toLocaleString("en-AU"), "wholesale investors on BHive", C.accent3, C.text2, "rev-s3");

// 10 FORECAST P&L TABLE
pres.addSection({ title: "Financials" });
s = content("Financials", `Five-year forecast: EBITDA positive from ${D.km.ebm}`,
  "Base case, consolidated group, intercompany fees eliminated. All costs included: people, set-up, licensing, technology, distribution, overheads, contingency");
const pl = [["$m", ...YRS],
  ["Capital placed through BHive", ...D.cap.map(v => m(v))], ["Managed AUM (year end)", ...D.aum.map(v => m(v))],
  ["Net revenue", ...D.rev.map(v => m(v))], ["Cost of sales (broker commissions etc.)", ...D.cogs.map(v => m(-v))], ["Gross profit", ...D.gp.map(v => m(v))],
  ["Staff costs", ...D.staff.map(v => m(-v))], ["Other operating costs", ...D.opex.map((v, i) => m(-(v - D.staff[i] - D.setup[i])))], ["One-off set-up costs (expensed)", ...D.setup.map(v => m(-v))],
  ["EBITDA", ...D.ebitda.map(v => m(v))], ["Net profit after tax", ...D.npat.map(v => m(v))], ["Closing cash", ...D.cash.map(v => m(v))], ["Group FTE (year end)", ...D.fte.map(v => v.toFixed(0))]];
const bold = new Set([0, 3, 5, 9, 10]);
s.addTable(pl.map((r, i) => r.map((c, j) => ({ text: String(c).replace(/^-(\d)/, "($1").replace(/^\((.*)$/, "($1)"), options: {
  bold: bold.has(i), fontSize: 13, align: j ? "right" : "left", color: i === 0 ? HEX.white : HEX.ink,
  fill: { color: i === 0 ? HEX.ink : i === 9 ? HEX.pale : (i % 2 ? HEX.white : HEX.mist) } } }))),
  { x: 0.6, y: 1.8, w: 12.1, colW: [4.6, 1.5, 1.5, 1.5, 1.5, 1.5], rowH: 0.36, border: { type: "solid", color: "E3E5EA", pt: 0.5 }, valign: "middle", objectName: "pl-table" });
s.addText("Brackets = negative. Source: BCatcher_Group_Integrated_Financial_Model.xlsx (Annual Summary).", { x: 0.6, y: 6.6, w: 12, h: 0.3, fontSize: 10, italic: true, color: C.accent5, margin: 0, isTextBox: true, objectName: "pl-src" });

// 11 COST TO BUILD
s = content("Financials", `It costs ${$m(D.km.setup, 2)} one-off to set up and build the group`,
  "Every set-up item is listed, timed and costed in the model; technology build is capitalised and amortised over 36 months");
const sc = Object.entries(D.setup_cat).sort((a, b) => b[1] - a[1]);
s.addChart(pres.charts.BAR, [{ name: "Set-up cost", labels: sc.map(x => x[0]), values: sc.map(x => Math.round(x[1] / 1000)) }], {
  ...chartBase, x: 0.6, y: 1.75, w: 7.6, h: 4.95, barDir: "bar", chartColors: [HEX.navy], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$#,##0\"k\"",
  showLegend: false, showTitle: true, title: "One-off set-up by category ($k)", valAxisHidden: true, catAxisOrientation: "maxMin", objectName: "su-chart" });
card(s, 8.6, 1.75, 4.1, 4.95, HEX.mist, "su-card");
s.addText("By entity", { x: 8.9, y: 1.9, w: 3.5, h: 0.45, fontSize: 17, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "su-h" });
[["B-Catcher", HEX.navy], ["Honey Pot", HEX.honey], ["BHive", HEX.teal]].forEach(([e, col], i) => {
  s.addShape(pres.shapes.OVAL, { x: 8.9, y: 2.55 + i * 0.75, w: 0.3, h: 0.3, fill: { color: col }, line: { color: col }, objectName: "su-dot" + i });
  s.addText([{ text: e + "  ", options: { bold: true } }, { text: $k(D.setup_ent[e] || 0) }], { x: 9.35, y: 2.45 + i * 0.75, w: 3.2, h: 0.5, fontSize: 15, color: C.text1, margin: 0, isTextBox: true, objectName: "su-e" + i });
});
s.addText("Largest items: MVP build partner ($480k), secondary market engine ($350k), own AFSL application ($150k), mobile apps ($180k), market-operator legal ($120k). Includes 10% contingency.",
  { x: 8.9, y: 4.85, w: 3.6, h: 1.7, fontSize: 13, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "su-note" });

// 12 TEAM / HEADCOUNT
s = content("Financials", `Three employees run all three companies`,
  `${D.roles.length} costed roles: 3 employees (CEO & CIO, Head of Risk & Ops, Head of Distribution) plus fractional contractors; everything else is outsourced, AI-driven or commission-only`);
s.addChart(pres.charts.BAR, [
  { name: "B-Catcher", labels: YRS, values: D.fte_bc.map(v => +v.toFixed(1)) },
  { name: "Honey Pot", labels: YRS, values: D.fte_hp.map(v => +v.toFixed(1)) },
  { name: "BHive", labels: YRS, values: D.fte_bh.map(v => +v.toFixed(1)) },
], { ...chartBase, x: 0.6, y: 1.75, w: 6.6, h: 4.95, barDir: "col", barGrouping: "stacked", chartColors: [HEX.navy, HEX.honey, HEX.teal],
  showLegend: true, legendPos: "b", showTitle: true, title: "FTE by entity (year end)", showValue: true, dataLabelPosition: "ctr", dataLabelColor: "FFFFFF", valAxisHidden: true, objectName: "hc-chart" });
const firstHires = D.roles.filter(r => r[4] <= 7).map(r => r[2].replace(/ \(.*\)/, ""));
card(s, 7.6, 1.75, 5.1, 4.95, HEX.mist, "hc-card");
s.addText("The whole team", { x: 7.9, y: 1.9, w: 4.6, h: 0.45, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "hc-h" });
s.addText(D.roles.map(r => r[2].replace(" (contract)", "") + (r[3] === "Contractor" ? " - contract" : "")).join("\n"), { x: 7.9, y: 2.4, w: 4.6, h: 3.0, fontSize: 13, paraSpaceAfter: 6, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "hc-list" });
s.addText(`Staff cost ${$m(D.staff[0])} in 2027, ${$m(D.staff[4])} in 2031. Outsourced service desk, accountants, analysts and build partner, plus AI agents and commission-only BDMs, do the rest.`,
  { x: 7.9, y: 5.5, w: 4.6, h: 1.0, fontSize: 13, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "hc-note" });

// 13 ENTITY PROFITABILITY
s = content("Financials", "B-Catcher earns the margin; Honey Pot and BHive feed it",
  "EBITDA by cost centre including intercompany shared-services and platform licence fees ($m)");
s.addChart(pres.charts.BAR, [
  { name: "B-Catcher", labels: YRS, values: D.eb_bc.map(v => +m(v, 2)) },
  { name: "Honey Pot", labels: YRS, values: D.eb_hp.map(v => +m(v, 2)) },
  { name: "BHive", labels: YRS, values: D.eb_bh.map(v => +m(v, 2)) },
], { ...chartBase, x: 0.6, y: 1.75, w: 8.0, h: 4.95, barDir: "col", barGrouping: "clustered", chartColors: [HEX.navy, HEX.honey, HEX.teal],
  showLegend: true, legendPos: "b", showTitle: true, title: "Entity EBITDA ($m)", valAxisLabelFormatCode: "$0.0", objectName: "eb-chart" });
card(s, 9.0, 1.75, 3.7, 4.95, HEX.mist, "eb-card");
s.addText([
  { text: "Why the split matters", options: { bold: true, fontSize: 16, breakLine: true } },
  { text: "Honey Pot pays away ~" + pct(I["Share of capital sourced via brokers"] * I["Broker share of placement fee (commission)"]) + " of its fees to brokers; its value is the AUM it feeds B-Catcher.", options: { breakLine: true } },
  { text: "BHive's platform cost is mostly fixed; margin grows with AUM and secondary turnover.", options: { breakLine: true } },
  { text: "Separate P&Ls let each entity be licensed, capitalised or sold independently." },
], { x: 9.3, y: 1.95, w: 3.2, h: 4.6, fontSize: 13, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true, objectName: "eb-text" });

// 14 CASH CURVE
s = content("Financials", `Peak funding need is ${$m(D.km.peak)}; cash flow turns positive in ${D.km.cfm}`,
  "Cumulative cash before financing vs closing cash with the planned equity rounds (quarter end, $m)");
const qlab = Array.from({ length: 20 }, (_, i) => "Q" + ((i % 4) + 1) + " " + String(27 + Math.floor(i / 4)));
s.addChart(pres.charts.LINE, [
  { name: "Closing cash (with equity rounds)", labels: qlab, values: D.cash_q.map(v => +m(v, 2)) },
  { name: "Cumulative cash before financing", labels: qlab, values: D.cum_q.map(v => +m(v, 2)) },
], { ...chartBase, x: 0.6, y: 1.75, w: 8.6, h: 4.95, chartColors: [HEX.teal, HEX.red], lineSize: 3, lineDataSymbol: "none",
  showLegend: true, legendPos: "b", showTitle: false, valAxisLabelFormatCode: "$0", catAxisLabelFontSize: 9, objectName: "cash-chart" });
card(s, 9.6, 1.75, 3.1, 4.95, HEX.mist, "cash-card");
stat(s, 9.85, 1.95, 2.7, $m(D.km.peak), "peak cumulative burn before break-even", C.accent4, C.text2, "cash-s1");
stat(s, 9.85, 3.5, 2.7, $m(D.km.mincash), "lowest cash balance (buffer $0.75m)", C.text1, C.text2, "cash-s2");
stat(s, 9.85, 5.05, 2.7, D.km.ebm, "first month of positive group EBITDA", C.accent3, C.text2, "cash-s3");

// 15 SCENARIOS
s = content("Financials", "Results hinge on placement volume, not cost overruns",
  "Results from running the full model under each case. A 25% tech overrun is absorbed; a 50% volume shortfall is not");
const sr = D.scen.rows, cases = D.scen.cases;
const pick = [0, 1, 2, 3, 4, 6, 7];
const short = ["Base", "Downside", "Upside", "-50% volume", "6-mth licence delay", "+25% tech cost", "Downside + 6-mth delay"];
const metr = [["Net revenue 2031", "Net revenue Y5", v => $m(v)], ["EBITDA 2031", "EBITDA Y5", v => $m(v)], ["Peak funding need", "Peak funding requirement", v => $m(v)],
  [`Extra equity beyond ${$m(D.km.raised, 1)}`, "Extra equity needed to hold buffer", v => (v > 0 ? $m(v) : "-")], ["First EBITDA-positive month", "First month EBITDA positive", v => (typeof v === "string" && v.includes("Beyond") ? "> 2031" : v)]];
const stbl = [[{ text: "", options: { fill: { color: HEX.ink } } }].concat(short.map(h => ({ text: h, options: { bold: true, color: HEX.white, fill: { color: HEX.ink }, fontSize: 12, align: "center" } })))];
metr.forEach(([lab, key, f], i) => stbl.push([{ text: lab, options: { bold: true, fontSize: 13, color: HEX.ink, fill: { color: i % 2 ? HEX.white : HEX.mist } } }].concat(
  pick.map(j => ({ text: String(f(sr[key][j])), options: { fontSize: 13, align: "center", color: (key.startsWith("Extra") && sr[key][j] > 0) ? HEX.red : HEX.ink, bold: key.startsWith("Extra") && sr[key][j] > 0, fill: { color: i % 2 ? HEX.white : HEX.mist } } })))));
s.addTable(stbl, { x: 0.6, y: 1.85, w: 12.1, colW: [2.75, 1.3, 1.3, 1.3, 1.3, 1.4, 1.3, 1.45], rowH: 0.55, border: { type: "solid", color: "E3E5EA", pt: 0.5 }, valign: "middle", objectName: "scen-table" });
card(s, 0.6, 5.35, 12.1, 1.25, "FBEDEC", "scen-gate");
s.addText([{ text: "Board pivot gate (month 18):  ", options: { bold: true, color: HEX.red } },
  { text: "if run-rate placements are below ~$40m p.a., cut outsourced and build-partner spend, defer the secondary-market build and the buy-back reserve top-up. At downside volumes the full three-company build does not pay for itself." }],
  { x: 0.85, y: 5.45, w: 11.6, h: 1.05, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true, objectName: "scen-gate-t" });

// 16 THE RAISE
pres.addSection({ title: "The raise" });
s = content("The raise", `Raising ${$m(D.km.raised, 1)} in ${["one","two","three"][D.rounds.filter(r => r[1] > 0).length - 1]} milestone-gated rounds`,
  `Covers the ${$m(D.km.peak)} peak burn plus a $0.75m regulatory-capital and cash buffer and raising costs`);
const ms = ["Build platform, appoint licensee and trustee, first 3 offers documented, launch to wholesale investors (month 7)",
  "Licensed distribution live, platform in production, first subscriptions reconciled, signed issuer pipeline",
  "Secondary market live with licensed operator, ~$50m AUM, 60+ accredited brokers, own AFSL"];
const RR = D.rounds.filter(r => r[1] > 0);
const cw = RR.length === 2 ? 5.95 : 3.8, cg = RR.length === 2 ? 6.15 : 4.1;
RR.forEach(([n, amt, pre, own], i) => {
  const x = 0.6 + i * cg;
  const fill = [HEX.honey, HEX.navy, HEX.teal][i], tc = i === 0 ? HEX.ink : HEX.white;
  card(s, x, 1.85, cw, 1.6, fill, "r-head" + i);
  s.addText(n.replace(" (optional)", ""), { x: x + 0.3, y: 1.95, w: 3.2, h: 0.45, fontSize: 18, bold: true, color: tc, margin: 0, isTextBox: true, objectName: "r-n" + i });
  s.addText($m(amt, 1), { x: x + 0.3, y: 2.4, w: 3.2, h: 0.75, fontSize: 36, bold: true, color: tc, margin: 0, isTextBox: true, objectName: "r-a" + i });
  s.addText(`${["Month 1", "Month 10", "Month 28"][i]}  ·  ${$m(pre, 0)} pre-money  ·  ${pct(own, 1)}`, { x: x + 0.3, y: 3.05, w: cw - 0.4, h: 0.35, fontSize: 12, color: tc, margin: 0, isTextBox: true, objectName: "r-d" + i });
  card(s, x, 3.55, cw, 1.55, HEX.mist, "r-ms" + i);
  s.addText([{ text: "Milestones funded: ", options: { bold: true } }, { text: ms[i] }], { x: x + 0.25, y: 3.65, w: cw - 0.45, h: 1.4, fontSize: 12.5, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "r-mst" + i });
});
s.addChart(pres.charts.BAR, [{ name: "Use of funds", labels: D.uses.map(u => u[0].replace(" (all entities)", "").replace(" (capitalised)", "").replace(" (recurring)", "")), values: D.uses.map(u => +(u[1] / 1e6).toFixed(2)) }], {
  ...chartBase, x: 0.6, y: 5.15, w: 12.1, h: 1.85, barDir: "bar", chartColors: [HEX.slate], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "$0.0\"m\"",
  showLegend: false, showTitle: false, valAxisHidden: true, catAxisOrientation: "maxMin", catAxisLabelFontSize: 9, dataLabelFontSize: 9, objectName: "uof-chart" });
s.addNotes("Use of funds = gross spend in the first 24 months by bucket. Founders and existing B-Catcher/BHive holders retain " + pct(D.founder_pct) + " after all three rounds, before any ESOP. Valuations are placeholders for negotiation.");

// 17 RETURNS
s = content("The raise", "Illustrative returns at the end of year 5",
  `Exit run-rate: ${$m(D.exit.rev_rr)} revenue and ${$m(D.exit.ebitda_rr)} EBITDA (Q4 2031 annualised); equity value ${$m(D.exit.eq, 0)} on blended ${I["Exit multiple - EBITDA (year 5)"]}x EBITDA / ${I["Exit multiple - revenue (year 5)"]}x revenue`);
const RT = D.returns.filter(r => r[1] > 0);
RT.forEach(([n, inv, moic, irr], i) => {
  const x = 0.6 + i * (RT.length === 2 ? 6.15 : 4.1);
  card(s, x, 1.9, RT.length === 2 ? 5.95 : 3.8, 3.4, HEX.mist, "ret-card" + i);
  s.addText(n, { x: x + 0.3, y: 2.05, w: 3.2, h: 0.45, fontSize: 17, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "ret-n" + i });
  s.addText(moic.toFixed(1) + "x", { x: x + 0.3, y: 2.6, w: 3.2, h: 1.1, fontSize: 60, bold: true, color: [C.accent1, C.accent2, C.accent3][i], margin: 0, isTextBox: true, objectName: "ret-m" + i });
  s.addText(`money multiple on ${$m(inv, 1)}\n~${pct(irr)} p.a. IRR`, { x: x + 0.3, y: 3.75, w: 3.2, h: 1.0, fontSize: 15, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "ret-d" + i });
});
s.addText("Illustrative only. Returns depend on achieving the forecast, exit multiples and further dilution (ESOP, later rounds) that is not modelled. Upside case reaches EBITDA break-even in " + D.scen.rows["First month EBITDA positive"][2] + ".",
  { x: 0.6, y: 5.6, w: 12.1, h: 0.7, fontSize: 13, italic: true, color: C.text2, margin: 0, isTextBox: true, objectName: "ret-note" });

// 18 REGULATORY GATES
pres.addSection({ title: "Execution" });
s = content("Execution", "Nothing regulated goes live until its gate is passed",
  "Seven launch gates from the Master Business Plan; costs for each are in the set-up budget");
const gates = [["1", "Structure", "Entities, ownership, IP, governance, conflicts"], ["2", "Legal opinion", "AFSL / MIS / market / credit perimeter mapped"],
  ["3", "Counterparties", "Issuer, trustee, custody, banking, registry signed"], ["4", "Platform", "MVP, security testing, compliance UAT"],
  ["5", "Wholesale launch", "Licensed distribution, first reconciled subscriptions"], ["6", "Secondary market", "Operator partner, rules, settlement, surveillance"],
  ["7", "New products", "Retail, payments or lending - separate approvals"]];
gates.forEach(([n, h, t], i) => {
  const y = 1.85 + i * 0.68;
  hexBadge(s, 0.6, y, 0.6, i < 5 ? HEX.honey : HEX.grey, n, HEX.ink, "g" + i);
  s.addText(h, { x: 1.4, y: y, w: 2.6, h: 0.53, fontSize: 16, bold: true, color: C.text1, valign: "middle", margin: 0, isTextBox: true, objectName: "gh" + i });
  s.addText(t, { x: 4.0, y: y, w: 4.6, h: 0.53, fontSize: 14, color: C.text2, valign: "middle", margin: 0, isTextBox: true, objectName: "gt" + i });
});
card(s, 9.0, 1.85, 3.7, 4.65, HEX.mist, "reg-card");
s.addText([
  { text: "Licensing route", options: { bold: true, fontSize: 16, breakLine: true } },
  { text: "Honey Pot starts as a corporate authorised representative under a principal licensee ($12k/month), then applies for its own AFSL in year 2 ($150k).", options: { breakLine: true } },
  { text: "BHive uses a licensed market operator rather than seeking its own market licence.", options: { breakLine: true } },
  { text: "Wholesale clients only at launch; retail is a separate future programme." },
], { x: 9.25, y: 2.0, w: 3.25, h: 4.4, fontSize: 13, color: C.text1, paraSpaceAfter: 10, valign: "top", margin: 0, isTextBox: true, objectName: "reg-text" });

// 19 ROADMAP
s = content("Execution", "From funding to a live marketplace in 18 months",
  "Execution Roadmap stages with model timing (month 1 = January 2027)");
const phases = [["Months 1-2", "Foundation", "Entities, IP, counsel, data room, budget owner", HEX.navy],
  ["Months 2-6", "Build & license", "MVP build, CAR appointment, trust templates, broker agreements", HEX.navy],
  ["Month 7", "Wholesale launch", "First offers to verified investors via BHive", HEX.honey],
  ["Months 10-12", "Series A", "Scale brokers; reconcile, report, tune commissions", HEX.honey],
  ["Months 13-20", "Own AFSL + market build", "Licence application, operator partner, mobile apps", HEX.teal],
  ["Month 19", "Secondary market live", "Controlled activation once authorised", HEX.teal]];
s.addShape(pres.shapes.LINE, { x: 0.8, y: 3.05, w: 11.7, h: 0, line: { color: HEX.grey, width: 2 }, objectName: "rm-line" });
phases.forEach(([when, h, t, col], i) => {
  const x = 0.6 + i * 2.05;
  s.addShape(pres.shapes.OVAL, { x: x + 0.2, y: 2.87, w: 0.36, h: 0.36, fill: { color: col }, line: { color: HEX.white, width: 2 }, objectName: "rm-dot" + i });
  s.addText(when, { x, y: 2.2, w: 1.95, h: 0.45, fontSize: 13, bold: true, color: col === HEX.honey ? "B07A0A" : col, margin: 0, isTextBox: true, objectName: "rm-w" + i });
  card(s, x, 3.5, 1.9, 2.3, HEX.mist, "rm-card" + i);
  s.addText(h, { x: x + 0.15, y: 3.6, w: 1.65, h: 0.75, fontSize: 15, bold: true, color: C.text1, margin: 0, valign: "top", isTextBox: true, objectName: "rm-h" + i });
  s.addText(t, { x: x + 0.15, y: 4.4, w: 1.65, h: 1.3, fontSize: 12.5, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "rm-t" + i });
});

// 20 RISKS
s = content("Execution", "Key risks and how the plan manages them", null);
const risks = [["Regulatory perimeter", "Written counsel opinions before launch; CAR + licensed partners; gated activation", HEX.red],
  ["Placement volume", "Hiring frozen until month-18 volume gate; broker pipeline tracked monthly", HEX.red],
  ["Broker misconduct", "Accreditation, supervision sampling, commission clawbacks, complaints register", HEX.honey],
  ["Buy-back commitments", "Ladder discounts, funded cash reserve, legal & capital-adequacy opinion before launch", HEX.honey],
  ["Client money & reconciliation", "Funds only to trust/custody accounts; daily exception review", HEX.honey],
  ["Cyber & data", "MFA, SOC 2, pen tests, MDR service, tested recovery", HEX.teal]];
risks.forEach(([h, t, col], i) => {
  const c = i % 2, r = Math.floor(i / 2);
  const x = 0.6 + c * 6.15, y = 1.4 + r * 1.85;
  card(s, x, y, 5.95, 1.6, HEX.mist, "risk-card" + i);
  s.addShape(pres.shapes.OVAL, { x: x + 0.3, y: y + 0.5, w: 0.5, h: 0.5, fill: { color: col }, line: { color: col }, objectName: "risk-dot" + i });
  s.addText(h, { x: x + 1.05, y: y + 0.15, w: 4.7, h: 0.5, fontSize: 17, bold: true, color: C.text1, margin: 0, isTextBox: true, objectName: "risk-h" + i });
  s.addText(t, { x: x + 1.05, y: y + 0.65, w: 4.7, h: 0.8, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "risk-t" + i });
});

// 21 CLOSE
s = pres.addSlide({ masterName: "DARK", sectionTitle: "Execution" });
s.addText(`The ask: ${$m(D.rounds[0][1], 1)} Seed to build, license and launch`, { placeholder: "title" });
s.addText(`Part of a ${$m(D.km.raised, 1)} milestone-gated plan to a ${$m(D.aum[4], 0)} AUM, ${$m(D.rev[4], 0)}-revenue group by 2031`, { placeholder: "body" });
const nexts = ["Confirm entities, cap table and merger ratio of legacy B-Catcher and BHive holders", "Appoint financial services counsel and principal licensee", "Select first two asset types and issuer partners", "Commission BHive build partner and custody/registry providers"];
s.addText(nexts.map((t, i) => ({ text: t, options: { bullet: true, breakLine: i < nexts.length - 1 } })),
  { x: 0.8, y: 5.0, w: 11.5, h: 1.6, fontSize: 15, color: C.background1, paraSpaceAfter: 6, margin: 0, isTextBox: true, objectName: "close-next" });
s.addText("Forecasts are planning assumptions, not guarantees. Founder bios, market evidence and legal opinions to be added before external use.", { x: 0.8, y: 6.75, w: 11.5, h: 0.4, fontSize: 11, color: C.accent5, margin: 0, isTextBox: true, objectName: "close-disc" });

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
