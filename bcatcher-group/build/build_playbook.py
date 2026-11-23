"""B-Catcher Group - Launch Playbook & Budget Book (docx), generated from the model's numbers."""
import json, sys, datetime as dt
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

D = json.load(open(sys.argv[1]))
OUT = sys.argv[2]
NAVY, HONEY, GREY = RGBColor(0x1F, 0x3A, 0x5F), RGBColor(0xB0, 0x7A, 0x0A), RGBColor(0x59, 0x59, 0x59)
YRS = ["2027", "2028", "2029", "2030", "2031"]
I = D["inputs"]


def money(v, k=False):
    if v is None or v == "":
        return "-"
    if isinstance(v, str):
        return v
    if abs(v) < 0.5:
        return "-"
    if k:
        s = f"${abs(v)/1000:,.0f}k"
    else:
        s = f"${abs(v):,.0f}"
    return f"({s})" if v < 0 else s


def mm(v):
    return "-" if abs(v) < 500 else (f"(${abs(v)/1e6:,.2f}m)" if v < 0 else f"${v/1e6:,.2f}m")


doc = Document()
st = doc.styles["Normal"]
st.font.name = "Arial"
st.font.size = Pt(10)
st.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
for lvl, size, col in [(1, 16, NAVY), (2, 13, NAVY), (3, 11, HONEY)]:
    h = doc.styles[f"Heading {lvl}"]
    h.font.name = "Arial"
    h.font.size = Pt(size)
    h.font.bold = True
    h.font.color.rgb = col
    h.element.rPr.rFonts.set(qn("w:eastAsia"), "Arial")
sec = doc.sections[0]
sec.left_margin = sec.right_margin = Cm(2)
sec.top_margin = sec.bottom_margin = Cm(1.8)

# header/footer
fp = sec.footer.paragraphs[0]
fp.text = "B-Catcher Group  |  Launch Playbook & Budget Book  |  Confidential planning document - not an offer or financial advice"
fp.runs[0].font.size = Pt(8)
fp.runs[0].font.color.rgb = GREY


def shade(cell, hexcol):
    tcPr = cell._tc.get_or_add_tcPr()
    sh = OxmlElement("w:shd")
    sh.set(qn("w:val"), "clear")
    sh.set(qn("w:color"), "auto")
    sh.set(qn("w:fill"), hexcol)
    tcPr.append(sh)


def table(headers, rows, widths=None, size=8, bold_rows=(), right_from=1):
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    for i, h in enumerate(headers):
        c = t.rows[0].cells[i]
        c.text = ""
        r = c.paragraphs[0].add_run(str(h))
        r.bold = True
        r.font.size = Pt(size)
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, "1B1F2A")
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        for i, v in enumerate(row):
            cells[i].text = ""
            p = cells[i].paragraphs[0]
            r = p.add_run("" if v is None else str(v))
            r.font.size = Pt(size)
            r.bold = ri in bold_rows
            if i >= right_from and right_from >= 0:
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            if ri % 2 == 1:
                shade(cells[i], "F3F4F7")
    if widths:
        t.autofit = False
        tblPr = t._tbl.tblPr
        lay = OxmlElement("w:tblLayout")
        lay.set(qn("w:type"), "fixed")
        tblPr.append(lay)
        grid = t._tbl.tblGrid
        for i, gc in enumerate(grid.findall(qn("w:gridCol"))):
            gc.set(qn("w:w"), str(int(widths[i] * 567)))
        for i, w in enumerate(widths):
            t.columns[i].width = Cm(w)
        for row in t.rows:
            for i, w in enumerate(widths):
                row.cells[i].width = Cm(w)
    doc.add_paragraph()
    return t


def para(text, bold=False, italic=False, size=10, color=None):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.bold, r.italic = bold, italic
    r.font.size = Pt(size)
    if color:
        r.font.color.rgb = color
    return p


def bullets(items):
    for it in items:
        p = doc.add_paragraph(style="List Bullet")
        if isinstance(it, tuple):
            r = p.add_run(it[0])
            r.bold = True
            p.add_run(it[1])
        else:
            p.add_run(it)


# ------------------------------------------------------------------ COVER
t = doc.add_paragraph()
r = t.add_run("B-CATCHER GROUP")
r.bold = True
r.font.size = Pt(26)
r.font.color.rgb = NAVY
t = doc.add_paragraph()
r = t.add_run("Launch Playbook & Budget Book")
r.bold = True
r.font.size = Pt(18)
r.font.color.rgb = HONEY
para("From funding to a live, licensed three-company group: every step, owner, month, cost, hire and budget.", size=11, color=GREY)
para(f"Version 1.0  |  {dt.date.today().strftime('%d %B %Y')}  |  Companion to BCatcher_Group_Integrated_Financial_Model.xlsx and Investor Deck v4", size=9, color=GREY)
para("Status: planning document. All figures are AUD planning assumptions from the Base case of the integrated model. Regulated activities require legal advice, licences and partner agreements before launch.", italic=True, size=9, color=GREY)

doc.add_heading("1. The plan on one page", 1)
km = D["km"]
table(["Measure", "Base case"], [
    ["Operating model", "Investors subscribe directly through BHive; Honey Pot manages brokers and distribution behind that process; B-Catcher originates and manages the investments"],
    ["One-off set-up and build cost (all three entities)", money(km["setup"])],
    ["Peak funding requirement (cumulative burn before break-even)", money(km["peak"])],
    ["Equity plan", "  ·  ".join(f"{r[0].replace(' (optional)','')} {money(r[1])}" for r in D["rounds"] if r[1])],
    ["Lowest cash balance with the equity plan", money(km["mincash"])],
    ["First month of positive group EBITDA", km["ebm"]],
    ["Wholesale launch / secondary market live", f"Month {I['Controlled wholesale launch (first subscriptions)']} (Jul 2027) / month {I['BHive secondary market activation']} (Jul 2029)"],
    ["2031: capital placed / managed AUM", mm(D["cap"][4]) + " / " + mm(D["aum"][4])],
    ["2031: net revenue / EBITDA / NPAT", mm(D["rev"][4]) + " / " + mm(D["ebitda"][4]) + " / " + mm(D["npat"][4])],
    ["Team: FTE end 2027 / 2029 / 2031", f"{D['fte'][0]:.0f} / {D['fte'][2]:.0f} / {D['fte'][4]:.0f}"],
], widths=[7, 10], size=9, right_from=-1)

doc.add_heading("Five-year summary (consolidated, $)", 3)
rows = []
for lab, key in [("Capital placed through BHive (client money, not revenue)", "cap"), ("Managed AUM - year end", "aum"), ("Honey Pot revenue", "rhp"), ("B-Catcher revenue", "rbc"), ("BHive revenue", "rbh"),
                 ("Group net revenue", "rev"), ("Cost of sales", "cogs"), ("Gross profit", "gp"), ("Staff costs", "staff"), ("Total operating expenses", "opex"), ("EBITDA", "ebitda"), ("Net profit after tax", "npat"), ("Closing cash", "cash")]:
    sign = -1 if key in ("cogs", "staff", "opex") else 1
    rows.append([lab] + [money(sign * v) for v in D[key]])
table(["$"] + YRS, rows, widths=[6.2] + [2.2] * 5, bold_rows=(5, 7, 10, 11))

# ------------------------------------------------------------------ PHASES
doc.add_heading("2. Start-to-finish launch sequence", 1)
para("Each phase lists what must be done, who owns it, when (model month; month 1 = January 2027), what it costs in the model and the evidence that closes the gate. Costs come straight from the Setup Costs and Opex registers.")
phases = [
    ("Phase 0 - Foundation (months 1-2)", "Gate 1: structure, ownership, IP, governance", [
        ("Close Seed round and open group bank accounts", "CEO", "1", "Raise costs 4%"),
        ("Confirm entity names; incorporate Honey Pot, BHive and template SPV", "CEO / counsel", "1", money(8000)),
        ("Legacy B-Catcher & BHive entity/IP due diligence; agree merger ratio of existing holders; assign IP", "CEO / counsel", "1-2", money(55000)),
        ("Shareholders' agreement, constitution, ESOP plan", "CEO / counsel", "2-3", money(40000)),
        ("Group tax structuring (trusts, GST, transfer pricing, R&D)", "Fractional CFO", "2-3", money(35000)),
        ("Appoint Chair, adopt board charter, delegation matrix, risk appetite", "Board", "1", "Chair $90k p.a."),
        ("Hire the 4 employees: CEO & CIO (m1), Head of Risk, Compliance & Ops (m1, 0.5 to m4), Head of Distribution (m3), Investor Communications & BDM (m5); contract fractional CFO, product owner, vCISO, UX. Two further hires only when volume triggers are met (Investment Analyst m13, Fund Operations m19). No other hires planned", "CEO", "1-3", "See section 4"),
    ]),
    ("Phase 1 - Legal perimeter & design (months 1-3)", "Gate 2: written regulatory opinion", [
        ("Financial services perimeter opinion: AFSL, MIS, market, credit, AML scope per entity", "External counsel", "1-3", money(90000)),
        ("Decide licensing route: Honey Pot as CAR under principal licensee; BHive via licensed market operator", "Board", "3", "Decision"),
        ("Brand identity & design system for all three brands", "Marketing", "2-3", money(65000)),
        ("BHive discovery, architecture & UX design; select build partner (RFP)", "Product owner + build partner", "2-3", money(95000)),
        ("Cloud landing zone, environments, domains", "Product owner + build partner", "2", money(15000)),
    ]),
    ("Phase 2 - Counterparties, products & build (months 3-6)", "Gates 3 & 4: counterparties signed, platform tested", [
        ("Appoint trustee / RE / custodian; onboard fund administrator & registry", "CEO/CIO / Fractional CFO", "5-6", money(30000)),
        ("Trust deed & IM templates: property, private credit, PE / pre-IPO (UTEGBO)", "CEO/CIO / counsel", "3-6", money(130000)),
        ("Due diligence on first 3 offers (valuation, QS, building, legal)", "CEO/CIO + outsourced analysts", "4-6", money(105000)),
        ("CAR appointment under principal licensee", "Head of Distribution", "4", money(25000) + " + $12k/month"),
        ("Compliance framework: compliance plan, breach, IDR, conflicts, DDO/TMD", "Head of Risk, Compliance & Ops", "3-5", money(65000)),
        ("AML/CTF program, AUSTRAC enrolment, independent review", "Head of Risk", "4-6", money(35000)),
        ("Broker agreement suite & commission policy", "Head of Distribution / counsel", "3-4", money(45000)),
        ("Broker CRM & commission system; LMS & accreditation content", "Head of Distribution", "3-6", money(110000)),
        ("BHive MVP: investor, broker & admin portals, wallet (build partner)", "Product owner + build partner", "3-8", money(650000)),
        ("Integrations: KYC, e-sign, registry, banking reconciliation", "Senior engineer + build partner", "4-7", money(155000)),
        ("Privacy impact assessment; pre-launch penetration test", "vCISO + build partner", "5-6", money(65000)),
        ("Launch campaign, broker roadshow, marketing compliance review", "Head of Distribution + Investor Comms & BDM", "6-8", money(105000)),
    ]),
    ("Phase 3 - Controlled wholesale launch (month 7)", "Gate 5: licensed distribution and complete audit trail", [
        ("First offers published to verified wholesale investors in BHive", "CEO/CIO / Head of Distribution", "7", "-"),
        ("Accredit first brokers; issue referral links; supervise first files", "Head of Distribution", "7+", "$1,500 per broker"),
        ("First subscriptions reconciled to register; first investor reports", "Head of Risk & Ops + outsourced accountants", "7-9", "-"),
        ("Insurance in force: PI, cyber, D&O, technology E&O", "Fractional CFO", "7", "~$11k/month"),
        ("Monthly KPI pack: capital placed, brokers, conversion, onboarding time, complaints", "CEO", "7+", "-"),
    ]),
    ("Phase 4 - Series A & scale distribution (months 10-18)", "Milestones: licence, live platform, signed pipeline", [
        ("Close Series A", "CEO / CFO", "10", money(D["rounds"][1][1])),
        ("Appoint commission-only regional BDMs (VIC/QLD first); no salaried hires", "Head of Distribution", "13-24", "0.20% sales commission pool"),
        ("Own AFSL application for Honey Pot (responsible managers, ASIC)", "Head of Risk, Compliance & Ops", "13-20", money(150000)),
        ("Hire Investment Analyst & Asset Manager (trigger: ~$50m placed / 10+ trusts)", "CEO/CIO", "13", "$140k + on-costs"),
        ("Mobile apps (iOS & Android)", "Product owner + build partner", "13-18", money(180000)),
        ("Buy-back windows open for PE trust units: secondary market first, group buys residual, capped at 5% a year", "CEO/CIO / Head of Risk", "13", "Funded from 1.5% reserve"),
        ("SOC 2 Type I readiness & audit", "vCISO", "10-15", money(70000)),
        ("Month-18 pivot gate: placements run-rate vs plan; hiring decision", "Board", "18", "Decision"),
    ]),
    ("Phase 5 - Own licence, Series B & marketplace build (months 19-30)", "Gate 6: own AFSL; Series B; AUM ~$150m", [
        ("Hire Fund Operations & Finance Officer (trigger: ~1,000 investors)", "Head of Risk, Compliance & Ops", "19", "$110k + on-costs"),
        ("Market-operator route: partner, market rules, operating agreement", "Product owner + build partner", "22-27", money(120000)),
        ("Secondary market engine; operator integration", "Product owner + build partner", "22-30", money(510000)),
        ("Own AFSL granted; exit CAR arrangement", "Head of Risk, Compliance & Ops", "24-25", "ASIC levy $2.5k/month"),
        ("Property-token buy-back ladder opens (15/10/5% below valuation; capped at 5% a year)", "CEO/CIO", "25", "Funded from 1.5% reserve"),
        ("Expand build-partner retainer to $50k/month (no in-house engineers)", "Product owner + build partner", "25", "$50k/month"),
        ("Close Series B (growth round, gated on volume)", "CEO / CFO", "28", money(D["rounds"][2][1])),
        ("Secondary-market pen test & surveillance validation", "vCISO", "30", money(50000)),
    ]),
    ("Phase 6 - Marketplace live & scale (months 31-60)", "Gate 7: secondary market authorised; any new product approved separately", [
        ("Controlled secondary-market activation with licensed operator", "Product owner + build partner", "31", "$15k/month + 40% fee share"),
        ("Digital token legal register opinion; tokenised register pilot", "Product owner / counsel", "30-39", money(130000)),
        ("Add commission-only BDMs in WA/SA and family-office channel", "Head of Distribution", "31-49", "See section 4"),
        ("Capacity review: scale outsourced providers; hire only if they cannot hold service levels", "CEO / Board", "37-49", "See section 4"),
        ("Retail, payments, lending or BHive card only as separately funded projects", "Board", "-", "Not in model"),
    ]),
]
for title, gate, steps in phases:
    doc.add_heading(title, 2)
    para(gate, bold=True, size=9, color=HONEY)
    table(["Step", "Owner", "Month", "Cost in model"], [list(s) for s in steps], widths=[9.5, 3.5, 1.5, 3], size=8, right_from=3)

# ------------------------------------------------------------------ SETUP REGISTER
doc.add_heading("3. Complete set-up & launch cost register", 1)
para(f"Every one-off cost, {money(km['setup'])} in total including 10% contingency. Capex is capitalised and amortised over 36 months; opex is expensed when incurred. Replace each estimate with a quote before external use.")
rows = [[e, c, it, money(a), str(int(s)), str(int(n)), t] for e, c, it, a, s, n, t in D["setup_items"]]
rows.append(["TOTAL", "", "", money(sum(x[3] for x in D["setup_items"])), "", "", ""])
table(["Entity", "Category", "Item", "Amount", "Start", "Months", "Type"], rows, widths=[1.8, 3.2, 7.2, 1.8, 1, 1.2, 1.2], size=7, right_from=3, bold_rows=(len(rows) - 1,))
para("By category", bold=True)
table(["Category", "Amount"], [[k, money(v)] for k, v in sorted(D["setup_cat"].items(), key=lambda x: -x[1])], widths=[9, 3], size=8)

# ------------------------------------------------------------------ HEADCOUNT
doc.add_heading("4. Headcount plan & people costs", 1)
para(f"{len(D['roles'])} costed roles. Loaded cost = salary x FTE x (1 + 12% super + 5.45% payroll tax + 1% workers comp + 3% other) for employees, or + {I['Contractor premium over salary equivalent (no on-costs)']*100:.0f}% premium for contractors. Recruitment at {I['Recruitment fee for employee hires']*100:.0f}% of salary for hires after month 1, plus ${I['IT equipment & onboarding per new starter']:,.0f} equipment per starter. Wages grow {I['Annual wage growth']*100:.1f}% p.a.")
for ent in ["B-Catcher", "Honey Pot", "BHive"]:
    doc.add_heading(ent, 3)
    rr = sorted([r for r in D["roles"] if r[0] == ent], key=lambda r: r[4])
    table(["Department", "Role", "Type", "Start month", "FTE", "Base salary", "Loaded p.a."],
          [[r[1], r[2], r[3], str(int(r[4])), f"{r[5]:.1f}", money(r[6]), money(r[7])] for r in rr],
          widths=[2.6, 6.4, 1.8, 1.4, 1, 2, 2], size=7, right_from=3)
table(["FTE at year end"] + YRS, [["B-Catcher"] + [f"{v:.1f}" for v in D["fte_bc"]], ["Honey Pot"] + [f"{v:.1f}" for v in D["fte_hp"]],
                                  ["BHive"] + [f"{v:.1f}" for v in D["fte_bh"]], ["Group"] + [f"{v:.1f}" for v in D["fte"]],
                                  ["Staff cost (salaries & on-costs)"] + [money(v) for v in D["staff"]]], widths=[5.5] + [2.3] * 5, size=8, bold_rows=(3,))

# ------------------------------------------------------------------ BUDGETS
doc.add_heading("5. Operating budgets - 2027 and 2028 by cost centre", 1)
para("Quarterly budgets per entity are in the 'Budget Y1-Y2' sheet. Annual totals below include intercompany fees, which net to zero in the group view.")
cur = None
rows = []
for b in D["budget"]:
    lab = b[0]
    if lab.endswith("COST CENTRE") or lab.startswith("GROUP") or lab.startswith("SET-UP"):
        if rows:
            table(["Line", "FY2027", "FY2028"], rows, widths=[9, 3.5, 3.5], size=8)
        rows = []
        doc.add_heading(lab.title().replace("B-Catcher", "B-Catcher").replace("Bhive", "BHive"), 3)
        continue
    rows.append([lab, money(b[9]), money(b[10])])
if rows:
    table(["Line", "FY2027", "FY2028"], rows, widths=[9, 3.5, 3.5], size=8)

# ------------------------------------------------------------------ OPEX REGISTER
doc.add_heading("6. Recurring operating cost register", 1)
para("Every recurring non-salary cost line and how it scales. Annual totals are Base case and include escalation.")
rows = []
for e, c, it, amt, drv, s, en, *tot in D["opex_items"]:
    a = f"{amt*100:.2f}%" if str(drv).startswith("%") else (f"${amt:,.2f}" if amt < 10 else f"${amt:,.0f}")
    rows.append([e, it, a, drv, str(int(float(s))) if s not in (None, "") else "", money(tot[0]), money(tot[2]), money(tot[4])])
table(["Entity", "Item", "Rate", "Driver", "Start", "2027", "2029", "2031"], rows, widths=[1.7, 5.8, 1.5, 2.8, 1, 1.7, 1.7, 1.7], size=7, right_from=4)
table(["Category"] + YRS, [[k] + [money(v) for v in vals] for k, vals in D["cats"].items()], widths=[6.2] + [2.2] * 5, size=8)

# ------------------------------------------------------------------ FUNDING
doc.add_heading("7. Funding plan, milestones and returns", 1)
ms = ["Build platform, appoint licensee and trustee, document first 3 offers, wholesale launch (month 7)",
      "Licensed distribution live, platform in production, first subscriptions reconciled, signed issuer pipeline",
      "Optional - not needed in Base; for Downside or faster expansion"]
table(["Round", "Amount", "Pre-money", "Post-round %", "Milestones it funds"],
      [[r[0], money(r[1]), money(r[2]), f"{r[3]*100:.1f}%", ms[i]] for i, r in enumerate(D["rounds"]) if r[1]], widths=[2.5, 2, 2, 2, 8.5], size=8, right_from=-1)
para(f"Founders and existing B-Catcher/BHive holders retain {D['founder_pct']*100:.1f}% after all planned rounds, before any ESOP. Valuations are negotiating placeholders.", size=9)
table(["Use of funds - first 24 months", "Amount"], [[u, money(v)] for u, v in D["uses"]], widths=[11, 3], size=8)
table(["Investor", "Invested", "MOIC at year 5", "IRR"], [[n, money(a), f"{x:.1f}x", f"{y*100:.0f}%"] for n, a, x, y in D["returns"] if a], widths=[5, 3, 3, 3], size=8)

# ------------------------------------------------------------------ SCENARIOS
doc.add_heading("8. Scenarios, stress tests and the pivot gate", 1)
cases, srows = D["scen"]["cases"], D["scen"]["rows"]
keys = ["Net revenue Y5", "EBITDA Y5", "Peak funding requirement", "Extra equity needed (vs buffer + buy-back reserve)", "First month EBITDA positive"]
rows = []
for j, cname in enumerate(cases):
    rows.append([cname] + [money(srows[k][j]) if not isinstance(srows[k][j], str) else ("> 2031" if "Beyond" in srows[k][j] else srows[k][j]) for k in keys])
table(["Case", "Revenue 2031", "EBITDA 2031", "Peak funding", "Extra equity", "EBITDA +ve"], rows, widths=[5.5, 2.3, 2.3, 2.3, 2.3, 2.3], size=8)
para("What this tells the board:", bold=True)
bullets([
    ("Volume is the risk. ", "A 25% technology overrun costs ~$0.4m; a 50% placement shortfall adds ~$14m to the funding need."),
    ("The cost base is mostly fixed. ", "Deferring post-launch hires by 18 months reduces the downside funding need, but does not make it viable on its own."),
    ("Month-18 pivot gate. ", "If run-rate placements are below ~$40m p.a. at month 18: freeze post-launch hiring, defer the secondary-market build, partner more of BHive's technology and operate Honey Pot under the principal licensee for longer."),
])

# ------------------------------------------------------------------ KPI TARGETS
doc.add_heading("9. KPI targets the board should track monthly", 1)
table(["KPI (Base case, year end)"] + YRS, [
    ["Capital placed in year"] + [mm(v) for v in D["cap"]],
    ["Managed AUM"] + [mm(v) for v in D["aum"]],
    ["Investors on BHive"] + [f"{v:,.0f}" for v in D["inv"]],
    ["Active accredited brokers"] + [f"{v:,.0f}" for v in D["brk"]],
    ["Group net revenue"] + [mm(v) for v in D["rev"]],
    ["EBITDA"] + [mm(v) for v in D["ebitda"]],
    ["Closing cash"] + [mm(v) for v in D["cash"]],
], widths=[5.5] + [2.3] * 5, size=8)
bullets(["Plus: conversion from referral to subscription, onboarding time, rejected applications, complaints and breaches, broker productivity, revenue per FTE, cost per $1 of capital placed (all on the Annual Summary sheet)."])

# ------------------------------------------------------------------ DECISIONS & VALIDATION
doc.add_heading("10. Decisions required and what still needs validating", 1)
doc.add_heading("Board decisions", 3)
bullets(["Final legal entity names, ownership and merger ratio between legacy B-Catcher and BHive shareholders",
         "Honey Pot licensing route: CAR first then own AFSL (modelled) vs own AFSL from day one",
         "Whether Honey Pot includes mortgage/credit brokers (separate ACL - not modelled)",
         "First two asset types and issuer partners; investment committee membership",
         "BHive build partner vs in-house; market-operator partner selection",
         "Seed amount, valuation and investor terms; ESOP pool size",
         "Approved hiring schedule with volume gates; month-18 pivot criteria"])
doc.add_heading("Replace estimates with evidence before external circulation", 3)
bullets(["Legal fee proposals: perimeter opinion, AFSL application, trust deeds, market-operator agreements",
         "Principal licensee, trustee/RE, custodian, registry and fund-administrator quotes",
         "Technology build partner quote and support retainer; KYC, e-sign, CRM, cloud pricing",
         "Insurance quotes: PI, cyber, D&O, technology E&O",
         "Signed issuer pipeline and broker letters of intent supporting the placement curve",
         "Independent review of the financial model; R&D eligibility opinion",
         "Founder and team bios; verified market sizing and competitor analysis"])
doc.add_heading("Deliberately excluded from the model", 3)
bullets(["Legacy B-Catcher DOGE mining, lottery, insurance and crypto-treasury businesses",
         "BHive Coin, BHive Card, token lending and balance-sheet lending (the buy-back ladder IS included, with a funded reserve)",
         "Retail investors; performance fees; GST timing and working-capital days; ESOP dilution"])

doc.save(OUT)
print("saved", OUT)
