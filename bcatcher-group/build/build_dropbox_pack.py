"""Generate text/CSV updates for every folder of the Dropbox workspace from the model (Dropbox connector accepts text only)."""
import openpyxl, json, csv, io, os, sys, datetime as dt

MODEL, NUMS, OUTDIR = sys.argv[1], sys.argv[2], sys.argv[3]
D = json.load(open(NUMS))
wb = openpyxl.load_workbook(MODEL, data_only=True)
I = D["inputs"]
YRS = ["2027", "2028", "2029", "2030", "2031"]
TODAY = dt.date.today().strftime("%d %B %Y")
F = {}  # relative path -> content


def money(v):
    if v is None or isinstance(v, str):
        return v or "-"
    if abs(v) < 0.5:
        return "-"
    s = f"${abs(v):,.0f}"
    return f"({s})" if v < 0 else s


def mm(v):
    return "-" if abs(v) < 500 else (f"(${abs(v)/1e6:,.2f}m)" if v < 0 else f"${v/1e6:,.2f}m")


def md_table(headers, rows):
    out = "| " + " | ".join(headers) + " |\n|" + "|".join(["---"] * len(headers)) + "|\n"
    for r in rows:
        out += "| " + " | ".join("" if c is None else str(c).replace("|", "/") for c in r) + " |\n"
    return out + "\n"


def to_csv(rows):
    b = io.StringIO()
    csv.writer(b).writerows(rows)
    return b.getvalue()


def monthly_sheet_csv(name):
    ws = wb[name]
    rows = [["Line item", "Unit", "60-month total"] + [ws.cell(4, c).value.strftime("%b-%Y") for c in range(5, 65)]]
    for r in range(8, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None:
            continue
        vals = [ws.cell(r, c).value for c in range(5, 65)]
        if all(v is None for v in vals):
            rows.append([a])
            continue
        rnd = lambda v: (round(v, 2) if abs(v) < 10 else round(v)) if isinstance(v, (int, float)) else v
        rows.append([a, ws.cell(r, 2).value or "", rnd(ws.cell(r, 3).value) if isinstance(ws.cell(r, 3).value, (int, float)) else ""] + [rnd(v) for v in vals])
    return to_csv(rows)


def annual(sheet, label, kind="flow"):
    ws = wb[sheet]
    for r in range(8, ws.max_row + 1):
        if ws.cell(r, 1).value == label:
            vals = [ws.cell(r, c).value or 0 for c in range(5, 65)]
            if kind == "stock":
                return [vals[12 * y + 11] for y in range(5)]
            return [sum(vals[12 * y:12 * y + 12]) for y in range(5)]
    raise KeyError(sheet + ":" + label)


HDR = lambda title, folder: (f"# {title}\n\n**B-Catcher Group** | {folder} | Version 10.0 | {TODAY}  \n"
                             "Confidential planning document. All figures are AUD Base-case planning assumptions from *BCatcher_Group_Integrated_Financial_Model.xlsx*. Not an offer or financial advice.\n\n")
km = D["km"]
rounds = D["rounds"]

# =====================================================================  ROOT REGISTER
reg = HDR("Document Register - B-Catcher Group workspace", "Root")
reg += ("## Operating model\nInvestors subscribe directly through **BHive**. **Honey Pot** manages brokers and distribution behind that process. **B-Catcher** originates and manages the investments.\n\n"
        "## Version 10 headline numbers (Base case)\n")
reg += md_table(["Measure", "Value"], [
    ["One-off set-up & build cost", money(km["setup"])], ["Peak funding requirement", money(km["peak"])],
    ["Equity plan", " / ".join(f"{r[0].replace(' (optional)','')} {mm(r[1])}" for r in rounds if r[1])],
    ["First month positive group EBITDA", km["ebm"]], ["2031 net revenue / EBITDA", f"{mm(D['rev'][4])} / {mm(D['ebitda'][4])}"],
    ["2031 managed AUM / investors / brokers", f"{mm(D['aum'][4])} / {D['inv'][4]:,.0f} / {D['brk'][4]:,.0f}"], ["Group FTE 2027 / 2031", f"{D['fte'][0]:.0f} / {D['fte'][4]:.0f}"]])
reg += "## What is current in each folder\n"
reg += md_table(["Folder", "Current (v10)", "Superseded / reference only"], [
    ["01 - Business Plans & Strategy", "Master Business Plan v10 (.md); Launch Playbook & Roadmap v10 (.md); Launch Playbook & Budget Book (.docx - to upload)", "Master Business Plan v3; Execution Roadmap v1"],
    ["02 - Corporate Structure & Legal", "Legal, Licensing & Corporate Workplan + Budget v10 (.md)", "Corporate & Regulatory Implementation v1 (still valid for regulatory scope)"],
    ["03 - Financial Models & Budgets", "Integrated Financial Model v10 (.xlsx - to upload) + CSV exports: assumptions, annual summary, entity P&Ls, quarterly budgets, monthly consolidated P&L, cash flow, balance sheet, set-up & opex registers, scenarios", "BCatcher_BHive_36_Month_Financial_Model.xlsx (single cost line, no entity split; runs out of cash month 16)"],
    ["04 - Organisation & Headcount", f"Headcount Plan v10 (.csv, {len(D['roles'])} roles); Organisation Structure v10 (.md)", "BCatcher_BHive_Headcount_Org_Structure_36_Months.xlsx"],
    ["05 - Honey Pot", "Honey Pot Business Case & Budget v10 (.md); Broker Commission Schedule (.csv)", "Broker Network Operating Manual v1 (still valid for procedures)"],
    ["06 - BHive", "BHive Build Plan & Budget v10 (.md)", "Product & Technical Requirements v1 (still valid for requirements)"],
    ["07 - Investor Deck & Capital Raise", "Investor Deck v10 (.pptx - to upload); Deck v10 slide script (.md); Capital Raise Summary v10 (.md)", "Investor Deck v3 Working; Content Brief v1 (evidence checklist still applies)"],
    ["08 - Historical Reference", "Archive Index (.md) - what was carried forward or excluded", "All originals, unchanged"]])
reg += f"Every v4-v8 Markdown/CSV file in these folders is superseded by the matching v10 file. v10 is the growth plan (~$750m AUM by 2031) built on the v9 fixes from the independent analyst review: {len(D['roles'])} roles - 4 core employees, 6 volume-triggered hires and fractional contractors (v4 had 49 roles, v5 had 27); fractional contractors, outsourced providers, AI tools, the build partner and commission-only sales replace every other role.\n\n"
reg += ("## Files to upload manually\nThe Dropbox connector can only write text files. Drag these three files (sent in the Claude session) into the folders shown:\n\n"
        "- `BCatcher_Group_Integrated_Financial_Model.xlsx` -> 03 - Financial Models & Budgets\n- `BCatcher_Group_Investor_Deck_v10.pptx` -> 07 - Investor Deck & Capital Raise\n"
        "- `BCatcher_Group_Launch_Playbook_and_Budget_Book.docx` -> 01 - Business Plans & Strategy\n")
F["00 - Document Register v10.md"] = reg

# =====================================================================  01 BUSINESS PLAN v9
f01 = "01 - Business Plans & Strategy"
bp = HDR("Master Business Plan - Version 10 (fully costed)", f01)
bp += ("## 1. Executive summary\n"
       "B-Catcher Group is an Australian wholesale private-markets group made of three companies with one controlled lifecycle: **B-Catcher** (parent) originates, structures and manages property trusts, private credit and private equity / pre-IPO investments; **Honey Pot** recruits, accredits and supervises the broker network that distributes approved offers; **BHive** is the single investor-facing platform where investors are verified, subscribe, hold, receive reporting and, once licensed, trade on a secondary market.\n\n"
       f"The group needs **{money(km['setup'])}** of one-off set-up and build spend and a total of **{mm(D['km']['peak'])}** of cumulative funding before it becomes cash-flow positive in **{km['cfm']}**. It reaches positive EBITDA in **{km['ebm']}** and, by 2031, **{mm(D['rev'][4])}** revenue, **{mm(D['ebitda'][4])}** EBITDA and **{mm(D['aum'][4])}** of managed AUM with **{D['fte'][4]:.0f}** FTE.\n\n")
bp += "## 2. Group structure\n"
bp += md_table(["Entity", "Role", "Revenue", "Regulatory route"], [
    ["B-Catcher (parent)", "Origination, investment committee, asset management, group finance, risk, shared services", "Entry fees (1.5%) and management fees (BHive schedule); intercompany shared-services fee", "Manager under external trustee/RE; no client money"],
    ["Honey Pot", "Broker recruitment, accreditation, supervision, offer approval, lead attribution, commissions", "Placement fees from issuers (net of broker commissions)", "CAR under principal licensee from month 4; own AFSL application months 13-20"],
    ["BHive", "Investor, broker and admin portals; wallet; reporting; secondary market", "Technology fee (0.25% p.a.), offer onboarding fees, exchange fees (0.25%), buy-back resale gains; intercompany licence fees", "Technology provider; trading via licensed market-operator partner from month 31"],
    ["Investment trusts / SPVs", "Hold assets, issue units, carry trust-level debt", "N/A - investor money and assets are never group revenue or group cash", "External trustee, custodian, registry"]])
bp += ("## 2a. Two kinds of investor\n")
bp += md_table(["", "Company (equity) investors", "BHive customers (product investors)"], [
    ["What they buy", "Shares in B-Catcher Group", "Units in property trusts and tokens, private-credit funds, PE / pre-IPO unit trusts, shares in Pty Ltd companies, ASX placements"],
    ["When", "Seed month 1 (Jan 2027) · Series A month 10 (no Series B - a loan facility covers the rest)", f"Buy from month {I['Controlled wholesale launch (first subscriptions)']} (Jul 2027); buy-back bridge from month 19; secondary market from month {I['BHive secondary market activation']}"],
    ["Their money goes to", "The group - pays for building and running the company", "The trust / custodian - never group money"],
    ["Return", "Growth in company value", "Rent, interest and growth from the assets, after fees"],
    ["Who can invest", "Sophisticated / professional investors", "Verified wholesale investors at launch; retail and overseas access only as licences allow"]])
ir = D["ir"]
bp += (f"**Rent:** each property trust receives **{I['Property gross rental yield on purchase price (rent received by the trust)']*100:.0f}% gross rent** less **{I['Less: vacancy, letting and non-recoverable outgoings']*100:.0f}%** vacancy and outgoings, paid to unit holders as distributions (not group revenue).\n\n")
bp += md_table(["Net return to investor per year (illustrative)"] + ir["cols"], [
    ["Gross return before group fees"] + [f"{x*100:.1f}%" for x in ir["gross"]],
    ["All group fees per year"] + [f"{x*100:.1f}%" for x in ir["fees"]],
    ["**Net to investor**"] + [f"**{x*100:.1f}%**" for x in ir["net"]],
    ["Investor keeps"] + [f"{x*100:.0f}%" for x in ir["kept"]],
    ["$150k invested: year-1 net return"] + [money(x) for x in ir["ex150"]]])
bp += "## 3. Products and fees\n"
bp += "v9 fees: the BHive entry and management fees are kept; the technology fee is cut to 0.25% and the property placement fee to 2.5%, so the ongoing fee on property is 1.5% p.a. all-in. Every offer shows one total-cost figure. Honey Pot's placement fee funds broker introduction commissions. See the Investor Returns sheet: investors keep about 70-79% of gross returns.\n\n"
bp += md_table(["Product", "Mix", "Honey Pot placement fee (broker commissions)", "B-Catcher entry fee", "B-Catcher management fee"], [
    ["Property trusts / tokens (geared and Debt-Free series)", "45%", "2.5% (brokers 1.5%)", "1.5% (BHive)", "1.25% p.a. (BHive)"],
    ["Private credit", "25%", "1.5% (brokers 0.9%)", "1.5%", "1.00% p.a."],
    ["Private equity / pre-IPO (UTEGBO trusts)", "20%", "5.0% (brokers 3.0%)", "1.5%", "1.50% p.a."],
    ["ASX placements / IPO", "10%", "4.0%", "-", "-"]])
bp += (f"BHive charges its {I['Technology / platform fee']*100:.2f}% p.a. technology fee, ${I['Offer onboarding fee (data room, workflows, registry set-up)']:,.0f} per offer onboarded and a {I['Exchange fee per trade (BHive schedule)']*100:.2f}% exchange fee per trade on {I['Secondary market turnover']*100:.0f}% annual turnover (40% to the licensed operator). "
       "**Buy-back structure (v9):** the buy-back is the investors' exit route until the secondary market opens - available from year 2 for property tokens (20% / 15% / 10% / 5% below valuation by holding year) and PE trust units (10% / 8% / 5%), discretionary and capped at 5% of eligible units a year. Once the secondary market is live (month 31) requests go there first and the group buys only the residual. Resale assumes a price haircut, and a 1.5% cash reserve is held on top of the minimum buffer. "
       f"Brokers source {I['Share of capital sourced via brokers']*100:.0f}% of capital and receive {I['Broker share of placement fee (commission)']*100:.0f}% of the placement fee. Performance fees are excluded.\n\n")
bp += "## 4. Five-year forecast (consolidated)\n"
rows = []
for lab, key, sg in [("Capital placed (client money)", "cap", 1), ("Managed AUM - year end", "aum", 1), ("Investors on BHive", "inv", 0), ("Active brokers", "brk", 0), ("Honey Pot revenue", "rhp", 1), ("B-Catcher revenue", "rbc", 1), ("BHive revenue", "rbh", 1),
                     ("Net revenue", "rev", 1), ("Cost of sales", "cogs", -1), ("Gross profit", "gp", 1), ("Staff costs", "staff", -1), ("Total operating expenses", "opex", -1), ("EBITDA", "ebitda", 1), ("NPAT", "npat", 1), ("Closing cash", "cash", 1), ("FTE", "fte", 0)]:
    rows.append([lab] + [(money(sg * v) if sg else f"{v:,.0f}") for v in D[key]])
bp += md_table(["$"] + YRS, rows)
bp += "## 5. Cost of setting up the business\n"
bp += md_table(["Set-up category", "Amount"], [[k, money(v)] for k, v in sorted(D["setup_cat"].items(), key=lambda x: -x[1])] + [["**Total**", f"**{money(km['setup'])}**"]])
bp += "Full line-by-line register: folder 03, `Setup_Cost_Register_v10.csv`.\n\n## 6. Operating costs by category\n"
bp += md_table(["Category"] + YRS, [[k] + [money(v) for v in vals] for k, vals in D["cats"].items()])
bp += "## 7. Funding\n"
bp += md_table(["Round", "Amount", "Month", "Pre-money", "Post-round %"], [[r[0], money(r[1]), m, money(r[2]), f"{r[3]*100:.1f}%"] for r, m in zip(rounds, ["1", "10", "28"]) if r[1]] + [["Loan facility (non-dilutive)", money(I['Loan facility (venture debt / working-capital line) - limit']), "16 and 22 (2 tranches)", "-", "0%"]])
bp += f"The loan facility is venture debt or a working-capital line at {I['Facility interest rate']*100:.0f}% p.a. plus a {I['Facility establishment fee']*100:.0f}% fee. It is interest-only until month {I['Facility repayment starts']} (Jan 2030), then repaid over {I['Facility repayment period']} months from operating cash flow, and fully repaid by Dec 2031. Total interest and fees: {money(sum(D['intx']))}. Replacing Series B with debt keeps the founders well above 51% and frees room for the key-people share plan. Ask the lender for a **$2m accordion** (extra limit on request) to absorb small shocks; a 6-month licensing or broker delay, or the Downside, would still need the board pivot gate and new money.\n\n"
bp += f"Founders and legacy B-Catcher/BHive holders keep {D['founder_pct']*100:.1f}% of the shares and **{D['founder_votes']*100:.1f}% of the votes** after Seed and Series A ({D['founder_check']}; floor 51% of votes). All key-people equity is **non-voting**: {I['Honey Pot key-people pool (non-voting B-class shares / options in Honey Pot Pty Ltd)']*100:.0f}% of Honey Pot, {I['BHive key-people pool (non-voting B-class shares / options in BHive Pty Ltd)']*100:.0f}% of BHive and a {I['Group executive share pool (parent) - % of fully diluted shares after Seed']*100:.0f}% group pool (see *04 - Organisation & Headcount / Key_People_Share_Plan_v3.md*). Lowest month-end cash with this plan: {money(km['mincash'])} (buffer $750k).\n\n## 8. Scenarios\n"
sr = D["scen"]["rows"]
bp += md_table(["Case", "Revenue 2031", "EBITDA 2031", "Peak funding", "Extra equity needed", "EBITDA positive"],
               [[c, money(sr["Net revenue Y5"][j]), money(sr["EBITDA Y5"][j]), money(sr["Peak funding requirement"][j]), money(sr["Extra equity needed (vs buffer + buy-back reserve)"][j]),
                 "> 2031" if "Beyond" in str(sr["First month EBITDA positive"][j]) else sr["First month EBITDA positive"][j]] for j, c in enumerate(D["scen"]["cases"])])
bp += ("**Board pivot gate (month 18):** if run-rate placements are below ~$25m p.a. (about 80% of plan), hold the volume-triggered hires, cut outsourced and build-partner spend, defer the secondary-market build and stay under the principal licensee longer. At downside volumes the full three-company build does not pay for itself.\n\n"
       "## 9. Governance and risk\nBoard with independent Chair (month 1) and two NEDs (month 13); independent investment-committee member (month 7); Head of Risk & Compliance reporting to the board. Key risks: regulatory perimeter, placement volume, broker misconduct, liquidity expectations (no guaranteed exits), client money reconciliation, cyber. See Launch Playbook for gates.\n\n"
       "## 10. Decisions required\n- Entity names, cap table and merger ratio of legacy B-Catcher and BHive shareholders\n- Licensing route (CAR then own AFSL - modelled)\n- First two asset types and issuer partners\n- BHive build partner; market-operator partner\n- Seed terms and ESOP pool\n- Approved hiring schedule with volume gates\n")
F[f"{f01}/BCatcher_Group_Master_Business_Plan_v10.md"] = bp

# Playbook & roadmap
pb = HDR("Launch Playbook & Execution Roadmap - Version 10", f01)
pb += "Month 1 = January 2027. Costs are from the model's set-up and opex registers. The full formatted version is `BCatcher_Group_Launch_Playbook_and_Budget_Book.docx`.\n\n"
phases = [("Phase 0 - Foundation (months 1-2) - Gate 1: structure, ownership, IP", [
    ["Close Seed; open group bank accounts", "CEO", "1", "raise costs 4%"], ["Incorporate Honey Pot, BHive, template SPV", "CEO / counsel", "1", "$8,000"],
    ["Legacy B-Catcher & BHive entity/IP due diligence; merger ratio; IP assignment", "CEO / counsel", "1-2", "$55,000"], ["Shareholders' agreement, constitution, ESOP", "CEO / counsel", "2-3", "$40,000"],
    ["Group tax structuring", "Fractional CFO", "2-3", "$35,000"], ["Appoint Chair; board charter; delegation matrix", "Board", "1", "Chair $90k p.a."]]),
    ("Phase 1 - Legal perimeter & design (months 1-3) - Gate 2: written regulatory opinion", [
    ["Perimeter opinion: AFSL, MIS, market, credit, AML", "External counsel", "1-3", "$90,000"], ["Licensing route decision", "Board", "3", "-"],
    ["Brand identity (3 brands)", "Marketing", "2-3", "$65,000"], ["BHive discovery & architecture; build partner RFP", "Product owner + build partner", "2-3", "$95,000"], ["Cloud landing zone", "Product owner + build partner", "2", "$15,000"]]),
    ("Phase 2 - Counterparties, products & build (months 3-6) - Gates 3-4", [
    ["Trustee/RE/custodian; fund admin & registry", "CEO/CIO / Fractional CFO", "5-6", "$30,000"], ["Trust deed & IM templates (property, credit, PE/UTEGBO)", "CEO/CIO / counsel", "3-6", "$130,000"],
    ["Due diligence on first 3 offers", "CEO/CIO + outsourced analysts", "4-6", "$105,000"], ["CAR appointment", "Head of Distribution", "4", "$25,000 + $12k/month"],
    ["Compliance framework; AML/CTF program", "Head of Risk, Compliance & Ops", "3-6", "$100,000"], ["Broker agreements; CRM & commission system; LMS", "Honey Pot", "3-6", "$155,000"],
    ["BHive core platform & wallet built up front (portals, admin, wallet, marketplace) and integrations", "Product owner + build partner", "3-8", "$805,000"], ["Privacy assessment; pre-launch pen test", "vCISO + build partner", "5-6", "$65,000"],
    ["Launch campaign & broker roadshow", "Head of Distribution", "6-8", "$105,000"]]),
    ("Phase 3 - Wholesale launch (month 7) - Gate 5", [
    ["First offers live to verified wholesale investors", "CEO/CIO / Head of Distribution", "7", "-"], ["Accredit first brokers", "Head of Distribution", "7+", "$1,500 per broker"],
    ["First reconciled subscriptions & reports", "Head of Risk & Ops + outsourced accountants", "7-9", "-"], ["Insurance in force (PI, cyber, D&O, tech E&O)", "Fractional CFO", "7", "~$11k/month"]]),
    ("Phase 4 - Series A & scale (months 10-18)", [
    ["Close Series A", "CEO", "10", money(rounds[1][1])], ["Own AFSL application", "Head of Risk, Compliance & Ops", "13-20", "$150,000"],
    ["Market-operator route & legal", "Product owner + build partner", "22-27", "$120,000"], ["Platform release 2 + mobile apps; Investment Analyst hire (m13); buy-back bridge opens (m19, capped)", "Product owner + CEO/CIO", "10-19", "$300,000"],
    ["SOC 2 Type I", "vCISO", "10-15", "$70,000"], ["Month-18 pivot gate", "Board", "18", "-"]]),
    ("Phase 5 - Own licence & marketplace build (months 19-30) - Gate 6", [
    ["Hire Fund Operations & Finance Officer (trigger ~1,000 investors) and Broker Relationship Manager (~40 brokers)", "Head of Risk / Head of Distribution", "19", "$250k + on-costs"], ["Hire Compliance & AML Officer (m22), Property Asset Manager (m25), Fund Accountant & Registry Officer (m31)", "Head of Risk / CEO", "22-31", "$410k + on-costs"], ["Own AFSL granted; exit CAR", "Head of Risk, Compliance & Ops", "24-25", "ASIC levy $2.5k/month"],
    ["Platform release 3 (SMSF/company/trust accounts, auto-reinvest, tax pack)", "Product owner + build partner", "16-21", "$120,000"], ["Expand build-partner retainer (no in-house engineers)", "Product owner + build partner", "25", "$50k/month"],
    ["Draw loan facility tranches (non-dilutive)", "CEO / CFO", "16 and 22", money(I['Loan facility (venture debt / working-capital line) - limit'])], ["Secondary engine, operator route & integration; pen test", "Product owner + build partner / vCISO", "22-30", "$680,000"]]),
    ("Phase 6 - Marketplace live & scale (months 31-60) - Gate 7: new products approved separately", [
    ["Activate secondary market with licensed operator", "Product owner + build partner", "31", "$15k/month + 40% fee share"], ["Platform release 4 (credit income marketplace, regular-investment plans, portfolio tools)", "Product owner + build partner", "28-33", "$150,000"], ["Platform release 5 (new offerings module)", "Product owner + build partner / Board", "40-48", "$200,000"], ["Token legal opinion; tokenised register pilot", "Product owner / counsel", "30-39", "$130,000"],
    ["Commission-only regional BDMs (VIC/QLD, WA/SA, family office)", "Head of Distribution", "13-49", "0.20% commission pool"], ["Capacity review: scale outsourced providers; hire only if they cannot hold service levels", "CEO / Board", "37-49", "-"],
    ["Retail, payments, lending: separate business cases", "Board", "-", "not modelled"]])]
for t, steps in phases:
    pb += f"## {t}\n" + md_table(["Step", "Owner", "Month", "Cost"], steps)
pb += "## Monthly KPI targets (year end, Base)\n" + md_table(["KPI"] + YRS, [["Capital placed"] + [mm(v) for v in D["cap"]], ["Managed AUM"] + [mm(v) for v in D["aum"]],
                                                                        ["Investors"] + [f"{v:,.0f}" for v in D["inv"]], ["Brokers"] + [f"{v:,.0f}" for v in D["brk"]],
                                                                        ["Revenue"] + [mm(v) for v in D["rev"]], ["EBITDA"] + [mm(v) for v in D["ebitda"]], ["Cash"] + [mm(v) for v in D["cash"]]])
F[f"{f01}/BCatcher_Group_Launch_Playbook_and_Roadmap_v10.md"] = pb

# =====================================================================  02 LEGAL
f02 = "02 - Corporate Structure & Legal"
lg = HDR("Legal, Licensing & Corporate Workplan and Budget - Version 10", f02)
lg += "Costs and timing for every legal, licensing and regulatory workstream. Regulatory scope remains as in *Corporate & Regulatory Implementation v1*; this document adds owners, months and money.\n\n## One-off legal, licensing & structuring costs\n"
legal_cats = ("Corporate & Legal Structure", "Licensing & Regulatory", "Product & Investment Structuring", "Security & Certification", "Corporate Raise Costs")
lrows = [[e, c, it, money(a), f"{int(s)}-{int(s)+int(n)-1}" if n > 1 else str(int(s))] for e, c, it, a, s, n, t in D["setup_items"] if c in legal_cats]
lg += md_table(["Entity", "Category", "Item", "Cost", "Months"], lrows)
lg += f"**Total legal, licensing, structuring, security & raise set-up: {money(sum(D['setup_cat'].get(c,0) for c in legal_cats))}**\n\n## Recurring regulatory, legal & governance costs\n"
orows = [[o[0], o[2], (f"{o[3]*100:.2f}% of capital" if str(o[4]).startswith('%') else f"${o[3]:,.0f} ({o[4]})"), int(float(o[5])) if o[5] not in (None, '') else '', money(o[7]), money(o[11])]
         for o in D["opex_items"] if o[1] in ("Licensing, Regulatory & Compliance", "Professional Fees (Legal, Audit, Tax)", "Board & Governance", "Insurance")]
lg += md_table(["Entity", "Item", "Rate", "Start month", "2027", "2031"], orows)
lg += ("## Licensing pathway (modelled)\n"
       "1. **Month 1-3** - perimeter opinion for each entity (AFSL authorisations, MIS, market licence/exemption, credit, AML/CTF reporting-entity allocation, privacy, payments).\n"
       "2. **Month 4** - Honey Pot appointed corporate authorised representative of a principal licensee ($25k + $12k/month to month 24).\n"
       "3. **Month 4-6** - AML/CTF program and AUSTRAC enrolment; wholesale-client verification procedures (s761G, s761GA, s708).\n"
       "4. **Month 13-20** - Honey Pot own AFSL application ($150k); responsible managers under RG 105; licence from month 25 (ASIC levy $2.5k/month).\n"
       "5. **Month 22-27** - BHive market-operator partner route, market rules and operating agreement ($120k); activation month 31 (deferred in v9 until AUM covers fixed costs).\n"
       "6. **Month 22-24** - digital-token legal register opinion ($40k) before any tokenised register pilot.\n"
       "7. **Not modelled** - ACL for mortgage/credit brokers, retail MIS, payments/card, lending. Each needs a separate business case.\n\n"
       "## Corporate actions\n- Incorporate Honey Pot Pty Ltd and BHive Pty Ltd as wholly owned subsidiaries of the B-Catcher parent\n- Confirm legacy entities (B-Catcher, BHive Capital Pty Ltd) are free of liabilities; assign IP\n- Agree merger ratio for legacy shareholders (cap table placeholder: 10,000,000 founder shares)\n- Shareholders' agreement, ESOP, board charter, delegation matrix, intercompany agreements (shared services $25k/month per subsidiary; BHive licence $20k + $10k/month)\n")
F[f"{f02}/BCatcher_Group_Legal_Licensing_Workplan_and_Budget_v10.md"] = lg

# =====================================================================  03 FINANCE (CSVs)
f03 = "03 - Financial Models & Budgets"
inp = wb["Inputs"]
irows = [["Assumption", "Live (Base)", "Unit", "Base", "Downside", "Upside", "Basis / source"]]
for r in range(9, inp.max_row + 1):
    a = inp.cell(r, 1).value
    if a:
        irows.append([a] + [inp.cell(r, c).value if not isinstance(inp.cell(r, c).value, (dt.date, dt.datetime)) else inp.cell(r, c).value.strftime("%d-%b-%Y") for c in range(2, 8)])
F[f"{f03}/Model_v10_Inputs_and_Assumptions.csv"] = to_csv(irows)
an = wb["Annual Summary"]
F[f"{f03}/Model_v10_Annual_Summary_2027-2031.csv"] = to_csv([[an.cell(r, c).value if not isinstance(an.cell(r, c).value, float) else round(an.cell(r, c).value, 2) for c in range(1, 8)] for r in range(4, an.max_row + 1) if an.cell(r, 1).value])
bu = wb["Budget Y1-Y2"]
F[f"{f03}/Budget_2027-2028_by_Cost_Centre_Quarterly_v10.csv"] = to_csv([[bu.cell(r, c).value if not isinstance(bu.cell(r, c).value, float) else round(bu.cell(r, c).value, 2) for c in range(1, 12)] for r in range(4, bu.max_row + 1) if bu.cell(r, 1).value])
for sh, fn in [("P&L Consolidated", "PL_Consolidated_Monthly"), ("Cash Flow", "Cash_Flow_Monthly"), ("Balance Sheet", "Balance_Sheet_Monthly")]:
    F[f"{f03}/Model_v10_{fn}.csv"] = monthly_sheet_csv(sh)
erows = [["Entity", "Line item"] + YRS + ["5-yr total"]]
for sh, ent in [("P&L B-Catcher", "B-Catcher"), ("P&L Honey Pot", "Honey Pot"), ("P&L BHive", "BHive")]:
    ws = wb[sh]
    for r in range(8, ws.max_row + 1):
        a = ws.cell(r, 1).value
        vals = [ws.cell(r, c).value for c in range(5, 65)]
        if a is None or all(v is None for v in vals) or "margin" in a:
            continue
        yrs = [round(sum(v or 0 for v in vals[12 * y:12 * y + 12])) for y in range(5)]
        erows.append([ent, a] + yrs + [sum(yrs)])
F[f"{f03}/Model_v10_PL_by_Entity_Annual.csv"] = to_csv(erows)
F[f"{f03}/Setup_Cost_Register_v10.csv"] = to_csv([["Entity", "Category", "Item", "Amount", "Start month", "Months", "Treatment"]] + D["setup_items"] + [["TOTAL", "", "", km["setup"]]])
F[f"{f03}/Opex_Register_v10.csv"] = to_csv([["Entity", "Category", "Item", "Rate", "Driver", "Start month", "End month"] + YRS] + [o[:3] + [f"{o[3]*100:.2f}%" if str(o[4]).startswith("%") else o[3]] + o[4:7] + [round(x) for x in o[7:]] for o in D["opex_items"]])
sc = wb["Scenarios"]
F[f"{f03}/Scenarios_and_Stress_Tests_v10.csv"] = to_csv([[sc.cell(r, c).value.strftime("%b-%Y") if isinstance(sc.cell(r, c).value, (dt.date, dt.datetime)) else (round(sc.cell(r, c).value, 0) if isinstance(sc.cell(r, c).value, float) else sc.cell(r, c).value) for c in range(1, sc.max_column + 1)] for r in range(4, 17) if sc.cell(r, 1).value])
fr = HDR("Financial Model v10 - Read Me", f03)
fr += ("## Status\n- **Current:** `BCatcher_Group_Integrated_Financial_Model.xlsx` (upload to this folder from the Claude session). 16 sheets, ~38,000 live formulas, zero errors, balance sheet balances.\n"
       "- **Superseded:** `BCatcher_BHive_36_Month_Financial_Model.xlsx` - one flat cost line, no entity split, no broker commissions or headcount link, cash negative from month 16.\n"
       "- The CSV files here are value exports of each sheet (Base case) so the numbers are readable in Dropbox without the workbook.\n\n"
       "## Sheets in the workbook\n" + md_table(["Sheet", "Contents"], [
           ["Cover", "Key outputs, model flow, colour code, integrity checks"], ["Inputs", "Scenario selector (Base/Downside/Upside), every assumption with source, stress levers"],
           ["Annual Summary", "Five-year P&L, volumes, entity EBITDA, cash, people, unit economics, key metrics"], ["Scenarios", "Nine cases run through the full model"],
           ["Funding & Returns", "Funding requirement, cap table, use of funds, investor returns"], ["Budget Y1-Y2", "Quarterly budget per cost centre + set-up budget by category"],
           ["P&L Consolidated / B-Catcher / Honey Pot / BHive", "Monthly P&Ls; intercompany eliminated on consolidation"], ["Cash Flow / Balance Sheet", "Monthly cash, funding, runway; balance check"],
           ["Revenue", "Capital placed, AUM roll-forward, investors, brokers, every fee line, cost drivers"], ["Headcount", f"{len(D['roles'])} roles; monthly FTE, loaded cost, hiring costs"],
           ["Setup Costs / Opex", "Every one-off and recurring cost with timing, driver and basis"]]) +
       "## How to use\nChange `Inputs!B4` to switch scenario; change blue cells to update assumptions; stress levers (placement multiplier, licensing delay, broker delay, tech overrun, hiring deferral) sit at the bottom of Inputs.\n")
F[f"{f03}/00 - Financial Model v10 Read Me.md"] = fr

# =====================================================================  04 HEADCOUNT
f04 = "04 - Organisation & Headcount"
hc = wb["Headcount"]
hrows = [["Entity", "Department", "Role", "Type", "Start month", "FTE", "Step month", "FTE from step", "End month", "Base salary", "Loaded cost p.a. (2027 $)", "Notes"] + [f"FTE end {y}" for y in YRS] + [f"Cost {y}" for y in YRS]]
N = len(D["roles"])
cost0 = None
for r in range(9, hc.max_row + 1):
    if hc.cell(r, 1).value == "=A9" or (hc.cell(r, 3).value and hc.cell(r, 3).value == hc.cell(9, 3).value and r > 9):
        cost0 = r
        break
for i in range(N):
    r = 9 + i
    fte = [hc.cell(r, 14 + 12 * y + 11).value or 0 for y in range(5)]
    cost = [sum((hc.cell(cost0 + i, 14 + mth).value or 0) for mth in range(12 * y, 12 * y + 12)) for y in range(5)] if cost0 else [""] * 5
    hrows.append([hc.cell(r, c).value for c in range(1, 11)] + [round(hc.cell(r, 12).value or 0), hc.cell(r, 11).value] + fte + [round(x) if isinstance(x, (int, float)) else x for x in cost])
F[f"{f04}/Headcount_Plan_v10_{len(D['roles'])}_Roles.csv"] = to_csv(hrows)
org = HDR("Organisation Structure & Headcount Plan - Version 10", f04)
org += ("Supersedes the v4 plan (49 roles) and `BCatcher_BHive_Headcount_Org_Structure_36_Months.xlsx`. Every role below is costed monthly in the model.\n\n"
        f"Loaded cost = salary x FTE x (1 + 12% super + 5.45% payroll tax + 1% workers comp + 3% other). Contractors: salary equivalent + 20%. Recruitment 15% of salary for hires after month 1; $3,500 equipment per starter; wages +3.5% p.a.\n\n## FTE and cost by entity\n")
org += md_table(["Year-end FTE"] + YRS, [["B-Catcher"] + [f"{v:.1f}" for v in D["fte_bc"]], ["Honey Pot"] + [f"{v:.1f}" for v in D["fte_hp"]], ["BHive"] + [f"{v:.1f}" for v in D["fte_bh"]],
                                         ["**Group**"] + [f"**{v:.1f}**" for v in D["fte"]], ["Staff cost (salaries + on-costs)"] + [money(v) for v in D["staff"]]])
org += ("## Operating model: four core employees (+6 volume-triggered hires), everything else bought in\n```\n"
        "Board (independent Chair; 2 NEDs from m13) + Investment Committee (independent member from m7)\n"
        "└── 1. Group CEO & CIO (acting CTO)          -> outsourced investment analysts, valuers & property managers; AI research tools\n"
        "    │                                          -> fractional CFO; fractional product owner + build partner; vCISO\n"
        "    ├── 2. Head of Risk, Compliance & Operations -> external compliance consultant; outsourced bookkeeping & fund accounting;\n"
        "    │                                             outsourced fund admin & registry; AI compliance monitoring & reconciliation\n"
        "    ├── 3. Head of Distribution (Honey Pot)    -> commission-only regional BDMs; external brokers (commission);\n"
        "    │                                          freelance marketing + AI content tools\n"
        "    └── 4. Investor Communications & BDM       -> investor & platform-customer updates, reporting, webinars, events;\n"
        "                                                  direct investors, family offices & new issuers; outsourced service desk + AI agent\n"
        "```\n\n"
        "## Why a small team can run $400-500m AUM\n"
        "Wholesale managers of this size commonly run with a handful of principals because the work is contracted out: the trustee/RE, custodian, fund administrator and registry hold assets and records; property managers run buildings; brokers are paid only on capital placed; and the platform is built and run by a technology partner. The four employees are the people who must exist inside the company: the accountable executive and investment decision-maker, the licensed responsible manager, the head of distribution, and the person investors, platform customers and new issuers actually hear from. Two of them (CEO and Head of Risk) are the AFSL responsible managers.\n\n"
        "## What replaced each removed role (v4 had 49 roles)\n")
org += md_table(["Removed role(s)", "Replaced by (annual cost in model, year 5)"], [
    ["CIO, CTO, Head of Product, Head of Secondary Market, all engineers, data/AI engineer, market-ops analyst", "CEO (combined) + fractional product owner + build partner retainer ($32k/month) + licensed market operator"],
    ["CFO, Financial Controller, Finance & Operations Manager, fund & commission accountants, EA", "Fractional CFO + Head of Risk & Ops + outsourced bookkeeping/fund accounting ($6k/month + $0.80/investor) + AI reconciliation"],
    ["Compliance managers, officers and analysts", "Head of Risk, Compliance & Ops + external compliance consultant + AI compliance monitoring"],
    ["Investment directors, analysts, asset managers", "CEO/CIO + outsourced analysis & asset-management support ($7k/month, $15k from m37) + external property managers (paid by trusts) + AI research"],
    ["Honey Pot MD, broker network manager, marketing staff, IR / communications manager", "Head of Distribution + Investor Communications & BDM ($150k) + freelance marketing ($4k/month) + AI content tools"],
    ["All other salaried BDMs", "Head of Distribution + Investor Comms & BDM; commission-only regional BDMs (0.20% of capital placed)"],
    ["Investor services, commission ops, platform support officers", "Outsourced service desk ($4k/month + $2/investor) + AI support agent"],
    ["People & culture; office", "Outsourced HR & payroll; co-working ($3.5k/month)"],
])
org += ("\n## Key-person and capacity risk\n"
        "With a core team of four (ten with the triggered hires) the plan is exposed to any one person leaving. Mitigations: key-person insurance, documented processes with each provider, a deputy responsible manager arranged with the compliance consultant, and service-level agreements with every outsourced provider. Add an employee only when a provider cannot hold its service levels.\n\n")
for ent in ["B-Catcher", "Honey Pot", "BHive"]:
    org += f"## {ent} roles\n" + md_table(["Role", "Type", "Start month", "FTE", "Base salary", "Loaded p.a."],
                                         [[r[2], r[3], int(r[4]), f"{r[5]:.1f}", money(r[6]), money(r[7])] for r in sorted([x for x in D["roles"] if x[0] == ent], key=lambda x: x[4])])
org += "## Outsourced / partner services\nPrincipal licensee (CAR), trustee/RE, custodian, fund administrator & registry, external audit & tax, legal counsel, build partner, penetration testing, MDR security service, licensed market operator, property managers. All costed in the Setup and Opex registers.\n"
F[f"{f04}/Organisation_Structure_and_Headcount_v10.md"] = org

# =====================================================================  05 HONEY POT
f05 = "05 - Honey Pot - Distribution & Brokers"
hp = HDR("Honey Pot - Business Case & Budget - Version 10", f05)
hpl = {k: annual("P&L Honey Pot", l) for k, l in [("ext", "External revenue"), ("comm", "Broker commissions"), ("gp", "GROSS PROFIT"), ("staff", "Salaries, superannuation & on-costs"),
                                                  ("opex", "TOTAL OPERATING EXPENSES"), ("ebitda", "EBITDA")]}
hp += ("Procedures remain as in the *Broker Network Operating Manual v1*. This document adds the economics, budget, staffing and targets.\n\n## Role in the group\nHoney Pot is the distribution engine: it pays away most of its fees to brokers and its strategic value is the capital and AUM it feeds to B-Catcher and BHive. It runs as a cost centre at a loss until scale, while the group consolidates to profit.\n\n## Five-year P&L (incl. intercompany charges)\n")
hp += md_table(["$"] + YRS, [["Capital placed (all channels)"] + [money(v) for v in D["cap"]], ["Placement fees"] + [money(v) for v in hpl["ext"]], ["Broker commissions"] + [money(-v) for v in hpl["comm"]],
                             ["Gross profit"] + [money(v) for v in hpl["gp"]], ["Staff costs"] + [money(-v) for v in hpl["staff"]], ["Total operating expenses"] + [money(-v) for v in hpl["opex"]],
                             ["EBITDA"] + [money(v) for v in hpl["ebitda"]], ["Active brokers (year end)"] + [f"{v:,.0f}" for v in D["brk"]]])
hp += "## Commission & fee schedule\n" + md_table(["Product", "Placement fee (paid by issuer)", "Broker share", "Broker receives (% of capital)", "Honey Pot retains"], [
    [p, f"{f*100:.1f}%", "60%", f"{f*0.6*100:.2f}%", f"{f*0.4*100:.2f}%"] for p, f in [("Property trusts", I["Placement fee - property trusts"]), ("Private credit", I["Placement fee - private credit"]), ("PE / pre-IPO", I["Placement fee - private equity / pre-IPO"]), ("ASX placements / IPO", I["Placement fee - ASX placements / IPO"])]])
hp += "Commissions are paid only on cleared, accepted subscriptions; clawback on refunds/cancellations; GST separately; dual authorisation and monthly reconciliation. 30% of capital is direct (no commission).\n\n"
hp += "## Broker growth targets\n" + md_table(["", *YRS], [["Active accredited brokers"] + [f"{v:,.0f}" for v in D["brk"]],
                                                              ["Capital per broker (broker channel)"] + [money(D["cap"][i] * 0.7 / D["brk"][i]) if D["brk"][i] else "-" for i in range(5)]])
hp += "## Set-up costs\n" + md_table(["Item", "Cost", "Month"], [[x[2], money(x[3]), int(x[4])] for x in D["setup_items"] if x[0] == "Honey Pot"])
hp += "## Recurring costs\n" + md_table(["Item", "Rate / driver", "2027", "2029", "2031"], [[o[2], (f"{o[3]*100:.2f}% of capital" if str(o[4]).startswith('%') else f"${o[3]:,.0f} {o[4]}"), money(o[7]), money(o[9]), money(o[11])] for o in D["opex_items"] if o[0] == "Honey Pot"])
hp += "## Team\n" + md_table(["Role", "Start month", "Base salary"], [[r[2], int(r[4]), money(r[6])] for r in sorted([x for x in D["roles"] if x[0] == "Honey Pot"], key=lambda x: x[4])])
hp += "## KPIs\nActive compliant brokers; productive brokers; qualified referrals; referral-to-subscription conversion; capital per broker; cost of acquisition; time-to-subscribe; complaints & breaches; clawbacks; net fee yield.\n"
hp += ("## Marketing: the B-SMART campaign\n"
       "Honey Pot runs the group's **B-SMART.** campaign (full plan: *09 - Notes / Marketing_Campaign_Plan_B-SMART_v2.md*): master line *B-SMART - the deals that build real wealth used to be invite-only. Be smart. Join the Hive.* "
       "Product lines: B-an Owner (property), B-the Bank (private credit), B-Early (pre-IPO), B-Informed (AI + experts), B-Free (buy-back bridge, then secondary market), B-Sure (every fee shown in dollars). "
       "Launch mechanic: *The First Hive* - 500 Founding Members. v10 adds a $15k/month national growth layer (media, PR, partnerships) from month 13 on top of performance marketing at 0.15% of capital placed.\n\n"
       "## Honey Pot Partner Program (broker tiers)\n")
hp += md_table(["Tier", "Who", "What they get"], [
    ["Worker", "Newly accredited brokers and referrers", "Portal access, training, standard commission on cleared money"],
    ["Builder", "Brokers with consistent placements", "Priority allocation on new offers, co-marketing support, quarterly reviews"],
    ["Queen", "Top-producing partners", "First look at deals, joint events, dedicated relationship manager"]])
hp += "Commissions are paid only on cleared, accepted subscriptions and stay inside the placement fee; tier benefits are service and access, never extra conflicted payments.\n\n"
F[f"{f05}/Honey_Pot_Business_Case_and_Budget_v10.md"] = hp
F[f"{f05}/Honey_Pot_Broker_Commission_Schedule_v10.csv"] = to_csv([["Product", "Mix of capital", "Placement fee % of capital", "Broker share of fee", "Broker commission % of capital", "Honey Pot retained % of capital", "Condition"]] +
                                                                [[p, mx, f, 0.6, round(f * 0.6, 4), round(f * 0.4, 4), "Cleared & accepted subscriptions only; clawback on refund"] for p, mx, f in [("Property trusts", 0.45, I["Placement fee - property trusts"]), ("Private credit", 0.25, I["Placement fee - private credit"]), ("PE / pre-IPO", 0.2, I["Placement fee - private equity / pre-IPO"]), ("ASX placements / IPO", 0.1, I["Placement fee - ASX placements / IPO"])]])

# =====================================================================  06 BHIVE
f06 = "06 - BHive - Wallet & Secondary Market"
bh = HDR("BHive - Build Plan & Budget - Version 10", f06)
bpl = {k: annual("P&L BHive", l) for k, l in [("plat", "Technology / platform fees"), ("offer", "Offer onboarding fees"), ("sec", "Exchange fees"), ("bb", "Buy-back resale gains"), ("ext", "External revenue"), ("ic", "Intercompany revenue"),
                                              ("cogs", "Total cost of sales"), ("staff", "Salaries, superannuation & on-costs"), ("opex", "TOTAL OPERATING EXPENSES"), ("ebitda", "EBITDA"), ("rd", "Other income - R&D tax incentive (accrued)"), ("capex", "Memo: capitalised set-up spend")]}
bh += "Requirements remain as in *Product & Technical Requirements v1*. This adds the build sequence, vendor stack, budget, team and platform economics.\n\n## Build sequence and cost\n"
bh += md_table(["Item", "Cost", "Months", "Treatment"], [[x[2], money(x[3]), f"{int(x[4])}-{int(x[4])+int(x[5])-1}" if x[5] > 1 else int(x[4]), x[6]] for x in D["setup_items"] if x[0] == "BHive"])
bh += f"**Total BHive set-up: {money(D['setup_ent']['BHive'])}** (capex amortised over 36 months). Plus Honey Pot CRM/commission system $75,000.\n\n"
bh += ("## Release plan\n- **Core platform & BHive wallet, built up front (months 3-8):** onboarding, KYC/AML, wholesale verification, offer marketplace & data room, e-sign subscriptions, registry reconciliation, investor wallet (holdings, documents, distributions, tax), broker portal, admin console\n- **Launch (month 7):** live with first offers; outsourced service desk + AI tier-1 agent\n- **Release 2 (months 10-15):** buy-back requests in the wallet, investor comms hub, referral & Founding Member features, analytics\n- **Mobile apps (months 13-18)**; SOC 2 Type I\n- **Release 3 (months 16-21):** SMSF, company & trust accounts; auto-reinvest; tax reporting pack\n- **Months 22-30:** secondary engine and operator integration built\n- **Release 4 (months 28-33):** private-credit income marketplace, regular-investment plans, portfolio tools\n- **Month 31:** secondary market activated with licensed operator (40% fee share + $15k/month), once AUM is ~$125m\n- **Months 34-39:** tokenised register pilot only after legal register opinion\n- **Release 5 (months 40-48):** new offerings module (managed portfolios, insurance; retail-ready module subject to licence)\n\n## Platform P&L (incl. intercompany licence income)\n")
bh += md_table(["$"] + YRS, [["Technology fees (0.25% p.a.)"] + [money(v) for v in bpl["plat"]], ["Offer onboarding fees"] + [money(v) for v in bpl["offer"]], ["Exchange fees (0.25%)"] + [money(v) for v in bpl["sec"]], ["Buy-back resale gains"] + [money(v) for v in bpl["bb"]],
                             ["Intercompany licence income"] + [money(v) for v in bpl["ic"]], ["Operator share (cost of sales)"] + [money(-v) for v in bpl["cogs"]], ["Staff costs"] + [money(-v) for v in bpl["staff"]],
                             ["Total operating expenses"] + [money(-v) for v in bpl["opex"]], ["EBITDA"] + [money(v) for v in bpl["ebitda"]], ["R&D tax incentive (accrued)"] + [money(v) for v in bpl["rd"]],
                             ["Capitalised build"] + [money(v) for v in bpl["capex"]], ["Value traded on secondary market"] + [money(v) for v in D["trade"]], ["Investors on platform"] + [f"{v:,.0f}" for v in D["inv"]]])
bh += "## Running costs\n" + md_table(["Item", "Rate / driver", "2027", "2029", "2031"], [[o[2], f"${o[3]:,.2f} {o[4]}" if o[3] < 10 else f"${o[3]:,.0f} {o[4]}", money(o[7]), money(o[9]), money(o[11])] for o in D["opex_items"] if o[0] == "BHive"])
bh += "## Team\n" + md_table(["Role", "Type", "Start month", "FTE", "Base salary"], [[r[2], r[3], int(r[4]), f"{r[5]:.1f}", money(r[6])] for r in sorted([x for x in D["roles"] if x[0] == "BHive"], key=lambda x: x[4])])
bh += ("## Vendor stack to procure (RFP list)\nBuild partner; AWS Sydney; identity/KYC/AML provider; e-signature; CRM (shared with Honey Pot); fund administrator & registry; banking/payment reconciliation; document vault; MDR/SIEM security; penetration tester; SOC 2 auditor; licensed market operator; AI support agent tooling.\n\n"
       "## Buy-back facility (the exit bridge until the secondary market)\nv9: available from year 2 for property tokens (20% yrs 1-2, 15% yr 3, 10% yr 4, 5% yr 5 below independent valuation) and PE trust units (10% / 8% / 5%). Discretionary and capped at 5% of eligible units a year. Before the secondary market opens (month 31) the model assumes 4% of eligible units are bought back a year (5% Downside = cap); after it opens, requests go to the secondary market first and the group buys only the residual (1% a year). Resale within ~4 months with a price haircut plus 1.5% costs; 1.5% cash reserve on top of the minimum buffer.\n\n"
       + md_table(["$"] + YRS, [["Units bought back (valuation)"] + [money(v) for v in D["bbnav"]], ["Resale gains"] + [money(v) for v in D["rbb"]], ["Units held for resale (year end)"] + [money(v) for v in D["bbinv"]], ["Cash reserve required (year end)"] + [money(v) for v in D["bbres"]]])
       + "Requires a legal and capital-adequacy opinion before launch ($40k in set-up).\n\n"
       "## Not in scope\nBHive Coin, BHive Card, token lending, retail access (all from the July 2026 BHive Capital pack) - excluded until separate legal and business cases.\n")
F[f"{f06}/BHive_Build_Plan_and_Budget_v10.md"] = bh

# =====================================================================  07 INVESTOR
f07 = "07 - Investor Deck & Capital Raise"
cr = HDR("Capital Raise Summary - Version 10", f07)
cr += f"## The ask\n**{mm(rounds[0][1])} Seed** to build, license and launch - part of a **{mm(km['raised'])}** equity plan (Seed + Series A) plus a **{mm(I['Loan facility (venture debt / working-capital line) - limit'])}** non-dilutive loan facility. No Series B.\n\n"
cr += md_table(["Round", "Amount", "Close", "Pre-money", "Post-round %", "Milestones funded"], [
    [rounds[0][0], money(rounds[0][1]), "Month 1 (Jan 2027)", money(rounds[0][2]), f"{rounds[0][3]*100:.1f}%", "Platform build, licensee & trustee, first 3 offers, wholesale launch month 7"],
    [rounds[1][0], money(rounds[1][1]), "Month 10 (Oct 2027)", money(rounds[1][2]), f"{rounds[1][3]*100:.1f}%", "Licensed distribution live, first reconciled subscriptions, signed pipeline"],
] + ([["Series B", money(rounds[2][1]), "Month 28 (Apr 2029)", money(rounds[2][2]), f"{rounds[2][3]*100:.1f}%", "Own AFSL granted, ~$100m AUM, ~90 brokers; funds the secondary-market build (live month 31)"]] if rounds[2][1] else []) + [["Loan facility", money(I['Loan facility (venture debt / working-capital line) - limit']), "Months 16 & 22", "-", "0% (debt)", f"Carries the burn to break-even; {I['Facility interest rate']*100:.0f}% p.a., repaid Jan 2030 - Dec 2031 from cash flow"]])
cr += "## Use of funds - first 24 months\n" + md_table(["Use", "Amount"], [[u, money(v)] for u, v in D["uses"]])
cr += "## Illustrative returns (end of year 5)\n" + md_table(["Investor", "Invested", "MOIC", "IRR"], [[n, money(a), f"{x:.1f}x", f"{y*100:.0f}%"] for n, a, x, y in D["returns"] if a])
cr += (f"Exit basis: Q4 2031 annualised revenue {mm(D['exit']['rev_rr'])} and EBITDA {mm(D['exit']['ebitda_rr'])}; equity value {mm(D['exit']['eq'])} (average of {I['Exit multiple - EBITDA (year 5)']:.0f}x EBITDA and {I['Exit multiple - revenue (year 5)']:.0f}x revenue, plus cash, less the 25% key-people stakes in Honey Pot and BHive; the facility is repaid by then). Illustrative only; later rounds not modelled.\n\n"
       "## Evidence still required before external circulation (from Content Brief)\n- Verified market sizing and competitors\n- Final legal names, ownership and merger ratio\n- Legal opinions on AFSL / market pathway\n- Signed issuer pipeline and broker letters of intent\n- Supplier quotes replacing model estimates\n- Founder and team bios\n- Independent review of the model and raise terms\n")
F[f"{f07}/Capital_Raise_Summary_v10.md"] = cr
# slide script from pptx
from pptx import Presentation
prs = Presentation(os.path.join(os.path.dirname(MODEL), "BCatcher_Group_Investor_Deck_v10.pptx"))
ss = HDR("Investor Deck v10 - Slide Script", f07) + "Text of every slide in `BCatcher_Group_Investor_Deck_v10.pptx` (upload the .pptx to this folder). Chart data is listed under each slide.\n\n"
for i, slide in enumerate(prs.slides, 1):
    texts, charts = [], []
    for shp in slide.shapes:
        if shp.has_text_frame and len(shp.text_frame.text.strip()) > 2:
            texts.append(shp.text_frame.text.strip().replace("\n", " / "))
        if getattr(shp, "has_chart", False) and shp.has_chart:
            ch = shp.chart
            for s in ch.plots[0].series:
                charts.append(f"{s.name}: " + ", ".join(f"{c}={v}" for c, v in zip(ch.plots[0].categories, s.values)))
        if shp.has_table:
            for row in shp.table.rows:
                texts.append(" | ".join(c.text.replace("\n", " ") for c in row.cells))
    ss += f"## Slide {i}: {texts[0] if texts else ''}\n" + "".join(f"- {t}\n" for t in texts[1:] if "Confidential planning document  |" not in t) + "".join(f"- *Chart* {c}\n" for c in charts) + "\n"
F[f"{f07}/Investor_Deck_v10_Slide_Script.md"] = ss

# =====================================================================  08 ARCHIVE
f08 = "08 - Historical Reference Documents"
ar = HDR("Archive Index - what was carried forward", f08)
ar += md_table(["Document", "Origin", "Carried into v10", "Excluded / not relied on"], [
    ["B-Catcher-9.pdf, B-Catcher 2.0.pdf", "B-Catcher 2025 investor memorandum", "Group name, tokenised RWA concept, wallet concept, governance intent", "DOGE mining (5.4 MW), Moonshot lottery, blockchain insurance, THORChain treasury, $100m valuation"],
    ["UTEGBO term sheets (Mac Archive / Deals / Crypto)", "B-Catcher 2025", "PE / pre-IPO unit trust with buy-back windows (15% / 10% / 8%) - kept, with a funded reserve", "-"],
    ["BHive Capital Finance Model.xlsx / Platform Forecast", "BHive July 2026", "Fee schedule adopted in full (entry, management, technology, exchange); investor tiers as reference", "$731m AUM, $24.5m revenue, 15x valuation; lending, BHive Coin revenue"],
    ["BHive Capital Executive Summary / IM / Legal Brief / Token Architecture / Retail Structure", "BHive July 2026", "Single platform, ERC-1400 tokens, CAR route, investor fee schedule (1.5% entry, 1.25% management, 0.25% exchange; technology fee cut from 0.5% to 0.25% in v9), exit ladder (25% / 15% / 10%), AI-native operating model", "Retail from day one, BHive Card/Coin, token lending"],
    ["BCatcher_Group_Operational_Business_Plan_v2.docx", "Nov 2026", "Entity boundaries, broker categories, stress tests, launch gates", "-"],
    ["BCatcher_Rebuilt_Investor_Deck_2026.pptx", "Nov 2026", "Slide sequence", "Superseded by Investor Deck v10"]])
ar += "Originals are unchanged. Legacy shareholder positions in B-Catcher and BHive Capital Pty Ltd must be reconciled into the group cap table (placeholder 10,000,000 founder shares).\n"
F[f"{f08}/00 - Archive Index v10.md"] = ar

os.makedirs(OUTDIR, exist_ok=True)
for rel, content in F.items():
    p = os.path.join(OUTDIR, rel)
    os.makedirs(os.path.dirname(p), exist_ok=True)
    open(p, "w").write(content)
    print(f"{len(content)/1024:7.1f} KB  {rel}")
