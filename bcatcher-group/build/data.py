"""Source data for the B-Catcher Group integrated financial model.

Every number here is a planning assumption (AUD, 2027 dollars) unless stated.
Basis notes are carried into the workbook next to each line so the reader can
see where each figure came from and replace it with a quote when one exists.
"""
import datetime as dt

N_MONTHS = 60
START_DATE = dt.date(2027, 1, 31)  # month 1 = Jan 2027 (month-end dates)
ENTITIES = ["B-Catcher", "Honey Pot", "BHive"]

# ---------------------------------------------------------------- INPUTS
# (name, label, unit, base, downside, upside, note, fmt)
# fmt: money | pct | num | dec | int | date | mult
INPUT_SECTIONS = [
    ("TIMING", [
        ("LaunchMonth", "Controlled wholesale launch (first subscriptions)", "model month", 7, 9, 6,
         "Execution Roadmap: build days 61-120, launch days 121-180. Month 7 = Jul 2027.", "int"),
        ("SecondaryMonth", "BHive secondary market activation", "model month", 19, 25, 16,
         "Roadmap months 6-18 for market-operator route; gated on licensed partner.", "int"),
    ]),
    ("CAPITAL PLACED (run-rate per operating year after launch)", [
        ("CapY1", "Capital placed - operating year 1", "$ p.a.", 25_000_000, 10_000_000, 40_000_000,
         "Prior model Y1 $12m. ~25 brokers + direct IR. Validate against signed issuer pipeline.", "money"),
        ("CapY2", "Capital placed - operating year 2", "$ p.a.", 75_000_000, 30_000_000, 110_000_000,
         "Prior model Y2 $30m. ~60 brokers x ~$0.8m + direct.", "money"),
        ("CapY3", "Capital placed - operating year 3", "$ p.a.", 150_000_000, 60_000_000, 210_000_000,
         "~100 brokers x ~$1m + direct.", "money"),
        ("CapY4", "Capital placed - operating year 4", "$ p.a.", 240_000_000, 95_000_000, 320_000_000,
         "BHive model Y4 new AUM $175m; ~140 brokers.", "money"),
        ("CapY5", "Capital placed - operating year 5", "$ p.a.", 340_000_000, 135_000_000, 450_000_000,
         "BHive model Y5 new AUM $481m (not relied on; deliberately lower). ~180 brokers.", "money"),
        ("RampMonths", "Months to reach run-rate after launch", "months", 6, 9, 4,
         "Linear ramp from launch; year-1 placements land below run-rate.", "int"),
    ]),
    ("PRODUCT MIX (% of capital placed)", [
        ("MixProp", "Residential & commercial property trusts", "%", 0.45, 0.45, 0.45, "Master Business Plan v3 phase-one catalogue.", "pct"),
        ("MixCredit", "Private credit", "%", 0.25, 0.25, 0.25, "", "pct"),
        ("MixPE", "Private company equity / pre-IPO (incl. UTEGBO-style unit trusts)", "%", 0.20, 0.20, 0.20, "B-Catcher UTEGBO term sheet (2025).", "pct"),
        ("MixASX", "ASX placements / IPO allocations (third-party issuers)", "%", 0.10, 0.10, 0.10, "Must sum to 100% with lines above.", "pct"),
    ]),
    ("HONEY POT - DISTRIBUTION ECONOMICS", [
        ("FeeProp", "Placement fee - property trusts", "% of capital", 0.030, 0.025, 0.030, "Paid by issuer/trust; disclosed in offer docs. Market range 2-4%.", "pct"),
        ("FeeCredit", "Placement fee - private credit", "% of capital", 0.015, 0.0125, 0.015, "Market range 1-2%.", "pct"),
        ("FeePE", "Placement fee - private equity / pre-IPO", "% of capital", 0.050, 0.040, 0.050, "Market range 4-6% for private raises.", "pct"),
        ("FeeASX", "Placement fee - ASX placements / IPO", "% of capital", 0.040, 0.035, 0.040, "Typical lead-manager/stamping fees 4-6%; Honey Pot share.", "pct"),
        ("BrokerChannel", "Share of capital sourced via brokers", "%", 0.70, 0.60, 0.75, "Remainder is direct/Honey Pot IR.", "pct"),
        ("BrokerPayout", "Broker share of placement fee (commission)", "% of fee", 0.60, 0.65, 0.55, "Honey Pot Operating Manual: commissions on cleared subscriptions only.", "pct"),
        ("AvgTicket", "Average investor subscription", "$", 150_000, 120_000, 175_000, "Wholesale investors; s708/s761G thresholds.", "money"),
        ("BrokerY1", "Active accredited brokers - end op. year 1", "brokers", 25, 12, 40, "", "int"),
        ("BrokerY2", "Active accredited brokers - end op. year 2", "brokers", 60, 30, 90, "", "int"),
        ("BrokerY3", "Active accredited brokers - end op. year 3", "brokers", 100, 50, 150, "", "int"),
        ("BrokerY4", "Active accredited brokers - end op. year 4", "brokers", 140, 70, 200, "", "int"),
        ("BrokerY5", "Active accredited brokers - end op. year 5", "brokers", 180, 90, 250, "", "int"),
        ("BrokerChurn", "Broker churn (replaced each year)", "% p.a.", 0.15, 0.20, 0.10, "Drives onboarding cost of replacement brokers.", "pct"),
    ]),
    ("B-CATCHER - INVESTMENT & ASSET MANAGEMENT ECONOMICS", [
        ("PropLVR", "Property trust gearing (debt / gross asset value)", "%", 0.45, 0.40, 0.50, "Typical wholesale property syndicate LVR 40-50%. Debt sits in the trust, not the group.", "pct"),
        ("AcqFee", "Property acquisition fee", "% of purchase price (GAV)", 0.015, 0.010, 0.020, "Market standard 1-2% of purchase price, paid by trust on settlement.", "pct"),
        ("StructFee", "Structuring / establishment fee - credit & PE products", "% of capital", 0.010, 0.0075, 0.0125, "Charged to fund/SPV (not ASX placements).", "pct"),
        ("MgmtProp", "Management fee - property (on gross asset value)", "% GAV p.a.", 0.0065, 0.0055, 0.0070, "Wholesale property trusts 0.5-0.85% of GAV; trustee/custody fees borne separately by trust.", "pct"),
        ("MgmtCredit", "Net management fee - private credit", "% AUM p.a.", 0.0100, 0.0085, 0.0110, "", "pct"),
        ("MgmtPE", "Net management fee - private equity", "% AUM p.a.", 0.0150, 0.0125, 0.0175, "BHive BPPF 1.5-2.0%.", "pct"),
        ("GrowthProp", "Property equity AUM revaluation growth", "% p.a.", 0.03, 0.0, 0.04, "BHive model used 5%; reduced.", "pct"),
        ("GrowthPE", "PE AUM revaluation growth", "% p.a.", 0.05, 0.0, 0.08, "", "pct"),
        ("RunoffProp", "Property AUM run-off (sales/redemptions)", "% p.a.", 0.05, 0.08, 0.04, "", "pct"),
        ("RunoffCredit", "Private credit run-off (loan repayments)", "% p.a.", 0.40, 0.50, 0.35, "Avg loan term ~2.5 years.", "pct"),
        ("RunoffPE", "PE AUM run-off (exits / buy-backs)", "% p.a.", 0.10, 0.12, 0.08, "UTEGBO buy-back windows yr1/2/4.", "pct"),
        ("MgrAdminCost", "Fund administration borne by manager", "% AUM p.a.", 0.0005, 0.0006, 0.0004, "Trust-level costs (RE, custody, valuations, trust audit) are borne by the vehicles, not the group.", "pct"),
    ]),
    ("BHIVE - PLATFORM & MARKETPLACE ECONOMICS", [
        ("PlatformFee", "Platform administration fee (charged to issuers/trusts)", "% AUM p.a.", 0.0030, 0.0020, 0.0035, "BHive model used 0.5%. Australian platforms charge ~0.25-0.5% admin.", "pct"),
        ("OfferSize", "Average offer size", "$", 5_000_000, 5_000_000, 5_000_000, "Drives number of offers onboarded.", "money"),
        ("OfferFee", "Offer onboarding fee (data room, workflows, registry set-up)", "$ per offer", 20_000, 12_500, 25_000, "Prior model $7,500 issuer fee; covers data room, workflows, registry set-up, investor reporting.", "money"),
        ("SecTurnover", "Secondary market turnover", "% of AUM p.a.", 0.08, 0.04, 0.12, "BHive model assumed 15%.", "pct"),
        ("SecFee", "Secondary transaction fee (buyer + seller combined)", "% of value", 0.010, 0.008, 0.012, "", "pct"),
        ("SecPartnerShare", "Licensed market-operator partner share", "% of fee", 0.40, 0.50, 0.35, "Partner route avoids BHive holding its own market licence.", "pct"),
        ("SecRamp", "Turnover achieved in first 12 months live", "% of run-rate", 0.50, 0.35, 0.60, "", "pct"),
    ]),
    ("INTERCOMPANY CHARGES (eliminated on consolidation)", [
        ("IC_SSF_HP", "B-Catcher shared-services fee charged to Honey Pot", "$ / month", 25_000, 25_000, 25_000, "Finance, compliance, office, executive. From month 3.", "money"),
        ("IC_SSF_BH", "B-Catcher shared-services fee charged to BHive", "$ / month", 25_000, 25_000, 25_000, "From month 1.", "money"),
        ("IC_LIC_HP", "BHive platform licence charged to Honey Pot", "$ / month", 20_000, 20_000, 20_000, "Broker portal, CRM, commission ledger. From launch.", "money"),
        ("IC_LIC_BC", "BHive platform licence charged to B-Catcher", "$ / month", 10_000, 10_000, 10_000, "Investor reporting, registry. From launch.", "money"),
    ]),
    ("PEOPLE COSTS", [
        ("Super", "Superannuation guarantee", "% of salary", 0.12, 0.12, 0.12, "SG rate 12% from 1 Jul 2025.", "pct"),
        ("PayrollTax", "Payroll tax", "% of salary", 0.0545, 0.0545, 0.0545, "NSW 5.45%; applied to all wages (conservative - ignores $1.2m threshold).", "pct"),
        ("WorkersComp", "Workers compensation", "% of salary", 0.01, 0.01, 0.01, "Office-based premium estimate.", "pct"),
        ("OtherOnCost", "Leave loading, benefits, training", "% of salary", 0.03, 0.03, 0.03, "", "pct"),
        ("ContractorPremium", "Contractor premium over salary equivalent (no on-costs)", "%", 0.20, 0.25, 0.15, "Fractional / contract roles.", "pct"),
        ("WageInflation", "Annual wage growth", "% p.a.", 0.035, 0.045, 0.030, "", "pct"),
        ("RecruitPct", "Recruitment fee for employee hires", "% of salary", 0.15, 0.18, 0.12, "Agency/search fee; founders excluded (month 1 hires).", "pct"),
        ("ITEquip", "IT equipment & onboarding per new starter", "$", 3_500, 4_000, 3_000, "Laptop, peripherals, licences set-up.", "money"),
    ]),
    ("OTHER COST & FINANCE ASSUMPTIONS", [
        ("CostEsc", "Non-salary cost escalation", "% p.a.", 0.03, 0.04, 0.025, "", "pct"),
        ("SetupFactor", "Set-up cost factor (vs quoted/estimated)", "x", 1.00, 1.15, 0.95, "Downside assumes 15% overrun on all set-up.", "dec"),
        ("OpexFactor", "Recurring opex factor", "x", 1.00, 1.10, 0.97, "", "dec"),
        ("ContPct", "Opex contingency", "% of non-staff opex", 0.05, 0.08, 0.05, "", "pct"),
        ("SetupContPct", "Set-up contingency", "% of set-up", 0.10, 0.10, 0.10, "", "pct"),
        ("AmortMonths", "Amortisation period - capitalised platform", "months", 36, 36, 36, "Straight line from month after spend.", "int"),
        ("TaxRate", "Company tax rate (base-rate entity)", "%", 0.25, 0.25, 0.25, "Tax-consolidated group; losses carried forward.", "pct"),
        ("InterestRate", "Interest earned on cash", "% p.a.", 0.035, 0.030, 0.040, "", "pct"),
        ("RDOn", "Claim R&D tax incentive? (1 = yes)", "flag", 1, 0, 1, "Requires AusIndustry registration each year.", "int"),
        ("RDEligible", "Eligible share of BHive technology staff cost", "%", 0.30, 0.30, 0.35, "Core/supporting R&D only; needs R&D adviser sign-off.", "pct"),
        ("RDRate", "Refundable R&D tax offset rate", "%", 0.435, 0.435, 0.435, "Turnover < $20m: 18.5% premium over 25% rate.", "pct"),
        ("RDLag", "Months until R&D refund received", "months", 12, 12, 12, "", "int"),
        ("MinCash", "Minimum cash buffer (regulatory capital + safety)", "$", 750_000, 750_000, 750_000, "AFSL financial requirements + 3 months fixed costs.", "money"),
    ]),
    ("CORPORATE FUNDING (equity into B-Catcher Group)", [
        ("FounderShares", "Existing shares (B-Catcher + BHive founders, merged)", "shares", 10_000_000, 10_000_000, 10_000_000, "Merger ratio between legacy B-Catcher and BHive holders TO BE CONFIRMED.", "num"),
        ("SeedAmt", "Seed round amount", "$", 5_500_000, 5_500_000, 5_500_000, "Funds build, licensing and launch to month 9 with the lean team.", "money"),
        ("SeedMonth", "Seed round close", "model month", 1, 1, 1, "", "int"),
        ("SeedPre", "Seed pre-money valuation", "$", 18_000_000, 18_000_000, 18_000_000, "BHive prior Series A pitch: $10.5m pre.", "money"),
        ("SeriesAAmt", "Series A amount", "$", 7_500_000, 7_500_000, 7_500_000, "Milestones: licences in place, platform live, first subscriptions, signed pipeline.", "money"),
        ("SeriesAMonth", "Series A close", "model month", 10, 10, 10, "Three months after launch.", "int"),
        ("SeriesAPre", "Series A pre-money valuation", "$", 40_000_000, 40_000_000, 40_000_000, "", "money"),
        ("SeriesBAmt", "Series B / growth round amount (0 = none)", "$", 0, 0, 0, "Not needed in Base with the lean team; use for Downside or faster expansion.", "money"),
        ("SeriesBMonth", "Series B close", "model month", 28, 28, 28, "", "int"),
        ("SeriesBPre", "Series B pre-money valuation", "$", 65_000_000, 65_000_000, 65_000_000, "", "money"),
        ("RaiseCostPct", "Capital-raising costs (legal, placement)", "% of raise", 0.04, 0.05, 0.04, "", "pct"),
        ("ExitEBITDAMult", "Exit multiple - EBITDA (year 5)", "x", 12, 10, 14, "Listed wealth/platform peers trade ~10-20x.", "mult"),
        ("ExitRevMult", "Exit multiple - revenue (year 5)", "x", 5, 3.5, 6, "", "mult"),
    ]),
]

# Stress levers: single value, applied on top of the scenario.
STRESS_INPUTS = [
    ("PlacementMult", "Capital-placed multiplier", "x", 1.0, "1.0 = scenario; 0.5 = 50% placement shortfall.", "dec"),
    ("LaunchDelay", "Licensing delay to launch & secondary market", "months", 0, "Revenue shifts right; costs do not.", "int"),
    ("BrokerDelay", "Broker recruitment ramp delay", "months", 0, "", "int"),
    ("DevOverrun", "Technology build overrun", "%", 0.0, "Applied to set-up lines flagged as tech build.", "pct"),
    ("HireDeferral", "Defer post-launch hires (roles starting month 13+)", "months", 0, "Management mitigation: hold hiring until volume milestones are met.", "int"),
]

CAP_NAMES = ["CapY1", "CapY2", "CapY3", "CapY4", "CapY5"]
BROKER_NAMES = ["BrokerY1", "BrokerY2", "BrokerY3", "BrokerY4", "BrokerY5"]

# ---------------------------------------------------------------- HEADCOUNT
# (entity, dept, role, type, start, fte1, step_month, fte2, end_month, salary, note)
E, C = "Employee", "Contractor"
ROLES = [
    # ONE shared executive team runs all three companies. Roles are combined until volume justifies specialists.
    # B-Catcher (parent) - executive, finance, risk, investments
    ("B-Catcher", "Executive", "Group CEO & Chief Investment Officer (acting CTO to month 24)", E, 1, 1, None, None, None, 300_000, "One executive for strategy, raise, investment committee and build-partner oversight. Replaces separate CEO, CIO and CTO roles."),
    ("B-Catcher", "Finance", "Fractional CFO (contract)", C, 1, 0.4, 25, 0.6, 48, 230_000, "Fractional for 4 years; replaces full-time CFO hire until year 5."),
    ("B-Catcher", "Finance", "Chief Financial Officer", E, 49, 1, None, None, None, 220_000, "Year 5 only, when AUM > $300m."),
    ("B-Catcher", "Finance", "Finance & Operations Manager", E, 7, 1, None, None, None, 145_000, "Combines financial controller, fund accounting, commission ledger and office management."),
    ("B-Catcher", "Finance", "Fund & Commission Accountant", E, 37, 1, None, None, None, 110_000, "Second finance hire once 30+ trusts and 100+ brokers."),
    ("B-Catcher", "Risk & Compliance", "Head of Risk & Compliance (Group; Honey Pot responsible manager)", E, 1, 0.5, 5, 1, None, 205_000, "One compliance head for all entities; RG 105 responsible manager for Honey Pot. Replaces separate distribution compliance manager."),
    ("B-Catcher", "Risk & Compliance", "Compliance & Risk Analyst", E, 49, 1, None, None, None, 110_000, "Year 5 only; external compliance consultant covers monitoring before."),
    ("B-Catcher", "Investments", "Investment Director (property, credit & PE)", E, 3, 1, None, None, None, 215_000, "One generalist director across all products; replaces separate property and credit/PE directors."),
    ("B-Catcher", "Investments", "Investment Analyst", E, 10, 1, None, None, None, 115_000, "AI-assisted due diligence and asset reporting."),
    ("B-Catcher", "Investments", "Asset Manager", E, 37, 1, None, None, None, 150_000, "When managed AUM passes ~$150m."),
    ("B-Catcher", "Investments", "Credit & PE Investment Director", E, 49, 1, None, None, None, 205_000, "Year 5 specialist hire."),
    # Honey Pot - distribution & broker network
    ("Honey Pot", "Executive", "Head of Distribution & Marketing (Honey Pot MD)", E, 3, 1, None, None, None, 235_000, "Combines MD, broker network manager and head of marketing."),
    ("Honey Pot", "Distribution", "BDM #1 (NSW/ACT)", E, 7, 1, None, None, None, 150_000, "Also runs broker accreditation with the LMS."),
    ("Honey Pot", "Distribution", "BDM #2 (VIC/QLD)", E, 19, 1, None, None, None, 150_000, ""),
    ("Honey Pot", "Distribution", "BDM #3 (WA/SA & family office)", E, 37, 1, None, None, None, 155_000, ""),
    ("Honey Pot", "Distribution", "BDM #4", E, 49, 1, None, None, None, 150_000, "Year 5 only."),
    ("Honey Pot", "Operations", "Client Services & Broker Operations Officer", E, 7, 1, None, None, None, 95_000, "Combines investor services, commission operations and tier-2 platform support; AI agent handles tier 1."),
    ("Honey Pot", "Operations", "Client Services Officer #2", E, 31, 1, None, None, None, 90_000, ""),
    ("Honey Pot", "Marketing", "Digital Marketing & Content (contract)", C, 7, 0.5, 37, 1, None, 110_000, "Part-time contractor; AI content tools."),
    # BHive - platform (build partner does most engineering)
    ("BHive", "Product", "Head of Product & Delivery", E, 2, 1, None, None, None, 195_000, "Runs the build partner, backlog, releases and the market-operator integration. Replaces separate CTO, head of product and head of secondary market until month 25."),
    ("BHive", "Technology", "Senior Full-stack Engineer", E, 4, 1, None, None, None, 185_000, "In-house owner of code, integrations and DevOps alongside build partner."),
    ("BHive", "Technology", "Chief Technology Officer", E, 25, 1, None, None, None, 240_000, "Hired once platform is live and in-house team grows."),
    ("BHive", "Technology", "Full-stack & Integrations Engineer", E, 25, 1, None, None, None, 175_000, "Takes over work from build partner."),
    ("BHive", "Technology", "Data & AI Engineer", E, 49, 1, None, None, None, 175_000, "Year 5 only."),
    ("BHive", "Technology", "Information Security Lead (fractional vCISO)", C, 4, 0.2, 19, 0.3, None, 260_000, "Fractional; covers security, SOC 2 and market-surveillance controls."),
    ("BHive", "Product", "UX/UI Designer (contract)", C, 2, 0.5, 13, 0.2, None, 140_000, ""),
    ("BHive", "Operations", "Platform & Market Operations Analyst", E, 49, 1, None, None, None, 120_000, "Year 5 only; licensed operator runs surveillance before."),
]

# ---------------------------------------------------------------- SET-UP COSTS
# (entity, category, item, amount, start, months, treatment, tech_flag, basis)
# start may be an int or a formula string (e.g. "=LaunchEff")
SETUP = [
    # B-Catcher - corporate & structuring
    ("B-Catcher", "Corporate & Legal Structure", "Incorporate Honey Pot, BHive & template SPV (ASIC fees, constitutions)", 8_000, 1, 1, "Opex", "N", "ASIC registration ~$611 per company + legal drafting."),
    ("B-Catcher", "Corporate & Legal Structure", "Group restructure: legacy B-Catcher & BHive entity/IP due diligence, IP assignment", 55_000, 1, 2, "Opex", "N", "Business Plan v2 s2: ASIC extracts, liabilities, IP assignment."),
    ("B-Catcher", "Corporate & Legal Structure", "Shareholders' agreement, constitution, ESOP plan", 40_000, 2, 2, "Opex", "N", ""),
    ("B-Catcher", "Corporate & Legal Structure", "Group tax structuring (trusts, GST, transfer pricing, R&D)", 35_000, 2, 2, "Opex", "N", ""),
    ("B-Catcher", "Licensing & Regulatory", "Financial services perimeter opinion (AFSL, MIS, market, credit, AML)", 90_000, 1, 3, "Opex", "N", "Corporate & Regulatory Implementation doc: counsel decision register."),
    ("B-Catcher", "Product & Investment Structuring", "Property trust deed & IM template (wholesale unit trust)", 45_000, 3, 3, "Opex", "N", ""),
    ("B-Catcher", "Product & Investment Structuring", "Private credit fund IM, loan & security templates", 45_000, 4, 3, "Opex", "N", ""),
    ("B-Catcher", "Product & Investment Structuring", "PE / pre-IPO unit trust with buy-back option deed (UTEGBO)", 40_000, 4, 3, "Opex", "N", "Builds on 2025 UTEGBO term sheet."),
    ("B-Catcher", "Product & Investment Structuring", "Investment committee charter, valuation & conflicts policies", 20_000, 3, 2, "Opex", "N", ""),
    ("B-Catcher", "Product & Investment Structuring", "Due diligence on first 3 offers (valuation, QS, building, legal)", 105_000, 4, 3, "Opex", "N", "~$35k per offer; recoverable from trust if offer proceeds."),
    ("B-Catcher", "Product & Investment Structuring", "Trustee / RE / custodian appointment & onboarding", 30_000, 5, 2, "Opex", "N", "External trustee for wholesale trusts."),
    ("B-Catcher", "Brand & Marketing Launch", "Brand identity & design system (B-Catcher, Honey Pot, BHive)", 65_000, 2, 2, "Opex", "N", ""),
    ("B-Catcher", "Brand & Marketing Launch", "Group corporate website", 30_000, 4, 2, "Opex", "N", ""),
    ("B-Catcher", "Office & Equipment", "Meeting-room AV & minor fit-out (serviced office)", 15_000, 3, 1, "Capex", "N", "Serviced office; no lease fit-out."),
    ("B-Catcher", "Corporate Raise Costs", "Data room, information memorandum & corporate deck", 25_000, 1, 2, "Opex", "N", ""),
    ("B-Catcher", "Corporate Raise Costs", "Independent financial model review / financial due diligence", 30_000, 2, 1, "Opex", "N", "Deck Content Brief: independently checked model."),
    ("B-Catcher", "Corporate Raise Costs", "Series A data room refresh & vendor due diligence", 40_000, 11, 2, "Opex", "N", ""),
    # Honey Pot - licensing & distribution infrastructure
    ("Honey Pot", "Licensing & Regulatory", "Corporate authorised representative appointment under principal licensee", 25_000, 4, 1, "Opex", "N", "Interim route while own AFSL is obtained."),
    ("Honey Pot", "Licensing & Regulatory", "Compliance framework: AFSL compliance plan, breach, IDR/complaints, conflicts, DDO/TMD", 65_000, 3, 3, "Opex", "N", ""),
    ("Honey Pot", "Licensing & Regulatory", "AML/CTF program, AUSTRAC enrolment, first independent review", 35_000, 4, 3, "Opex", "N", ""),
    ("Honey Pot", "Licensing & Regulatory", "Wholesale/sophisticated investor verification procedures (s761G, s761GA, s708)", 12_000, 4, 1, "Opex", "N", ""),
    ("Honey Pot", "Licensing & Regulatory", "Own AFSL application (legal, RG 105 responsible managers, ASIC fees)", 150_000, 13, 8, "Opex", "N", "Transition from CAR to own licence in year 2."),
    ("Honey Pot", "Distribution Infrastructure", "Broker agreement suite (AR, licensee-to-licensee, referral, commission policy)", 45_000, 3, 2, "Opex", "N", ""),
    ("Honey Pot", "Distribution Infrastructure", "Broker CRM & commission system implementation", 75_000, 3, 4, "Capex", "Y", "e.g. Salesforce Financial Services Cloud / HubSpot configuration."),
    ("Honey Pot", "Distribution Infrastructure", "Learning management system & broker accreditation content", 35_000, 5, 2, "Opex", "N", ""),
    ("Honey Pot", "Brand & Marketing Launch", "Launch campaign & national broker roadshow", 90_000, 6, 3, "Opex", "N", ""),
    ("Honey Pot", "Brand & Marketing Launch", "Marketing compliance review of launch materials", 15_000, 6, 1, "Opex", "N", ""),
    # BHive - platform build
    ("BHive", "Technology Build (Capitalised)", "Discovery, architecture & UX design", 95_000, 2, 2, "Capex", "Y", "Product & Technical Requirements doc."),
    ("BHive", "Technology Build (Capitalised)", "MVP build partner: investor, broker & admin portals, wallet", 480_000, 3, 6, "Capex", "Y", "Prior model $350k platform build; uplifted for 3 portals."),
    ("BHive", "Technology Build (Capitalised)", "Identity/KYC/AML provider integration", 25_000, 4, 3, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "E-signature, document vault & CRM integration", 30_000, 5, 2, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "Registry / fund administrator integration", 60_000, 5, 3, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "Banking & payment reconciliation integration", 40_000, 5, 3, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "Mobile apps (iOS & Android)", 180_000, 13, 6, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "Secondary market engine: listings, order book, transfer & settlement workflow", 350_000, 10, 9, "Capex", "Y", "Built and tested before activation."),
    ("BHive", "Technology Build (Capitalised)", "Market-operator partner integration", 60_000, 15, 4, "Capex", "Y", ""),
    ("BHive", "Technology Build (Capitalised)", "Tokenised register pilot: smart contracts & audit", 90_000, 25, 6, "Capex", "Y", "Only after token legal register opinion."),
    ("BHive", "Licensing & Regulatory", "Market operator route: legal, market rules, operating agreement", 120_000, 13, 6, "Opex", "N", "Partner route; own market licence not modelled."),
    ("BHive", "Licensing & Regulatory", "Digital token legal register opinion", 40_000, 22, 3, "Opex", "N", ""),
    ("BHive", "Security & Certification", "Pre-launch penetration test & security audit", 45_000, 6, 1, "Opex", "N", ""),
    ("BHive", "Security & Certification", "Privacy impact assessment & data processing agreements", 20_000, 5, 1, "Opex", "N", ""),
    ("BHive", "Security & Certification", "SOC 2 Type I readiness & audit", 70_000, 10, 6, "Opex", "N", ""),
    ("BHive", "Security & Certification", "Secondary market pen test & surveillance validation", 50_000, 18, 1, "Opex", "N", ""),
    ("BHive", "Technology Build (Capitalised)", "Cloud landing zone, environments, domains & app store accounts", 15_000, 2, 1, "Capex", "Y", "AWS Sydney region."),
]

# ---------------------------------------------------------------- RECURRING OPEX
CATEGORIES = [
    "Premises & Office",
    "Board & Governance",
    "Insurance",
    "Professional Fees (Legal, Audit, Tax)",
    "Licensing, Regulatory & Compliance",
    "Technology, Hosting & Software",
    "Marketing & Investor Acquisition",
    "Broker Network & Distribution",
    "Investor Servicing & Onboarding",
    "Investment Origination & Due Diligence",
    "Travel & General",
]
DRIVERS = [
    "Fixed $/month",
    "$/FTE/month",
    "$/active broker/month",
    "$/new broker",
    "$/new investor",
    "$/investor/month",
    "% of capital placed",
    "% of entity revenue",
]
# (entity, category, item, amount, driver, start, end, basis)
OPEX = [
    # B-Catcher
    ("B-Catcher", "Premises & Office", "Serviced office (Sydney CBD, ~12 desks)", 9_000, "Fixed $/month", 1, 36, "Serviced/co-working space; no lease while team < 25."),
    ("B-Catcher", "Premises & Office", "Serviced office (expanded, ~22 desks)", 14_000, "Fixed $/month", 37, None, "Still serviced space at year-5 team size."),
    ("B-Catcher", "Premises & Office", "Utilities, cleaning, consumables", 120, "$/FTE/month", 1, None, "B-Catcher per-FTE lines use total group FTE (shared services)."),
    ("B-Catcher", "Board & Governance", "Independent Chair fee", 7_500, "Fixed $/month", 1, None, "$90k p.a."),
    ("B-Catcher", "Board & Governance", "Two independent non-executive directors", 10_000, "Fixed $/month", 13, None, "$60k p.a. each."),
    ("B-Catcher", "Board & Governance", "Independent investment committee member", 2_500, "Fixed $/month", 7, None, "$30k p.a."),
    ("B-Catcher", "Insurance", "Directors & officers insurance", 3_000, "Fixed $/month", 1, None, ""),
    ("B-Catcher", "Insurance", "General liability, office & contents", 600, "Fixed $/month", 1, None, ""),
    ("B-Catcher", "Professional Fees (Legal, Audit, Tax)", "External audit & tax (group, 3 entities)", 7_000, "Fixed $/month", 1, None, "~$85k p.a."),
    ("B-Catcher", "Professional Fees (Legal, Audit, Tax)", "Ongoing corporate & investment legal", 8_000, "Fixed $/month", 7, None, ""),
    ("B-Catcher", "Professional Fees (Legal, Audit, Tax)", "Bookkeeping & payroll bureau (pre controller)", 3_500, "Fixed $/month", 1, 9, ""),
    ("B-Catcher", "Licensing, Regulatory & Compliance", "ASIC annual review fees, registers, company secretarial", 600, "Fixed $/month", 1, None, ""),
    ("B-Catcher", "Technology, Hosting & Software", "Business software (M365, Xero, HRIS, payroll)", 180, "$/FTE/month", 1, None, ""),
    ("B-Catcher", "Technology, Hosting & Software", "AI productivity tools (all staff)", 120, "$/FTE/month", 1, None, "AI-native operating model."),
    ("B-Catcher", "Investment Origination & Due Diligence", "Pipeline due diligence on deals not proceeding", 8_000, "Fixed $/month", 7, None, ""),
    ("B-Catcher", "Investment Origination & Due Diligence", "Property & market data (CoreLogic, research)", 3_500, "Fixed $/month", 4, None, ""),
    ("B-Catcher", "Travel & General", "Travel, accommodation & entertainment", 5_000, "Fixed $/month", 1, 12, ""),
    ("B-Catcher", "Travel & General", "Travel, accommodation & entertainment", 6_000, "Fixed $/month", 13, None, ""),
    ("B-Catcher", "Travel & General", "Bank fees, subscriptions, memberships, sundry", 2_000, "Fixed $/month", 1, None, ""),
    ("B-Catcher", "Travel & General", "Staff training & conferences", 100, "$/FTE/month", 1, None, ""),
    ("B-Catcher", "Professional Fees (Legal, Audit, Tax)", "Outsourced HR, payroll & recruitment support", 1_500, "Fixed $/month", 10, None, "Replaces in-house people & culture and EA roles."),
    # Honey Pot
    ("Honey Pot", "Licensing, Regulatory & Compliance", "Principal licensee fee (CAR arrangement)", 12_000, "Fixed $/month", 4, 24, "Ends when own AFSL granted."),
    ("Honey Pot", "Licensing, Regulatory & Compliance", "ASIC industry funding levy (own AFSL)", 2_500, "Fixed $/month", 25, None, ""),
    ("Honey Pot", "Licensing, Regulatory & Compliance", "External compliance consultant & monitoring", 6_000, "Fixed $/month", 5, None, ""),
    ("Honey Pot", "Licensing, Regulatory & Compliance", "AFCA membership & complaint levies", 700, "Fixed $/month", 7, None, ""),
    ("Honey Pot", "Licensing, Regulatory & Compliance", "Sanctions / PEP screening & AML monitoring tools", 1_500, "Fixed $/month", 7, None, ""),
    ("Honey Pot", "Licensing, Regulatory & Compliance", "Annual independent AML/CTF review", 2_000, "Fixed $/month", 13, None, ""),
    ("Honey Pot", "Insurance", "Professional indemnity insurance (AFSL)", 4_500, "Fixed $/month", 7, None, "RG 126 adequate PI cover."),
    ("Honey Pot", "Insurance", "PI uplift with distribution volume", 0.0002, "% of capital placed", 7, None, ""),
    ("Honey Pot", "Marketing & Investor Acquisition", "Brand & content (always-on)", 10_000, "Fixed $/month", 6, None, ""),
    ("Honey Pot", "Marketing & Investor Acquisition", "Performance marketing - wholesale investors", 0.0015, "% of capital placed", 7, None, "0.15% of capital = CAC on direct channel."),
    ("Honey Pot", "Marketing & Investor Acquisition", "Investor events & webinars", 6_000, "Fixed $/month", 7, None, ""),
    ("Honey Pot", "Broker Network & Distribution", "Broker onboarding: due diligence, accreditation, training", 1_500, "$/new broker", 1, None, "Manual: verify AFSL/ACL, training, portal access."),
    ("Honey Pot", "Broker Network & Distribution", "Broker CRM, LMS & portal seats", 60, "$/active broker/month", 1, None, ""),
    ("Honey Pot", "Broker Network & Distribution", "Broker roadshows, PD events & annual conference", 8_000, "Fixed $/month", 7, None, ""),
    ("Honey Pot", "Broker Network & Distribution", "BDM incentive pool", 0.0010, "% of capital placed", 7, None, "Paid on cleared subscriptions; on-costs ignored."),
    ("Honey Pot", "Broker Network & Distribution", "Broker file reviews & supervision sampling (external)", 2_500, "Fixed $/month", 10, None, ""),
    ("Honey Pot", "Technology, Hosting & Software", "CRM & marketing automation licences", 200, "$/FTE/month", 1, None, ""),
    ("Honey Pot", "Travel & General", "Interstate travel (BDMs & roadshows)", 6_000, "Fixed $/month", 5, None, ""),
    ("Honey Pot", "Professional Fees (Legal, Audit, Tax)", "Distribution legal (offer & marketing reviews)", 5_000, "Fixed $/month", 7, None, ""),
    # BHive
    ("BHive", "Technology, Hosting & Software", "Cloud hosting - base environments (AWS Sydney)", 7_000, "Fixed $/month", 3, None, "Prod, staging, dev, DR."),
    ("BHive", "Technology, Hosting & Software", "Cloud hosting - variable per investor", 0.60, "$/investor/month", 1, None, ""),
    ("BHive", "Technology, Hosting & Software", "Engineering tooling (GitHub, CI, monitoring, logging)", 250, "$/FTE/month", 1, None, ""),
    ("BHive", "Technology, Hosting & Software", "Build partner retainer: maintenance, DevOps & QA", 16_000, "Fixed $/month", 9, None, "Covers DevOps, QA and overflow development instead of in-house hires."),
    ("BHive", "Technology, Hosting & Software", "Registry / fund-admin platform fees", 4_000, "Fixed $/month", 7, None, ""),
    ("BHive", "Technology, Hosting & Software", "AI support agent & AI ops tooling", 3_000, "Fixed $/month", 7, None, "Tier-1 support automation."),
    ("BHive", "Technology, Hosting & Software", "Security operations (MDR/SIEM service)", 6_000, "Fixed $/month", 7, None, ""),
    ("BHive", "Technology, Hosting & Software", "Annual penetration testing", 4_000, "Fixed $/month", 13, None, "~$48k p.a."),
    ("BHive", "Technology, Hosting & Software", "SOC 2 Type II annual audit", 5_000, "Fixed $/month", 19, None, ""),
    ("BHive", "Investor Servicing & Onboarding", "Identity verification & AML onboarding checks", 35, "$/new investor", 1, None, "Incl. wholesale certificate validation."),
    ("BHive", "Investor Servicing & Onboarding", "E-signature & document processing", 12, "$/new investor", 1, None, ""),
    ("BHive", "Investor Servicing & Onboarding", "Ongoing AML monitoring & notifications", 1.00, "$/investor/month", 1, None, ""),
    ("BHive", "Investor Servicing & Onboarding", "Helpdesk & support tooling", 1_200, "Fixed $/month", 7, None, ""),
    ("BHive", "Licensing, Regulatory & Compliance", "Market-operator partner fixed platform fee", 15_000, "Fixed $/month", "=SecEff", None, "Plus revenue share in cost of sales."),
    ("BHive", "Licensing, Regulatory & Compliance", "Market surveillance tooling & transaction reporting", 3_000, "Fixed $/month", "=SecEff", None, ""),
    ("BHive", "Insurance", "Cyber liability insurance", 3_000, "Fixed $/month", 7, None, ""),
    ("BHive", "Insurance", "Technology PI / errors & omissions", 1_500, "Fixed $/month", 7, None, ""),
    ("BHive", "Marketing & Investor Acquisition", "Product marketing, app store & UX research", 3_000, "Fixed $/month", 7, None, ""),
    ("BHive", "Professional Fees (Legal, Audit, Tax)", "Technology & privacy legal", 3_000, "Fixed $/month", 7, None, ""),
]
