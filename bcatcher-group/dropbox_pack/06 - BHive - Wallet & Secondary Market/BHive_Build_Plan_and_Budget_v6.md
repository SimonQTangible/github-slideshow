# BHive - Build Plan & Budget - Version 6

**B-Catcher Group** | 06 - BHive - Wallet & Secondary Market | Version 6.0 | 23 November 2026  
Confidential planning document. All figures are AUD Base-case planning assumptions from *BCatcher_Group_Integrated_Financial_Model.xlsx*. Not an offer or financial advice.

Requirements remain as in *Product & Technical Requirements v1*. This adds the build sequence, vendor stack, budget, team and platform economics.

## Build sequence and cost
| Item | Cost | Months | Treatment |
|---|---|---|---|
| Discovery, architecture & UX design | $95,000 | 2-3 | Capex |
| MVP build partner: investor, broker & admin portals, wallet | $480,000 | 3-8 | Capex |
| Identity/KYC/AML provider integration | $25,000 | 4-6 | Capex |
| E-signature, document vault & CRM integration | $30,000 | 5-6 | Capex |
| Registry / fund administrator integration | $60,000 | 5-7 | Capex |
| Banking & payment reconciliation integration | $40,000 | 5-7 | Capex |
| Mobile apps (iOS & Android) | $180,000 | 13-18 | Capex |
| Secondary market engine: listings, order book, transfer & settlement workflow | $350,000 | 10-18 | Capex |
| Market-operator partner integration | $60,000 | 15-18 | Capex |
| Tokenised register pilot: smart contracts & audit | $90,000 | 25-30 | Capex |
| Market operator route: legal, market rules, operating agreement | $120,000 | 13-18 | Opex |
| Digital token legal register opinion | $40,000 | 22-24 | Opex |
| Pre-launch penetration test & security audit | $45,000 | 6 | Opex |
| Privacy impact assessment & data processing agreements | $20,000 | 5 | Opex |
| SOC 2 Type I readiness & audit | $70,000 | 10-15 | Opex |
| Secondary market pen test & surveillance validation | $50,000 | 18 | Opex |
| Cloud landing zone, environments, domains & app store accounts | $15,000 | 2 | Capex |

**Total BHive set-up: $1,770,000** (capex amortised over 36 months). Plus Honey Pot CRM/commission system $75,000.

## Release plan
- **MVP (months 3-8):** onboarding, KYC/AML, wholesale verification, offers & data room, e-sign subscriptions, registry reconciliation, wallet/portfolio, broker portal, admin portal
- **Launch (month 7):** live with first offers; platform support + AI tier-1 agent
- **Months 10-18:** secondary engine (listings, order book, transfer & settlement workflow), operator integration, mobile apps, SOC 2 Type I
- **Month 19:** secondary market activated with licensed operator (40% fee share + $15k/month)
- **Months 25-30:** tokenised register pilot only after legal register opinion

## Platform P&L (incl. intercompany licence income)
| $ | 2027 | 2028 | 2029 | 2030 | 2031 |
|---|---|---|---|---|---|
| Technology fees (0.5% p.a.) | $5,851 | $112,224 | $435,138 | $1,040,493 | $1,970,182 |
| Offer onboarding fees | $29,167 | $200,000 | $450,000 | $780,000 | $1,160,000 |
| Exchange fees (0.25%) | - | $3,104 | $26,629 | $78,037 | $147,764 |
| Buy-back resale gains | - | $4,351 | $113,620 | $617,468 | $1,259,381 |
| Intercompany licence income | $180,000 | $360,000 | $360,000 | $360,000 | $360,000 |
| Operator share (cost of sales) | - | ($1,242) | ($10,651) | ($31,215) | ($59,105) |
| Staff costs | ($249,200) | ($182,574) | ($164,540) | ($170,299) | ($176,260) |
| Total operating expenses | ($1,013,530) | ($1,713,895) | ($1,764,365) | ($1,867,290) | ($1,996,704) |
| EBITDA | ($798,513) | ($1,035,458) | ($389,631) | $977,494 | $2,841,517 |
| R&D tax incentive (accrued) | $32,521 | $23,826 | $21,472 | $22,224 | $23,002 |
| Capitalised build | $861,667 | $473,333 | $90,000 | - | - |
| Value traded on secondary market | - | $1,241,612 | $10,651,416 | $31,214,803 | $59,105,457 |
| Investors on platform | 49 | 382 | 1,132 | 2,432 | 4,365 |

## Running costs
| Item | Rate / driver | 2027 | 2029 | 2031 |
|---|---|---|---|---|
| Cloud hosting - base environments (AWS Sydney) | $7,000 Fixed $/month | $70,000 | $89,116 | $94,543 |
| Cloud hosting - variable per investor | $0.60 $/investor/month | $78 | $5,543 | $27,519 |
| Engineering tooling (GitHub, CI, monitoring, logging) | $250 $/FTE/month | $3,475 | $1,910 | $2,026 |
| Build partner retainer: engineering, DevOps & QA (no in-house engineers) | $24,000 Fixed $/month | $96,000 | - | - |
| Build partner retainer (expanded): secondary market, mobile & scale | $32,000 Fixed $/month | - | $407,386 | $432,195 |
| Registry / fund-admin platform fees | $4,000 Fixed $/month | $24,000 | $50,923 | $54,024 |
| AI support agent & AI ops tooling | $5,000 Fixed $/month | $30,000 | $63,654 | $67,531 |
| AI support agent - volume usage | $0.40 $/investor/month | $52 | $3,695 | $18,346 |
| Security operations (MDR/SIEM service) | $6,000 Fixed $/month | $36,000 | $76,385 | $81,037 |
| Annual penetration testing | $4,000 Fixed $/month | - | $50,923 | $54,024 |
| SOC 2 Type II annual audit | $5,000 Fixed $/month | - | $63,654 | $67,531 |
| Identity verification & AML onboarding checks | $35 $/new investor | $1,701 | $27,849 | $76,159 |
| E-signature & document processing | $12 $/new investor | $583 | $9,548 | $26,112 |
| Ongoing AML monitoring & notifications | $1.00 $/investor/month | $130 | $9,239 | $45,864 |
| Helpdesk & support tooling | $1,200 Fixed $/month | $7,200 | $15,277 | $16,207 |
| Market-operator partner fixed platform fee | $15,000 Fixed $/month | - | $190,962 | $202,592 |
| Market surveillance tooling & transaction reporting | $3,000 Fixed $/month | - | $38,192 | $40,518 |
| Cyber liability insurance | $3,000 Fixed $/month | $18,000 | $38,192 | $40,518 |
| Technology PI / errors & omissions | $1,500 Fixed $/month | $9,000 | $19,096 | $20,259 |
| Product marketing, app store & UX research | $3,000 Fixed $/month | $18,000 | $38,192 | $40,518 |
| Technology & privacy legal | $3,000 Fixed $/month | $18,000 | $38,192 | $40,518 |

## Team
| Role | Type | Start month | FTE | Base salary |
|---|---|---|---|---|
| Fractional Product Owner (contract) | Contractor | 2 | 0.6 | $190,000 |
| UX/UI Designer (contract) | Contractor | 2 | 0.5 | $140,000 |
| Information Security Lead (fractional vCISO) | Contractor | 4 | 0.2 | $260,000 |

## Vendor stack to procure (RFP list)
Build partner; AWS Sydney; identity/KYC/AML provider; e-signature; CRM (shared with Honey Pot); fund administrator & registry; banking/payment reconciliation; document vault; MDR/SIEM security; penetration tester; SOC 2 auditor; licensed market operator; AI support agent tooling.

## Buy-back facility (kept from BHive and B-Catcher)
Property tokens: BHive exit ladder - buy-back at 25% (year 3), 15% (year 4) and 10% (year 5) below independent valuation. Trust units (PE / pre-IPO): UTEGBO windows at 15% (year 1), 10% (year 2) and 8% (year 4). The group buys exercised units, resells them through BHive / Honey Pot at valuation within ~3 months, and holds a cash reserve of 2% of eligible AUM. Model assumes 4% of eligible holdings exercise each year (8% Downside).

| $ | 2027 | 2028 | 2029 | 2030 | 2031 |
|---|---|---|---|---|---|
| Units bought back (valuation) | - | $139,079 | $1,891,031 | $6,280,847 | $12,194,843 |
| Resale gains | - | $4,351 | $113,620 | $617,468 | $1,259,381 |
| Units held for resale (year end) | - | $79,104 | $798,177 | $1,775,551 | $3,242,816 |
| Cash reserve required (year end) | - | $224,253 | $2,160,933 | $4,604,665 | $8,197,770 |

Requires a legal and capital-adequacy opinion before launch ($40k in set-up).

## Not in scope
BHive Coin, BHive Card, token lending, retail access (all from the July 2026 BHive Capital pack) - excluded until separate legal and business cases.
