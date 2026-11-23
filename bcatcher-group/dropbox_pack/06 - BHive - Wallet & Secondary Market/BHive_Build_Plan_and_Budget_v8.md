# BHive - Build Plan & Budget - Version 8

**B-Catcher Group** | 06 - BHive - Wallet & Secondary Market | Version 8.0 | 23 November 2026  
Confidential planning document. All figures are AUD Base-case planning assumptions from *BCatcher_Group_Integrated_Financial_Model.xlsx*. Not an offer or financial advice.

Requirements remain as in *Product & Technical Requirements v1*. This adds the build sequence, vendor stack, budget, team and platform economics.

## Build sequence and cost
| Item | Cost | Months | Treatment |
|---|---|---|---|
| Discovery, architecture & UX design | $99,750 | 2-3 | Capex |
| MVP build partner: investor, broker & admin portals, wallet | $682,500 | 3-8 | Capex |
| Identity/KYC/AML provider integration | $26,250 | 4-6 | Capex |
| E-signature, document vault & CRM integration | $31,500 | 5-6 | Capex |
| Registry / fund administrator integration | $63,000 | 5-7 | Capex |
| Banking & payment reconciliation integration | $42,000 | 5-7 | Capex |
| Mobile apps (iOS & Android) | $189,000 | 13-18 | Capex |
| Secondary market engine: listings, order book, transfer & settlement workflow | $472,500 | 22-30 | Capex |
| Market-operator partner integration | $63,000 | 27-30 | Capex |
| Tokenised register pilot: smart contracts & audit | $94,500 | 34-39 | Capex |
| Market operator route: legal, market rules, operating agreement | $126,000 | 22-27 | Opex |
| Digital token legal register opinion | $42,000 | 30-32 | Opex |
| Pre-launch penetration test & security audit | $47,250 | 6 | Opex |
| Privacy impact assessment & data processing agreements | $21,000 | 5 | Opex |
| SOC 2 Type I readiness & audit | $73,500 | 10-15 | Opex |
| Secondary market pen test & surveillance validation | $52,500 | 30 | Opex |
| Cloud landing zone, environments, domains & app store accounts | $15,750 | 2 | Capex |

**Total BHive set-up: $2,142,000** (capex amortised over 36 months). Plus Honey Pot CRM/commission system $75,000.

## Release plan
- **MVP (months 3-8):** onboarding, KYC/AML, wholesale verification, offers & data room, e-sign subscriptions, registry reconciliation, wallet/portfolio, broker portal, admin portal
- **Launch (month 7):** live with first offers; platform support + AI tier-1 agent
- **Months 10-18:** secondary engine (listings, order book, transfer & settlement workflow), operator integration, mobile apps, SOC 2 Type I
- **Months 22-30:** secondary engine and operator integration built
- **Month 31:** secondary market activated with licensed operator (40% fee share + $15k/month), once AUM is ~$150m
- **Months 34-39:** tokenised register pilot only after legal register opinion

## Platform P&L (incl. intercompany licence income)
| $ | 2027 | 2028 | 2029 | 2030 | 2031 |
|---|---|---|---|---|---|
| Technology fees (0.25% p.a.) | $2,340 | $44,889 | $174,055 | $418,410 | $805,004 |
| Offer onboarding fees | $17,500 | $120,000 | $270,000 | $480,000 | $720,000 |
| Exchange fees (0.25%) | - | - | $8,249 | $50,030 | $120,751 |
| Buy-back resale gains | - | $630 | $20,871 | $141,170 | $297,618 |
| Intercompany licence income | $270,000 | $540,000 | $540,000 | $540,000 | $540,000 |
| Operator share (cost of sales) | - | - | ($3,299) | ($20,012) | ($48,300) |
| Staff costs | ($249,200) | ($182,574) | ($164,540) | ($170,299) | ($176,260) |
| Total operating expenses | ($1,110,375) | ($1,711,500) | ($2,122,817) | ($2,190,419) | ($2,320,138) |
| EBITDA | ($820,535) | ($1,005,980) | ($1,112,942) | ($580,820) | $114,935 |
| R&D tax incentive (accrued) | $32,521 | $23,826 | $21,472 | $22,224 | $23,002 |
| Capitalised build | $960,750 | $346,500 | $425,250 | $47,250 | - |
| Value traded on secondary market | - | - | $3,299,482 | $20,012,137 | $48,300,264 |
| Investors on platform | 39 | 306 | 906 | 1,972 | 3,572 |

## Running costs
| Item | Rate / driver | 2027 | 2029 | 2031 |
|---|---|---|---|---|
| Cloud hosting - base environments (AWS Sydney) | $7,000 Fixed $/month | $74,200 | $94,463 | $100,215 |
| Cloud hosting - variable per investor | $0.60 $/investor/month | $66 | $4,701 | $23,813 |
| Engineering tooling (GitHub, CI, monitoring, logging) | $250 $/FTE/month | $3,684 | $2,024 | $2,147 |
| Build partner retainer: engineering, DevOps & QA (no in-house engineers) | $40,000 Fixed $/month | $169,600 | - | - |
| Build partner retainer (expanded): secondary market, mobile & scale | $50,000 Fixed $/month | - | $674,732 | $715,824 |
| Registry / fund-admin platform fees | $4,000 Fixed $/month | $25,440 | $53,979 | $57,266 |
| AI support agent & AI ops tooling | $5,000 Fixed $/month | $31,800 | $67,473 | $71,582 |
| AI support agent - volume usage | $0.40 $/investor/month | $44 | $3,134 | $15,875 |
| Security operations (MDR/SIEM service) | $6,000 Fixed $/month | $38,160 | $80,968 | $85,899 |
| Annual penetration testing | $4,000 Fixed $/month | - | $53,979 | $57,266 |
| SOC 2 Type II annual audit | $5,000 Fixed $/month | - | $67,473 | $71,582 |
| Identity verification & AML onboarding checks | $35 $/new investor | $1,443 | $23,616 | $66,810 |
| E-signature & document processing | $12 $/new investor | $495 | $8,097 | $22,906 |
| Ongoing AML monitoring & notifications | $1.00 $/investor/month | $110 | $7,834 | $39,688 |
| Helpdesk & support tooling | $1,200 Fixed $/month | $7,632 | $16,194 | $17,180 |
| Market-operator partner fixed platform fee | $15,000 Fixed $/month | - | $101,210 | $214,747 |
| Market surveillance tooling & transaction reporting | $3,000 Fixed $/month | - | $20,242 | $42,949 |
| Cyber liability insurance | $3,000 Fixed $/month | $19,080 | $40,484 | $42,949 |
| Technology PI / errors & omissions | $1,500 Fixed $/month | $9,540 | $20,242 | $21,475 |
| Product marketing, app store & UX research | $3,000 Fixed $/month | $19,080 | $40,484 | $42,949 |
| Technology & privacy legal | $3,000 Fixed $/month | $19,080 | $40,484 | $42,949 |

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
| Units bought back (valuation) | - | $83,447 | $1,134,618 | $3,784,705 | $7,464,329 |
| Resale gains | - | $630 | $20,871 | $141,170 | $297,618 |
| Units held for resale (year end) | - | $60,199 | $639,999 | $1,455,921 | $2,713,363 |
| Cash reserve required (year end) | - | $134,552 | $1,296,560 | $2,801,562 | $5,033,277 |

Requires a legal and capital-adequacy opinion before launch ($40k in set-up).

## Not in scope
BHive Coin, BHive Card, token lending, retail access (all from the July 2026 BHive Capital pack) - excluded until separate legal and business cases.
