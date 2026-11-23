# BHive - Build Plan & Budget - Version 10

**B-Catcher Group** | 06 - BHive - Wallet & Secondary Market | Version 10.0 | 23 November 2026  
Confidential planning document. All figures are AUD Base-case planning assumptions from *BCatcher_Group_Integrated_Financial_Model.xlsx*. Not an offer or financial advice.

Requirements remain as in *Product & Technical Requirements v1*. This adds the build sequence, vendor stack, budget, team and platform economics.

## Build sequence and cost
| Item | Cost | Months | Treatment |
|---|---|---|---|
| Discovery, architecture & UX design | $99,750 | 2-3 | Capex |
| Core platform & BHive wallet (built up front): investor & broker portals, admin console, wallet (holdings, documents, distributions, tax), offer marketplace | $682,500 | 3-8 | Capex |
| Identity/KYC/AML provider integration | $26,250 | 4-6 | Capex |
| E-signature, document vault & CRM integration | $31,500 | 5-6 | Capex |
| Registry / fund administrator integration | $63,000 | 5-7 | Capex |
| Banking & payment reconciliation integration | $42,000 | 5-7 | Capex |
| Platform release 2: buy-back requests in wallet, investor comms hub, referral & Founding Member features, analytics | $126,000 | 10-15 | Capex |
| Mobile apps (iOS & Android) | $189,000 | 13-18 | Capex |
| Platform release 3: SMSF, company & trust accounts; auto-reinvest; tax reporting pack | $126,000 | 16-21 | Capex |
| Platform release 4: private-credit income marketplace, regular-investment plans, portfolio tools | $157,500 | 28-33 | Capex |
| Platform release 5: new offerings module (managed portfolios, insurance, retail-ready module subject to licence) | $210,000 | 40-48 | Capex |
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

**Total BHive set-up: $2,761,500** (capex amortised over 36 months). Plus Honey Pot CRM/commission system $75,000.

## Release plan
- **Core platform & BHive wallet, built up front (months 3-8):** onboarding, KYC/AML, wholesale verification, offer marketplace & data room, e-sign subscriptions, registry reconciliation, investor wallet (holdings, documents, distributions, tax), broker portal, admin console
- **Launch (month 7):** live with first offers; outsourced service desk + AI tier-1 agent
- **Release 2 (months 10-15):** buy-back requests in the wallet, investor comms hub, referral & Founding Member features, analytics
- **Mobile apps (months 13-18)**; SOC 2 Type I
- **Release 3 (months 16-21):** SMSF, company & trust accounts; auto-reinvest; tax reporting pack
- **Months 22-30:** secondary engine and operator integration built
- **Release 4 (months 28-33):** private-credit income marketplace, regular-investment plans, portfolio tools
- **Month 31:** secondary market activated with licensed operator (40% fee share + $15k/month), once AUM is ~$125m
- **Months 34-39:** tokenised register pilot only after legal register opinion
- **Release 5 (months 40-48):** new offerings module (managed portfolios, insurance; retail-ready module subject to licence)

## Platform P&L (incl. intercompany licence income)
| $ | 2027 | 2028 | 2029 | 2030 | 2031 |
|---|---|---|---|---|---|
| Technology fees (0.25% p.a.) | $3,510 | $72,867 | $308,944 | $755,960 | $1,451,259 |
| Offer onboarding fees | $26,250 | $210,000 | $495,000 | $870,000 | $1,290,000 |
| Exchange fees (0.25%) | - | - | $14,751 | $90,408 | $217,689 |
| Buy-back resale gains | - | $7,631 | $121,454 | $108,788 | $230,090 |
| Intercompany licence income | $300,000 | $600,000 | $600,000 | $600,000 | $600,000 |
| Operator share (cost of sales) | - | - | ($5,901) | ($36,163) | ($87,076) |
| Staff costs | ($249,200) | ($182,574) | ($164,540) | ($170,299) | ($176,260) |
| Total operating expenses | ($1,111,508) | ($1,724,874) | ($2,163,278) | ($2,273,202) | ($2,461,525) |
| EBITDA | ($781,748) | ($834,376) | ($629,029) | $115,791 | $1,240,437 |
| R&D tax incentive (accrued) | $32,521 | $23,826 | $21,472 | $22,224 | $23,002 |
| Capitalised build | $1,023,750 | $535,500 | $582,750 | $257,250 | - |
| Value traded on secondary market | - | - | $5,900,510 | $36,163,390 | $87,075,514 |
| Investors on platform | 58 | 525 | 1,625 | 3,558 | 6,425 |

## Running costs
| Item | Rate / driver | 2027 | 2029 | 2031 |
|---|---|---|---|---|
| Cloud hosting - base environments (AWS Sydney) | $7,000 Fixed $/month | $74,200 | $94,463 | $100,215 |
| Cloud hosting - variable per investor | $0.60 $/investor/month | $99 | $8,333 | $42,902 |
| Engineering tooling (GitHub, CI, monitoring, logging) | $250 $/FTE/month | $3,684 | $2,024 | $2,147 |
| Build partner retainer: engineering, DevOps & QA (no in-house engineers) | $40,000 Fixed $/month | $169,600 | - | - |
| Build partner retainer (expanded): secondary market, mobile & scale | $50,000 Fixed $/month | - | $674,732 | $715,824 |
| Registry / fund-admin platform fees | $4,000 Fixed $/month | $25,440 | $53,979 | $57,266 |
| AI support agent & AI ops tooling | $5,000 Fixed $/month | $31,800 | $67,473 | $71,582 |
| AI support agent - volume usage | $0.40 $/investor/month | $66 | $5,555 | $28,601 |
| Security operations (MDR/SIEM service) | $6,000 Fixed $/month | $38,160 | $80,968 | $85,899 |
| Annual penetration testing | $4,000 Fixed $/month | - | $53,979 | $57,266 |
| SOC 2 Type II annual audit | $5,000 Fixed $/month | - | $67,473 | $71,582 |
| Identity verification & AML onboarding checks | $35 $/new investor | $2,164 | $43,295 | $119,702 |
| E-signature & document processing | $12 $/new investor | $742 | $14,844 | $41,041 |
| Ongoing AML monitoring & notifications | $1.00 $/investor/month | $165 | $13,888 | $71,503 |
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

## Buy-back facility (the exit bridge until the secondary market)
v9: available from year 2 for property tokens (20% yrs 1-2, 15% yr 3, 10% yr 4, 5% yr 5 below independent valuation) and PE trust units (10% / 8% / 5%). Discretionary and capped at 5% of eligible units a year. Before the secondary market opens (month 31) the model assumes 4% of eligible units are bought back a year (5% Downside = cap); after it opens, requests go to the secondary market first and the group buys only the residual (1% a year). Resale within ~4 months with a price haircut plus 1.5% costs; 1.5% cash reserve on top of the minimum buffer.

| $ | 2027 | 2028 | 2029 | 2030 | 2031 |
|---|---|---|---|---|---|
| Units bought back (valuation) | - | $600,568 | $1,875,367 | $2,276,309 | $4,484,926 |
| Resale gains | - | $7,631 | $121,454 | $108,788 | $230,090 |
| Units held for resale (year end) | - | $427,381 | $375,378 | $862,829 | $1,603,273 |
| Cash reserve required (year end) | - | $758,780 | $2,328,287 | $5,056,540 | $9,053,842 |

Requires a legal and capital-adequacy opinion before launch ($40k in set-up).

## Not in scope
BHive Coin, BHive Card, token lending, retail access (all from the July 2026 BHive Capital pack) - excluded until separate legal and business cases.
