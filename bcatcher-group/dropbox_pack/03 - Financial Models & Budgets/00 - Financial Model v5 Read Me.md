# Financial Model v5 - Read Me

**B-Catcher Group** | 03 - Financial Models & Budgets | Version 5.0 | 23 November 2026  
Confidential planning document. All figures are AUD Base-case planning assumptions from *BCatcher_Group_Integrated_Financial_Model.xlsx*. Not an offer or financial advice.

## Status
- **Current:** `BCatcher_Group_Integrated_Financial_Model.xlsx` (upload to this folder from the Claude session). 16 sheets, ~38,000 live formulas, zero errors, balance sheet balances.
- **Superseded:** `BCatcher_BHive_36_Month_Financial_Model.xlsx` - one flat cost line, no entity split, no broker commissions or headcount link, cash negative from month 16.
- The CSV files here are value exports of each sheet (Base case) so the numbers are readable in Dropbox without the workbook.

## Sheets in the workbook
| Sheet | Contents |
|---|---|
| Cover | Key outputs, model flow, colour code, integrity checks |
| Inputs | Scenario selector (Base/Downside/Upside), every assumption with source, stress levers |
| Annual Summary | Five-year P&L, volumes, entity EBITDA, cash, people, unit economics, key metrics |
| Scenarios | Nine cases run through the full model |
| Funding & Returns | Funding requirement, cap table, use of funds, investor returns |
| Budget Y1-Y2 | Quarterly budget per cost centre + set-up budget by category |
| P&L Consolidated / B-Catcher / Honey Pot / BHive | Monthly P&Ls; intercompany eliminated on consolidation |
| Cash Flow / Balance Sheet | Monthly cash, funding, runway; balance check |
| Revenue | Capital placed, AUM roll-forward, investors, brokers, every fee line, cost drivers |
| Headcount | 27 roles; monthly FTE, loaded cost, hiring costs |
| Setup Costs / Opex | Every one-off and recurring cost with timing, driver and basis |

## How to use
Change `Inputs!B4` to switch scenario; change blue cells to update assumptions; stress levers (placement multiplier, licensing delay, broker delay, tech overrun, hiring deferral) sit at the bottom of Inputs.
