# B-Catcher Group — financial model, investor deck & launch playbook

Three linked deliverables for the B-Catcher / Honey Pot / BHive group, all generated from one model so every number agrees.

| File | What it is |
|---|---|
| `BCatcher_Group_Integrated_Financial_Model.xlsx` | 60-month, formula-driven model (~38,000 formulas, zero errors). Inputs with Base/Downside/Upside, every set-up cost, 49-role headcount plan, every recurring cost, revenue by entity, three cost-centre P&Ls, consolidated P&L, cash flow, balance sheet, annual summary, Y1–Y2 quarterly budgets, funding/cap table/returns, scenario & stress results. |
| `BCatcher_Group_Investor_Deck_v8.pptx` | 21-slide investor deck built from the model's Base case. |
| `BCatcher_Group_Launch_Playbook_and_Budget_Book.docx` | Start-to-finish launch sequence (step, owner, month, cost), full set-up register, headcount plan, budgets, opex register, funding, scenarios, KPIs, open decisions. |

## Rebuild after changing assumptions

Edit `build/data.py` (inputs, roles, set-up items, opex lines), then:

```bash
cd build
python3 build_all.py                                   # model + scenario runs + recalc
python3 extract.py ../BCatcher_Group_Integrated_Financial_Model.xlsx ../out/numbers.json
NODE_PATH=<dir with pptxgenjs> node build_deck.js ../out/numbers.json ../BCatcher_Group_Investor_Deck_v8.pptx
python3 build_playbook.py ../out/numbers.json ../BCatcher_Group_Launch_Playbook_and_Budget_Book.docx
```

For quick what-ifs, just open the workbook and change `Inputs!B4` (scenario) or the stress levers — everything recalculates.

## Sources merged
- `/TITCH CIP/B-Catcher Group/` v3 pack (business plan, regulatory, Honey Pot manual, BHive requirements, roadmap, deck brief, earlier 36-month model and headcount workbook)
- `/TITCH CIP/BHIVE/` (finance model, platform forecast, executive summary, AI-native operating model, retail structure)
- Legacy B-Catcher materials (B-Catcher-9 investor memorandum, UTEGBO term sheets, wallet spec, meeting minutes)

All figures are planning assumptions — not an offer or financial advice.
