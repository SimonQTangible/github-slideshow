"""Pull headline numbers from the recalculated model into JSON for the deck and playbook."""
import openpyxl, json, sys, datetime as dt
src = sys.argv[1]; out = sys.argv[2]
wb = openpyxl.load_workbook(src, data_only=True)
an = wb['Annual Summary']
lab = {}
for r in range(1, an.max_row+1):
    v = an.cell(r,1).value
    if v and v not in lab: lab[v] = r
def yrs(label): return [an.cell(lab[label], c).value or 0 for c in range(2,7)]
def km(label): return an.cell(lab[label], 2).value
D = {}
for k,l in [("cap","Capital placed through BHive"),("aum","Managed AUM - year end"),("inv","Investors on BHive - year end"),
            ("brk","Active accredited brokers - year end"),("rhp","Honey Pot - distribution"),("rbc","B-Catcher - structuring & management"),
            ("rbh","BHive - platform & marketplace"),("rev","GROUP NET REVENUE"),("gp","Gross profit"),("staff","Staff costs (salaries & on-costs)"),
            ("opex","Total operating expenses"),("ebitda","EBITDA"),("npat","NET PROFIT AFTER TAX"),("cash","Closing cash - year end"),
            ("fte","FTE - GROUP (year end)"),("fte_bc","FTE - B-Catcher (year end)"),("fte_hp","FTE - Honey Pot (year end)"),("fte_bh","FTE - BHive (year end)"),
            ("eb_bc","B-Catcher"),("eb_hp","Honey Pot"),("eb_bh","BHive"),("cogs","Cost of sales (broker commissions, partner share, fund admin)"),
            ("setup","Set-up & launch costs (expensed)"),("rbb","BHive - buy-back resale gains"),("bbnav","Units bought back under buy-back facility (at valuation)"),("bbinv","Bought-back units held for resale - year end"),("bbres","Buy-back liquidity reserve required - year end"),("capex","Capital expenditure"),("trade","Secondary market value traded")]:
    D[k] = yrs(l)
D["cats"] = {}
from data_cats import CATS
for c in CATS: D["cats"][c] = yrs(c)
def s(v): return v.strftime("%b %Y") if isinstance(v,(dt.date,dt.datetime)) else v
D["km"] = {k: s(km(l)) for k,l in [("peak","Peak funding requirement (lowest cumulative pre-financing cash)"),("raised","Total equity raised"),
           ("mincash","Lowest month-end cash balance"),("ebm","First month of positive group EBITDA"),("cfm","First month of positive pre-financing cash flow"),
           ("setup","Total one-off set-up & build cost")]}
# setup by category & entity
su = wb['Setup Costs']
D["setup_cat"] = {}; D["setup_ent"] = {}; D["setup_items"] = []
for r in range(9, su.max_row+1):
    cat, ent, amt = su.cell(r,2).value, su.cell(r,1).value, su.cell(r,9).value
    if not cat or not isinstance(amt,(int,float)) or su.cell(r,3).value in (None,"TOTAL SET-UP"): continue
    if su.cell(r,7).value not in ("Opex","Capex"): continue
    D["setup_cat"][cat] = D["setup_cat"].get(cat,0)+amt
    D["setup_ent"][ent] = D["setup_ent"].get(ent,0)+amt
    D["setup_items"].append([ent, cat, su.cell(r,3).value, amt, su.cell(r,5).value, su.cell(r,6).value, su.cell(r,7).value])
# funding & returns
fr = wb['Funding & Returns']
rows = {fr.cell(r,1).value: r for r in range(1, fr.max_row+1) if fr.cell(r,1).value}
D["rounds"] = [[n, fr.cell(rows[n],2).value, fr.cell(rows[n],3).value, fr.cell(rows[n],6).value] for n in ["Seed","Series A","Series B (optional)"]]
D["founder_pct"] = fr.cell(rows["Founders & existing holders (B-Catcher + BHive)"],6).value
D["esop_pct"] = fr.cell(rows["Group executive share pool (parent ESOP, created before Seed)"],6).value
D["founder_check"] = fr.cell(rows["Founder control check: founders' ownership vs 51% floor"],2).value
D["minority"] = fr.cell(rows["Less: key-people minority stakes in Honey Pot & BHive (at EBITDA multiple)"],2).value
uses = ["Team - salaries, on-costs & hiring (all entities)","Platform & technology build (capitalised)","Licensing, legal, structuring & compliance set-up",
        "Technology, hosting & security (recurring)","Marketing, broker network & investor servicing","Regulatory, insurance, professional fees & governance",
        "Premises, origination, travel & general"]
D["uses"] = [[u, fr.cell(rows[u],2).value] for u in uses]
D["returns"] = [[n, fr.cell(rows[n],2).value, fr.cell(rows[n],5).value, fr.cell(rows[n],6).value] for n in ["Seed investors","Series A investors","Series B investors"]]
D["exit"] = {k: fr.cell(rows[l],2).value for k,l in [("ebitda_rr","Exit run-rate EBITDA (Q4 year 5 x 4)"),("rev_rr","Exit run-rate net revenue (Q4 year 5 x 4)"),("eq","Equity value to B-Catcher Group shareholders")]}
# scenarios
sc = wb['Scenarios']
cases = [sc.cell(4,c).value for c in range(2, sc.max_column+1) if sc.cell(4,c).value]
D["scen"] = {"cases": cases, "rows": {}}
for r in range(5, 17):
    m = sc.cell(r,1).value
    if m: D["scen"]["rows"][m] = [s(sc.cell(r,c).value) for c in range(2, 2+len(cases))]
# quarterly cash & cumulative pre-financing
cf = wb['Cash Flow']
cr = {cf.cell(r,1).value: r for r in range(1, cf.max_row+1) if cf.cell(r,1).value}
D["cash_q"] = [cf.cell(cr["CLOSING CASH"], 4+m).value for m in range(3,61,3)]
D["cum_q"] = [cf.cell(cr["Cumulative cash flow before financing"], 4+m).value for m in range(3,61,3)]
# headcount by entity people count
hc = wb['Headcount']
D["roles"] = []
for r in range(9, 200):
    if hc.cell(r,1).value in ("B-Catcher","Honey Pot","BHive") and hc.cell(r,10).value and isinstance(hc.cell(r,10).value,(int,float)) and hc.cell(r,5).value:
        D["roles"].append([hc.cell(r,1).value, hc.cell(r,2).value, hc.cell(r,3).value, hc.cell(r,4).value, hc.cell(r,13).value, hc.cell(r,6).value, hc.cell(r,10).value, hc.cell(r,12).value])
    elif D["roles"] and hc.cell(r,1).value is None: break
inp = wb['Inputs']

ir = wb['Investor Returns']
D["ir"] = {"cols": ["Property - geared", "Property - Debt-Free", "Private credit", "PE / pre-IPO"]}
for key, r in [("gross", 10), ("mgmt", 11), ("tech", 12), ("upfront", 13), ("upfront_pa", 14), ("net", 15), ("fees", 16), ("kept", 17), ("ex150", 18)]:
    D["ir"][key] = [ir.cell(r, c).value for c in range(2, 6)]
D["inputs"] = {inp.cell(r,1).value: inp.cell(r,2).value for r in range(9, inp.max_row+1) if inp.cell(r,1).value and inp.cell(r,2).value is not None}
json.dump(D, open(out,"w"), indent=1, default=str)
print("ok", len(D["roles"]), "roles", len(D["setup_items"]), "setup items")
# --- extras for playbook
bu = wb['Budget Y1-Y2']
D["budget"] = []
for r in range(5, bu.max_row+1):
    a = bu.cell(r,1).value
    if a is None: continue
    D["budget"].append([a] + [bu.cell(r,c).value for c in range(2,12)])
ox = wb['Opex']
D["opex_items"] = []
for r in range(9, ox.max_row+1):
    if ox.cell(r,1).value in ("B-Catcher","Honey Pot","BHive") and ox.cell(r,5).value in [x for x in ["Fixed $/month","$/FTE/month","$/active broker/month","$/new broker","$/new investor","$/investor/month","% of capital placed","% of entity revenue"]]:
        tot = [sum((ox.cell(r, 11+m).value or 0) for m in range(1+12*y, 13+12*y)) for y in range(5)]
        D["opex_items"].append([ox.cell(r,1).value, ox.cell(r,2).value, ox.cell(r,3).value, ox.cell(r,4).value, ox.cell(r,5).value, ox.cell(r,6).value, ox.cell(r,7).value] + tot)
json.dump(D, open(out,"w"), indent=1, default=str)
print("extras", len(D["budget"]), len(D["opex_items"]))
