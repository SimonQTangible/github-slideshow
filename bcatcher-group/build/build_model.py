"""Build the B-Catcher Group integrated 60-month financial model (formula-driven)."""
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.comments import Comment
from openpyxl.chart import BarChart, LineChart, Reference

from data import *

OUT = sys.argv[1] if len(sys.argv) > 1 else "model.xlsx"
NM = N_MONTHS

# ------------------------------------------------------------------ styles
FONT = "Arial"
NAVY, GOLD, GREY, LIGHT = "1F2A44", "E0A526", "F2F2F2", "FFF4D6"
f_base = Font(name=FONT, size=9)
f_bold = Font(name=FONT, size=9, bold=True)
f_input = Font(name=FONT, size=9, color="0000FF")
f_link = Font(name=FONT, size=9, color="008000")
f_title = Font(name=FONT, size=14, bold=True, color=NAVY)
f_sub = Font(name=FONT, size=9, italic=True, color="595959")
f_hdr = Font(name=FONT, size=9, bold=True, color="FFFFFF")
fill_hdr = PatternFill("solid", fgColor=NAVY)
fill_sec = PatternFill("solid", fgColor="D9E1F2")
fill_tl = PatternFill("solid", fgColor=GREY)
fill_key = PatternFill("solid", fgColor="FFFF00")
fill_tot = PatternFill("solid", fgColor=LIGHT)
thin = Side(style="thin", color="808080")
b_top = Border(top=thin)
b_tot = Border(top=thin, bottom=Side(style="double", color="808080"))

FMT = {
    "money": '$#,##0;($#,##0);"-"',
    "pct": '0.0%;(0.0%);"-"',
    "num": '#,##0;(#,##0);"-"',
    "dec": '#,##0.0;(#,##0.0);"-"',
    "dec2": '0.00;(0.00);"-"',
    "int": '0;(0);"-"',
    "date": 'mmm-yy',
    "mult": '0.0"x"',
    "text": "@",
}

wb = Workbook()
wb.remove(wb.active)


def name(nm, ref):
    wb.defined_names[nm] = DefinedName(nm, attr_text=ref)


def ws_new(title, tab_color=None):
    ws = wb.create_sheet(title)
    ws.sheet_view.showGridLines = False
    if tab_color:
        ws.sheet_properties.tabColor = tab_color
    return ws


def q(sheet):
    return "'" + sheet + "'" if any(ch in sheet for ch in " &-") else sheet


def put(ws, ref, val, font=f_base, fmt=None, fill=None, bold=False, align=None, border=None):
    c = ws[ref]
    c.value = val
    c.font = Font(name=FONT, size=font.size, bold=bold or font.bold, color=font.color, italic=font.italic)
    if fmt:
        c.number_format = FMT.get(fmt, fmt)
    if fill:
        c.fill = fill
    if align:
        c.alignment = align
    if border:
        c.border = border
    return c


def title(ws, t, sub):
    put(ws, "A1", t, f_title)
    put(ws, "A2", sub, f_sub)


def section(ws, row, label, last_col):
    for ci in range(1, last_col + 1):
        ws.cell(row=row, column=ci).fill = fill_hdr
    put(ws, f"A{row}", label, f_hdr)


# ------------------------------------------------------------------ month-grid helper
class Grid:
    """A sheet whose months run across columns starting at `first`."""

    def __init__(self, ws, first, label_cols=4):
        self.ws, self.first = ws, first

    def c(self, m):  # column letter for month m (1-based)
        return L(self.first + m - 1)

    @property
    def c1(self):
        return self.c(1)

    @property
    def cN(self):
        return self.c(NM)

    def rng(self, row):
        return f"${self.c1}${row}:${self.cN}${row}"

    def timeline(self, label_col="A"):
        ws = self.ws
        labels = {3: "Model month", 4: "Period ending", 5: "Model year", 6: "Quarter"}
        for r, lab in labels.items():
            put(ws, f"{label_col}{r}", lab, f_bold, fill=fill_tl)
        for m in range(1, NM + 1):
            c = self.c(m)
            put(ws, f"{c}3", 1 if m == 1 else f"={self.c(m-1)}3+1", f_bold, "int", fill_tl,
                align=Alignment(horizontal="center"))
            put(ws, f"{c}4", f"=EOMONTH(StartDate,{c}3-1)", f_bold, "date", fill_tl, align=Alignment(horizontal="center"))
            put(ws, f"{c}5", f"=ROUNDUP({c}3/12,0)", f_base, "int", fill_tl, align=Alignment(horizontal="center"))
            put(ws, f"{c}6", f"=ROUNDUP({c}3/3,0)", f_base, "int", fill_tl, align=Alignment(horizontal="center"))
            ws.column_dimensions[c].width = 11
        ws.freeze_panes = ws[f"{self.c1}7"]

    def row(self, r, label, fn, fmt="money", unit="", total="sum", bold=False, font=f_base,
            border=None, fill=None, note=None):
        """fn(c, p, m) -> formula string (with leading '=') for month m in column c; p = prev col or None."""
        ws = self.ws
        put(ws, f"A{r}", label, f_bold if bold else f_base, fill=fill, border=border)
        if unit:
            put(ws, f"B{r}", unit, f_sub, fill=fill, border=border)
        elif fill or border:
            put(ws, f"B{r}", None, fill=fill, border=border)
        for m in range(1, NM + 1):
            c = self.c(m)
            p = self.c(m - 1) if m > 1 else None
            put(ws, f"{c}{r}", fn(c, p, m), font, fmt, fill, bold=bold, border=border)
        if total == "sum":
            put(ws, f"C{r}", f"=SUM({self.c1}{r}:{self.cN}{r})", f_bold if bold else f_base, fmt, fill, border=border)
        elif total == "last":
            put(ws, f"C{r}", f"={self.cN}{r}", f_base, fmt, fill, border=border)
        elif total == "max":
            put(ws, f"C{r}", f"=MAX({self.c1}{r}:{self.cN}{r})", f_base, fmt, fill, border=border)
        elif total == "min":
            put(ws, f"C{r}", f"=MIN({self.c1}{r}:{self.cN}{r})", f_base, fmt, fill, border=border)
        elif fill or border:
            put(ws, f"C{r}", None, fill=fill, border=border)
        if note:
            ws[f"A{r}"].comment = Comment(note, "Model")
        return r


def std_cols(ws, first=5):
    ws.column_dimensions["A"].width = 58
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 15
    if first == 5:
        ws.column_dimensions["D"].width = 2
        put(ws, "C7", "60-mth total", f_bold, align=Alignment(horizontal="right"))


# =====================================================================================
# COVER (filled at end)
cover = ws_new("Cover", NAVY)

# =====================================================================================
# INPUTS
inp = ws_new("Inputs", "0000FF")
title(inp, "INPUTS & SCENARIOS", "Blue = hard-coded input. Change the scenario in B4 (yellow). Live values in column B drive every sheet. All AUD, 2027 dollars.")
put(inp, "A4", "ACTIVE SCENARIO  >>", f_bold)
put(inp, "B4", "Base", Font(name=FONT, size=10, bold=True, color="0000FF"), fill=fill_key)
put(inp, "C4", "Choose Base / Downside / Upside", f_sub)
dv = DataValidation(type="list", formula1='"Base,Downside,Upside"', allow_blank=False)
inp.add_data_validation(dv)
dv.add("B4")
put(inp, "A5", "Scenario index", f_base)
put(inp, "B5", "=MATCH(B4,$D$7:$F$7,0)", f_base, "int")
name("ScenarioIdx", "Inputs!$B$5")
for col, h in zip("ABCDEFG", ["Assumption", "LIVE value", "Unit", "Base", "Downside", "Upside", "Basis / source / note"]):
    put(inp, f"{col}7", h, f_hdr, fill=fill_hdr)
widths = {"A": 62, "B": 15, "C": 14, "D": 14, "E": 14, "F": 14, "G": 90}
for k, v in widths.items():
    inp.column_dimensions[k].width = v

r = 8
INPUT_CELL = {}
for sec, items in INPUT_SECTIONS:
    r += 1
    section(inp, r, sec, 7)
    for (nm, label, unit, b, d, u, note, fmt) in items:
        r += 1
        put(inp, f"A{r}", label)
        put(inp, f"B{r}", f"=CHOOSE(ScenarioIdx,D{r},E{r},F{r})", f_bold, fmt)
        put(inp, f"C{r}", unit, f_sub)
        for col, v in zip("DEF", (b, d, u)):
            put(inp, f"{col}{r}", v, f_input, fmt)
        put(inp, f"G{r}", note, f_sub)
        name(nm, f"Inputs!$B${r}")
        INPUT_CELL[nm] = r
name("CapTargets", f"Inputs!$B${INPUT_CELL['CapY1']}:$B${INPUT_CELL['CapY5']}")
name("BrokerTargets", f"Inputs!$B${INPUT_CELL['BrokerY1']}:$B${INPUT_CELL['BrokerY5']}")

r += 2
section(inp, r, "STRESS LEVERS (single value - applied on top of the active scenario)", 7)
for (nm, label, unit, v, note, fmt) in STRESS_INPUTS:
    r += 1
    put(inp, f"A{r}", label)
    put(inp, f"B{r}", v, f_input, fmt, fill_key)
    put(inp, f"C{r}", unit, f_sub)
    put(inp, f"G{r}", note, f_sub)
    name(nm, f"Inputs!$B${r}")
    INPUT_CELL[nm] = r

r += 2
section(inp, r, "CALCULATED & STRUCTURAL", 7)
calc = [
    ("StartDate", "First model month (period end)", START_DATE, "date", "Month 1 = January 2027.", True),
    ("LaunchEff", "Effective launch month (incl. delay)", "=LaunchMonth+LaunchDelay", "int", "", False),
    ("SecEff", "Effective secondary-market month (incl. delay)", "=SecondaryMonth+LaunchDelay", "int", "", False),
    ("OnCostRate", "Total employee on-cost rate", "=Super+PayrollTax+WorkersComp+OtherOnCost", "pct", "", False),
    ("MixCheck", "Product mix check (must equal 100%)", "=MixProp+MixCredit+MixPE+MixASX", "pct", "", False),
]
for nm, label, v, fmt, note, is_input in calc:
    r += 1
    put(inp, f"A{r}", label)
    put(inp, f"B{r}", v, f_input if is_input else f_base, fmt)
    put(inp, f"G{r}", note, f_sub)
    name(nm, f"Inputs!$B${r}")
r += 2
section(inp, r, "LISTS", 7)
r += 1
put(inp, f"A{r}", "Entities", f_bold)
for i, e in enumerate(ENTITIES):
    put(inp, f"{L(2+i)}{r}", e)
name("EntityList", f"Inputs!$B${r}:$D${r}")
r += 1
put(inp, f"A{r}", "Opex drivers", f_bold)
drv_row = r
for i, d in enumerate(DRIVERS):
    put(inp, f"A{r+1+i}", f"  {i+1}", f_sub)
    put(inp, f"B{r+1+i}", d)
name("DriverList", f"Inputs!$B${r+1}:$B${r+len(DRIVERS)}")
inp.freeze_panes = "A8"

# =====================================================================================
# REVENUE & DRIVERS
rev = ws_new("Revenue", "008000")
title(rev, "REVENUE & OPERATING DRIVERS", "Capital placed, AUM, investors, brokers and fee revenue by entity. Monthly, 60 months.")
std_cols(rev)
G = Grid(rev, 5)
G.timeline()
R = {}
row = 8


def sec(g, label):
    global row
    row += 1
    section(g.ws, row, label, g.first + NM - 1)
    row += 1


def add(g, key, label, fn, **kw):
    global row
    g.row(row, label, fn, **kw)
    R[key] = row
    row += 1
    return row - 1


sec(G, "OPERATING TIMELINE")
add(G, "k", "Operating month (since launch)", lambda c, p, m: f"=MAX(0,{c}$3-LaunchEff+1)", fmt="int", total=None)
add(G, "oy", "Operating year", lambda c, p, m: f"=IF({c}{R['k']}=0,0,MIN(5,ROUNDUP({c}{R['k']}/12,0)))", fmt="int", total=None)
add(G, "ramp", "Launch ramp factor", lambda c, p, m: f"=IF({c}{R['k']}=0,0,MIN(1,{c}{R['k']}/RampMonths))", fmt="pct", total=None)
add(G, "seclive", "Secondary market live (1 = yes)", lambda c, p, m: f"=IF({c}$3>=SecEff,1,0)", fmt="int", total=None)
add(G, "secm", "Months secondary market live", lambda c, p, m: f"=IF({c}{R['seclive']}=1,{c}$3-SecEff+1,0)", fmt="int", total=None)

add(G, "bplan", "Planned active brokers (no recruitment delay)",
    lambda c, p, m: (f"=IF({c}{R['oy']}=0,0,IF({c}{R['oy']}=1,0,INDEX(BrokerTargets,MAX(1,{c}{R['oy']}-1)))"
                     f"+(INDEX(BrokerTargets,{c}{R['oy']})-IF({c}{R['oy']}=1,0,INDEX(BrokerTargets,MAX(1,{c}{R['oy']}-1))))"
                     f"*MIN(12,{c}{R['k']}-({c}{R['oy']}-1)*12)/12)"), fmt="dec", total=None)
add(G, "bfac", "Broker capacity factor (actual / planned brokers)", lambda c, p, m: "=1", fmt="pct", total=None)
sec(G, "CAPITAL PLACED (investor subscriptions through BHive - NOT group revenue)")
add(G, "cap", "Total capital placed", lambda c, p, m: f"=IF({c}{R['oy']}=0,0,INDEX(CapTargets,{c}{R['oy']})/12*{c}{R['ramp']}*PlacementMult*((1-BrokerChannel)+BrokerChannel*{c}{R['bfac']}))", bold=True)
add(G, "capP", "  Property trusts", lambda c, p, m: f"={c}{R['cap']}*MixProp")
add(G, "capC", "  Private credit", lambda c, p, m: f"={c}{R['cap']}*MixCredit")
add(G, "capE", "  Private equity / pre-IPO", lambda c, p, m: f"={c}{R['cap']}*MixPE")
add(G, "capA", "  ASX placements / IPO", lambda c, p, m: f"={c}{R['cap']}*MixASX")
add(G, "capB", "Broker-sourced capital", lambda c, p, m: f"={c}{R['cap']}*BrokerChannel")
add(G, "capD", "Direct / Honey Pot IR capital", lambda c, p, m: f"={c}{R['cap']}-{c}{R['capB']}")
add(G, "capCum", "Cumulative capital placed", lambda c, p, m: f"={c}{R['cap']}" + (f"+{p}{row}" if p else ""), total="last")

sec(G, "INVESTORS, BROKERS & OFFERS")
add(G, "newInv", "New investors", lambda c, p, m: f"={c}{R['cap']}/AvgTicket", fmt="dec")
add(G, "cumInv", "Cumulative investors on BHive", lambda c, p, m: f"={c}{R['newInv']}" + (f"+{p}{row}" if p else ""), fmt="num", total="last")
add(G, "kb", "Broker ramp month", lambda c, p, m: f"=MAX(0,{c}{R['k']}-BrokerDelay)", fmt="int", total=None)
add(G, "yb", "Broker ramp year", lambda c, p, m: f"=IF({c}{R['kb']}=0,0,MIN(5,ROUNDUP({c}{R['kb']}/12,0)))", fmt="int", total=None)
add(G, "brk", "Active accredited brokers",
    lambda c, p, m: (f"=IF({c}{R['yb']}=0,0,IF({c}{R['yb']}=1,0,INDEX(BrokerTargets,MAX(1,{c}{R['yb']}-1)))"
                     f"+(INDEX(BrokerTargets,{c}{R['yb']})-IF({c}{R['yb']}=1,0,INDEX(BrokerTargets,MAX(1,{c}{R['yb']}-1))))"
                     f"*MIN(12,{c}{R['kb']}-({c}{R['yb']}-1)*12)/12)"), fmt="dec", total="last")
add(G, "newBrk", "New brokers onboarded (growth + churn replacement)",
    lambda c, p, m: (f"=MAX(0,{c}{R['brk']}-{p}{R['brk']})+{p}{R['brk']}*BrokerChurn/12" if p else f"={c}{R['brk']}"), fmt="dec")
for m in range(1, NM + 1):
    c = G.c(m)
    put(rev, f"{c}{R['bfac']}", f"=IF({c}{R['bplan']}=0,1,MIN(1,{c}{R['brk']}/{c}{R['bplan']}))", f_base, "pct")
add(G, "offers", "Offers launched", lambda c, p, m: f"={c}{R['cap']}/OfferSize", fmt="dec2")

sec(G, "ASSETS UNDER MANAGEMENT (held in trusts/SPVs - NOT group assets)")
for key, lab, capk, g_nm, ro_nm in [("P", "Property trusts", "capP", "GrowthProp", "RunoffProp"),
                                    ("C", "Private credit", "capC", None, "RunoffCredit"),
                                    ("E", "Private equity / pre-IPO", "capE", "GrowthPE", "RunoffPE")]:
    o = row
    add(G, f"aumO{key}", f"{lab} - opening AUM", lambda c, p, m, o=o: (f"={p}{o+3}" if p else "=0"), total=None)
    add(G, f"aumN{key}", f"  + new capital", lambda c, p, m, capk=capk: f"={c}{R[capk]}")
    gexpr = f"{g_nm}-{ro_nm}" if g_nm else f"-{ro_nm}"
    add(G, f"aumG{key}", f"  + revaluation less run-off", lambda c, p, m, o=o, gexpr=gexpr: f"={c}{o}*({gexpr})/12")
    add(G, f"aumC{key}", f"{lab} - closing AUM", lambda c, p, m, o=o: f"=SUM({c}{o}:{c}{o+2})", bold=True, total="last")
add(G, "gav", "Memo: property trusts gross asset value (incl. trust debt)", lambda c, p, m: f"={c}{R['aumCP']}/(1-PropLVR)", total="last")
add(G, "aum", "Total managed AUM (closing)", lambda c, p, m: f"={c}{R['aumCP']}+{c}{R['aumCC']}+{c}{R['aumCE']}", bold=True, total="last")
add(G, "aumAvg", "Average managed AUM in month",
    lambda c, p, m: (f"=({p}{R['aum']}+{c}{R['aum']})/2" if p else f"={c}{R['aum']}/2"), total=None)

sec(G, "HONEY POT - DISTRIBUTION REVENUE")
add(G, "hpP", "Placement fees - property trusts", lambda c, p, m: f"={c}{R['capP']}*FeeProp")
add(G, "hpC", "Placement fees - private credit", lambda c, p, m: f"={c}{R['capC']}*FeeCredit")
add(G, "hpE", "Placement fees - private equity / pre-IPO", lambda c, p, m: f"={c}{R['capE']}*FeePE")
add(G, "hpA", "Placement fees - ASX placements / IPO", lambda c, p, m: f"={c}{R['capA']}*FeeASX")
add(G, "hpRev", "Total placement fees", lambda c, p, m: f"=SUM({c}{R['hpP']}:{c}{R['hpA']})", bold=True, border=b_top)
add(G, "hpComm", "Broker commissions (cost of sales)", lambda c, p, m: f"={c}{R['hpRev']}*BrokerChannel*BrokerPayout")

sec(G, "B-CATCHER - INVESTMENT & ASSET MANAGEMENT REVENUE")
add(G, "bcAcq", "Entry fees - property (BHive 1.5% of equity raised)", lambda c, p, m: f"={c}{R['capP']}*AcqFee")
add(G, "bcStruct", "Entry fees - credit & PE", lambda c, p, m: f"=({c}{R['capC']}+{c}{R['capE']})*StructFee")
add(G, "bcMP", "Management fees - property (BHive 1.25% of AUM)", lambda c, p, m: f"=({c}{R['aumOP']}+{c}{R['aumCP']})/2*MgmtProp/12")
add(G, "bcMC", "Management fees - private credit", lambda c, p, m: f"=({c}{R['aumOC']}+{c}{R['aumCC']})/2*MgmtCredit/12")
add(G, "bcME", "Management fees - private equity", lambda c, p, m: f"=({c}{R['aumOE']}+{c}{R['aumCE']})/2*MgmtPE/12")
add(G, "bcRev", "Total B-Catcher external revenue", lambda c, p, m: f"={c}{R['bcAcq']}+{c}{R['bcStruct']}+SUM({c}{R['bcMP']}:{c}{R['bcME']})", bold=True, border=b_top)
add(G, "bcAdmin", "Manager-borne fund administration (cost of sales)", lambda c, p, m: f"={c}{R['aumAvg']}*MgrAdminCost/12")

sec(G, "BHIVE - PLATFORM & MARKETPLACE REVENUE")
add(G, "bhPlat", "Platform administration fees", lambda c, p, m: f"={c}{R['aumAvg']}*PlatformFee/12")
add(G, "bhOffer", "Offer onboarding fees", lambda c, p, m: f"={c}{R['offers']}*OfferFee")
add(G, "bhTraded", "Secondary market value traded", lambda c, p, m: f"={c}{R['seclive']}*{c}{R['aumAvg']}*SecTurnover/12*IF({c}{R['secm']}<=12,SecRamp,1)")
add(G, "bhSec", "Secondary transaction fees", lambda c, p, m: f"={c}{R['bhTraded']}*SecFee")
add(G, "bhRev", "Total BHive external revenue (excl. buy-back gains)", lambda c, p, m: f"={c}{R['bhPlat']}+{c}{R['bhOffer']}+{c}{R['bhSec']}", bold=True, border=b_top)
add(G, "bhPartner", "Market-operator partner share (cost of sales)", lambda c, p, m: f"={c}{R['bhSec']}*SecPartnerShare")

sec(G, "BUY-BACK FACILITY (exit ladder / UTEGBO) - group buys units at a discount and resells at valuation")
add(G, "bbP", "Property tokens bought back (at valuation)", lambda c, p, m: f"=IF({c}{R['k']}>=BBStartProp,{c}{R['aumOP']}*IF({c}$3>=SecEff,BBRatePost,BBRate)/12,0)")
add(G, "bbE", "PE / trust units bought back (at valuation)", lambda c, p, m: f"=IF({c}{R['k']}>=BBStartPE,{c}{R['aumOE']}*IF({c}$3>=SecEff,BBRatePost,BBRate)/12,0)")
add(G, "bbNAV", "Total bought back (at valuation)", lambda c, p, m: f"={c}{R['bbP']}+{c}{R['bbE']}", bold=True)
add(G, "bbCost", "Cash paid to investors (after ladder discount)", lambda c, p, m: f"={c}{R['bbP']}*(1-BBDiscProp)+{c}{R['bbE']}*(1-BBDiscPE)")
add(G, "bbSoldNAV", "Units resold (valuation, after resale lag)", lambda c, p, m: f"=IF({c}$3>BBResaleMonths,INDEX($E${R['bbNAV']}:$BL${R['bbNAV']},{c}$3-BBResaleMonths),0)")
add(G, "bbSoldCost", "Cost of units resold", lambda c, p, m: f"=IF({c}$3>BBResaleMonths,INDEX($E${R['bbCost']}:$BL${R['bbCost']},{c}$3-BBResaleMonths),0)")
add(G, "bbProceeds", "Resale proceeds (net of resale costs)", lambda c, p, m: f"={c}{R['bbSoldNAV']}*(1-BBResaleCost)")
add(G, "bbGain", "Buy-back resale gain (BHive revenue)", lambda c, p, m: f"={c}{R['bbProceeds']}-{c}{R['bbSoldCost']}", bold=True)
add(G, "bbInv", "Units held for resale (group balance sheet)", lambda c, p, m: f"={c}{R['bbCost']}-{c}{R['bbSoldCost']}" + (f"+{p}{row}" if p else ""), total="last")
add(G, "bbElig", "Eligible AUM for buy-back", lambda c, p, m: f"=IF({c}{R['k']}>=BBStartProp,{c}{R['aumCP']},0)+IF({c}{R['k']}>=BBStartPE,{c}{R['aumCE']},0)", total="last")
add(G, "bbRes", "Required buy-back liquidity reserve", lambda c, p, m: f"={c}{R['bbElig']}*BBReserve", total="max")

sec(G, "INTERCOMPANY (eliminated on consolidation)")
add(G, "icSSFhp", "B-Catcher shared-services fee <- Honey Pot", lambda c, p, m: f"=IF({c}$3>=3,IC_SSF_HP,0)")
add(G, "icSSFbh", "B-Catcher shared-services fee <- BHive", lambda c, p, m: f"=IC_SSF_BH")
add(G, "icLICbh", "BHive platform licence <- Honey Pot", lambda c, p, m: f"=IF({c}$3>=LaunchEff,IC_LIC_HP,0)")
add(G, "icLICbc", "BHive platform licence <- B-Catcher", lambda c, p, m: f"=IF({c}$3>=LaunchEff,IC_LIC_BC,0)")

sec(G, "COST DRIVERS (feed the Opex register)")
DRV_START = row
add(G, "dFTEbc", "FTE driver - B-Catcher (total group FTE, shared services)", lambda c, p, m: "=0", fmt="dec", total=None)  # filled later
add(G, "dFTEhp", "FTE driver - Honey Pot", lambda c, p, m: "=0", fmt="dec", total=None)
add(G, "dFTEbh", "FTE driver - BHive", lambda c, p, m: "=0", fmt="dec", total=None)
add(G, "dBrk", "Active brokers", lambda c, p, m: f"={c}{R['brk']}", fmt="dec", total=None)
add(G, "dNewBrk", "New brokers", lambda c, p, m: f"={c}{R['newBrk']}", fmt="dec", total=None)
add(G, "dNewInv", "New investors", lambda c, p, m: f"={c}{R['newInv']}", fmt="dec", total=None)
add(G, "dCumInv", "Cumulative investors", lambda c, p, m: f"={c}{R['cumInv']}", fmt="num", total=None)
add(G, "dCap", "Capital placed", lambda c, p, m: f"={c}{R['cap']}", total=None)
add(G, "dRevbc", "External revenue - B-Catcher", lambda c, p, m: f"={c}{R['bcRev']}", total=None)
add(G, "dRevhp", "External revenue - Honey Pot", lambda c, p, m: f"={c}{R['hpRev']}", total=None)
add(G, "dRevbh", "External revenue - BHive", lambda c, p, m: f"={c}{R['bhRev']}", total=None)

# =====================================================================================
# HEADCOUNT
hc = ws_new("Headcount", "7030A0")
title(hc, "HEADCOUNT & PEOPLE COSTS", "Role register (blue = input). FTE, loaded cost and one-off hiring costs calculated monthly. Salaries are 2027 AUD base; wage growth from Inputs.")
HF = 14
HG = Grid(hc, HF)
HG.timeline(label_col="C")
hdrs = ["Entity", "Department", "Role", "Type", "Start month", "FTE", "Step month", "FTE from step", "End month", "Base salary", "Notes", "Loaded cost p.a. (yr-1 $)", "Effective start"]
for i, h in enumerate(hdrs):
    put(hc, f"{L(i+1)}8", h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True, vertical="center"))
for ci in range(HF - 1, HF + NM):
    hc.cell(row=8, column=ci).fill = fill_hdr
for k, v in zip("ABCDEFGHIJKLM", [12, 16, 46, 11, 8, 6, 8, 8, 8, 12, 34, 13, 9]):
    hc.column_dimensions[k].width = v
hc.row_dimensions[8].height = 30
NR = len(ROLES)
R0 = 9
put(hc, f"{HG.c1}8", "FTE by month >>", f_hdr)
for i, (ent, dept, role, typ, st, f1, stp, f2, end, sal, note) in enumerate(ROLES):
    r = R0 + i
    vals = [ent, dept, role, typ, st, f1, stp, f2, end, sal, note]
    fmts = [None, None, None, None, "int", "dec2", "int", "dec2", "int", "money", None]
    for j, (v, fm) in enumerate(zip(vals, fmts)):
        put(hc, f"{L(j+1)}{r}", v, f_input if j in (3, 4, 5, 6, 7, 8, 9) else f_base, fm)
    put(hc, f"L{r}", f'=J{r}*F{r}*(1+IF(D{r}="Employee",OnCostRate,ContractorPremium))', f_base, "money")
    put(hc, f"M{r}", f'=IF(E{r}>=13,E{r}+HireDeferral,E{r})', f_base, "int")
    for m in range(1, NM + 1):
        c = HG.c(m)
        put(hc, f"{c}{r}", f'=IF(AND({c}$3>=$M{r},{c}$3<=IF($I{r}="",9999,$I{r})),IF(AND($G{r}<>"",{c}$3>=$G{r}),$H{r},$F{r}),0)', f_base, "dec2")
RL = R0 + NR - 1
# example row comment
hc[f"C{R0}"].comment = Comment("Example row: entity, department, role, type (Employee/Contractor), start month, FTE, optional step-up month & FTE, optional end month, base salary.", "Model")

# cost block
CB = RL + 3
section(hc, CB, "LOADED SALARY COST (salary x FTE x wage growth x on-costs or contractor premium)", HF + NM - 1)
put(hc, f"{HG.c1}{CB}", "Cost by month >>", f_hdr)
for i in range(NR):
    r, s = CB + 1 + i, R0 + i
    put(hc, f"A{r}", f"=A{s}", f_link)
    put(hc, f"B{r}", f"=B{s}", f_link)
    put(hc, f"C{r}", f"=C{s}", f_link)
    for m in range(1, NM + 1):
        c = HG.c(m)
        put(hc, f"{c}{r}", f'={c}{s}*$J{s}/12*(1+WageInflation)^({c}$5-1)*(1+IF($D{s}="Employee",OnCostRate,ContractorPremium))', f_base, "money")
CL = CB + NR
# recruitment block
RB = CL + 2
section(hc, RB, "ONE-OFF HIRING COSTS (recruitment fee for employees hired after month 1 + IT equipment)", HF + NM - 1)
for i in range(NR):
    r, s = RB + 1 + i, R0 + i
    put(hc, f"A{r}", f"=A{s}", f_link)
    put(hc, f"B{r}", f"=B{s}", f_link)
    put(hc, f"C{r}", f"=C{s}", f_link)
    for m in range(1, NM + 1):
        c = HG.c(m)
        put(hc, f"{c}{r}", f'=IF({c}$3=$M{s},IF(AND($D{s}="Employee",$M{s}>1),RecruitPct*$J{s},0)+ITEquip,0)', f_base, "money")
RLL = RB + NR
# totals
TB = RLL + 2
section(hc, TB, "TOTALS BY ENTITY", HF + NM - 1)
HCT = {}
r = TB + 1
for kind, lbl, a, b, fmt in [("fte", "FTE", R0, RL, "dec"), ("ppl", "Headcount (people)", R0, RL, "int"),
                             ("cost", "Loaded salary cost", CB + 1, CL, "money"), ("rec", "Hiring costs", RB + 1, RLL, "money")]:
    for e in ENTITIES:
        put(hc, f"C{r}", f"{lbl} - {e}", f_bold if kind == "fte" else f_base)
        put(hc, f"A{r}", e, f_sub)
        for m in range(1, NM + 1):
            c = HG.c(m)
            if kind == "ppl":
                f = f'=COUNTIFS($A${a}:$A${b},"{e}",{c}${a}:{c}${b},">0")'
            else:
                f = f'=SUMIF($A${a}:$A${b},"{e}",{c}${a}:{c}${b})'
            put(hc, f"{c}{r}", f, f_base, fmt)
        HCT[(kind, e)] = r
        r += 1
    put(hc, f"C{r}", f"{lbl} - GROUP TOTAL", f_bold, fill=fill_tot)
    for m in range(1, NM + 1):
        c = HG.c(m)
        put(hc, f"{c}{r}", f"=SUM({c}{r-3}:{c}{r-1})", f_bold, fmt, fill_tot)
    HCT[(kind, "Group")] = r
    r += 2
put(hc, f"C{r}", "BHive technology & product staff cost (R&D base)", f_bold)
for m in range(1, NM + 1):
    c = HG.c(m)
    put(hc, f"{c}{r}", f'=SUMIFS({c}${CB+1}:{c}${CL},$A${CB+1}:$A${CL},"BHive",$B${CB+1}:$B${CL},"Technology")+SUMIFS({c}${CB+1}:{c}${CL},$A${CB+1}:$A${CL},"BHive",$B${CB+1}:$B${CL},"Product")', f_base, "money")
HCT["rd"] = r
hc.freeze_panes = hc[f"D9"]

# fill FTE driver rows on Revenue
for key, e in [("dFTEbc", "Group"), ("dFTEhp", "Honey Pot"), ("dFTEbh", "BHive")]:
    rr = R[key]
    for m in range(1, NM + 1):
        put(rev, f"{G.c(m)}{rr}", f"=Headcount!{HG.c(m)}{HCT[('fte', e)]}", f_link, "dec")

# =====================================================================================
# SET-UP COSTS
su = ws_new("Setup Costs", "C00000")
title(su, "SET-UP & LAUNCH COSTS (ONE-OFF)", "Every pre-launch and build cost by entity. Blue = input estimate (replace with quotes). Capex is capitalised and amortised; opex hits the P&L when incurred.")
SF = 12
SG = Grid(su, SF)
SG.timeline(label_col="C")
for i, h in enumerate(["Entity", "Category", "Item", "Estimate (base)", "Start month", "Months", "Treatment", "Tech build?", "Live amount", "Basis / source"]):
    put(su, f"{L(i+1)}8", h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True, vertical="center"))
for ci in range(SF - 1, SF + NM):
    su.cell(row=8, column=ci).fill = fill_hdr
for k, v in zip("ABCDEFGHIJK", [11, 28, 62, 13, 7, 7, 9, 7, 13, 50, 2]):
    su.column_dimensions[k].width = v
su.row_dimensions[8].height = 30
S0 = 9
items = list(SETUP)
for i, (ent, cat, item, amt, st, mo, tr, tf, basis) in enumerate(items):
    r = S0 + i
    for j, (v, fm) in enumerate(zip([ent, cat, item, amt, st, mo, tr, tf], [None, None, None, "money", "int", "int", None, None])):
        put(su, f"{L(j+1)}{r}", v, f_input if j >= 3 else f_base, fm)
    put(su, f"J{r}", basis, f_sub)
SL0 = S0 + len(items) - 1
rc = SL0 + 1
for j, v in enumerate(["B-Catcher", "Contingency", "Set-up contingency (on all lines above)", f"=SetupContPct*SUM(D{S0}:D{SL0})", 1, 12, "Opex", "N"]):
    put(su, f"{L(j+1)}{rc}", v, f_input if j in (4, 5, 6, 7) else f_base, ["", "", "", "money", "int", "int", "", ""][j] or None)
put(su, f"J{rc}", "Unallocated contingency spread over year 1.", f_sub)
SL = rc
for r in range(S0, SL + 1):
    put(su, f"I{r}", f'=D{r}*SetupFactor*IF(H{r}="Y",1+DevOverrun,1)', f_base, "money")
    for m in range(1, NM + 1):
        c = SG.c(m)
        put(su, f"{c}{r}", f"=IF(AND({c}$3>=$E{r},{c}$3<$E{r}+$F{r}),$I{r}/$F{r},0)", f_base, "money")
put(su, f"C{SL+1}", "TOTAL SET-UP", f_bold, fill=fill_tot, border=b_tot)
put(su, f"D{SL+1}", f"=SUM(D{S0}:D{SL})", f_bold, "money", fill_tot, border=b_tot)
put(su, f"I{SL+1}", f"=SUM(I{S0}:I{SL})", f_bold, "money", fill_tot, border=b_tot)
for m in range(1, NM + 1):
    c = SG.c(m)
    put(su, f"{c}{SL+1}", f"=SUM({c}{S0}:{c}{SL})", f_bold, "money", fill_tot, border=b_tot)
SUT = {}
r = SL + 3
section(su, r, "TOTALS BY ENTITY & TREATMENT", SF + NM - 1)
r += 1
for e in ENTITIES:
    for t in ("Opex", "Capex"):
        put(su, f"C{r}", f"{e} - {t}", f_base)
        put(su, f"I{r}", f'=SUMIFS($I${S0}:$I${SL},$A${S0}:$A${SL},"{e}",$G${S0}:$G${SL},"{t}")', f_base, "money")
        for m in range(1, NM + 1):
            c = SG.c(m)
            put(su, f"{c}{r}", f'=SUMIFS({c}${S0}:{c}${SL},$A${S0}:$A${SL},"{e}",$G${S0}:$G${SL},"{t}")', f_base, "money")
        SUT[(e, t)] = r
        r += 1
su.freeze_panes = su["D9"]
SETUP_CATS = []
for it in items:
    if it[1] not in SETUP_CATS:
        SETUP_CATS.append(it[1])
SETUP_CATS.append("Contingency")

# =====================================================================================
# OPEX REGISTER
ox = ws_new("Opex", "C00000")
title(ox, "RECURRING OPERATING COSTS", "Every recurring non-salary cost by entity and category. Blue = input. Driver column sets how the amount scales (fixed, per FTE, per broker, per investor, % of capital or revenue).")
OF = 12
OG = Grid(ox, OF)
OG.timeline(label_col="C")
for i, h in enumerate(["Entity", "Category", "Item", "Amount", "Driver", "Start month", "End month", "Basis / note", "Ent #", "Drv #"]):
    put(ox, f"{L(i+1)}8", h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True, vertical="center"))
for ci in range(OF - 1, OF + NM):
    ox.cell(row=8, column=ci).fill = fill_hdr
for k, v in zip("ABCDEFGHIJK", [11, 30, 52, 11, 18, 7, 7, 40, 5, 5, 2]):
    ox.column_dimensions[k].width = v
ox.row_dimensions[8].height = 30
O0 = 9
dv2 = DataValidation(type="list", formula1="=DriverList", allow_blank=False)
ox.add_data_validation(dv2)
for i, (ent, cat, item, amt, drv, st, end, basis) in enumerate(OPEX):
    r = O0 + i
    afmt = "pct" if drv.startswith("%") else ("money" if amt >= 10 else "$#,##0.00")
    for j, (v, fm) in enumerate(zip([ent, cat, item, amt, drv, st, end], [None, None, None, afmt, None, "int", "int"])):
        put(ox, f"{L(j+1)}{r}", v, f_input if j >= 3 else f_base, fm)
    if afmt == "pct":
        ox[f"D{r}"].number_format = '0.00%'
    put(ox, f"H{r}", basis, f_sub)
    put(ox, f"I{r}", f"=MATCH(A{r},EntityList,0)", f_base, "int")
    put(ox, f"J{r}", f"=MATCH(E{r},DriverList,0)", f_base, "int")
    dv2.add(f"E{r}")
OL = O0 + len(OPEX) - 1
for r in range(O0, OL + 1):
    for m in range(1, NM + 1):
        c, rc_ = OG.c(m), G.c(m)
        drv = (f"CHOOSE($J{r},1,INDEX(Revenue!{rc_}${R['dFTEbc']}:{rc_}${R['dFTEbh']},$I{r}),Revenue!{rc_}${R['dBrk']},"
               f"Revenue!{rc_}${R['dNewBrk']},Revenue!{rc_}${R['dNewInv']},Revenue!{rc_}${R['dCumInv']},Revenue!{rc_}${R['dCap']},"
               f"INDEX(Revenue!{rc_}${R['dRevbc']}:{rc_}${R['dRevbh']},$I{r}))")
        put(ox, f"{c}{r}", f'=IF(AND({c}$3>=$F{r},{c}$3<=IF($G{r}="",9999,$G{r})),$D{r}*IF($J{r}>=7,1,(1+CostEsc)^({c}$5-1))*OpexFactor*{drv},0)', f_base, "money")
OXT = {}
r = OL + 3
section(ox, r, "TOTALS BY ENTITY & CATEGORY", OF + NM - 1)
r += 1
for e in ENTITIES:
    for cat in CATEGORIES:
        put(ox, f"A{r}", e, f_sub)
        put(ox, f"C{r}", f"{e} - {cat}")
        for m in range(1, NM + 1):
            c = OG.c(m)
            put(ox, f"{c}{r}", f'=SUMIFS({c}${O0}:{c}${OL},$A${O0}:$A${OL},"{e}",$B${O0}:$B${OL},"{cat}")', f_base, "money")
        OXT[(e, cat)] = r
        r += 1
    put(ox, f"C{r}", f"{e} - TOTAL", f_bold, fill=fill_tot)
    for m in range(1, NM + 1):
        c = OG.c(m)
        put(ox, f"{c}{r}", f"=SUM({c}{r-len(CATEGORIES)}:{c}{r-1})", f_bold, "money", fill_tot)
    OXT[(e, "TOTAL")] = r
    r += 2
ox.freeze_panes = ox["D9"]

# =====================================================================================
# ENTITY P&Ls
PL = {}


def link(sheet, grid, rownum):
    return lambda c, p, m: f"={q(sheet)}!{grid.c(m)}{rownum}"


def entity_pl(ent, sheetname, revlines, iclines_in, cogs, icout, rd=False):
    ws = ws_new(sheetname, {"B-Catcher": "1F2A44", "Honey Pot": "E0A526", "BHive": "F4B183"}[ent])
    title(ws, f"{ent.upper()} - MONTHLY P&L (COST CENTRE)", "Entity view including intercompany charges. Eliminated in the consolidated P&L.")
    std_cols(ws)
    g = Grid(ws, 5)
    g.timeline()
    rr = {}
    global row
    row = 8
    sec(g, "REVENUE")
    ext = []
    for k, lab in revlines:
        add(g, f"x{k}", lab, lambda c, p, m, k=k: f"=Revenue!{c}{R[k]}", font=f_link)
        rr[k] = R[f"x{k}"]
        ext.append(row - 1)
    add(g, "extTot", "External revenue", lambda c, p, m, a=ext[0], b=ext[-1]: f"=SUM({c}{a}:{c}{b})", bold=True, border=b_top)
    rr["ext"] = row - 1
    ic = []
    for k, lab in iclines_in:
        add(g, f"i{k}", lab, lambda c, p, m, k=k: f"=Revenue!{c}{R[k]}", font=f_link)
        ic.append(row - 1)
    if ic:
        add(g, "icTot", "Intercompany revenue", lambda c, p, m, a=ic[0], b=ic[-1]: f"=SUM({c}{a}:{c}{b})", border=b_top)
        rr["ic"] = row - 1
    add(g, "revTot", "TOTAL REVENUE", lambda c, p, m: f"={c}{rr['ext']}" + (f"+{c}{rr['ic']}" if ic else ""), bold=True, fill=fill_tot)
    rr["rev"] = row - 1
    sec(g, "COST OF SALES")
    cg = []
    for k, lab in cogs:
        add(g, f"c{k}", lab, lambda c, p, m, k=k: f"=Revenue!{c}{R[k]}", font=f_link)
        cg.append(row - 1)
        rr[k] = row - 1
    add(g, "cogsTot", "Total cost of sales", lambda c, p, m, a=cg[0], b=cg[-1]: f"=SUM({c}{a}:{c}{b})", bold=True, border=b_top)
    rr["cogs"] = row - 1
    add(g, "gp", "GROSS PROFIT", lambda c, p, m: f"={c}{rr['rev']}-{c}{rr['cogs']}", bold=True, fill=fill_tot)
    rr["gp"] = row - 1
    sec(g, "OPERATING EXPENSES")
    add(g, "staff", "Salaries, superannuation & on-costs", lambda c, p, m: f"=Headcount!{HG.c(m)}{HCT[('cost', ent)]}", font=f_link)
    rr["staff"] = row - 1
    add(g, "hire", "Recruitment & onboarding", lambda c, p, m: f"=Headcount!{HG.c(m)}{HCT[('rec', ent)]}", font=f_link)
    rr["hire"] = row - 1
    cat_rows = []
    for cat in CATEGORIES:
        add(g, f"o{cat}", cat, lambda c, p, m, cat=cat: f"=Opex!{OG.c(m)}{OXT[(ent, cat)]}", font=f_link)
        rr[cat] = row - 1
        cat_rows.append(row - 1)
    icx = []
    for k, lab in icout:
        add(g, f"e{k}", lab, lambda c, p, m, k=k: f"=Revenue!{c}{R[k]}", font=f_link)
        icx.append(row - 1)
    rr["icx_rows"] = icx
    add(g, "setup", "Set-up & launch costs (one-off, expensed)", lambda c, p, m: f"={q('Setup Costs')}!{SG.c(m)}{SUT[(ent, 'Opex')]}", font=f_link)
    rr["setup"] = row - 1
    add(g, "cont", "Contingency", lambda c, p, m: f"=ContPct*(SUM({c}{cat_rows[0]}:{c}{cat_rows[-1]})+{c}{rr['setup']})")
    rr["cont"] = row - 1
    add(g, "opexTot", "TOTAL OPERATING EXPENSES", lambda c, p, m: f"=SUM({c}{rr['staff']}:{c}{rr['cont']})", bold=True, fill=fill_tot)
    rr["opex"] = row - 1
    row += 1
    add(g, "ebitda", "EBITDA", lambda c, p, m: f"={c}{rr['gp']}-{c}{rr['opex']}", bold=True, border=b_tot, fill=fill_tot)
    rr["ebitda"] = row - 1
    add(g, "ebitdaM", "EBITDA margin (on external revenue)", lambda c, p, m: f"=IF({c}{rr['ext']}=0,0,{c}{rr['ebitda']}/{c}{rr['ext']})", fmt="pct", total=None)
    row += 1
    if rd:
        add(g, "rd", "Other income - R&D tax incentive (accrued)", lambda c, p, m: f"=RDOn*RDEligible*RDRate*Headcount!{HG.c(m)}{HCT['rd']}")
    else:
        add(g, "rd", "Other income - R&D tax incentive (accrued)", lambda c, p, m: "=0")
    rr["rd"] = row - 1
    add(g, "capex", "Memo: capitalised set-up spend", lambda c, p, m: f"={q('Setup Costs')}!{SG.c(m)}{SUT[(ent, 'Capex')]}", font=f_link)
    rr["capex"] = row - 1
    add(g, "da", "Depreciation & amortisation",
        lambda c, p, m: f"=SUMPRODUCT(($E$3:$BL$3<{c}$3)*($E$3:$BL$3>={c}$3-AmortMonths)*($E${rr['capex']}:$BL${rr['capex']}))/AmortMonths")
    rr["da"] = row - 1
    add(g, "ebit", "EBIT (before interest & tax - taxed at group level)", lambda c, p, m: f"={c}{rr['ebitda']}+{c}{rr['rd']}-{c}{rr['da']}", bold=True, border=b_tot)
    rr["ebit"] = row - 1
    rr["cat_rows"] = cat_rows
    PL[ent] = (sheetname, g, rr)


entity_pl("B-Catcher", "P&L B-Catcher",
          [("bcAcq", "Entry fees - property"), ("bcStruct", "Entry fees - credit & PE"), ("bcMP", "Management fees - property"),
           ("bcMC", "Management fees - private credit"), ("bcME", "Management fees - private equity")],
          [("icSSFhp", "Shared-services fee from Honey Pot"), ("icSSFbh", "Shared-services fee from BHive")],
          [("bcAdmin", "Manager-borne fund administration")],
          [("icLICbc", "BHive platform licence (intercompany)")])
entity_pl("Honey Pot", "P&L Honey Pot",
          [("hpP", "Placement fees - property trusts"), ("hpC", "Placement fees - private credit"),
           ("hpE", "Placement fees - private equity / pre-IPO"), ("hpA", "Placement fees - ASX placements / IPO")],
          [],
          [("hpComm", "Broker commissions")],
          [("icSSFhp", "B-Catcher shared-services fee (intercompany)"), ("icLICbh", "BHive platform licence (intercompany)")])
entity_pl("BHive", "P&L BHive",
          [("bhPlat", "Technology / platform fees"), ("bhOffer", "Offer onboarding fees"), ("bhSec", "Exchange fees"), ("bbGain", "Buy-back resale gains")],
          [("icLICbh", "Platform licence from Honey Pot"), ("icLICbc", "Platform licence from B-Catcher")],
          [("bhPartner", "Market-operator partner share")],
          [("icSSFbh", "B-Catcher shared-services fee (intercompany)")], rd=True)

# =====================================================================================
# CONSOLIDATED P&L
cs = ws_new("P&L Consolidated", "000000")
title(cs, "B-CATCHER GROUP - CONSOLIDATED MONTHLY P&L", "Sum of the three cost centres with intercompany revenue and charges eliminated. Client subscriptions and trust assets are excluded.")
std_cols(cs)
CG = Grid(cs, 5)
CG.timeline()
CR = {}
row = 8


def sumpl(key):
    def f(c, p, m):
        parts = []
        for e in ENTITIES:
            sh, g, rr = PL[e]
            if key in rr:
                parts.append(f"{q(sh)}!{c}{rr[key]}")
        return "=" + "+".join(parts) if parts else "=0"
    return f


sec(CG, "EXTERNAL REVENUE")
s_bc, _, r_bc = PL["B-Catcher"]
s_hp, _, r_hp = PL["Honey Pot"]
s_bh, _, r_bh = PL["BHive"]
lines = [("Honey Pot - placement fees", s_hp, [r_hp[k] for k in ("hpP", "hpC", "hpE", "hpA")]),
         ("B-Catcher - entry fees", s_bc, [r_bc["bcAcq"], r_bc["bcStruct"]]),
         ("B-Catcher - management fees", s_bc, [r_bc[k] for k in ("bcMP", "bcMC", "bcME")]),
         ("BHive - technology / platform fees", s_bh, [r_bh["bhPlat"]]),
         ("BHive - offer onboarding fees", s_bh, [r_bh["bhOffer"]]),
         ("BHive - exchange fees", s_bh, [r_bh["bhSec"]]),
         ("BHive - buy-back resale gains", s_bh, [r_bh["bbGain"]])]
first_rev = row
for lab, sh, rows_ in lines:
    add(CG, lab, lab, lambda c, p, m, sh=sh, rows_=rows_: "=" + "+".join(f"{q(sh)}!{c}{x}" for x in rows_), font=f_link)
CR["rev"] = add(CG, "c_rev", "NET REVENUE", lambda c, p, m: f"=SUM({c}{first_rev}:{c}{row-1})", bold=True, fill=fill_tot)
CR["icrev"] = add(CG, "c_icrev", "Memo: intercompany revenue eliminated", sumpl("ic"), font=f_link)
sec(CG, "COST OF SALES")
a = row
add(CG, "c1", "Broker commissions (Honey Pot)", lambda c, p, m: f"={q(s_hp)}!{c}{r_hp['hpComm']}", font=f_link)
add(CG, "c2", "Market-operator partner share (BHive)", lambda c, p, m: f"={q(s_bh)}!{c}{r_bh['bhPartner']}", font=f_link)
add(CG, "c3", "Manager-borne fund administration (B-Catcher)", lambda c, p, m: f"={q(s_bc)}!{c}{r_bc['bcAdmin']}", font=f_link)
CR["cogs"] = add(CG, "c_cogs", "Total cost of sales", lambda c, p, m: f"=SUM({c}{a}:{c}{row-1})", bold=True, border=b_top)
CR["gp"] = add(CG, "c_gp", "GROSS PROFIT", lambda c, p, m: f"={c}{CR['rev']}-{c}{CR['cogs']}", bold=True, fill=fill_tot)
add(CG, "c_gm", "Gross margin", lambda c, p, m: f"=IF({c}{CR['rev']}=0,0,{c}{CR['gp']}/{c}{CR['rev']})", fmt="pct", total=None)
sec(CG, "OPERATING EXPENSES (external only)")
a = row
CR["staff"] = add(CG, "c_staff", "Salaries, superannuation & on-costs", sumpl("staff"), font=f_link)
CR["hire"] = add(CG, "c_hire", "Recruitment & onboarding", sumpl("hire"), font=f_link)
for cat in CATEGORIES:
    CR[cat] = add(CG, f"c_{cat}", cat, sumpl(cat), font=f_link)
CR["setup"] = add(CG, "c_setup", "Set-up & launch costs (one-off, expensed)", sumpl("setup"), font=f_link)
CR["cont"] = add(CG, "c_cont", "Contingency", sumpl("cont"), font=f_link)
CR["opex"] = add(CG, "c_opex", "TOTAL OPERATING EXPENSES", lambda c, p, m: f"=SUM({c}{a}:{c}{row-1})", bold=True, fill=fill_tot)
row += 1
CR["ebitda"] = add(CG, "c_ebitda", "EBITDA", lambda c, p, m: f"={c}{CR['gp']}-{c}{CR['opex']}", bold=True, border=b_tot, fill=fill_tot)
add(CG, "c_ebitdam", "EBITDA margin", lambda c, p, m: f"=IF({c}{CR['rev']}=0,0,{c}{CR['ebitda']}/{c}{CR['rev']})", fmt="pct", total=None)
CR["chk"] = add(CG, "c_chk", "Check: sum of entity EBITDA less consolidated (should be 0)", lambda c, p, m: "=" + "+".join(f"{q(PL[e][0])}!{c}{PL[e][2]['ebitda']}" for e in ENTITIES) + f"-{c}{CR['ebitda']}", total="sum")
row += 1
CR["rd"] = add(CG, "c_rd", "Other income - R&D tax incentive", sumpl("rd"), font=f_link)
CR["da"] = add(CG, "c_da", "Depreciation & amortisation", sumpl("da"), font=f_link)
CR["ebit"] = add(CG, "c_ebit", "EBIT", lambda c, p, m: f"={c}{CR['ebitda']}+{c}{CR['rd']}-{c}{CR['da']}", bold=True, border=b_top)
CR["int"] = add(CG, "c_int", "Interest income on group cash", lambda c, p, m: "=0")  # filled after cash flow
CR["pbt"] = add(CG, "c_pbt", "PROFIT BEFORE TAX", lambda c, p, m: f"={c}{CR['ebit']}+{c}{CR['int']}", bold=True, fill=fill_tot)
CR["cumpbt"] = add(CG, "c_cumpbt", "Memo: cumulative PBT (tax losses carried forward)", lambda c, p, m: f"={c}{CR['pbt']}" + (f"+{p}{row}" if p else ""), total="last")
CR["cumtax"] = add(CG, "c_cumtax", "Memo: cumulative tax payable", lambda c, p, m: f"=MAX(0,{c}{CR['cumpbt']})*TaxRate", total="last")
CR["tax"] = add(CG, "c_tax", "Income tax expense", lambda c, p, m: f"={c}{CR['cumtax']}" + (f"-{p}{CR['cumtax']}" if p else ""))
CR["npat"] = add(CG, "c_npat", "NET PROFIT AFTER TAX", lambda c, p, m: f"={c}{CR['pbt']}-{c}{CR['tax']}", bold=True, border=b_tot, fill=fill_tot)
sec(CG, "ENTITY EBITDA (incl. intercompany)")
for e in ENTITIES:
    CR[f"eb_{e}"] = add(CG, f"c_eb_{e}", f"EBITDA - {e}", lambda c, p, m, e=e: f"={q(PL[e][0])}!{c}{PL[e][2]['ebitda']}", font=f_link)

# =====================================================================================
# CASH FLOW
cf = ws_new("Cash Flow", "000000")
title(cf, "GROUP CASH FLOW & FUNDING", "Corporate cash only. Investor subscription money flows to issuer/trust/custody accounts and never sits in group cash.")
std_cols(cf)
FG = Grid(cf, 5)
FG.timeline()
FR = {}
row = 8
P = q("P&L Consolidated")
sec(FG, "OPERATING CASH FLOW")
FR["ebitda"] = add(FG, "f_ebitda", "EBITDA", lambda c, p, m: f"={P}!{c}{CR['ebitda']}", font=f_link)
FR["int"] = add(FG, "f_int", "Interest received", lambda c, p, m: f"={P}!{c}{CR['int']}", font=f_link)
FR["tax"] = add(FG, "f_tax", "Income tax paid", lambda c, p, m: f"=-{P}!{c}{CR['tax']}", font=f_link)
FR["rd"] = add(FG, "f_rd", "R&D tax refund received (lagged)", lambda c, p, m: f"=IF({c}$3>RDLag,INDEX({P}!$E${CR['rd']}:$BL${CR['rd']},{c}$3-RDLag),0)")
FR["ocf"] = add(FG, "f_ocf", "Operating cash flow", lambda c, p, m: f"=SUM({c}{FR['ebitda']}:{c}{FR['rd']})", bold=True, border=b_top)
sec(FG, "INVESTING CASH FLOW")
FR["capex"] = add(FG, "f_capex", "Platform build & capitalised set-up", lambda c, p, m: "=-(" + "+".join(f"{q(PL[e][0])}!{c}{PL[e][2]['capex']}" for e in ENTITIES) + ")", font=f_link)
FR["bbwc"] = add(FG, "f_bbwc", "Buy-back units purchased less cost of units resold", lambda c, p, m: f"=-(Revenue!{c}{R['bbCost']}-Revenue!{c}{R['bbSoldCost']})", font=f_link)
sec(FG, "FINANCING CASH FLOW")
FR["eq"] = add(FG, "f_eq", "Equity raised", lambda c, p, m: f"=IF({c}$3=SeedMonth,SeedAmt,0)+IF({c}$3=SeriesAMonth,SeriesAAmt,0)+IF({c}$3=SeriesBMonth,SeriesBAmt,0)", font=f_base)
FR["eqc"] = add(FG, "f_eqc", "Capital-raising costs", lambda c, p, m: f"=-{c}{FR['eq']}*RaiseCostPct")
FR["fcf"] = add(FG, "f_fcf", "Financing cash flow", lambda c, p, m: f"={c}{FR['eq']}+{c}{FR['eqc']}", bold=True, border=b_top)
row += 1
FR["net"] = add(FG, "f_net", "NET CASH FLOW", lambda c, p, m: f"={c}{FR['ocf']}+{c}{FR['capex']}+{c}{FR['bbwc']}+{c}{FR['fcf']}", bold=True, fill=fill_tot)
FR["open"] = add(FG, "f_open", "Opening cash", lambda c, p, m: (f"={p}{row+1}" if p else "=0"), total=None)
FR["close"] = add(FG, "f_close", "CLOSING CASH", lambda c, p, m: f"={c}{FR['open']}+{c}{FR['net']}", bold=True, border=b_tot, fill=fill_tot, total="min")
put(cf, f"B{FR['close']}", "min >", f_sub)
sec(FG, "FUNDING ANALYSIS")
FR["pre"] = add(FG, "f_pre", "Cash flow before financing", lambda c, p, m: f"={c}{FR['ocf']}+{c}{FR['capex']}+{c}{FR['bbwc']}")
FR["cumpre"] = add(FG, "f_cumpre", "Cumulative cash flow before financing", lambda c, p, m: f"={c}{FR['pre']}" + (f"+{p}{row}" if p else ""), total="min")
put(cf, f"B{FR['cumpre']}", "min >", f_sub)
FR["burn"] = add(FG, "f_burn", "Average monthly burn (3-month)", lambda c, p, m: f"=-MIN(0,AVERAGE({FG.c(max(1,m-2))}{FR['pre']}:{c}{FR['pre']}))", total=None)
FR["runway"] = add(FG, "f_runway", "Runway at current burn (months)", lambda c, p, m: f"=IF({c}{FR['burn']}=0,99,MIN(99,{c}{FR['close']}/{c}{FR['burn']}))", fmt="dec", total=None)
FR["req"] = add(FG, "f_req", "Required minimum cash (buffer + buy-back reserve)", lambda c, p, m: f"=MinCash+Revenue!{c}{R['bbRes']}", total="max")
FR["buf"] = add(FG, "f_buf", "Headroom over required minimum cash", lambda c, p, m: f"={c}{FR['close']}-{c}{FR['req']}", total="min")
put(cf, f"B{FR['buf']}", "min >", f_sub)
FR["ebflag"] = add(FG, "f_ebflag", "EBITDA positive (1 = yes)", lambda c, p, m: f"=IF({P}!{c}{CR['ebitda']}>0,1,0)", fmt="int", total=None)
FR["cfflag"] = add(FG, "f_cfflag", "Cash flow before financing positive (1 = yes)", lambda c, p, m: f"=IF({c}{FR['pre']}>0,1,0)", fmt="int", total=None)
# interest back into consolidated P&L
for m in range(1, NM + 1):
    c = CG.c(m)
    put(cs, f"{c}{CR['int']}", f"=MAX(0,{q('Cash Flow')}!{c}{FR['open']})*InterestRate/12", f_link, "money")

# =====================================================================================
# BALANCE SHEET
bs = ws_new("Balance Sheet", "000000")
title(bs, "GROUP BALANCE SHEET (SIMPLIFIED)", "Trust assets and investor money are off balance sheet. GST, receivables and payables assumed settled in month.")
std_cols(bs)
BG = Grid(bs, 5)
BG.timeline()
BR = {}
row = 8
CFq = q("Cash Flow")
sec(BG, "ASSETS")
BR["cash"] = add(BG, "b_cash", "Cash at bank", lambda c, p, m: f"={CFq}!{c}{FR['close']}", font=f_link, total="last")
BR["cap"] = add(BG, "b_cap", "Capitalised platform & systems - cost", lambda c, p, m: f"=-{CFq}!{c}{FR['capex']}" + (f"+{p}{row}" if p else ""), total="last")
BR["acc"] = add(BG, "b_acc", "Less accumulated amortisation", lambda c, p, m: f"=-{P}!{c}{CR['da']}" + (f"+{p}{row}" if p else ""), total="last")
BR["rdr"] = add(BG, "b_rdr", "R&D tax incentive receivable", lambda c, p, m: f"={P}!{c}{CR['rd']}-{CFq}!{c}{FR['rd']}" + (f"+{p}{row}" if p else ""), total="last")
BR["bbi"] = add(BG, "b_bbi", "Buy-back units held for resale", lambda c, p, m: f"=Revenue!{c}{R['bbInv']}", font=f_link, total="last")
BR["ta"] = add(BG, "b_ta", "TOTAL ASSETS", lambda c, p, m: f"=SUM({c}{BR['cash']}:{c}{BR['bbi']})", bold=True, border=b_tot, fill=fill_tot, total="last")
sec(BG, "EQUITY")
BR["sc"] = add(BG, "b_sc", "Share capital (net of raising costs)", lambda c, p, m: f"={CFq}!{c}{FR['fcf']}" + (f"+{p}{row}" if p else ""), total="last")
BR["re"] = add(BG, "b_re", "Retained earnings / (accumulated losses)", lambda c, p, m: f"={P}!{c}{CR['npat']}" + (f"+{p}{row}" if p else ""), total="last")
BR["te"] = add(BG, "b_te", "TOTAL EQUITY", lambda c, p, m: f"={c}{BR['sc']}+{c}{BR['re']}", bold=True, border=b_tot, fill=fill_tot, total="last")
row += 1
BR["chk"] = add(BG, "b_chk", "Balance check: assets less equity (must be 0)", lambda c, p, m: f"=ROUND({c}{BR['ta']}-{c}{BR['te']},2)", total="max")

# =====================================================================================
# ANNUAL SUMMARY
an = ws_new("Annual Summary", GOLD)
title(an, "ANNUAL SUMMARY - FIVE-YEAR FORECAST", "Calendar years (Year 1 = 2027). Flows summed from monthly sheets; balances taken at year end.")
an.column_dimensions["A"].width = 56
for col in "BCDEFG":
    an.column_dimensions[col].width = 15
yrs = ["Year 1\n2027", "Year 2\n2028", "Year 3\n2029", "Year 4\n2030", "Year 5\n2031", "5-yr total"]
for i, h in enumerate(yrs):
    put(an, f"{L(2+i)}4", h, f_hdr, fill=fill_hdr, align=Alignment(horizontal="center", wrap_text=True))
put(an, "A4", "AUD", f_hdr, fill=fill_hdr)
an.row_dimensions[4].height = 28
AR = {}
ar = 5


def an_sec(label):
    global ar
    ar += 1
    section(an, ar, label, 7)
    ar += 1


def an_row(key, label, sheet, srow, kind="flow", fmt="money", bold=False, fill=None, border=None):
    """kind: flow (sum by year) | stock (year-end value) | formula (custom lambda y-> formula)"""
    global ar
    put(an, f"A{ar}", label, f_bold if bold else f_base, fill=fill, border=border)
    for y in range(1, 6):
        col = L(1 + y)
        if kind == "flow":
            f = f"=SUMIF({q(sheet)}!$E$5:$BL$5,{y},{q(sheet)}!$E${srow}:$BL${srow})"
        elif kind == "stock":
            f = f"=INDEX({q(sheet)}!$E${srow}:$BL${srow},{12*y})"
        else:
            f = srow(col, y)
        put(an, f"{col}{ar}", f, f_link if kind in ("flow", "stock") else f_base, fmt, fill, bold=bold, border=border)
    if kind == "flow":
        put(an, f"G{ar}", f"=SUM(B{ar}:F{ar})", f_bold if bold else f_base, fmt, fill, border=border)
    elif fill or border:
        put(an, f"G{ar}", None, fill=fill, border=border)
    AR[key] = ar
    ar += 1


an_sec("VOLUMES (client capital - not group revenue)")
an_row("cap", "Capital placed through BHive", "Revenue", R["cap"])
an_row("aum", "Managed AUM - year end", "Revenue", R["aum"], "stock")
an_row("inv", "Investors on BHive - year end", "Revenue", R["cumInv"], "stock", "num")
an_row("brk", "Active accredited brokers - year end", "Revenue", R["brk"], "stock", "num")
an_row("offers", "Offers launched", "Revenue", R["offers"], "flow", "dec")
an_row("trade", "Secondary market value traded", "Revenue", R["bhTraded"])
an_row("bbnav", "Units bought back under buy-back facility (at valuation)", "Revenue", R["bbNAV"])
an_row("bbinv", "Bought-back units held for resale - year end", "Revenue", R["bbInv"], "stock")
an_row("bbres", "Buy-back liquidity reserve required - year end", "Revenue", R["bbRes"], "stock")
an_sec("REVENUE BY ENTITY (external)")
an_row("rhp", "Honey Pot - distribution", "Revenue", R["hpRev"])
an_row("rbc", "B-Catcher - structuring & management", "Revenue", R["bcRev"])
an_row("rbh", "BHive - platform & marketplace", "Revenue", R["bhRev"])
an_row("rbb", "BHive - buy-back resale gains", "Revenue", R["bbGain"])
an_row("rev", "GROUP NET REVENUE", "P&L Consolidated", CR["rev"], bold=True, fill=fill_tot)
an_sec("CONSOLIDATED P&L")
an_row("cogs", "Cost of sales (broker commissions, partner share, fund admin)", "P&L Consolidated", CR["cogs"])
an_row("gp", "Gross profit", "P&L Consolidated", CR["gp"], bold=True)
an_row("staff", "Staff costs (incl. recruitment)", "P&L Consolidated", CR["staff"])
put(an, f"A{AR['staff']}", "Staff costs (salaries & on-costs)")
an_row("hire", "Recruitment & onboarding", "P&L Consolidated", CR["hire"])
for cat in CATEGORIES:
    an_row(f"cat_{cat}", cat, "P&L Consolidated", CR[cat])
an_row("setup", "Set-up & launch costs (expensed)", "P&L Consolidated", CR["setup"])
an_row("cont", "Contingency", "P&L Consolidated", CR["cont"])
an_row("opex", "Total operating expenses", "P&L Consolidated", CR["opex"], bold=True, border=b_top)
an_row("ebitda", "EBITDA", "P&L Consolidated", CR["ebitda"], bold=True, fill=fill_tot, border=b_tot)
an_row("ebitdam", "EBITDA margin", None, lambda col, y: f"=IF({col}{AR['rev']}=0,0,{col}{AR['ebitda']}/{col}{AR['rev']})", "formula", "pct")
an_row("rd", "R&D tax incentive", "P&L Consolidated", CR["rd"])
an_row("da", "Depreciation & amortisation", "P&L Consolidated", CR["da"])
an_row("int", "Interest income", "P&L Consolidated", CR["int"])
an_row("tax", "Income tax", "P&L Consolidated", CR["tax"])
an_row("npat", "NET PROFIT AFTER TAX", "P&L Consolidated", CR["npat"], bold=True, fill=fill_tot, border=b_tot)
an_sec("ENTITY EBITDA (incl. intercompany charges)")
for e in ENTITIES:
    an_row(f"eb_{e}", f"{e}", "P&L Consolidated", CR[f"eb_{e}"])
an_sec("CASH & FUNDING")
an_row("ocf", "Operating cash flow", "Cash Flow", FR["ocf"])
an_row("capex", "Capital expenditure", "Cash Flow", FR["capex"])
an_row("eq", "Equity raised (net)", "Cash Flow", FR["fcf"])
an_row("cash", "Closing cash - year end", "Cash Flow", FR["close"], "stock", bold=True)
an_sec("PEOPLE")
for e in ENTITIES + ["Group"]:
    an_row(f"fte_{e}", f"FTE - {e if e != 'Group' else 'GROUP'} (year end)", "Headcount", lambda col, y: "=0", "formula", "dec", bold=(e == "Group"))
    for y in range(1, 6):
        an[f"{L(1+y)}{AR[f'fte_{e}']}"].value = f"=INDEX(Headcount!${HG.c1}${HCT[('fte', e)]}:${HG.cN}${HCT[('fte', e)]},{12*y})"
        an[f"{L(1+y)}{AR[f'fte_{e}']}"].font = f_link
an_row("ppl", "People employed or contracted (year end)", "Headcount", lambda col, y: "=0", "formula", "num")
for y in range(1, 6):
    an[f"{L(1+y)}{AR['ppl']}"].value = f"=INDEX(Headcount!${HG.c1}${HCT[('ppl', 'Group')]}:${HG.cN}${HCT[('ppl', 'Group')]},{12*y})"
an_sec("UNIT ECONOMICS")
an_row("yield", "Group revenue / capital placed", None, lambda col, y: f"=IF({col}{AR['cap']}=0,0,{col}{AR['rev']}/{col}{AR['cap']})", "formula", "pct")
an_row("netyield", "Gross profit / capital placed", None, lambda col, y: f"=IF({col}{AR['cap']}=0,0,{col}{AR['gp']}/{col}{AR['cap']})", "formula", "pct")
an_row("revfte", "Revenue per FTE", None, lambda col, y: f"=IF({col}{AR['fte_Group']}=0,0,{col}{AR['rev']}/{col}{AR['fte_Group']})", "formula")
an_row("capbrk", "Capital placed per active broker", None, lambda col, y: f"=IF({col}{AR['brk']}=0,0,{col}{AR['cap']}*BrokerChannel/{col}{AR['brk']})", "formula")
an_row("costcap", "Operating cost per $1 of capital placed", None, lambda col, y: f"=IF({col}{AR['cap']}=0,0,{col}{AR['opex']}/{col}{AR['cap']})", "formula", "$0.000")
an_sec("KEY METRICS")
KM = {}
for key, lab, f, fmt in [
    ("peak", "Peak funding requirement (lowest cumulative pre-financing cash)", f"=-MIN(0,{CFq}!C{FR['cumpre']})", "money"),
    ("raised", "Total equity raised", "=SeedAmt+SeriesAAmt+SeriesBAmt", "money"),
    ("mincash", "Lowest month-end cash balance", f"={CFq}!C{FR['close']}", "money"),
    ("minmonth", "Month of lowest cash", f"=INDEX({CFq}!$E$4:$BL$4,MATCH({CFq}!C{FR['close']},{CFq}!$E${FR['close']}:$BL${FR['close']},0))", "mmm-yy"),
    ("ebm", "First month of positive group EBITDA", f'=IFERROR(INDEX({CFq}!$E$4:$BL$4,MATCH(1,{CFq}!$E${FR["ebflag"]}:$BL${FR["ebflag"]},0)),"Beyond 5 yrs")', "mmm-yy"),
    ("cfm", "First month of positive pre-financing cash flow", f'=IFERROR(INDEX({CFq}!$E$4:$BL$4,MATCH(1,{CFq}!$E${FR["cfflag"]}:$BL${FR["cfflag"]},0)),"Beyond 5 yrs")', "mmm-yy"),
    ("buf", "Lowest headroom over minimum cash buffer", f"={CFq}!C{FR['buf']}", "money"),
    ("fund_ok", "Funding adequate? (cash never below buffer)", f'=IF({CFq}!C{FR["buf"]}>=0,"YES","NO - RAISE MORE")', None),
    ("setup", "Total one-off set-up & build cost", f"={q('Setup Costs')}!I{SL+1}", "money"),
]:
    put(an, f"A{ar}", lab, f_bold)
    put(an, f"B{ar}", f, f_bold, fmt, fill_key if key == "fund_ok" else None)
    KM[key] = ar
    ar += 1
an.freeze_panes = "B5"

# charts on annual summary
ch = BarChart()
ch.type = "col"
ch.title = "Group revenue vs operating expenses ($)"
ch.y_axis.title = None
data = Reference(an, min_col=1, max_col=6, min_row=AR["rev"], max_row=AR["rev"])
ch.add_data(Reference(an, min_col=1, max_col=6, min_row=AR["rev"]), from_rows=True, titles_from_data=True)
ch.add_data(Reference(an, min_col=1, max_col=6, min_row=AR["opex"]), from_rows=True, titles_from_data=True)
ch.set_categories(Reference(an, min_col=2, max_col=6, min_row=4))
ch.height, ch.width = 7.5, 16
ch.roundedCorners = False
ch.legend.position = "b"
ch.legend.overlay = False
ch.title.overlay = False
ch.x_axis.delete = False
ch.y_axis.delete = False
ch.y_axis.numFmt = '$#,##0'
an.add_chart(ch, "I5")

# =====================================================================================
# BUDGET Y1-Y2 (quarterly)
bu = ws_new("Budget Y1-Y2", GOLD)
title(bu, "OPERATING BUDGETS - YEARS 1 & 2 BY QUARTER AND COST CENTRE", "Approved-budget format: each cost centre's revenue, staff, operating categories, set-up and capex by quarter.")
bu.column_dimensions["A"].width = 50
qhdr = [f"Q{(i-1)%4+1} {2027 + (i-1)//4}" for i in range(1, 9)]
cols = [L(2 + i) for i in range(8)] + ["J", "K"]
for i, h in enumerate(qhdr + ["FY2027", "FY2028"]):
    put(bu, f"{cols[i]}4", h, f_hdr, fill=fill_hdr, align=Alignment(horizontal="center"))
    bu.column_dimensions[cols[i]].width = 13
put(bu, "A4", "AUD", f_hdr, fill=fill_hdr)
br = 5


def bu_line(label, sheet, srow, bold=False, fill=None, border=None):
    global br
    put(bu, f"A{br}", label, f_bold if bold else f_base, fill=fill, border=border)
    for i in range(8):
        put(bu, f"{cols[i]}{br}", f"=SUMIF({q(sheet)}!$E$6:$BL$6,{i+1},{q(sheet)}!$E${srow}:$BL${srow})", f_link, "money", fill, bold=bold, border=border)
    put(bu, f"J{br}", f"=SUM(B{br}:E{br})", f_bold, "money", fill, border=border)
    put(bu, f"K{br}", f"=SUM(F{br}:I{br})", f_bold, "money", fill, border=border)
    br += 1


for e in ENTITIES:
    sh, g, rr = PL[e]
    br += 1
    section(bu, br, f"{e.upper()} COST CENTRE", 11)
    br += 1
    bu_line("External revenue", sh, rr["ext"])
    if "ic" in rr:
        bu_line("Intercompany revenue", sh, rr["ic"])
    bu_line("Cost of sales", sh, rr["cogs"])
    bu_line("Gross profit", sh, rr["gp"], bold=True, border=b_top)
    bu_line("Salaries, super & on-costs", sh, rr["staff"])
    bu_line("Recruitment & onboarding", sh, rr["hire"])
    for cat in CATEGORIES:
        bu_line(cat, sh, rr[cat])
    for x in rr["icx_rows"]:
        bu_line(sh and g.ws[f"A{x}"].value, sh, x)
    bu_line("Set-up & launch costs (expensed)", sh, rr["setup"])
    bu_line("Contingency", sh, rr["cont"])
    bu_line("Total operating expenses", sh, rr["opex"], bold=True, border=b_top)
    bu_line("EBITDA", sh, rr["ebitda"], bold=True, fill=fill_tot, border=b_tot)
    bu_line("Capital expenditure (capitalised build)", sh, rr["capex"])
br += 1
section(bu, br, "GROUP (CONSOLIDATED, INTERCOMPANY ELIMINATED)", 11)
br += 1
bu_line("Net revenue", "P&L Consolidated", CR["rev"], bold=True)
bu_line("Gross profit", "P&L Consolidated", CR["gp"])
bu_line("Total operating expenses", "P&L Consolidated", CR["opex"])
bu_line("EBITDA", "P&L Consolidated", CR["ebitda"], bold=True, fill=fill_tot, border=b_tot)
bu_line("Capital expenditure", "Cash Flow", FR["capex"])
bu_line("Net cash flow (incl. equity raised)", "Cash Flow", FR["net"])
br += 1
section(bu, br, "SET-UP & LAUNCH BUDGET BY CATEGORY (opex + capex, all entities)", 11)
br += 1
for cat in SETUP_CATS:
    put(bu, f"A{br}", cat)
    for i in range(8):
        put(bu, f"{cols[i]}{br}", f"=SUMPRODUCT(({q('Setup Costs')}!$B${S0}:$B${SL}=\"{cat}\")*({q('Setup Costs')}!${SG.c1}$6:${SG.cN}$6={i+1})*{q('Setup Costs')}!${SG.c1}${S0}:${SG.cN}${SL})", f_link, "money")
    put(bu, f"J{br}", f"=SUM(B{br}:E{br})", f_bold, "money")
    put(bu, f"K{br}", f"=SUM(F{br}:I{br})", f_bold, "money")
    br += 1
put(bu, f"A{br}", "Total set-up & launch", f_bold, fill=fill_tot, border=b_tot)
for col in cols:
    put(bu, f"{col}{br}", f"=SUM({col}{br-len(SETUP_CATS)}:{col}{br-1})", f_bold, "money", fill_tot, border=b_tot)
bu.freeze_panes = "B5"

# =====================================================================================
# FUNDING & RETURNS
fr = ws_new("Funding & Returns", GOLD)
title(fr, "CORPORATE CAPITAL RAISE, USE OF FUNDS & INVESTOR RETURNS", "Corporate equity only. Not an offer. Valuations and exit multiples are planning assumptions.")
fr.column_dimensions["A"].width = 52
for col in "BCDEF":
    fr.column_dimensions[col].width = 16
r = 4
section(fr, r, "FUNDING REQUIREMENT", 6)
r += 1
for lab, f, fmt in [("Peak funding requirement (from cash flow)", f"='Annual Summary'!B{KM['peak']}", "money"),
                    ("Minimum cash incl. buy-back reserve (peak)", f"={CFq}!C{FR['req']}", "money"),
                    ("Required equity incl. buffer & raise costs", f"=(B{r}+B{r+1})/(1-RaiseCostPct)", "money"),
                    ("Equity planned (Seed + Series A + Series B)", "=SeedAmt+SeriesAAmt+SeriesBAmt", "money"),
                    ("Surplus / (shortfall)", f"=B{r+3}-B{r+2}", "money")]:
    put(fr, f"A{r}", lab, f_bold if "Surplus" in lab else f_base)
    put(fr, f"B{r}", f, f_bold if "Surplus" in lab else f_link, fmt, fill_key if "Surplus" in lab else None)
    r += 1
FUND_SURPLUS = r - 1
r += 1
section(fr, r, "CAP TABLE", 6)
r += 1
for i, h in enumerate(["Round", "Amount", "Pre-money", "Price / share", "New shares", "Post-round %"]):
    put(fr, f"{L(1+i)}{r}", h, f_bold, fill=fill_sec)
r += 1
c0 = r
put(fr, f"A{r}", "Founders & existing holders (B-Catcher + BHive)")
put(fr, f"E{r}", "=FounderShares", f_link, "num")
r += 1
put(fr, f"A{r}", "Group executive share pool (parent ESOP, created before Seed)")
put(fr, f"E{r}", f"=EsopPct*(1+SeedAmt/SeedPre)*E{c0}/(1-EsopPct*(1+SeedAmt/SeedPre))", f_base, "num")
r += 1
put(fr, f"A{r}", "Seed")
put(fr, f"B{r}", "=SeedAmt", f_link, "money")
put(fr, f"C{r}", "=SeedPre", f_link, "money")
put(fr, f"D{r}", f"=C{r}/SUM(E{c0}:E{r-1})", f_base, "$0.000")
put(fr, f"E{r}", f"=B{r}/D{r}", f_base, "num")
r += 1
put(fr, f"A{r}", "Series A")
put(fr, f"B{r}", "=SeriesAAmt", f_link, "money")
put(fr, f"C{r}", "=SeriesAPre", f_link, "money")
put(fr, f"D{r}", f"=C{r}/SUM(E{c0}:E{r-1})", f_base, "$0.000")
put(fr, f"E{r}", f"=B{r}/D{r}", f_base, "num")
r += 1
put(fr, f"A{r}", "Series B (optional)")
put(fr, f"B{r}", "=SeriesBAmt", f_link, "money")
put(fr, f"C{r}", "=SeriesBPre", f_link, "money")
put(fr, f"D{r}", f"=IF(B{r}=0,0,C{r}/SUM(E{c0}:E{r-1}))", f_base, "$0.000")
put(fr, f"E{r}", f"=IF(B{r}=0,0,B{r}/D{r})", f_base, "num")
r += 1
put(fr, f"A{r}", "Total (fully diluted, incl. ESOP pool)", f_bold, border=b_tot)
put(fr, f"B{r}", f"=SUM(B{c0}:B{r-1})", f_bold, "money", border=b_tot)
put(fr, f"E{r}", f"=SUM(E{c0}:E{r-1})", f_bold, "num", border=b_tot)
ctot = r
for rr_ in range(c0, ctot + 1):
    put(fr, f"F{rr_}", f"=E{rr_}/$E${ctot}", f_bold if rr_ == ctot else f_base, "pct", border=b_tot if rr_ == ctot else None)
put(fr, f"A{r+1}", "Post-money valuation after Series A", f_base)
put(fr, f"B{r+1}", f"=SeriesAPre+SeriesAAmt", f_base, "money")
put(fr, f"A{r+2}", "Post-money valuation after Series B (if raised)", f_base)
put(fr, f"B{r+2}", f"=IF(SeriesBAmt=0,0,SeriesBPre+SeriesBAmt)", f_base, "money")
put(fr, f"A{r+3}", "Founder control check: founders' ownership vs 51% floor", f_bold)
put(fr, f"B{r+3}", f"=IF(F{c0}>=FounderFloor,\"OK - founders \"&TEXT(F{c0},\"0.0%\"),\"BREACH - founders \"&TEXT(F{c0},\"0.0%\"))", f_bold)
r += 4
section(fr, r, "USE OF FUNDS - FIRST 24 MONTHS (gross spend)", 6)
r += 1
uf0 = r
CFS = q("Cash Flow")


def sum24(sheet, srow):
    return f"SUMIF({q(sheet)}!$E$3:$BL$3,\"<=24\",{q(sheet)}!$E${srow}:$BL${srow})"


uses = [
    ("Team - salaries, on-costs & hiring (all entities)", f"={sum24('P&L Consolidated', CR['staff'])}+{sum24('P&L Consolidated', CR['hire'])}"),
    ("Platform & technology build (capitalised)", f"=-{sum24('Cash Flow', FR['capex'])}"),
    ("Licensing, legal, structuring & compliance set-up", f"={sum24('P&L Consolidated', CR['setup'])}"),
    ("Technology, hosting & security (recurring)", f"={sum24('P&L Consolidated', CR['Technology, Hosting & Software'])}"),
    ("Marketing, broker network & investor servicing", f"={sum24('P&L Consolidated', CR['Marketing & Investor Acquisition'])}+{sum24('P&L Consolidated', CR['Broker Network & Distribution'])}+{sum24('P&L Consolidated', CR['Investor Servicing & Onboarding'])}"),
    ("Regulatory, insurance, professional fees & governance", "=" + "+".join(sum24('P&L Consolidated', CR[c]) for c in ["Licensing, Regulatory & Compliance", "Insurance", "Professional Fees (Legal, Audit, Tax)", "Board & Governance"])),
    ("Premises, origination, travel & general", "=" + "+".join(sum24('P&L Consolidated', CR[c]) for c in ["Premises & Office", "Investment Origination & Due Diligence", "Travel & General"]) + f"+{sum24('P&L Consolidated', CR['cont'])}"),
]
for lab, f in uses:
    put(fr, f"A{r}", lab)
    put(fr, f"B{r}", f, f_link, "money")
    put(fr, f"C{r}", f"=B{r}/B${uf0+len(uses)}", f_base, "pct")
    r += 1
put(fr, f"A{r}", "Total gross spend (months 1-24)", f_bold, border=b_tot)
put(fr, f"B{r}", f"=SUM(B{uf0}:B{r-1})", f_bold, "money", border=b_tot)
put(fr, f"C{r}", f"=SUM(C{uf0}:C{r-1})", f_bold, "pct", border=b_tot)
r += 1
put(fr, f"A{r}", "Less: gross profit, R&D refunds & interest earned (months 1-24)")
put(fr, f"B{r}", f"=-({sum24('P&L Consolidated', CR['gp'])}+{sum24('Cash Flow', FR['rd'])}+{sum24('Cash Flow', FR['int'])})", f_link, "money")
r += 1
put(fr, f"A{r}", "Net cash consumed (months 1-24)", f_bold)
put(fr, f"B{r}", f"=B{r-2}+B{r-1}", f_bold, "money")
r += 1
put(fr, f"A{r}", "Cash on hand at month 24", f_bold)
put(fr, f"B{r}", f"={CFS}!{FG.c(24)}{FR['close']}", f_link, "money")
r += 2
section(fr, r, "ILLUSTRATIVE INVESTOR RETURNS AT END OF YEAR 5", 6)
r += 1
ret0 = r
for lab, f, fmt in [("Exit run-rate EBITDA (Q4 year 5 x 4)", f"=4*SUMIF({q('P&L Consolidated')}!$E$6:$BL$6,20,{q('P&L Consolidated')}!$E${CR['ebitda']}:$BL${CR['ebitda']})", "money"),
                    ("Exit run-rate net revenue (Q4 year 5 x 4)", f"=4*SUMIF({q('P&L Consolidated')}!$E$6:$BL$6,20,{q('P&L Consolidated')}!$E${CR['rev']}:$BL${CR['rev']})", "money"),
                    ("Enterprise value - EBITDA multiple", f"=MAX(0,B{r})*ExitEBITDAMult", "money"),
                    ("Enterprise value - revenue multiple", f"=B{r+1}*ExitRevMult", "money"),
                    ("Equity value (average of methods + year-5 cash)", f"=AVERAGE(B{r+2},B{r+3})+'Annual Summary'!F{AR['cash']}", "money"),
                    ("Less: key-people minority stakes in Honey Pot & BHive (at EBITDA multiple)", f"=-(HPPool*MAX(0,'Annual Summary'!F{AR['eb_Honey Pot']})+BHPool*MAX(0,'Annual Summary'!F{AR['eb_BHive']}))*ExitEBITDAMult", "money"),
                    ("Equity value to B-Catcher Group shareholders", f"=B{r+4}+B{r+5}", "money")]:
    put(fr, f"A{r}", lab, f_bold if "Equity" in lab else f_base)
    put(fr, f"B{r}", f, f_bold if "Equity" in lab else f_link, fmt)
    r += 1
eqv = r - 1
r += 1
for i, h in enumerate(["Investor", "Invested", "Ownership", "Value at Y5", "MOIC", "IRR (approx.)"]):
    put(fr, f"{L(1+i)}{r}", h, f_bold, fill=fill_sec)
r += 1
for lab, crow, mon in [("Seed investors", c0 + 2, "SeedMonth"), ("Series A investors", c0 + 3, "SeriesAMonth"), ("Series B investors", c0 + 4, "SeriesBMonth")]:
    put(fr, f"A{r}", lab)
    put(fr, f"B{r}", f"=B{crow}", f_base, "money")
    put(fr, f"C{r}", f"=F{crow}", f_base, "pct")
    put(fr, f"D{r}", f"=C{r}*$B${eqv}", f_base, "money")
    put(fr, f"E{r}", f"=IF(B{r}=0,0,D{r}/B{r})", f_bold, "mult")
    put(fr, f"F{r}", f"=IF(E{r}<=0,0,E{r}^(12/(60-{mon}+1))-1)", f_base, "pct")
    r += 1
put(fr, f"A{r+1}", "Illustrative only. Exit value depends on achieving forecasts, market multiples and dilution from later rounds. The 12% ESOP pool is included in the cap table.", f_sub)

# =====================================================================================
# INVESTOR RETURNS (product-level net return to investors after all group fees)
ir = ws_new("Investor Returns", GOLD)
title(ir, "INVESTOR NET RETURNS - what an investor keeps after all fees (illustrative, per $1 invested)",
      "Product-level disclosure. Uses the live fee inputs. Targets, not forecasts; capital at risk. Does not drive the group P&L.")
for col, w in zip("ABCDEF", [58, 17, 17, 17, 17, 70]):
    ir.column_dimensions[col].width = w
hdrs = ["Per $1 of investor capital", "Property - geared", "Property - Debt-Free series", "Private credit", "PE / pre-IPO", "Notes"]
for col, h in zip("ABCDEF", hdrs):
    put(ir, f"{col}4", h, f_hdr, fill=fill_hdr, align=Alignment(wrap_text=True, horizontal="center"))
ir.row_dimensions[4].height = 30
rows = [
    ("Gross asset exposure (GAV per $1 equity)", ["=1/(1-PropLVR)", "=1", "=1", "=1"], "dec2", "Geared trust borrows inside the trust; investor has no personal loan."),
    ("Income: net rent (7% gross less vacancy & outgoings) / interest", ["=(IRPropYield-IRPropOutgoings)*B5", "=(IRPropYield-IRPropOutgoings)*C5", "=IRCreditRate", "=0"], "pct", "Rent is received by the trust and paid to unit holders; it is not group revenue."),
    ("Less: interest on trust debt", ["=-IRPropDebtRate*(B5-1)", "=0", "=0", "=0"], "pct", "Debt-Free series carries no interest cost."),
    ("Less: trust-level costs / expected losses", ["=-IRPropTrustCost*B5", "=-IRPropTrustCost*C5", "=-(IRCreditLoss+IRCreditTrustCost)", "=0"], "pct", ""),
    ("Capital growth / gross return", ["=GrowthProp*B5", "=GrowthProp*C5", "=0", "=IRPEGross"], "pct", "PE is a high-risk target; outcomes range widely, including loss."),
    ("GROSS RETURN BEFORE GROUP FEES", ["=SUM(B6:B9)", "=SUM(C6:C9)", "=SUM(D6:D9)", "=SUM(E6:E9)"], "pct", ""),
    ("Ongoing: B-Catcher management fee", ["=-MgmtProp", "=-MgmtProp", "=-MgmtCredit", "=-MgmtPE"], "pct", "On equity (investor) capital."),
    ("Ongoing: BHive technology fee", ["=-PlatformFee", "=-PlatformFee", "=-PlatformFee", "=-PlatformFee"], "pct", ""),
    ("Up-front: placement + entry + onboarding (total)", ["=-(FeeProp+AcqFee+OfferFee/OfferSize)", "=-(FeeProp+AcqFee+OfferFee/OfferSize)", "=-(FeeCredit+StructFee+OfferFee/OfferSize)", "=-(FeePE+StructFee+OfferFee/OfferSize)"], "pct", "Placement fee funds broker commissions (60%)."),
    ("Up-front fees spread over hold (per year)", ["=B13/IRHold", "=C13/IRHold", "=D13/IRHoldCredit", "=E13/IRHold"], "pct", "Property & PE over IRHold years; credit over IRHoldCredit years."),
    ("NET RETURN TO INVESTOR (per year, approx.)", ["=B10+B11+B12+B14", "=C10+C11+C12+C14", "=D10+D11+D12+D14", "=E10+E11+E12+E14"], "pct", "Before personal tax. Target only - not guaranteed."),
    ("Total group fees per year (ongoing + spread up-front)", ["=-(B11+B12+B14)", "=-(C11+C12+C14)", "=-(D11+D12+D14)", "=-(E11+E12+E14)"], "pct", "Disclose this all-in figure on every offer and in the BHive wallet."),
    ("Share of gross return kept by investor", ["=IF(B10=0,0,B15/B10)", "=IF(C10=0,0,C15/C10)", "=IF(D10=0,0,D15/D10)", "=IF(E10=0,0,E15/E10)"], "pct", ""),
    ("Example: $150,000 invested - net income/return in year 1 ($)", ["=150000*B15", "=150000*C15", "=150000*D15", "=150000*E15"], "money", "Illustrative only."),
]
for i, (lab, fs, fmt, note) in enumerate(rows):
    r_ = 5 + i
    bold = lab.isupper() or lab.startswith("NET")
    put(ir, f"A{r_}", lab, f_bold if bold else f_base, fill=fill_tot if bold else None)
    for col, f in zip("BCDE", fs):
        put(ir, f"{col}{r_}", f, f_bold if bold else f_base, fmt, fill=fill_tot if bold else None)
    put(ir, f"F{r_}", note, f_sub)
put(ir, "A21", "Comparators (approx., for context): Fundrise ~1.0% p.a. all-in; unlisted property syndicates ~1-2% acquisition + 0.6-1.0% of GAV p.a.; private credit managers retain 2-6% margins (ASIC Rep 846).", f_sub)
put(ir, "A22", "Buy-back: if used, units are bought at a discount to valuation (property 15/10/5%, trust units 10/8/5%) - an exit cost, not shown above.", f_sub)
IRROW = {"net": 15, "fees": 16}

# =====================================================================================
# SCENARIOS (placeholder; filled by run_scenarios.py)
sc = ws_new("Scenarios", GOLD)
title(sc, "SCENARIO & STRESS-TEST RESULTS", "Snapshot values produced by running the model under each case. Re-run build to refresh. Live model reflects Inputs!B4 only.")

# =====================================================================================
# COVER
ws = cover
ws.column_dimensions["A"].width = 3
ws.column_dimensions["B"].width = 34
ws.column_dimensions["C"].width = 100
put(ws, "B2", "B-CATCHER GROUP", Font(name=FONT, size=20, bold=True, color=NAVY))
put(ws, "B3", "Integrated Financial Model - Set-up, Headcount, Budgets, P&L, Cash Flow & Forecasts (60 months)", Font(name=FONT, size=11, bold=True, color="595959"))
put(ws, "B4", "B-Catcher (parent: origination & asset management)  |  Honey Pot (distribution & broker network)  |  BHive (investor platform, wallet & secondary market)", f_sub)
put(ws, "B5", "CONFIDENTIAL PLANNING MODEL - all figures are assumptions in AUD, not forecasts of certainty. Not an offer or financial advice.", Font(name=FONT, size=9, bold=True, color="C00000"))
r = 7
section(ws, r, "  KEY OUTPUTS (live - active scenario)", 3)
r += 1
put(ws, f"B{r}", "Active scenario", f_bold)
put(ws, f"C{r}", "=Inputs!B4", f_link)
r += 1
for lab, ref, fmt in [("Year-5 group net revenue", f"='Annual Summary'!F{AR['rev']}", "money"),
                      ("Year-5 group EBITDA", f"='Annual Summary'!F{AR['ebitda']}", "money"),
                      ("5-year capital placed through BHive", f"='Annual Summary'!G{AR['cap']}", "money"),
                      ("Managed AUM at end of year 5", f"='Annual Summary'!F{AR['aum']}", "money"),
                      ("Total one-off set-up & build cost", f"='Annual Summary'!B{KM['setup']}", "money"),
                      ("Peak funding requirement", f"='Annual Summary'!B{KM['peak']}", "money"),
                      ("Equity planned (all rounds)", f"='Annual Summary'!B{KM['raised']}", "money"),
                      ("First month of positive EBITDA", f"='Annual Summary'!B{KM['ebm']}", "mmm-yy"),
                      ("Funding adequate?", f"='Annual Summary'!B{KM['fund_ok']}", None),
                      ("Group FTE at end of year 1 / year 3 / year 5", f"=TEXT('Annual Summary'!B{AR['fte_Group']},\"0.0\")&\" / \"&TEXT('Annual Summary'!D{AR['fte_Group']},\"0.0\")&\" / \"&TEXT('Annual Summary'!F{AR['fte_Group']},\"0.0\")", None),
                      ("Model integrity checks", f"=IF(AND(ABS({q('Balance Sheet')}!C{BR['chk']})<1,ABS({q('P&L Consolidated')}!C{CR['chk']})<1,ABS(MixCheck-1)<0.0001),\"ALL CHECKS OK\",\"CHECK MODEL\")", None)]:
    put(ws, f"B{r}", lab)
    put(ws, f"C{r}", ref, f_link, fmt, align=Alignment(horizontal="left"))
    r += 1
r += 1
section(ws, r, "  HOW THE MODEL FLOWS", 3)
for t in ["Inputs  ->  Revenue & drivers (capital placed, AUM, investors, brokers, fees)  ->  Headcount / Set-up Costs / Opex registers",
          "  ->  P&L B-Catcher, P&L Honey Pot, P&L BHive (cost centres incl. intercompany)  ->  P&L Consolidated (intercompany eliminated)",
          "  ->  Cash Flow (funding, runway)  ->  Balance Sheet (check)  ->  Annual Summary, Budget Y1-Y2, Funding & Returns, Scenarios",
          "Operating model: investors subscribe directly through BHive; Honey Pot manages brokers and distribution behind that process; B-Catcher originates and manages the investments.",
          "Client subscription money goes to issuer / trust / custody accounts. Capital placed and AUM are volume drivers only - they are never group revenue or group cash."]:
    r += 1
    put(ws, f"B{r}", t)
r += 2
section(ws, r, "  SHEET INDEX", 3)
for sh, d in [("Inputs", "Scenario selector, every assumption with Base / Downside / Upside values and source notes, stress levers."),
              ("Revenue", "Monthly capital placed by product, AUM roll-forward, investors, brokers, offers, every fee line by entity, intercompany charges, cost drivers."),
              ("Headcount", f"{len(ROLES)} roles across three entities: start month, FTE, salary, on-costs; monthly FTE, loaded cost, hiring costs, totals by entity."),
              ("Setup Costs", f"{len(SETUP)+1} one-off set-up & build items with timing, capex/opex treatment and basis."),
              ("Opex", f"{len(OPEX)} recurring cost lines with drivers (fixed, per FTE, per broker, per investor, % of capital)."),
              ("P&L B-Catcher / Honey Pot / BHive", "Monthly P&L per cost centre down to EBIT, including intercompany fees."),
              ("P&L Consolidated", "Group P&L, intercompany eliminated, interest, tax with losses carried forward, NPAT."),
              ("Cash Flow", "Operating, investing and financing cash flows; equity rounds; runway and buffer tests."),
              ("Balance Sheet", "Simplified group balance sheet with balance check."),
              ("Annual Summary", "Five-year P&L, volumes, entity EBITDA, cash, people, unit economics and key metrics."),
              ("Budget Y1-Y2", "Quarterly operating budget per cost centre and set-up budget by category."),
              ("Funding & Returns", "Funding requirement, cap table, use of funds, illustrative investor returns."),
              ("Investor Returns", "Net return to investors per product after every group fee (disclosure)."),
              ("Scenarios", "Base / Downside / Upside and stress-test results side by side.")]:
    r += 1
    put(ws, f"B{r}", sh, f_bold)
    put(ws, f"C{r}", d)
r += 2
section(ws, r, "  COLOUR CODE", 3)
for lab, fnt, fill in [("Blue text", f_input, None), ("Black text", f_base, None), ("Green text", f_link, None), ("Yellow fill", f_bold, fill_key)]:
    r += 1
    put(ws, f"B{r}", lab, fnt, fill=fill)
    put(ws, f"C{r}", {"Blue text": "Hard-coded input - change these", "Black text": "Formula - do not overwrite",
                      "Green text": "Link from another sheet", "Yellow fill": "Key selector / result to check"}[lab])
r += 2
section(ws, r, "  STATUS & LIMITATIONS", 3)
for t in ["Planning model for board and investor discussion. Replace estimates with supplier quotes, signed issuer pipeline and legal fee proposals before external use.",
          "Regulated activities (AFSL, market operation, custody, credit) require external legal advice and authorisation before launch; costs here assume the partner/CAR routes in the v3 plan.",
          "Historic BHive projections ($731m AUM, $24.5m revenue) are NOT relied on. Historic B-Catcher mining, lottery and crypto-treasury activities are excluded from this group plan.",
          "Irrecoverable GST is included via the cost factors (+6% opex, +5% set-up). Excluded: GST timing, debtor/creditor days, share-based payment expense (non-cash), performance fees, trust-level costs borne by investment vehicles, retail/lending/card products."]:
    r += 1
    put(ws, f"B{r}", t)
    ws[f"B{r}"].alignment = Alignment(wrap_text=False)

# sheet order
order = ["Cover", "Inputs", "Annual Summary", "Scenarios", "Investor Returns", "Funding & Returns", "Budget Y1-Y2", "P&L Consolidated",
         "P&L B-Catcher", "P&L Honey Pot", "P&L BHive", "Cash Flow", "Balance Sheet", "Revenue", "Headcount", "Setup Costs", "Opex"]
wb._sheets = [wb[s] for s in order]
wb.active = 0

# export key row refs for scenario runner
import json
json.dump({"AR": AR, "KM": KM, "FUND_SURPLUS": FUND_SURPLUS, "INPUT_CELL": INPUT_CELL, "eqv": eqv,
           "ret_rows": [eqv + 3, eqv + 4, eqv + 5]}, open(OUT + ".refs.json", "w"))
wb.save(OUT)
print("saved", OUT)
