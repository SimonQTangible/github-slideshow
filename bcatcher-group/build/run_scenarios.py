"""Run the model under each scenario / stress case and write a results table into the Scenarios sheet."""
import json, subprocess, sys, shutil, os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

MODEL = sys.argv[1]
RECALC = "/root/.claude/skills/synced/e3a10683-0f6f-4df7-9d97-1f3de7e0ea41_a12ffef0-b97c-4fce-adde-69bad4852925/xlsx/scripts/recalc.py"
refs = json.load(open(MODEL + ".refs.json"))
AR, KM, IC = refs["AR"], refs["KM"], refs["INPUT_CELL"]

CASES = [
    ("Base case", "Base", {}),
    ("Downside", "Downside", {}),
    ("Upside", "Upside", {}),
    ("Stress: 50% placement shortfall", "Base", {"PlacementMult": 0.5}),
    ("Stress: 6-month licensing delay", "Base", {"LaunchDelay": 6}),
    ("Stress: 6-month broker ramp delay", "Base", {"BrokerDelay": 6}),
    ("Stress: 25% technology build overrun", "Base", {"DevOverrun": 0.25}),
    ("Combined: Downside + 6-month delay", "Downside", {"LaunchDelay": 6}),
    ("Downside + mitigation (defer post-launch hires 18 mths)", "Downside", {"HireDeferral": 18}),
]
METRICS = [
    ("5-yr capital placed", ("AR", "cap", "G"), "money"),
    ("Managed AUM end Y5", ("AR", "aum", "F"), "money"),
    ("Net revenue Y3", ("AR", "rev", "D"), "money"),
    ("Net revenue Y5", ("AR", "rev", "F"), "money"),
    ("EBITDA Y3", ("AR", "ebitda", "D"), "money"),
    ("EBITDA Y5", ("AR", "ebitda", "F"), "money"),
    ("Cumulative 5-yr NPAT", ("AR", "npat", "G"), "money"),
    ("Peak funding requirement", ("KM", "peak", "B"), "money"),
    ("Lowest month-end cash", ("KM", "mincash", "B"), "money"),
    ("Extra equity needed to hold buffer", None, "money"),
    ("First month EBITDA positive", ("KM", "ebm", "B"), "mmm-yy"),
    ("Group FTE end Y5", ("AR", "fte_Group", "F"), "#,##0.0"),
]

tmpdir = os.path.dirname(os.path.abspath(MODEL))
results = []
for label, scen, levers in CASES:
    wb = openpyxl.load_workbook(MODEL)
    ws = wb["Inputs"]
    ws["B4"] = scen
    for k in ("PlacementMult", "LaunchDelay", "BrokerDelay", "DevOverrun", "HireDeferral"):
        ws[f"B{IC[k]}"] = levers.get(k, {"PlacementMult": 1.0}.get(k, 0))
    tmp = os.path.join(tmpdir, "_case.xlsx")
    wb.save(tmp)
    out = subprocess.run([sys.executable, RECALC, tmp, "240"], capture_output=True, text=True).stdout
    assert '"total_errors": 0' in out, out
    v = openpyxl.load_workbook(tmp, data_only=True)
    an = v["Annual Summary"]
    row = []
    for mlabel, ref, fmt in METRICS:
        if ref is None:
            row.append(None)
            continue
        kind, key, col = ref
        r = (AR if kind == "AR" else KM)[key]
        row.append(an[f"{col}{r}"].value)
    lowest = row[8]
    row[9] = max(0.0, 750000 - lowest) if isinstance(lowest, (int, float)) else None
    results.append((label, row))
    print(label, [round(x) if isinstance(x, float) else x for x in row])
os.remove(tmp)

# write table into model
wb = openpyxl.load_workbook(MODEL)
sc = wb["Scenarios"]
F = "Arial"
hdr = Font(name=F, size=9, bold=True, color="FFFFFF")
fill = PatternFill("solid", fgColor="1F2A44")
sc.column_dimensions["A"].width = 38
sc["A4"] = "Metric"
sc["A4"].font = hdr
sc["A4"].fill = fill
for j, (label, _) in enumerate(results):
    c = sc.cell(row=4, column=2 + j, value=label)
    c.font = hdr
    c.fill = fill
    c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    sc.column_dimensions[openpyxl.utils.get_column_letter(2 + j)].width = 17
sc.row_dimensions[4].height = 42
for i, (mlabel, ref, fmt) in enumerate(METRICS):
    r = 5 + i
    sc.cell(row=r, column=1, value=mlabel).font = Font(name=F, size=9, bold=True)
    for j, (_, vals) in enumerate(results):
        v = vals[i]
        c = sc.cell(row=r, column=2 + j, value=v)
        c.font = Font(name=F, size=9)
        c.number_format = {"money": '$#,##0;($#,##0);"-"'}.get(fmt, fmt)
        if mlabel.startswith("Extra equity") and isinstance(v, (int, float)) and v > 0:
            c.font = Font(name=F, size=9, bold=True, color="C00000")
r = 5 + len(METRICS) + 1
notes = [
    "Values are a snapshot produced by build_all.py, which runs the full model once per case (recalculated in LibreOffice). They do not change when you edit Inputs.",
    "To explore live, change Inputs!B4 (scenario) and the stress levers on the Inputs sheet; Cover and Annual Summary update immediately.",
    "Extra equity needed = amount by which the lowest month-end cash falls below the $750k minimum buffer with the planned Seed + Series A + Series B.",
    "Downside: lower placements, fees and broker growth, later launch and secondary market, 15% set-up overrun, 10% opex overrun, no R&D claim.",
]
for n in notes:
    sc.cell(row=r, column=1, value=n).font = Font(name=F, size=9, italic=True, color="595959")
    r += 1
wb.save(MODEL)
json.dump(results, open(MODEL + ".scenarios.json", "w"), default=str)
print("written")
