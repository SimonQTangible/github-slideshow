import sys, openpyxl
wb=openpyxl.load_workbook(sys.argv[1],data_only=True)
ws=wb['Annual Summary']
for row in ws.iter_rows(min_row=4,max_row=ws.max_row):
    a=row[0].value
    if a is None: continue
    vals=[]
    for c in row[1:7]:
        v=c.value
        if isinstance(v,(int,float)): vals.append(f"{v:>14,.2f}" if abs(v)<10 else f"{v:>14,.0f}")
        elif v is None: vals.append(" "*14)
        else: vals.append(f"{str(v):>14}")
    print(f"{str(a)[:55]:55}"+"".join(vals))
ws=wb['Cover']
for r in range(8,21): print(ws[f'B{r}'].value, '|', ws[f'C{r}'].value)
ws=wb['Funding & Returns']
for r in range(4,50):
    v=[ws.cell(r,c).value for c in range(1,7)]
    if any(x is not None for x in v): print(v)
