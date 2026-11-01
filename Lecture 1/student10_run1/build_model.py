from pathlib import Path
import os, math, json
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter as L
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.chart import LineChart, Reference
from openpyxl.workbook.properties import CalcProperties
ROOT=Path(__file__).resolve().parent
wb=Workbook(); wb.remove(wb.active)
wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
read=wb.create_sheet('Start'); inp=wb.create_sheet('Inputs'); dash=wb.create_sheet('Dashboard')
read.append(['STUDENT HOUSING | IXELLES / ELSENE','20-year investment appraisal'])
for r in [
 ['Prepared for','Jennifer, Sara, Madiha and team'],
 ['Purpose','Illustrative feasibility model; replace blue assumptions with property-specific evidence.'],
 ['Use','Inputs → scenarios → Dashboard. All cash flows in nominal EUR; annual end-year convention.'],
 ['Horizon','Year 0 acquisition plus years 1–20 including development; sale at end of year 20.'],
 ['Affordability','Rent is capped by an editable policy ceiling; utilities shown separately. Ceiling is a project assumption, not a statutory rent limit.'],
 ['Status','Pre-tax, unlevered property returns. Debt KPIs are indicative; equity cash flows include debt repayment on sale.'],
 ['Timing','Acquisition in year 0; construction equally spread over development years; rent begins the following year.'],
 ['Limitations','No corporate/personal income tax, VAT recovery, depreciation, subsidies, refinancing or monthly draw schedules. Land/build cost include irrecoverable VAT.'],
 ['Decision','Do not commit until planning, housing, fire, energy, tenancy and affordability checks are cleared.'],
 ['Sources','Regulations and evidence are documented on Due diligence and Sources; unknown checks remain open.']]: read.append(r)
rows=[
('Rooms',30,30,30,'beds','Illustrative; verify permitted capacity'),
('Acquisition price',1800000,1800000,1800000,'EUR','Illustrative property cost'),
('Acquisition costs',.15,.15,.15,'ratio','Allowance incl. duties/notary/advice; not a tax determination'),
('Base construction budget',1500000,1500000,1500000,'EUR','Gross including irrecoverable VAT; quote required'),
('Construction overrun',.25,.1,0,'ratio','Stress assumption'),
('Professional fees / works',.15,.12,.1,'ratio','Architect, engineering and permits allowance'),
('Development years',3,2,1,'years','Full non-operating years after acquisition'),
('Requested monthly room rent',550,550,550,'EUR/month','Illustrative target rent excluding utilities'),
('Affordable monthly rent ceiling',550,550,550,'EUR/month','Owner policy; validate against student budgets'),
('Rent annual growth',.01,.02,.025,'ratio','Only if legally allowed; lease/PEB restrictions must be checked'),
('Ceiling annual growth',.01,.02,.02,'ratio','Affordability policy assumption'),
('Occupancy',.85,.94,.98,'ratio','Economic occupancy; 12 billed months assumption'),
('Bad debt / gross occupied rent',.03,.015,.005,'ratio','Collection losses beyond vacancy'),
('Fixed owner operating costs',45000,40000,36000,'EUR/year','First operating-year tax, insurance, admin, nonrecoverable utilities'),
('Variable operating costs / revenue',.2,.17,.15,'ratio','Management, routine repairs and turnover'),
('Cost inflation',.035,.025,.02,'ratio','All owner costs grow from first operating year'),
('Annual capital reserve per room',600,500,400,'EUR/bed/year','Treated as annual cash expenditure; no reserve balance recovered'),
('Exit capitalization rate',.065,.055,.0475,'ratio','Stress yield on year-21 stabilized NOI'),
('Sale costs',.03,.025,.02,'ratio','Selling costs on gross asset value'),
('Unlevered discount rate',.085,.075,.065,'ratio','Nominal required return; project risk assumption'),
('Debt / total development cost',.55,.55,.55,'ratio','Debt drawn at opening; equity bridges all development costs'),
('Debt interest rate',.055,.045,.04,'ratio','Fixed illustrative rate'),
('Debt amortization term',20,20,20,'years','Annual annuity; first payment in first operating year'),
('Student monthly budget',1100,1100,1100,'EUR/month','Illustrative available resources, not researched evidence'),
('Monthly utilities paid by student',100,90,80,'EUR/month','First operating-year allowance'),
('Maximum housing budget share',.6,.6,.6,'ratio','Owner policy; rent plus utilities / student resources')]
inp.append(['ASSUMPTION','Bad','Intermediate','Optimal','Unit','Basis / interpretation'])
for r in rows: inp.append(r)
for r in range(2,len(rows)+2):
 for c in range(2,5):
  inp.cell(r,c).font=Font(color='0070C0'); inp.cell(r,c).number_format='0.0%' if rows[r-2][4]=='ratio' else '#,##0.00'
  dv=DataValidation(type='decimal',operator='greaterThanOrEqual',formula1=0); inp.add_data_validation(dv); dv.add(inp.cell(r,c))
# input row helper
idx={r[0]:i+2 for i,r in enumerate(rows)}
scenario_names=['Bad','Intermediate','Optimal']
for s,sc in enumerate(scenario_names,2):
 ws=wb.create_sheet(sc); C=L(s)
 def p(name):return f"Inputs!${C}${idx[name]}"
 ws.append(['YEAR','Build spend','Collected rent','Owner costs','NOI','Capital reserve','Exit proceeds','Project cash flow','Discounted CF','Cumulative CF','Cumulative discounted','Debt draw','Debt service','Closing debt','Equity cash flow','DSCR'])
 for y in range(21):
  r=y+2; ws.cell(r,1,y)
  formulas={
  2:f'=IF(A{r}=0,{p("Acquisition price")}*(1+{p("Acquisition costs")}),IF(A{r}<={p("Development years")},{p("Base construction budget")}*(1+{p("Construction overrun")})*(1+{p("Professional fees / works")})/{p("Development years")},0))',
  3:f'=IF(A{r}>{p("Development years")},{p("Rooms")}*12*MIN({p("Requested monthly room rent")}*(1+{p("Rent annual growth")})^(A{r}-{p("Development years")}-1),{p("Affordable monthly rent ceiling")}*(1+{p("Ceiling annual growth")})^(A{r}-{p("Development years")}-1))*{p("Occupancy")}*(1-{p("Bad debt / gross occupied rent")}),0)',
  4:f'=IF(C{r}>0,{p("Fixed owner operating costs")}*(1+{p("Cost inflation")})^(A{r}-{p("Development years")}-1)+C{r}*{p("Variable operating costs / revenue")},0)',
  5:f'=C{r}-D{r}',
  6:f'=IF(C{r}>0,{p("Rooms")}*{p("Annual capital reserve per room")}*(1+{p("Cost inflation")})^(A{r}-{p("Development years")}-1),0)',
  7:f'=IF(A{r}=20,MAX(0,({p("Rooms")}*12*MIN({p("Requested monthly room rent")}*(1+{p("Rent annual growth")})^(20-{p("Development years")}),{p("Affordable monthly rent ceiling")}*(1+{p("Ceiling annual growth")})^(20-{p("Development years")}))*{p("Occupancy")}*(1-{p("Bad debt / gross occupied rent")}))*(1-{p("Variable operating costs / revenue")})-{p("Fixed owner operating costs")}*(1+{p("Cost inflation")})^(20-{p("Development years")}))/{p("Exit capitalization rate")}*(1-{p("Sale costs")}),0)',
  8:f'=E{r}-F{r}-B{r}+G{r}',
  9:f'=H{r}/(1+{p("Unlevered discount rate")})^A{r}',
  10:f'=SUM(H$2:H{r})',11:f'=SUM(I$2:I{r})',
  12:f'=IF(A{r}={p("Development years")},SUM(B$2:B$22)*{p("Debt / total development cost")},0)',
  13:f'=IF(AND(A{r}>{p("Development years")},A{r}<={p("Development years")}+{p("Debt amortization term")}),-PMT({p("Debt interest rate")},{p("Debt amortization term")},SUM(B$2:B$22)*{p("Debt / total development cost")}),0)',
  14: '=L2' if y==0 else f'=MAX(0,N{r-1}*(1+IF(M{r}>0,{p("Debt interest rate")},0))+L{r}-M{r})',
  15:f'=H{r}+L{r}-M{r}-IF(A{r}=20,N{r},0)',
  16:f'=IF(M{r}>0,(E{r}-F{r})/M{r},0)'}
  for c,f in formulas.items():ws.cell(r,c,f).number_format='#,##0.00;[Red](#,##0.00)'
 ws['A25']='Payback uses first nonnegative cumulative end-year; includes sale at year 20 if needed.'
 ws['A26']='Debt drawn at completion; bridging interest and loan fees excluded. Equity IRR is indicative.'
 chart=LineChart();chart.title='Project cumulative cash flow';chart.add_data(Reference(ws,min_col=10,max_col=11,min_row=1,max_row=22),titles_from_data=True);chart.set_categories(Reference(ws,min_col=1,min_row=2,max_row=22));ws.add_chart(chart,'A29')
dash.append(['INVESTMENT KPI','Bad','Intermediate','Optimal','Interpretation'])
metrics=[('Project NPV','SUM(\'{sc}\'!I2:I22)','Discounted unlevered cash flows, including exit'),('Project IRR','IFERROR(IRR(\'{sc}\'!H2:H22),"No IRR")','Annual unlevered return'),('Payback, years','IF(COUNTIF(\'{sc}\'!J3:J22,">=0")=0,"Beyond 20",MATCH(TRUE,INDEX(\'{sc}\'!J3:J22>=0,0),0))','First full year recovered; includes sale'),('Discounted payback, years','IF(COUNTIF(\'{sc}\'!K3:K22,">=0")=0,"Beyond 20",MATCH(TRUE,INDEX(\'{sc}\'!K3:K22>=0,0),0))','At scenario discount rate; includes sale'),('Total development cost','SUM(\'{sc}\'!B2:B22)','Acquisition plus works, overrun and fees'),('Cost per bed','B6/Inputs!B2','Total development cost / rooms'),('First operating NOI','INDEX(\'{sc}\'!E2:E22,Inputs!B8+2)','Before capital reserve and financing'),('Initial yield on cost','B8/B6','First operating NOI / cost'),('Minimum operating DSCR','MIN(\'{sc}\'!P{first}:P22)','NOI less capital reserve / debt service'),('Equity IRR','IFERROR(IRR(\'{sc}\'!O2:O22),"No IRR")','Includes loan draw, service and exit repayment; no bridge costs'),('Exit value net of selling costs','\'{sc}\'!G22','Capitalized year-21 NOI'),('Exit share of PV inflows','\'{sc}\'!G22/(1+Inputs!B21)^20/SUMPRODUCT((\'{sc}\'!E2:E22-\'{sc}\'!F2:F22+\'{sc}\'!G2:G22)/(1+Inputs!B21)^\'{sc}\'!A2:A22)','High share signals dependence on exit value'),('Student housing budget share','(MIN(Inputs!B9,Inputs!B10)+Inputs!B26)/Inputs!B25','First operating-year rent plus utilities / assumed budget'),('Affordability policy','IF(B14<=Inputs!B27,"Pass","Fail")','Budget share ceiling; not proof of legal rent compliance'),('Operating break-even occupancy','(Inputs!B15+Inputs!B2*Inputs!B18)/(Inputs!B2*12*MIN(Inputs!B9,Inputs!B10)*(1-Inputs!B14)*(1-Inputs!B16))','First operating year; includes capital reserve, excludes debt'),('Development / operating years','Inputs!B8&" / "&(20-Inputs!B8)','20-year total excludes year 0'),('Investment screen','IF(AND(B2>0,ISNUMBER(B3),B3>Inputs!B21,B10>=1.2,B15="Pass"),"Financial screen passes","Review / reject")','Illustrative thresholds; ALL regulatory checks still required')]
for i,(name,formula,note) in enumerate(metrics,2):
 dash.cell(i,1,name);dash.cell(i,5,note)
 for j,sc in enumerate(scenario_names,2):
  col=L(j); f=formula.format(sc=sc,first=rows[6][j-1]+3)
  if j!=2:
   import re
   f=re.sub(r'(?<![A-Za-z])B(?=\d)',col,f)
  dash.cell(i,j,'='+f).number_format='0.0%' if i in [3,9,11,13,14,16] else '#,##0.00'
# Validation of scenario inputs
for c in ['B','C','D']:
 dv=DataValidation(type='whole',operator='between',formula1=1,formula2=19);inp.add_data_validation(dv);dv.add(inp[f'{c}8'])
for title,headers in [('Due diligence',['Area','Requirement / question','Evidence needed','Status','Impact']),('Sources',['Topic','URL','Verified date','Finding / limitation']),('Prompt',['REUSABLE PROJECT PROMPT','Instructions'])]:
 ws=wb.create_sheet(title);ws.append(headers)
prompt='''Act as a real estate investment analyst specializing in affordable student housing in Ixelles/Elsene, Brussels. Build a transparent, editable 20-year nominal EUR model with Bad, Intermediate and Optimal scenarios. Ask for property address, title, permitted use, rooms and areas, acquisition terms, gross works quotes, student rent/budget evidence, operating costs, financing, legal entity and tax treatment. Clearly label any missing data and illustrative assumptions. Verify current rules using official Brussels and Ixelles sources; cite dated links, distinguish mandatory obligations from voluntary labels, and never assume a permit or compliance. Cover change of use, planning, housing quality, fire safety, energy/PEB, accessibility, leases, rent reference/excessive rent controls, indexation, charges, deposits, registration and taxation. Apply an explicit affordable rent ceiling and test rent plus utilities against student resources. Include construction delays, cost overruns, vacancy, bad debt, maintenance, capex and exit yield risk. Calculate annual project and equity cash flows, NPV, IRR, simple and discounted payback, yield on cost, cost per bed, DSCR, operating break-even occupancy, affordability share and dependence on terminal value. Define timing, discount rates, tax/VAT assumptions, financing and terminal value; show operating-only versus sale-assisted payback. Add rent/works-cost/exit-yield sensitivities, qualitative risks, evidence gaps and go/no-go conditions. Explain market complexity including student demand, seasonality, substitutes, neighbourhood fit, transport, university proximity and supply pipeline. Do not improve returns by violating affordability or regulatory requirements. Deliver a usable workbook with traceable formulas, sources and limitations; state whether results justify proceeding to property-specific due diligence.'''
wb['Prompt'].append(['Copy and reuse',prompt])
(ROOT/'project_prompt.txt').write_text(prompt)
for ws in wb:
 ws.freeze_panes='B2';ws.auto_filter.ref=ws.dimensions;ws.sheet_view.showGridLines=False
 ws.row_dimensions[1].height=32
 for cell in ws[1]:cell.fill=PatternFill('solid',fgColor='16324F');cell.font=Font(color='FFFFFF',bold=True);cell.alignment=Alignment(wrap_text=True)
 for row in ws.iter_rows(min_row=2):
  for cell in row:
   if cell.data_type=='f':cell.font=Font(color='183D30')
   cell.alignment=Alignment(vertical='top',wrap_text=True)
 for c in range(1,ws.max_column+1):ws.column_dimensions[L(c)].width=21 if c>1 else 39
 if ws.title in ['Start','Prompt']:ws.column_dimensions['B'].width=110
 if ws.title=='Inputs':ws.column_dimensions['F'].width=70
 if ws.title=='Dashboard':ws.column_dimensions['E'].width=65
 for r in range(2,ws.max_row+1):
  if ws.title in ['Start','Inputs','Dashboard']:ws.row_dimensions[r].height=34
wb['Prompt'].row_dimensions[2].height=340
path=ROOT/'model.tmp.xlsx';wb.save(path);os.replace(path,ROOT/'model.xlsx')
print('Saved model.xlsx', (ROOT/'model.xlsx').stat().st_size,'bytes; SAVE_NOW=',(ROOT/'SAVE_NOW.txt').exists())
