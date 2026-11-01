from pathlib import Path
import openpyxl,os,datetime
from openpyxl.chart import LineChart,BarChart,Reference
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L
w=openpyxl.load_workbook('model.xlsx');e=w['Analysis_engine'];src=Path('build_model.py').read_text();exec(src[src.index('def case('):src.index("base=case(")])
def sh(n,h):s=w.create_sheet(n);s.append(h);return s
# Bounds remain feasible when quantities are edited downward.
for r in range(2,123):w['Optimization'].cell(r,4,f'=IF(AND(A{r}>=Required_Social_Quota*Total_Rooms,A{r}<=Total_Rooms),1,0)')
w['Model_checks']['B15']='=AND(B2:B12)'
s=sh('Financing_sensitivity',['Debt rate',.3,.4,.5,.6,.7])
for r,rate in enumerate([.03,.04,.045,.05,.06,.07],2):
 s.cell(r,1,rate)
 for col in range(2,7):
  letter=L(col);formula=f'=-Initial_Investment*(1-{letter}$1)'
  for year in range(1,21):formula+=f'+(Cash_flow!{L(year+2)}19+PMT($A{r},Debt_Term,Initial_Investment*{letter}$1))/(1+Equity_Discount)^{year}'
  s.cell(r,col,formula)
s['A10']='Equity NPV heatmap: debt term fixed 20 years; project NPV is unaffected by financing in this pre-tax unlevered model.'
a=sh('Affordability_path',['Measure']+list(range(1,21)))
for label in ['Market total EUR/month','Social total EUR/month','Market policy ceiling','Social policy ceiling','Social policy pass','Market policy pass']:a.append([label])
for year in range(1,21):
 col=L(year+1)
 for r,f in [(2,'(Market_Base_Rent+Monthly_Charges)'),(3,'(Social_Base_Rent+Monthly_Charges)'),(4,'Market_Total_Ceiling'),(5,'Social_Total_Ceiling')]:a.cell(r,year+1,f'={f}*(1+Rent_Growth)^({col}$1-1)')
 a.cell(6,year+1,f'={col}3<={col}5');a.cell(7,year+1,f'={col}2<={col}4')
a['A10']='Policy ceilings grow at the same assumed rate as total housing cost. Actual lawful indexation and eligibility must be checked annually; not asserted as legal permission.'
sc=sh('Scenarios',['Scenario','Acquisition EUR','Works factor','Market occupancy','Project NPV EUR','Probabilities'])
for name,acq,cost,occ in [('Downside',5000000,1.2,.85),('Base',4500000,1,.95),('Upside',4000000,.9,.98)]:
 rr=case(name,acq=acq,cost=cost,occ=occ);sc.append([name,acq,cost,occ,f'=Analysis_engine!L{rr}','Not assigned: no evidence-based probabilities'])
risk=sh('Risk_register',['Gap / risk','Consequence','Mitigation','Status'])
for row in [('Parcel and land quotes unavailable','Location ranking may reverse','Commission comparable parcel appraisal','Unverified'),('Quota / permit unknown','Capacity, social share and finance may change','Municipal pre-application review and university agreement','Hold'),('VAT and tax structure unknown','Upfront cost/after-tax returns may differ','Written transaction-specific tax advice','Unverified'),('Demand and occupancy assumptions','Receipts and exit NOI overstated','Comparable achieved rents and pre-letting','Assumption'),('Terminal cap rate','Positive base NPV depends on sale','Independent appraisal and no-sale comparison','Assumption')]:risk.append(row)
for name,sheet,cell in [('Project_IRR','Results','B3'),('Equity_NPV','Results','B4'),('Equity_IRR','Results','B5'),('Simple_Payback','Results','B6'),('Discounted_Payback','Results','B7'),('Maximum_Acquisition','Results','B14')]:w.defined_names.add(DefinedName(name,attr_text=f"'{sheet}'!${cell[0]}${cell[1:]}"))
cf=w['Cash_flow'];ch=LineChart();ch.title='Cumulative discounted project cash flow (EUR): sale dependency';ch.add_data(Reference(cf,min_col=1,max_col=22,min_row=24,max_row=25),from_rows=True,titles_from_data=True);ch.set_categories(Reference(cf,min_col=2,max_col=22,min_row=1,max_row=1));cf.add_chart(ch,'K33')
for sheet in [s,a,sc,risk]:sheet.freeze_panes='B2';sheet.column_dimensions['A'].width=38;sheet.column_dimensions['B'].width=28
w.save('.model.extend.xlsx');os.replace('.model.extend.xlsx','model.xlsx')
with open('activity.log','a') as f:f.write(datetime.datetime.now(datetime.timezone.utc).isoformat()+' Enhanced MarieLaure model with commune land-driven acquisition/NPV comparison, current guidance publication source, debt-rate/leverage equity heatmap, 20-year affordability path, three scenarios, sale-dependence chart, risk register and additional named outputs.\n')
