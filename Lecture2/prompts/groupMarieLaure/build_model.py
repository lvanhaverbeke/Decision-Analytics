from pathlib import Path
import os,json,time,datetime,random,math,shutil
import openpyxl
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from openpyxl.utils import get_column_letter as L
from openpyxl.chart import BarChart,LineChart,Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import ColorScaleRule
os.chdir(Path(__file__).parent)
def log(s):
 with open('activity.log','a') as f:f.write(datetime.datetime.now(datetime.timezone.utc).isoformat()+' '+s+'\n')
w=openpyxl.Workbook();w.remove(w.active)
def sh(n,headers):s=w.create_sheet(n);s.append(headers);return s
read=sh('README',['Topic','Instructions / assumptions'])
for row in [('Purpose','Illustrative purchase and upfront refurbishment of an existing student residence; no secured property. 20 years operations after period-0 investment.'),('Sign convention','Nonnegative inputs; costs subtracted in formulas. Genuine negative cash flows, NPVs and losses are preserved.'),('Timing','Year 0 investment; years 1-20 operations. All annual cash flows at period end. No detailed construction schedule in this version.'),('Use','Blue = editable input, black = formula. Cached results provided; automatic full calculation in Excel. Fixed 20-column horizon; do not change Horizon_Years.'),('Legal','City of Brussels 2021 recommendations are municipal guidance, not a universal regional quota. Required_Social_Quota is a management assumption pending permit confirmation. No site has legal clearance.'),('Tax','Pre-tax project and equity model. 12.5% acquisition duty is a verified general Brussels rate; transaction-specific VAT/duty treatment unverified. Nonrecoverable works VAT 21% is a conservative assumption. Corporate taxes/tax shields excluded.'),('Analyses','Dynamic formula sensitivities; exhaustive integer social mix optimization including no-build; 1000 seeded triangular Monte Carlo trials. Solver instructions supplied; native Solver not executed.'),('Evidence','Official legal sources and market comparator checked 29 October 2026. Site costs, occupancy, travel times, and land cost comparisons are unverified assumptions.'),('Limitations','No native Microsoft Excel application execution. Formula evaluator supports the functions used. Site planning, university agreement, fire/accessibility, housing standards, valuation and loan terms require diligence.')]:read.append(row)
inputs=[('Market_Rooms',96,'rooms','Control','Assumption','Illustrative room allocation'),('Social_Rooms',24,'rooms','Control','Assumption','20% conservative management policy'),('Market_Base_Rent',500,'EUR/room/month','Control','Verified comparator','Brukot ordinary room median, Oct 28 2026; site asking rent unverified'),('Social_Base_Rent',359.3,'EUR/room/month','Control','Assumption','VUB benchmark less EUR90 utility allowance'),('Monthly_Charges',90,'EUR/occupied room/month','Approximate uncertainty','Assumption','Editable utilities/services allowance'),('Market_Occupancy',.95,'fraction','Inherent uncertainty','Assumption','No site absorption evidence'),('Social_Occupancy',.98,'fraction','Inherent uncertainty','Assumption','No waiting-list evidence'),('Rent_Growth',.02,'fraction/year','Inherent uncertainty','Assumption','Financial scenario; not a legal indexing entitlement'),('Expense_Growth',.025,'fraction/year','Inherent uncertainty','Assumption','Expense inflation separate from rent inflation'),('Acquisition_Price',4500000,'EUR','Control','Assumption','No secured plot/asking price'),('Works_Cost',1200000,'EUR ex VAT','Approximate uncertainty','Assumption','No contractor quote'),('Acquisition_Duty',.125,'fraction acquisition','Constant','Verified general rate','S4; transaction applicability unverified'),('Legal_Fees',25000,'EUR','Approximate uncertainty','Assumption','Notary/diligence allowance'),('Soft_Fee_Rate',.15,'fraction works','Approximate uncertainty','Assumption','Professional fees allowance'),('Contingency_Rate',.1,'fraction works+soft','Approximate uncertainty','Assumption','Contingency retained even in stress draws'),('Works_VAT',.21,'fraction works package','Constant','Unverified','Conservative nonrecoverable VAT; tax adviser confirmation'),('Furniture_Per_Room',1800,'EUR/room','Approximate uncertainty','Assumption','Supplier quote required'),('Working_Capital',100000,'EUR','Control','Assumption','Liquidity reserve recovered at exit; no double counting'),('Management_Rate',.06,'fraction receipts','Approximate uncertainty','Assumption','Operator fee'),('Maintenance_Per_Room',250,'EUR/room/year','Approximate uncertainty','Assumption','Repairs allowance'),('Property_Tax_Per_Room',150,'EUR/room/year','Approximate uncertainty','Unverified','Owner tax/charges allowance, not a statutory tax estimate'),('Insurance_Per_Room',100,'EUR/room/year','Approximate uncertainty','Assumption','Insurance allowance'),('Fixed_Admin',30000,'EUR/year','Approximate uncertainty','Assumption','Administration allowance'),('Replacement_Per_Room',150,'EUR/room/year','Approximate uncertainty','Assumption','Cash replacement capex separate from NOI'),('Project_Discount',.08,'fraction/year','Control','Assumption','Nominal pre-tax investor hurdle; not estimated WACC'),('Equity_Discount',.1,'fraction/year','Control','Assumption','Nominal levered equity hurdle'),('Exit_Cap_Rate',.055,'fraction','Inherent uncertainty','Assumption','NOI year21 / cap; appraiser required'),('Selling_Fee',.02,'fraction gross sale','Approximate uncertainty','Assumption','Broker/legal sale allowance'),('Debt_Ratio',.5,'fraction investment','Control','Assumption','Debt funds 50% total upfront uses'),('Debt_Rate',.045,'fraction/year','Inherent uncertainty','Assumption','No bank quote'),('Debt_Term',20,'years','Control','Assumption','Amortizing loan; <=20 required'),('Required_Social_Quota',.2,'fraction rooms','Control','Assumption','Conservative policy only; local permit applicability to confirm'),('Quota_Verified',0,'1 applies / 2 does not apply / 0 to confirm','Constant','Unverified','Selected municipality Ixelles; City of Brussels reference not assumed applicable'),('Budget_Limit',8000000,'EUR','Control','Assumption','Management development ceiling'),('VUB_Student_Weight',.6,'weight','Control','Assumption','ULB weight = 1 minus this input'),('Horizon_Years',20,'years fixed','Constant','Assumption','Fixed workbook period count'),('Months_Per_Year',12,'months/year','Constant','Verified','Calendar convention'),('Minimum_DSCR',1.2,'ratio','Control','Assumption','Investor covenant screen; no agreed loan covenant'),('Social_Total_Ceiling',449.3,'EUR/month incl charges','Control','Verified comparator','S2 institutional comparator, not private legal rent maximum'),('Market_Total_Ceiling',650,'EUR/month incl charges','Control','Assumption','Management affordability ceiling, not law')]
inputs += [('Selected_Location',1,'1 Ixelles,2 Etterbeek,3 Auderghem,4 City','Control','Assumption','Select hypothetical commune land budget'),('Building_Value_Ex_Land',2700000,'EUR','Approximate uncertainty','Assumption','Illustrative building value; not a valuation')]
i=sh('Inputs',['Input name','Value','Unit','Category','Status','Source / justification']);ref={}
for row in inputs:
 i.append(row);r=i._current_row;ref[row[0]]=r;i.cell(r,2).font=Font(color='1565C0');w.defined_names.add(DefinedName(row[0],attr_text=f"'Inputs'!$B${r}"))
i.cell(ref['Acquisition_Price'],2,'=Building_Value_Ex_Land+INDEX(Locations!E2:E5,Selected_Location)')
i.cell(ref['Acquisition_Price'],4,'Calculated from controls');i.cell(ref['Acquisition_Price'],5,'Assumption')
w.defined_names.add(DefinedName('Total_Rooms',attr_text="'Cost_summary'!$B$2"))
val=DataValidation(type='decimal',operator='greaterThanOrEqual',formula1=0);i.add_data_validation(val);val.add(f'B2:B{len(inputs)+1}')
legal=sh('Legal_and_sources',['ID / subject','Finding','Source URL','Source date','Checked','Status / project action'])
sources=[('S1 City 2021 guidance','City of Brussels reference: 20% social student units from 30 created units. Guidance not transplanted to Ixelles.','https://www.bruxelles.be/sites/default/files/bxl/Recommandations_logements_etudiants_VBX_17juin2021.pdf','17 June 2021','2026-10-29','Historical official text search-index verified; full PDF returned 403. Current site/permit applicability unverified.'),('S2 VUB benchmark','Standard room EUR449.30/month including services in 2026-27.','https://www.vub.be/en/studying-vub/practical-info-for-students/student-housing/practical-housing-information','2026-27','2026-10-29','Verified comparator; institutional access restricted'),('S3 Market anchor','Brukot ordinary room median EUR500 base / EUR590 incl charges.','https://www.brukot.be/en/student-guide/rent-stats','28 Oct 2026','2026-10-29','Verified asking-rent statistic, not site valuation'),('S4 Acquisition duty','General Brussels acquisition duty 12.5%; VAT regime may replace duty for particular transactions.','https://be.brussels/en/taxes-finances/taxes/property-related-taxes/registration-duties','Live page; date unstated','2026-10-29','General rate verified; asset structure unverified'),('S5 Quality standards','Regional safety, sanitary and equipment requirements apply; 2026 rules must be checked against measured plans.','https://be.brussels/fr/logement/location/conformite-logement/normes-et-securite','2026 requirements','2026-10-29','Framework verified; building inspection outstanding'),('S6 Reference rents','Reference disclosure and qualified excessive-rent framework; no universal price cap assumed.','https://be.brussels/fr/logement/location/bail-dhabitation/loyers-de-reference','Effective May 2025','2026-10-29','Framework verified; exact legal reference unverified'),('S7 Commune price data','IBSA publishes commune dwelling sale prices, not verified buildable land EUR/m2 for our sites.','https://ibsa.brussels/themes/logement/prix-des-logements','May 2026 tables','2026-10-29','No land price substitution; location land budgets are illustrative')]
for row in sources:legal.append(row)
legal.append(['S8 Current City guidance page','City still lists 17 June 2021 student housing recommendations; guidance is not itself a regulation.','https://www.bruxelles.be/plans-reglements-urbanisme','Live page 29 Oct 2026','2026-10-29','Verified publication; selected permit applicability remains to confirm.'])
for name in ['Planning permit and lawful use','University agreement / social allocation contract','Fire safety and accessibility','Property tax and VAT treatment','Room count / shared facilities / net usable area']:
 legal.append([name,'No property-specific assessment completed','Municipality / university / qualified professional','Pending','2026-10-29','UNVERIFIED: condition precedent to commitment'])
co=sh('Cost_summary',['Cost or named output','EUR / quantity'])
for row in [('Total rooms','=Market_Rooms+Social_Rooms'),('Acquisition','=Acquisition_Price'),('Acquisition duty','=Acquisition_Price*Acquisition_Duty'),('Works','=Works_Cost'),('Soft fees','=Works_Cost*Soft_Fee_Rate'),('Contingency','=(B5+B6)*Contingency_Rate'),('Works VAT','=SUM(B5:B7)*Works_VAT'),('Furniture','=Total_Rooms*Furniture_Per_Room'),('Legal/diligence','=Legal_Fees'),('Working capital','=Working_Capital'),('Total investment','=SUM(B3:B11)'),('Annual utility cost','=Total_Rooms*Monthly_Charges*Months_Per_Year'),('Annual maintenance','=Total_Rooms*Maintenance_Per_Room'),('Annual owner property tax allowance','=Total_Rooms*Property_Tax_Per_Room'),('Annual insurance','=Total_Rooms*Insurance_Per_Room'),('Annual admin','=Fixed_Admin'),('Annual replacement cash capex','=Total_Rooms*Replacement_Per_Room')]:co.append(row)
w.defined_names.add(DefinedName('Initial_Investment',attr_text="'Cost_summary'!$B$12"))
# Time across columns; line items down rows.
cf=sh('Cash_flow',['Line item']+list(range(21)))
labels=['Market base rent receipts','Social base rent receipts','Mandatory charge receipts','Total receipts','Utilities','Management','Maintenance','Property taxes allowance','Insurance','Administration','NOI','Replacement capex','Development spend','Terminal gross sale','Selling costs','Working capital release','Project CF without sale','Project CF with sale','Cumulative without sale','Cumulative with sale','Discounted without sale','Discounted with sale','Cumulative disc without sale','Cumulative disc with sale','Simple crossing no sale','Simple crossing with sale','Disc crossing no sale','Disc crossing with sale']
for label in labels:cf.append([label])
for t in range(21):
 col=L(t+2);prev=L(t+1)
 for r in range(2,30):cf.cell(r,t+2,0)
 if t==0:cf.cell(14,2,'=Initial_Investment')
 else:
  g=f'(1+Rent_Growth)^({col}$1-1)';eg=f'(1+Expense_Growth)^({col}$1-1)'
  fs={2:f'Market_Rooms*Market_Base_Rent*Market_Occupancy*Months_Per_Year*{g}',3:f'Social_Rooms*Social_Base_Rent*Social_Occupancy*Months_Per_Year*{g}',4:f'(Market_Rooms*Market_Occupancy+Social_Rooms*Social_Occupancy)*Monthly_Charges*Months_Per_Year*{g}',5:f'SUM({col}2:{col}4)',6:f'Cost_summary!B13*{eg}',7:f'{col}5*Management_Rate',8:f'Cost_summary!B14*{eg}',9:f'Cost_summary!B15*{eg}',10:f'Cost_summary!B16*{eg}',11:f'Cost_summary!B17*{eg}',12:f'{col}5-SUM({col}6:{col}11)',13:f'Cost_summary!B18*{eg}'}
  for r,f in fs.items():cf.cell(r,t+2,'='+f)
  if t==20:
   cf.cell(15,t+2,f'=({col}5*(1+Rent_Growth)*(1-Management_Rate)-SUM({col}6,{col}8:{col}11)*(1+Expense_Growth))/Exit_Cap_Rate');cf.cell(16,t+2,f'={col}15*Selling_Fee');cf.cell(17,t+2,'=Working_Capital')
 for r,f in {18:f'{col}12-{col}13-{col}14+{col}17',19:f'{col}18+{col}15-{col}16',20:f'SUM($B18:{col}18)',21:f'SUM($B19:{col}19)',22:f'{col}18/(1+Project_Discount)^{col}$1',23:f'{col}19/(1+Project_Discount)^{col}$1',24:f'SUM($B22:{col}22)',25:f'SUM($B23:{col}23)'}.items():cf.cell(r,t+2,'='+f)
 if t>0:
  for r,cum,flow in [(26,20,18),(27,21,19),(28,24,22),(29,25,23)]:cf.cell(r,t+2,f'=IF(AND({col}{cum}>=0,{prev}{cum}<0),{col}$1-1-{prev}{cum}/{col}{flow},0)')
f=sh('Financing',['Line item']+list(range(21)))
for lab in ['Opening debt','Debt draws','Interest','Principal','Debt service','Closing debt','Equity CF without sale','Equity CF with sale','Cumulative equity CF','DSCR']:f.append([lab])
for t in range(21):
 col=L(t+2);prev=L(t+1)
 for r in range(2,12):f.cell(r,t+2,0)
 if t==0:f.cell(3,2,'=Initial_Investment*Debt_Ratio');f.cell(7,2,'=B3')
 else:
  for r,formula in {2:f'{prev}7',4:f'{col}2*Debt_Rate',5:f'IF({col}$1<=Debt_Term,MIN({col}2,MAX(0,-PMT(Debt_Rate,Debt_Term,Initial_Investment*Debt_Ratio)-{col}4)),0)',6:f'{col}4+{col}5',7:f'{col}2-{col}5',11:f'IF({col}6>0,Cash_flow!{col}12/{col}6,1E99)'}.items():f.cell(r,t+2,'='+formula)
 for r,formula in {8:f'Cash_flow!{col}18+{col}3-{col}6'+(f'-{col}7' if t==20 else ''),9:f'Cash_flow!{col}19+{col}3-{col}6'+(f'-{col}7' if t==20 else ''),10:f'SUM($B9:{col}9)'}.items():f.cell(r,t+2,'='+formula)
res=sh('Results',['Metric','With terminal sale','Without terminal sale','Interpretation'])
rows=[('Project NPV','=NPV(Project_Discount,Cash_flow!C19:V19)+Cash_flow!B19','=NPV(Project_Discount,Cash_flow!C18:V18)+Cash_flow!B18','Positive creates value at nominal pre-tax hurdle; no-sale excludes residual owned asset value.'),('Project IRR','=IFERROR(IRR(Cash_flow!B19:V19),"Undefined")','=IFERROR(IRR(Cash_flow!B18:V18),"Undefined")','Annual discount rate making NPV zero; multiple sign changes can mislead.'),('Equity NPV','=NPV(Equity_Discount,Financing!C9:V9)+Financing!B9','=NPV(Equity_Discount,Financing!C8:V8)+Financing!B8','Use levered hurdle only on equity flows.'),('Equity IRR','=IFERROR(IRR(Financing!B9:V9),"Undefined")','=IFERROR(IRR(Financing!B8:V8),"Undefined")','IRR undefined if no meaningful crossing.'),('Simple project payback','=IF(COUNTIF(Cash_flow!B21:V21,">=0")=0,"not achieved within 20 years",SUM(Cash_flow!C27:V27))','=IF(COUNTIF(Cash_flow!B20:V20,">=0")=0,"not achieved within 20 years",SUM(Cash_flow!C26:V26))','Interpolated annual crossing; sale-dependent result separated.'),('Discounted project payback','=IF(COUNTIF(Cash_flow!B25:V25,">=0")=0,"not achieved within 20 years",SUM(Cash_flow!C29:V29))','=IF(COUNTIF(Cash_flow!B24:V24,">=0")=0,"not achieved within 20 years",SUM(Cash_flow!C28:V28))','Discounted capital recovery.'),('Year1 NOI','=Cash_flow!C12',None,'Before financing, income taxes and replacement capex.'),('Minimum operating DSCR','=MIN(Financing!C11:V11)',None,'Compare assumed 1.2 minimum; excludes years after loan maturity only if term remains fixed20.'),('Total development cost','=Initial_Investment',None,'Sources and uses must reconcile.'),('Social share','=Social_Rooms/Total_Rooms',None,'Compare required quota, subject to legal verification.'),('Market total housing cost','=Market_Base_Rent+Monthly_Charges',None,'Base rent + mandatory charges, EUR/month.'),('Social total housing cost','=Social_Base_Rent+Monthly_Charges',None,'Base rent + mandatory charges, EUR/month.'),('Maximum viable acquisition price','=Acquisition_Price+B2/(1+Acquisition_Duty)',None,'Holding all operations/works fixed; negative ceiling means even free acquisition insufficient.'),('Terminal PV share','=Cash_flow!V15*(1-Selling_Fee)/(1+Project_Discount)^Horizon_Years/(B2+Initial_Investment)',None,'Share of total discounted positive project inflows; no-sale is not a full hold valuation.')]
for row in rows:res.append(row)
w.defined_names.add(DefinedName('Project_NPV',attr_text="'Results'!$B$2"))
log('Core model built: named inputs, distinct project/equity finance, horizontal years 0-20, with/without sale KPI paths. Saving checkpoint.')
w.save('intermediate/model_core.xlsx')
# Common analytical engine: every row is the full 20-year project model.
e=sh('Analysis_engine',['Case','Market rooms','Social rooms','Market rent','Market occupancy','Social occupancy','Acquisition','Works multiplier','Discount','Exit multiplier','Investment','NPV']+['CF '+str(t) for t in range(21)])
def case(name,nm='=Market_Rooms',ns='=Social_Rooms',rent='=Market_Base_Rent',occ='=Market_Occupancy',occs='=Social_Occupancy',acq='=Acquisition_Price',cost=1,discount='=Project_Discount',exitmul=1):
 r=e._current_row+1;e.append([name,nm,ns,rent,occ,occs,acq,cost,discount,exitmul,f'=G{r}*(1+Acquisition_Duty)+Works_Cost*H{r}*(1+Soft_Fee_Rate)*(1+Contingency_Rate)*(1+Works_VAT)+(B{r}+C{r})*Furniture_Per_Room+Legal_Fees+Working_Capital',f'=NPV(I{r},N{r}:AG{r})+M{r}'])
 for t in range(21):
  col=13+t
  if t==0:formula=f'=-K{r}'
  else:
   receipt=f'(B{r}*E{r}*(D{r}+Monthly_Charges)+C{r}*F{r}*(Social_Base_Rent+Monthly_Charges))*Months_Per_Year'
   expense=f'((B{r}+C{r})*(Monthly_Charges*Months_Per_Year+Maintenance_Per_Room+Property_Tax_Per_Room+Insurance_Per_Room+Replacement_Per_Room)+Fixed_Admin)'
   formula=f'={receipt}*(1-Management_Rate)*(1+Rent_Growth)^{t-1}-{expense}*(1+Expense_Growth)^{t-1}'
   if t==20:
    terminal=f'({receipt}*(1-Management_Rate)*(1+Rent_Growth)^Horizon_Years-({expense}-(B{r}+C{r})*Replacement_Per_Room)*(1+Expense_Growth)^Horizon_Years)/Exit_Cap_Rate'
    formula+=f'+{terminal}*(1-Selling_Fee)*J{r}+Working_Capital'
  e.cell(r,col,formula)
 return r
base=case('Base reconciliation')
se=sh('Sensitivity',['Variable','Low','High','NPV low','NPV high','Swing','Range justification'])
for name,lo,hi,kw in [('Market base rent',395,630,'rent'),('Market occupancy',.75,.99,'occ'),('Discount rate',.06,.12,'discount'),('Works multiplier',.85,1.25,'cost'),('Exit value multiplier',.7,1.2,'exitmul')]:
 a=case(name+' low',**{kw:lo});b=case(name+' high',**{kw:hi});r=se._current_row+1;se.append([name,lo,hi,f'=Analysis_engine!L{a}',f'=Analysis_engine!L{b}',f'=ABS(D{r}-E{r})','Rent range S3; all other envelopes subjective, not confidence bounds.'])
se['A9']='Two-way market base rent x occupancy';
for col,occ in enumerate([.75,.85,.95,.99],2):se.cell(10,col,occ)
for r,rent in enumerate([395,450,500,560,630],11):
 se.cell(r,1,rent)
 for col in range(2,6):rr=case('Two-way',rent=rent,occ=f'=Sensitivity!{L(col)}10');se.cell(r,col,f'=Analysis_engine!L{rr}')
se.conditional_formatting.add('B11:E15',ColorScaleRule(start_type='min',start_color='F8696B',end_type='max',end_color='63BE7B'))
opt=sh('Optimization',['Social rooms','Market rooms','NPV','Quota pass','Budget pass','Score'])
for n in range(121):
 r=opt._current_row+1;rr=case('Mix optimization',nm=f'=Total_Rooms-Optimization!A{r}',ns=n)
 opt.append([n,f'=Total_Rooms-A{r}',f'=Analysis_engine!L{rr}',f'=IF(A{r}>=Required_Social_Quota*Total_Rooms,1,0)',f'=IF(Analysis_engine!K{rr}<=Budget_Limit,1,0)',f'=IF(AND(D{r}=1,E{r}=1),C{r},-1E99)'])
opt['H1']='Best construction NPV';opt['I1']='=MAX(F2:F122)';opt['H2']='Selected NPV including no build';opt['I2']='=MAX(0,I1)';opt['H3']='Best social rooms';opt['I3']='=INDEX(A2:A122,MATCH(I1,F2:F122,0))';opt['H4']='Best market rooms';opt['I4']='=Total_Rooms-I3';opt['H5']='Solver setup';opt['I5']='Max Results!B2 by Inputs Market_Rooms and Social_Rooms; integer >=0; sum=120; social>=required quota*total; Cost_summary!B12<=budget. Simplex LP applies with total fixed. Native Solver NOT run. Exact enumeration implemented for 120 rooms; regenerate domain if total changes.'
mc=sh('Monte_Carlo',['Trial','Works factor','Market occupancy','Rent factor','Exit factor','NPV EUR']);rng=random.Random(29102026)
for trial in range(1,1001):
 r=mc._current_row+1;draw=[rng.triangular(.85,1.25,1),rng.triangular(.75,.99,.95),rng.triangular(.85,1.1,1),rng.triangular(.7,1.2,1)]
 rr=case('MC',rent=f'=Market_Base_Rent*Monte_Carlo!D{r}',occ=f'=Monte_Carlo!C{r}',cost=f'=Monte_Carlo!B{r}',exitmul=f'=Monte_Carlo!E{r}');mc.append([trial]+draw+[f'=Analysis_engine!L{rr}'])
for r,(label,fm) in enumerate([('Mean','=AVERAGE(F2:F1001)'),('Standard deviation','=STDEV(F2:F1001)'),('P(NPV<0)','=COUNTIF(F2:F1001,"<0")/1000'),('P5','=PERCENTILE(F2:F1001,0.05)'),('P50','=PERCENTILE(F2:F1001,0.5)'),('P95','=PERCENTILE(F2:F1001,0.95)')],2):mc.cell(r,8,label);mc.cell(r,9,fm)
mc['H10']='1000 seeded independent triangular draws; subjective min/mode/max parameters. Draws fixed; NPV formulas update with inputs. Budget and affordability are not guaranteed by stress draws; test before commitment.'
mc['K1']='NPV upper bound';mc['L1']='Count'
for r in range(2,14):mc.cell(r,11,f'=MIN(F$2:F$1001)+(MAX(F$2:F$1001)-MIN(F$2:F$1001))*{r-1}/12');mc.cell(r,12,f'=COUNTIF(F$2:F$1001,"<="&K{r})'+(f'-COUNTIF(F$2:F$1001,"<="&K{r-1})' if r>2 else ''))
loc=sh('Locations',['Municipality','VUB commute min','ULB commute min','Weighted commute','Land budget EUR','Land area m2','EUR/m2 implied','Status'])
for r,(name,vub,ulb,land,area) in enumerate([('Ixelles',10,12,1800000,600),('Etterbeek',12,20,1500000,600),('Auderghem',18,25,1200000,600),('City of Brussels',30,35,1300000,600)],2):loc.append([name,vub,ulb,f'=B{r}*VUB_Student_Weight+C{r}*(1-VUB_Student_Weight)',land,area,f'=E{r}/F{r}','UNVERIFIED illustrative travel/land budgets. Not listings or appraisals; obtain parcel and transit evidence.'])
loc['A8']='Commute and land comparison are screening assumptions only. Cheapest land may lose rent/occupancy; no financial location ranking without a site.'
lp=sh('Location_finance',['Municipality','Assumed land budget EUR','Assumed building value EUR','Acquisition EUR','Project NPV EUR','Interpretation'])
for rr in range(2,6):
 er=case('Location land comparison',acq=f'=Building_Value_Ex_Land+Locations!E{rr}')
 lp.append([f'=Locations!A{rr}',f'=Locations!E{rr}','=Building_Value_Ex_Land',f'=B{rr}+C{rr}',f'=Analysis_engine!L{er}','All parcels, land budgets and constant building value are UNVERIFIED. Not a market ranking; same room plan and operating economics held fixed.'])
lp['A8']='No verified current comparable buildable land sales were found within the time budget. Statbel / IBSA dwelling medians are not land EUR/m2 and were not substituted.'
ch=BarChart();ch.title='Illustrative land-cost effect on NPV (EUR)';ch.add_data(Reference(lp,min_col=5,min_row=1,max_row=5),titles_from_data=True);ch.set_categories(Reference(lp,min_col=1,min_row=2,max_row=5));lp.add_chart(ch,'H2')
check=sh('Model_checks',['Check','Result'])
for row in [('Base annual vs analytical engine','=ABS(Results!B2-Analysis_engine!L2)<0.01'),('Debt balance no negative','=MIN(Financing!B7:V7)>=0'),('Sources and uses','=ABS(-Financing!B9+Financing!B3-Initial_Investment)<0.01'),('Final debt if term20','=ABS(Financing!V7)<0.01'),('Social quota','=Social_Rooms>=Required_Social_Quota*Total_Rooms'),('Social total cost','=Social_Base_Rent+Monthly_Charges<=Social_Total_Ceiling'),('Market total cost','=Market_Base_Rent+Monthly_Charges<=Market_Total_Ceiling'),('Budget','=Initial_Investment<=Budget_Limit'),('Occupancy bounds','=AND(Market_Occupancy>=0,Market_Occupancy<=1,Social_Occupancy>=0,Social_Occupancy<=1)'),('1000 trials','=COUNT(Monte_Carlo!F2:F1001)=1000'),('Histogram total','=SUM(Monte_Carlo!L2:L13)=1000'),('Quota verified','=Quota_Verified>0'),('Inputs nonnegative',f'=MIN(Inputs!B2:B{len(inputs)+1})>=0'),('Financial formulas OK','=COUNTIF(B2:B12,0)=0'),('Master model OK','=AND(B14,B15)'),('Investment approval gate','=IF(AND(B16,B13,Results!B2>0,Results!B9>=Minimum_DSCR),"CONDITIONAL PROCEED after site diligence","HOLD / REDESIGN: financial or legal conditions unresolved")')]:check.append(row)
con=sh('Conclusion',['Managerial question','Technique','Calculated answer / recommendation'])
for row in [('Should we invest?','NPV + legal + financing gates','=Model_checks!B17'),('Does the sale drive returns?','With/without sale comparison','=Results!B2-Results!C2'),('What acquisition price is supportable?','Zero-NPV residual acquisition ceiling','=Results!B14'),('What social mix is preferred?','Exact integer enumeration','=Optimization!I3'),('What is the downside risk?','1000 Monte Carlo trials','=Monte_Carlo!I4'),('Legal conditions','Municipal guidance verification','City 2021 guidance verified historically; selected Ixelles project applicability, planning permit and university agreement remain unverified.')]:con.append(row)
for s in w:
 s.freeze_panes='B2';s.column_dimensions['A'].width=34
 for col in range(2,min(s.max_column,35)+1):s.column_dimensions[L(col)].width=17
 for c in s[1]:c.fill=PatternFill('solid',fgColor='163A55');c.font=Font(color='FFFFFF',bold=True)
 for row in s:
  for c in row:
   if c.data_type=='f':c.font=Font(color='000000');c.number_format='#,##0.00;[Red](#,##0.00)'
for s in [read,legal,con]:
 s.column_dimensions['B'].width=90
 for row in s:
  s.row_dimensions[row[0].row].height=45
  for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
for s,xcol,ycol,end,title,pos in [(se,1,6,6,'NPV tornado: swing by stress range (EUR)','H2'),(mc,11,12,13,'Monte Carlo NPV histogram (EUR)','N2')]:
 ch=BarChart();ch.title=title;ch.add_data(Reference(s,min_col=ycol,min_row=1,max_row=end),titles_from_data=True);ch.set_categories(Reference(s,min_col=xcol,min_row=2,max_row=end));s.add_chart(ch,pos)
ch=LineChart();ch.title='Annual project cash flow with sale (EUR)';ch.add_data(Reference(cf,min_col=2,max_col=22,min_row=19,max_row=19),from_rows=True);ch.set_categories(Reference(cf,min_col=2,max_col=22,min_row=1,max_row=1));cf.add_chart(ch,'B33')
dictsh=sh('Named_cells',['Name','Sheet','Cell','Meaning'])
for name,dn in w.defined_names.items():
 for sn,cell in dn.destinations:dictsh.append([name,sn,cell,name.replace('_',' ')])
w.calculation=CalcProperties(fullCalcOnLoad=True);w.save('.model.tmp.xlsx');os.replace('.model.tmp.xlsx','model.xlsx');shutil.copy2('model.xlsx','intermediate/model_analysis.xlsx')
log('Saved model.xlsx: 1000 simulations, 121 room-mix candidates, one/two-way sensitivities, location comparison, legal register and charts. Starting formula evaluation.')
t=Path('../groupInas/recalculate.py').read_text().replace("if fn=='AVERAGE':", "if fn=='STDEV':return statistics.stdev(ns)\n if fn=='AVERAGE':")
Path('recalculate.py').write_text(t)
