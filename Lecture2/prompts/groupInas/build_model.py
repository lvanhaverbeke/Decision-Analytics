import openpyxl,math,random,json,os,time,shutil,zipfile,sys
from pathlib import Path
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import LineChart,BarChart,Reference,ScatterChart,Series
from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.properties import CalcProperties
folder=Path(__file__).parent;os.chdir(folder);(folder/'intermediate').mkdir(exist_ok=True)
w=openpyxl.Workbook();w.remove(w.active)
def sheet(n,headers):
 s=w.create_sheet(n);s.append(headers);return s
s=sheet('Read_Me',['Topic','Instructions'])
for a,b in [('Purpose','Evaluate a hypothetical student housing project near VUB. Site and all financial costs require quotations.'),('Currency and timing','Nominal EUR; year 0 capex at valuation 29 October 2026; first operations at year-end 2027. One development period; delay risk not modelled.'),('Use','Blue inputs editable. Black formulas. Recalculate automatically in Excel. Seeded simulation changes when assumptions change; rerun build_model.py for new seed.'),('Cash flows','Pre-tax unlevered and equity cash flows; 20 operations years, annual equal debt payments, sale at end year 20. No corporate tax, depreciation tax shield or grants.'),('Sensitivity','Equivalent dynamic formula grid, not native Excel What-If Data Tables.'),('Optimization','Exhaustive discrete integer capacity / EUR 5 rent grid. No Solver add-in required. Continuous rent optimum is not claimed.'),('Market evidence','Sources verified through web tools 29 October 2026. Asking rents do not establish local absorption or achievable rents.'),('Affordability','At least 10% rooms at no more than VUB standard room all-inclusive benchmark. Management policy, not a statutory rent cap; enrolled means-tested students assumed eligible.'),('Limits','No plot selected; legal reference rents and permits unverified. Demand elasticity, construction/energy/financing inputs are management assumptions. No native Excel desktop execution.'),('AI disclosure','AI built analysis and researched sources; student validation and contributions have not occurred and must not be represented as completed.')]:s.append([a,b])
source=sheet('Sources',['ID','Evidence / value','URL','Published / period','Retrieved','Use / limitation'])
sources=[['S1','VUB standard room EUR 449.30/month, includes utilities','https://www.vub.be/en/studying-vub/practical-info-for-students/student-housing/practical-housing-information','2026-27; page updated 18 Sep 2026','2026-10-29','Institutional comparator; limited allocation, not private achievable rent'],['S2','Brukot ordinary student room median EUR 590 including charges; listing statistics','https://www.brukot.be/en/student-guide/rent-stats','Rolling last 12 months; publication date unstated','2026-10-29','Anchor for initial rent; city listings, not VUB transactions'],['S3','Brussels excessive-rent framework; reference disclosure; 20% presumption with qualifications','https://be.brussels/fr/logement/location/bail-dhabitation/loyers-de-reference','Rule effective 1 May 2025; page publication unstated','2026-10-29','No universal rent ceiling inferred. Property-specific legal review outstanding'],['S4','ABEX construction price index 1071 at 1 July 2026','https://abex.be/en/abex-index/?v=910955a907e7','1 July 2026','2026-10-29','Index provides escalation context, not EUR/m2 quote'],['S5','NBB monthly agreed bank lending rate dataset','https://www.nbb.be/en/statistics/topics/financial-institutions/banks/interest-rates-loans-and-deposits','Latest release 1 Oct 2026; Aug 2026 data','2026-10-29','No project loan quote extracted; 4.5% assumption unverified']]
for row in sources:source.append(row)
inputs=[('Total_Rooms',120,'rooms','control','Hypothetical site capacity design'),('Affordable_Units_Count',12,'rooms','control','At least 10%'),('Base_Rent_EUR_Month',590,'EUR/month all inclusive','control','S2 initial anchor'),('Affordable_Rent_EUR_Month',449.3,'EUR/month all inclusive','control','S1 comparator; policy'),('Reference_Affordability_Rent_EUR',449.3,'EUR/month','fact','S1 2026-27 institutional benchmark'),('Reference_Market_Rent',590,'EUR/month','fact','S2 city median; not site-specific'),('Reference_Demand',125,'students at reference rent','assumption','Unobserved local annual demand pool; survey needed'),('Demand_Elasticity',1.2,'parameter','assumption','Uncalibrated exponential response; 0.6-2 stress'),('Gross_Area_Per_Room',28,'m2/room','assumption','All common areas allocated; no verified plan'),('Hard_Cost_Per_m2',2200,'EUR/m2 ex VAT','assumption','Quote required; stress 1800-2800'),('Acquisition_EUR',1500000,'EUR','assumption','No property selected; vendor quote required'),('Transaction_Rate',.15,'fraction land','assumption','Allowance, not asserted tax rate'),('Soft_Cost_Rate',.15,'fraction hard cost','assumption','Architect/permit/engineering allowance'),('Contingency_Rate',.1,'fraction hard+soft','assumption','Cost allowance'),('Nonrecoverable_VAT',.21,'fraction development package','assumption','Conservative simplified VAT; professional tax review needed'),('Furniture_Per_Room',1800,'EUR/room','assumption','Supplier quote needed'),('Working_Capital',50000,'EUR','assumption','Initial cash permanently committed'),('Fixed_Opex',60000,'EUR/year','assumption','Admin, insurance and owner tax allowance'),('Variable_Opex',700,'EUR/room/year','assumption','Maintenance, water, waste, internet, cleaning allowance'),('Energy_kWh_Per_Room',1800,'kWh/room/year','assumption','Illustrative annual energy model'),('Energy_Price',.25,'EUR/kWh','assumption','No supplier quote; includes modelled energy only'),('Management_Fee',.06,'fraction revenue','assumption','Operator quote required'),('Replacement_Reserve',150,'EUR/room/year','assumption','Recurring cash replacement allowance'),('Annual_Growth',.02,'annual fraction','assumption','Common nominal rent/opex growth; legal indexing not guaranteed'),('Project_Discount_Rate',.08,'annual fraction','assumption','Investor hurdle, not observed WACC'),('Equity_Discount_Rate',.1,'annual fraction','assumption','Riskier levered hurdle; illustrative'),('Investment_Horizon_Years',20,'years','assumption','Model implemented for 20 annual periods; keep fixed'),('Terminal_Cap_Rate',.055,'fraction','assumption','Exit quote required; 4.5%-7% stress'),('Selling_Cost_Rate',.02,'fraction sale','assumption','Broker/legal allowance'),('Debt_Ratio',.55,'fraction capex','control','Financing assumption'),('Debt_Interest_Rate',.045,'annual fraction','assumption','Unverified loan pricing; S5 context only'),('Loan_Term',20,'years','assumption','Amortizing; fixed 20-year horizon'),('Budget_EUR',14000000,'EUR','constraint','Management ceiling'),('Area_Capacity',4000,'m2 gross','constraint','Unverified hypothetical envelope'),('Rent_Policy_Max',650,'EUR/month','constraint','Management affordability ceiling; not law'),('Minimum_Rooms',40,'rooms','constraint','Operating scale assumption'),('CO2_Factor',.15,'kgCO2e/kWh','assumption','Illustrative operational factor; no verified energy mix'),('Cost_Multiplier',1,'multiplier','uncertainty','Base 1; triangular 0.85,1,1.25'),('Demand_Multiplier',1,'multiplier','uncertainty','Base 1; triangular 0.75,1,1.15'),('Opex_Multiplier',1,'multiplier','uncertainty','Base 1; triangular 0.9,1,1.2')]
i=sheet('Inputs',['Named input','Value','Unit','Category','Evidence / rationale']);refs={}
for row in inputs:
 i.append(row);r=i.max_row;refs[row[0]]=r;w.defined_names.add(DefinedName(row[0],attr_text=f"'Inputs'!$B${r}"))
 for c in [i.cell(r,2)]:c.font=Font(color='1565C0')
dv=DataValidation(type='decimal',operator='greaterThanOrEqual',formula1=0);i.add_data_validation(dv);dv.add('B2:B41')
# Development, demand, operations use one shared logic; all analyses below use algebraically equivalent engine.
d=sheet('Development_Costs',['Cost','EUR']);
for row in [('Land','=Acquisition_EUR'),('Transactions','=B2*Transaction_Rate'),('Hard works','=Total_Rooms*Gross_Area_Per_Room*Hard_Cost_Per_m2*Cost_Multiplier'),('Soft costs','=B4*Soft_Cost_Rate'),('Contingency','=(B4+B5)*Contingency_Rate'),('Nonrecoverable VAT','=SUM(B4:B6)*Nonrecoverable_VAT'),('Furniture','=Total_Rooms*Furniture_Per_Room'),('Working capital','=Working_Capital'),('Total investment','=SUM(B2:B9)')]:d.append(row)
w.defined_names.add(DefinedName('Initial_Investment_EUR',attr_text="'Development_Costs'!$B$10"))
dm=sheet('Demand_Model',['Calculation','Value'])
for row in [('Market rooms','=Total_Rooms-Affordable_Units_Count'),('Market demand','=Reference_Demand*Demand_Multiplier*EXP(-Demand_Elasticity*(Base_Rent_EUR_Month/Reference_Market_Rent-1))'),('Market occupied','=MIN(B2,MAX(0,B3))'),('Affordable occupied','=Affordable_Units_Count*MIN(1,Demand_Multiplier)'),('Occupied rooms','=B4+B5'),('Occupancy','=B6/Total_Rooms'),('Annual revenue','=12*(B4*Base_Rent_EUR_Month+B5*Affordable_Rent_EUR_Month)')]:dm.append(row)
op=sheet('Operating_Model',['Year 1 calculation','EUR/year'])
for row in [('Revenue','=Demand_Model!B8'),('Fixed opex','=Fixed_Opex*Opex_Multiplier'),('Variable opex','=Total_Rooms*Variable_Opex*Opex_Multiplier'),('Energy','=Total_Rooms*Energy_kWh_Per_Room*Energy_Price*Opex_Multiplier'),('Management','=B2*Management_Fee'),('Net operating income','=B2-SUM(B3:B6)'),('Replacement capex','=Total_Rooms*Replacement_Reserve'),('Recurring project cash flow','=B7-B8')]:op.append(row)
f=sheet('Financing',['Year','Opening debt','Interest','Principal','Debt service','Closing debt']);
for t in range(1,21):
 r=t+1;opening='Initial_Investment_EUR*Debt_Ratio' if t==1 else f'F{r-1}'
 f.append([t,'='+opening,f'=B{r}*Debt_Interest_Rate',f'=MIN(B{r},MAX(0,-PMT(Debt_Interest_Rate,Loan_Term,Initial_Investment_EUR*Debt_Ratio)-C{r}))',f'=C{r}+D{r}',f'=B{r}-D{r}'])
c=sheet('Cash_Flows',['Year','Revenue','NOI','Replacement','Project CF','Debt service','Debt balance','Equity CF','Cum project','Discounted CF','Cum discounted','Payback crossing','Disc crossing'])
c.append([0,0,0,0,'=-Initial_Investment_EUR',0,'=Initial_Investment_EUR*Debt_Ratio','=-Initial_Investment_EUR*(1-Debt_Ratio)','=E2','=E2','=J2',0,0])
for t in range(1,21):
 r=t+2;cf=f'=C{r}-D{r}'
 if t==20:cf+=f'+(C{r}-D{r})*(1+Annual_Growth)/Terminal_Cap_Rate*(1-Selling_Cost_Rate)'
 c.append([t,f'=Operating_Model!B2*(1+Annual_Growth)^{t-1}',f'=Operating_Model!B7*(1+Annual_Growth)^{t-1}',f'=Operating_Model!B8*(1+Annual_Growth)^{t-1}',cf,f'=Financing!E{t+1}',f'=Financing!F{t+1}',f'=E{r}-F{r}'+(f'-G{r}' if t==20 else ''),f'=I{r-1}+E{r}',f'=E{r}/(1+Project_Discount_Rate)^A{r}',f'=K{r-1}+J{r}',f'=IF(AND(I{r}>=0,I{r-1}<0),A{r}-1-I{r-1}/E{r},0)',f'=IF(AND(K{r}>=0,K{r-1}<0),A{r}-1-K{r-1}/J{r},0)'])
k=sheet('KPI_Dashboard',['Metric','Value'])
for row in [('Project NPV','=NPV(Project_Discount_Rate,Cash_Flows!E3:E22)+Cash_Flows!E2'),('Project IRR','=IFERROR(IRR(Cash_Flows!E2:E22),"No meaningful IRR")'),('Equity NPV','=NPV(Equity_Discount_Rate,Cash_Flows!H3:H22)+Cash_Flows!H2'),('Equity IRR','=IFERROR(IRR(Cash_Flows!H2:H22),"No meaningful IRR")'),('Payback including sale','=IF(COUNTIF(Cash_Flows!I2:I22,">=0")=0,"Beyond horizon",SUM(Cash_Flows!L2:L22))'),('Discounted payback','=IF(COUNTIF(Cash_Flows!K2:K22,">=0")=0,"Beyond horizon",SUM(Cash_Flows!M2:M22))'),('Annual revenue','=Operating_Model!B2'),('NOI','=Operating_Model!B7'),('Occupancy','=Demand_Model!B7'),('DSCR year 1','=Operating_Model!B7/Financing!E2'),('Affordable share','=Affordable_Units_Count/Total_Rooms'),('Carbon tonnes/year','=Total_Rooms*Energy_kWh_Per_Room*CO2_Factor/1000'),('Energy kWh/year','=Total_Rooms*Energy_kWh_Per_Room'),('Profitability index','=1+B2/Initial_Investment_EUR'),('Maximum viable development spend','=Initial_Investment_EUR+B2'),('Annual student rent savings','=Affordable_Units_Count*12*(Base_Rent_EUR_Month-Affordable_Rent_EUR_Month)')]:k.append(row)
w.defined_names.add(DefinedName('Project_NPV_EUR',attr_text="'KPI_Dashboard'!$B$2"))
w.calculation=CalcProperties(fullCalcOnLoad=True)
w.save('intermediate/model_core.xlsx')
# Engine rows calculate full same-model NPV via growing annuity, with explicit year 1 bridge.
e=sheet('Analysis_Engine',['Case','Rooms','Market rent','Cost factor','Demand factor','Opex factor','Cap rate','Affordable units','Investment EUR','Occupied market','Occupied affordable','Revenue','Cash NOI after reserve','PV multiplier','Project NPV','Feasible'])
def engine_row(label,n,rent,cost=1,demand=1,opex=1,cap='=Terminal_Cap_Rate',aff=None):
 r=e.max_row+1
 if aff is None:aff=f'=INT((B{r}+9)/10)'
 e.append([label,n,rent,cost,demand,opex,cap,aff,f'=Acquisition_EUR*(1+Transaction_Rate)+B{r}*Gross_Area_Per_Room*Hard_Cost_Per_m2*D{r}*(1+Soft_Cost_Rate)*(1+Contingency_Rate)*(1+Nonrecoverable_VAT)+B{r}*Furniture_Per_Room+Working_Capital',f'=MIN(B{r}-H{r},Reference_Demand*E{r}*EXP(-Demand_Elasticity*(C{r}/Reference_Market_Rent-1)))',f'=H{r}*MIN(1,E{r})',f'=12*(J{r}*C{r}+K{r}*Affordable_Rent_EUR_Month)',f'=L{r}*(1-Management_Fee)-(Fixed_Opex+B{r}*(Variable_Opex+Energy_kWh_Per_Room*Energy_Price))*F{r}-B{r}*Replacement_Reserve',f'=IF(Project_Discount_Rate=Annual_Growth,Investment_Horizon_Years/(1+Project_Discount_Rate),(1-((1+Annual_Growth)/(1+Project_Discount_Rate))^Investment_Horizon_Years)/(Project_Discount_Rate-Annual_Growth))+(1+Annual_Growth)^Investment_Horizon_Years/G{r}*(1-Selling_Cost_Rate)/(1+Project_Discount_Rate)^Investment_Horizon_Years',f'=-I{r}+M{r}*N{r}',f'=IF(AND(B{r}>=Minimum_Rooms,B{r}*Gross_Area_Per_Room<=Area_Capacity,I{r}<=Budget_EUR,C{r}<=Rent_Policy_Max,H{r}>=B{r}*0.1,Affordable_Rent_EUR_Month<=Reference_Affordability_Rent_EUR),1,0)'])
 return r
base=engine_row('Base','=Total_Rooms','=Base_Rent_EUR_Month','=Cost_Multiplier','=Demand_Multiplier','=Opex_Multiplier',aff='=Affordable_Units_Count')
se=sheet('Sensitivity_Analysis',['Variable','Low','High','NPV low','NPV high','NPV swing','Range evidence'])
for label,lo,hi,idx,basis in [('Monthly market rent',480,730,1,'S2 city all-inclusive broad room range; 730 exceeds management policy for stress'),('Demand factor',.75,1.15,3,'Judgment; no empirical local range'),('Hard cost EUR/m2',1800,2800,2,'Unverified bid envelope'),('Opex factor',.9,1.2,4,'Judgment; quotes needed')]:
 args=['=Total_Rooms','=Base_Rent_EUR_Month',1,1,1];args[idx]=lo/2200 if idx==2 else lo;a=engine_row(label+' low',*args)
 args[idx]=hi/2200 if idx==2 else hi;b=engine_row(label+' high',*args)
 r=se.max_row+1;se.append([label,lo,hi,f'=Analysis_Engine!O{a}',f'=Analysis_Engine!O{b}',f'=ABS(D{r}-E{r})',basis])
se.append(['Greatest swing','=INDEX(A2:A5,MATCH(MAX(F2:F5),F2:F5,0))'])
se.append(['Two-way rent / cost multiplier NPV; dynamic formulas'])
for j,cm in enumerate([.85,1,1.15,1.25],2):se.cell(8,j,cm)
for r,rent in enumerate([480,530,590,650,730],9):
 se.cell(r,1,rent)
 for j in range(2,6):rr=engine_row('two-way','=Total_Rooms',f'=Sensitivity_Analysis!A{r}',f'=Sensitivity_Analysis!{L(j)}8');se.cell(r,j,f'=Analysis_Engine!O{rr}')
se.conditional_formatting.add('B9:E13',ColorScaleRule(start_type='min',start_color='F8696B',end_type='max',end_color='63BE7B'))
# Fully enumerate integer capacities and EUR 5 rents, all feasible alternatives; objective includes do-not-invest.
opt=sheet('Optimization',['Engine row','Rooms','Rent EUR','NPV','Feasible','Selection score'])
for n in range(40,141):
 for rent in range(400,651,5):
  rr=engine_row('optimization',n,rent);r=opt.max_row+1
  opt.append([rr,f'=Analysis_Engine!B{rr}',f'=Analysis_Engine!C{rr}',f'=Analysis_Engine!O{rr}',f'=Analysis_Engine!P{rr}',f'=IF(E{r}=1,D{r},-1E99)'])
end=opt.max_row
opt['H1']='Objective: maximize project NPV; area/budget/affordability/capacity constraints';opt['H2']='Best built NPV';opt['I2']=f'=MAX(F2:F{end})';opt['H3']='Selected NPV incl no-build';opt['I3']='=MAX(0,I2)';opt['H4']='Best row';opt['I4']=f'=MATCH(I2,F2:F{end},0)+1';opt['H5']='Best built rooms';opt['I5']=f'=INDEX(B1:B{end},I4)';opt['H6']='Best built rent';opt['I6']=f'=INDEX(C1:C{end},I4)';opt['H7']='Recommendation';opt['I7']='=IF(I2>0,"Conditional investment; obtain permits and bids","Do not invest: zero NPV dominates")';opt['H8']='Scope: exhaustive 40..140 rooms; EUR400..650 in EUR5 steps. Above 140 requires regeneration.'
w.save('intermediate/model_analysis.xlsx')
mc=sheet('Monte_Carlo',['Trial','Cost factor','Demand factor','Opex factor','Exit cap','NPV EUR']);rng=random.Random(20261029)
for trial in range(1,1001):
 draws=[rng.triangular(.85,1.25,1),rng.triangular(.75,1.15,1),rng.triangular(.9,1.2,1),rng.triangular(.045,.07,.055)];r=mc.max_row+1
 rr=engine_row('MC','=Total_Rooms','=Base_Rent_EUR_Month',f'=Monte_Carlo!B{r}',f'=Monte_Carlo!C{r}',f'=Monte_Carlo!D{r}',f'=Monte_Carlo!E{r}',aff='=Affordable_Units_Count')
 mc.append([trial]+draws+[f'=Analysis_Engine!O{rr}'])
for r,(label,formula) in enumerate([('Mean','=AVERAGE(F2:F1001)'),('Median','=MEDIAN(F2:F1001)'),('Minimum','=MIN(F2:F1001)'),('Maximum','=MAX(F2:F1001)'),('P5','=PERCENTILE(F2:F1001,0.05)'),('P95','=PERCENTILE(F2:F1001,0.95)'),('Probability negative','=COUNTIF(F2:F1001,"<0")/1000')],2):mc.cell(r,8,label);mc.cell(r,9,formula)
mc['H10']='Independent persistent triangular shocks, subjective endpoints; static seeded draws, formula-driven NPV. Correlations and delays omitted.'
mc['K1']='NPV upper bound EUR';mc['L1']='Frequency'
for r in range(2,15):
 mc.cell(r,11,f'=MIN($F$2:$F$1001)+(MAX($F$2:$F$1001)-MIN($F$2:$F$1001))*{r-1}/13');mc.cell(r,12,f'=COUNTIF($F$2:$F$1001,"<="&K{r})'+(f'-COUNTIF($F$2:$F$1001,"<="&K{r-1})' if r>2 else ''))
checks=sheet('Model_Checks',['Check','Result / status'])
for row in [('Engine matches annual base','=ABS(Analysis_Engine!O2-KPI_Dashboard!B2)<0.01'),('Affordability count','=Affordable_Units_Count>=Total_Rooms*0.1'),('Affordable price','=Affordable_Rent_EUR_Month<=Reference_Affordability_Rent_EUR'),('Capacity','=Total_Rooms*Gross_Area_Per_Room<=Area_Capacity'),('Debt final zero','=ABS(Financing!F21)<0.01'),('MC trial count','=COUNT(Monte_Carlo!F2:F1001)'),('Histogram count','=SUM(Monte_Carlo!L2:L14)'),('No native Excel test','Not performed; formulas checked by purpose-built local evaluator'),('Source/legal/quote validation','Not complete: exact site, legal reference rents, construction bids, EPC, fire safety and lender terms'),('Data Tables / Solver','Equivalent formula grid and finite enumeration implemented; native features not configured')]:checks.append(row)
summary=sheet('Decision_Summary',['Managerial question','Technique','Result','Interpretation / action'])
for row in [('Financial feasibility?','20-year pre-tax DCF','=KPI_Dashboard!B2','Reject if NPV <= 0; all unverified inputs need diligence'),('Minimum investment?','Break-even spend envelope','=KPI_Dashboard!B16','This is maximum affordable development spend, not a minimum; no unique minimum without secured site/design'),('Affordable sustainable rent?','Demand-aware grid','=Optimization!I6','Best built rent subject to management cap; affordable rent separately capped at VUB comparator'),('Optimal capacity?','Integer capacity and rent enumeration','=Optimization!I5','Best constructed project, not approval to invest; compare zero alternative'),('How long to recover?','Cumulative CF incl sale','=KPI_Dashboard!B6','Sale may dominate payback; see discounted payback'),('Main profitability drivers?','Stress ranges','=Sensitivity_Analysis!B6','Depends on assumed range widths; not a universal ranking'),('Probability of negative NPV?','1000 joint scenarios','=Monte_Carlo!I8','Conditional on subjective independent distributions'),('Invest?','NPV incl no build','=Optimization!I7','Require finance, legal approval, permits, bids, and demand validation')]:summary.append(row)
summary.append(['Accommodation mix','Scope','Standard private rooms and reserved affordable rooms','No studios introduced without site floor plan; configuration remains a limitation.'])
for sh,col1,col2,title,pos in [(c,1,5,'Annual project cash flow (EUR)','O2'),(se,1,4,'NPV sensitivity (EUR)','I2'),(mc,11,12,'Simulated project NPV histogram','N2')]:
 ch=BarChart();ch.title=title;ch.add_data(Reference(sh,min_col=col2,max_col=col2,min_row=1,max_row=21 if sh==c else 5 if sh==se else 14),titles_from_data=True);ch.set_categories(Reference(sh,min_col=col1,min_row=2,max_row=22 if sh==c else 5 if sh==se else 14));sh.add_chart(ch,pos)
# Named-cell dictionary / audit trail.
nd=sheet('Named_Cell_Dictionary',['Name','Sheet','Reference','Definition'])
for name,dn in w.defined_names.items():
 for sn,ref in dn.destinations:nd.append([name,sn,ref,name.replace('_',' ')])
for sh in w:
 sh.freeze_panes='B2';sh.auto_filter.ref=sh.dimensions
 for cell in sh[1]:cell.fill=PatternFill('solid',fgColor='143D59');cell.font=Font(color='FFFFFF',bold=True)
 sh.column_dimensions['A'].width=34
 for col in range(2,min(sh.max_column,16)+1):sh.column_dimensions[L(col)].width=22
 if sh.title in ['Read_Me','Sources','Decision_Summary','Model_Checks']:sh.column_dimensions['B'].width=75
 if sh.title in ['Read_Me','Sources','Decision_Summary','Model_Checks']:
  for row in sh:
   for cell in row:cell.alignment=Alignment(vertical='top',wrap_text=True)
   sh.row_dimensions[row[0].row].height=44
 for row in sh:
  for cell in row:
   if cell.data_type=='f':cell.number_format='#,##0.00;[Red](#,##0.00)'
w.save('.model.tmp.xlsx');os.replace('.model.tmp.xlsx','model.xlsx');print('MODEL SAVED',len(e['A']),len(opt['A']))
# Generic evaluator inherited from the inspected local implementation; add named reference support.
t=Path('../groupMSN/recalculate.py').read_text();a=t.index('for sheet,refs in ');b=t.index('if errors:',a);t=t[:a]+t[b:]
t=t.replace("if ss not in w.sheetnames:raise ValueError('sheet '+ss)","if ss not in w.sheetnames:raise ValueError('sheet '+ss)\n  if ref in w.defined_names:\n   target=list(w.defined_names[ref].destinations)[0]\n   return get(target[0],target[1].replace('$',''))")
Path('recalculate.py').write_text(t)
