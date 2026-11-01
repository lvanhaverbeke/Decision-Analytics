import openpyxl,os,math,json
from pathlib import Path
from openpyxl.chart import BarChart,LineChart,Reference
from openpyxl.formatting.rule import ColorScaleRule
w=openpyxl.load_workbook('model.xlsx');e=w['Analysis_Engine']
src=Path('build_model.py').read_text();exec(src[src.index('def engine_row('):src.index('base=engine_row')])
def sheet(n,h):
 s=w.create_sheet(n);s.append(h);return s
# Four coherent scenarios, explicitly subjective equal weighting for illustration only.
dc=sheet('Decision_Criteria',['Strategy','Optimistic NPV','Base NPV','Pessimistic NPV','Severe NPV','Illustrative equal-weight EV','Maximax','Maximin','Hurwicz alpha 0.5','Maximum regret'])
for strategy in range(1,4):
 r=strategy+1;row=[['A Conventional','B Efficient','C Ambitious'][strategy-1]]
 for label,cost,demand,opex,cap,growth in [('Optimistic',.9,1.15,.9,.045,.03),('Base',1,1,1,.055,.02),('Pessimistic',1.15,.85,1.1,.065,.01),('Severe',1.25,.75,1.2,.07,.005)]:
  rr=engine_row(label,120,590,cost,demand,opex,cap,strategy=strategy,growth=growth);row.append(f'=Analysis_Engine!O{rr}')
 row += [f'=AVERAGE(B{r}:E{r})',f'=MAX(B{r}:E{r})',f'=MIN(B{r}:E{r})',f'=0.5*G{r}+0.5*H{r}',f'=MAX(MAX(B$2:B$5)-B{r},MAX(C$2:C$5)-C{r},MAX(D$2:D$5)-D{r},MAX(E$2:E$5)-E{r})'];dc.append(row)
dc.append(['No investment',0,0,0,0,0,0,0,0,0]);dc['J5']='=MAX(MAX(B$2:B$5),MAX(C$2:C$5),MAX(D$2:D$5),MAX(E$2:E$5))'
dc['A7']='Scenario weights 25% each are illustrative; not evidence-based event probabilities. EV must not be interpreted as a market expectation.'
for j,(label,col,mode) in enumerate([('Expected value','F','MAX'),('Maximax','G','MAX'),('Maximin','H','MAX'),('Hurwicz','I','MAX'),('Minimax regret','J','MIN')],9):dc.cell(j,1,label);dc.cell(j,2,f'=INDEX(A2:A5,MATCH({mode}({col}2:{col}5),{col}2:{col}5,0))')
dc['L1']='Optimism alpha';dc['M1']='A';dc['N1']='B';dc['O1']='C';dc['P1']='No investment'
for r,alpha in enumerate([0,.25,.5,.75,1],2):
 dc.cell(r,12,alpha)
 for j,strategy in enumerate(range(2,5),13):dc.cell(r,j,f'=$L{r}*G{strategy}+(1-$L{r})*H{strategy}')
 dc.cell(r,16,0)
comp=sheet('Strategy_Comparison',['Strategy','Investment EUR','NPV EUR','Occupancy','Annual energy kWh','Annual CO2 tonnes','Affordable rooms','Year1 revenue','NPV vs A'])
for strategy in range(1,4):
 rr=engine_row('strategy compare','=Total_Rooms','=Base_Rent_EUR_Month',strategy=strategy);r=strategy+1
 comp.append([strategy,f'=Analysis_Engine!I{rr}',f'=Analysis_Engine!O{rr}',f'=(Analysis_Engine!J{rr}+Analysis_Engine!K{rr})/Total_Rooms',f'=Total_Rooms*Analysis_Engine!R{rr}',f'=E{r}*CO2_Factor/1000',f'=Analysis_Engine!H{rr}',f'=Analysis_Engine!L{rr}',f'=C{r}-C$2'])
comp['K1']='All energy performance, premiums and appeal factors are assumptions, not measured building results.'
comp['K2']='Sustainability energy threshold: <=2000 kWh/room/year; management policy, not legal standard.'
comp['K3']='NPV trade-off vs A includes premium, energy savings, demand appeal and additional affordable rooms.'
# Second dynamic two-way table, rent and elasticity.
two=sheet('Demand_Sensitivity',['Market rent EUR',.6,1.2,1.8,2.4])
for r,rent in enumerate([480,530,590,650,730],2):
 two.cell(r,1,rent)
 for col,el in enumerate([.6,1.2,1.8,2.4],2):rr=engine_row('rent elasticity','=Total_Rooms',rent,elasticity=el);two.cell(r,col,f'=Analysis_Engine!O{rr}')
two.conditional_formatting.add('B2:E6',ColorScaleRule(start_type='min',start_color='F8696B',end_type='max',end_color='63BE7B'))
# Additional sensitivity cases: affordable rent, elasticity, growth and terminal cap.
add=sheet('Additional_Sensitivity',['Variable','Low','High','NPV low','NPV high','Comment'])
for label,lo,hi,kw in [('Elasticity',.6,2.4,'elasticity'),('Growth',.005,.035,'growth'),('Terminal cap',.045,.07,'cap')]:
 a=engine_row(label+' low','=Total_Rooms','=Base_Rent_EUR_Month',**{kw:lo});b=engine_row(label+' high','=Total_Rooms','=Base_Rent_EUR_Month',**{kw:hi});add.append([label,lo,hi,f'=Analysis_Engine!O{a}',f'=Analysis_Engine!O{b}','Judgmental management stress endpoints'])
# Affordable rent affects its own revenue directly; exact linear NPV difference at fixed demand.
add.append(['Affordable rent',359.44,449.3,'=Project_NPV_EUR+(B5-Affordable_Rent_EUR_Month)*Affordable_Units_Count*12*(1-Management_Fee)*Analysis_Engine!N2','=Project_NPV_EUR+(C5-Affordable_Rent_EUR_Month)*Affordable_Units_Count*12*(1-Management_Fee)*Analysis_Engine!N2','80%-100% VUB benchmark policy stress; not legal threshold'])
mc=w['Monte_Carlo'];mc['J1']='Within budget (1/0)';mc['M1']='Affordability (1/0)';mc['N1']='DSCR >=1.2 (1/0)';mc['O1']='All objectives (1/0)'
# Locate each simulation engine row from its NPV formula.
for r in range(2,5002):
 rr=int(mc.cell(r,6).value.split('O')[-1]);mc.cell(r,10,f'=IF(Analysis_Engine!I{rr}<=Budget_EUR,1,0)');mc.cell(r,13,f'=IF(AND(Affordable_Units_Count>=Total_Rooms*Required_Affordable_Share,Affordable_Rent_EUR_Month<=Reference_Affordability_Rent_EUR),1,0)');mc.cell(r,14,f'=IF((Analysis_Engine!M{rr}+Total_Rooms*Replacement_Reserve)/(-PMT(Debt_Interest_Rate,Loan_Term,Analysis_Engine!I{rr}*Debt_Ratio))>=1.2,1,0)');mc.cell(r,15,f'=IF(AND(F{r}>0,J{r}=1,M{r}=1,N{r}=1,Energy_kWh_Per_Room<=2000),1,0)')
for r,label,col in [(12,'Probability within budget','J'),(13,'Probability affordability compliance','M'),(14,'Probability DSCR >=1.2','N'),(15,'Probability all objectives','O')]:mc.cell(r,8,label);mc.cell(r,9,f'=AVERAGE({col}2:{col}5001)')
for r,p in [(17,.1),(18,.5),(19,.9)]:mc.cell(r,8,f'NPV P{int(p*100)}');mc.cell(r,9,f'=PERCENTILE(F2:F5001,{p})')
mc['H20']='NPV standard deviation';mc['I20']='=STDEV(F2:F5001)'
risk=sheet('Risk_Drivers',['Uncertain factor','Pearson correlation with NPV','Meaning'])
for label,col in [('Construction factor','B'),('Demand factor','C'),('Opex factor','D'),('Exit cap','E'),('Annual growth','G')]:risk.append([label,f'=CORREL(Monte_Carlo!{col}2:{col}5001,Monte_Carlo!F2:F5001)','Association in this simulation; not causal attribution. Nonlinear effects and assumed independence matter.'])
# Exactly five requested managerial questions.
del w['Decision_Summary'];q=sheet('Management_Conclusions',['Decision question','Technique','Results','Interpretation','Management action'])
for row in [('1. Optimal rental strategy?','Demand model + finite optimization','=Optimization!I6','Best built rent; affordable rent separately capped. Rent-demand curve not calibrated.','Use conditional result only; no-build dominates if NPV <=0.'),('2. Which development strategy?','Compare A/B/C + optimize','=Optimization!I9','Strategy index 1=A,2=B,3=C. Investment cash flows remain negative if no-build chosen.','Review costs, subsidies and land pricing before construction.'),('3. Robustness under uncertainty?','5000 five-input simultaneous trials','=Monte_Carlo!I15','Joint NPV/budget/affordability/DSCR/energy success, conditional on subjective distributions.','Mitigate costs and exit risk; do not rely on expected NPV alone.'),('4. Main investment risk drivers?','Sensitivity + Pearson driver correlations','=Sensitivity_Analysis!B6','Stress ranking depends on ranges; see Risk_Drivers for statistical association.','Obtain bids, demand evidence, energy and lender quotes.'),('5. Choice when objectives conflict?','EV/Maximax/Maximin/Hurwicz/Minimax regret','=Decision_Criteria!B9','No-build has no student housing benefit; compare mandate against financial affordability.','Document required subsidy or acceptable social cost before approval.')]:q.append(row)
setup=sheet('Excel_Setup',['Feature','Exact settings / limitation'])
ref={r[0].value:r[1].coordinate for r in w['Inputs'].iter_rows(min_row=2)}
for a,b in [('Native Data Table 1','Not configured; formula equivalent in Sensitivity_Analysis. On blank sheet: A1 = Project_NPV_EUR; B1:F1 rent values; A2:A6 cost multipliers. Select A1:F6, Data > What-If Analysis > Data Table. Row input Inputs!'+ref['Base_Rent_EUR_Month']+'; Column input Inputs!'+ref['Cost_Multiplier']+'.'),('Native Data Table 2','Same layout, rent across row, elasticity down column. Row input Inputs!'+ref['Base_Rent_EUR_Month']+'; Column input Inputs!'+ref['Demand_Elasticity']+'. Output Project_NPV_EUR.'),('Solver objective','Maximize KPI_Dashboard!B2. Change Inputs!'+ref['Total_Rooms']+', '+ref['Base_Rent_EUR_Month']+', '+ref['Selected_Strategy']+'. Affordable quantity is derived from minimum policy share; affordable rent capped at benchmark.'),('Solver constraints','Rooms integer 40..140; strategy integer 1..3; market rent 400..650; Development_Costs!B10<=Inputs budget; rooms*gross area<=area capacity; affordable quantity>=10% rooms; affordable rent<=benchmark; energy<=2000 kWh/room. Need reject strategy A under this energy mandate.'),('Solver method','Evolutionary for nonsmooth MIN/INT and discrete strategy; no global continuous guarantee. Enumeration result is exact only within finite grid. Solver has NOT been executed.'),('Incomplete requested work','No construction delay or multi-year construction schedule; no tax computation; no rent-by-unit-type model; no native Data Tables/Solver execution; no Excel desktop opening; no empirical correlation calibration; no simulation IRR distribution yet.'),('Rubric status','Structure/names/formulas: implemented and locally checked. Sensitivity: dynamic equivalents implemented, native tables manual. Optimization: finite search implemented, Solver manual. Monte Carlo: 5000 five-input trials implemented, distributions subjective.'),('Three student roles','Suggested only: student 1 evidence/model; student 2 sensitivity/optimization/criteria; student 3 simulation/report. No names or contributions fabricated.')]:setup.append([a,b])
for sh,col1,col2,maxrow,title,pos in [(comp,1,3,4,'Strategy NPV (EUR)','K5'),(comp,1,6,4,'Operational carbon (tCO2e/year)','K20'),(risk,1,2,6,'Simulation risk-driver association','E2'),(dc,1,6,5,'Illustrative scenario expected NPV','L10')]:
 ch=BarChart();ch.title=title;ch.add_data(Reference(sh,min_col=col2,min_row=1,max_row=maxrow),titles_from_data=True);ch.set_categories(Reference(sh,min_col=col1,min_row=2,max_row=maxrow));sh.add_chart(ch,pos)
for sh in [dc,comp,two,add,risk,q,setup]:
 sh.freeze_panes='B2';sh.column_dimensions['A'].width=42;sh.column_dimensions['B'].width=65
w['Read_Me']['B7']='Exact finite search across 3 concepts, integer rooms 40..140 and EUR10 rent grid; add energy mandate if required. Solver setup in Excel_Setup; native Solver not run.'
w.save('intermediate/model_extended.xlsx');os.replace('intermediate/model_extended.xlsx','model.xlsx')
p=Path('recalculate.py');t=p.read_text().replace("if fn=='AVERAGE':", "if fn=='STDEV':return statistics.stdev(ns)\n if fn=='CORREL':\n  x=nums([vals[0]]);y=nums([vals[1]]);mx=statistics.mean(x);my=statistics.mean(y);return sum((a-mx)*(b-my) for a,b in zip(x,y))/math.sqrt(sum((a-mx)**2 for a in x)*sum((b-my)**2 for b in y))\n if fn=='AVERAGE':")
p.write_text(t);print('Extended model saved')
