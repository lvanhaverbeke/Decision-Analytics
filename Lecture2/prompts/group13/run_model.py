from pathlib import Path
import os,math,random,time,json,csv,statistics,datetime,shutil,zipfile,xml.etree.ElementTree as ET,itertools
import openpyxl
from openpyxl.styles import Font,PatternFill
from openpyxl.chart import BarChart,LineChart,Reference
from openpyxl.utils import get_column_letter as L
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
from document_tools import pdf
os.chdir(Path(__file__).parent);START=json.load(open('execution_started.json'))['start_epoch']
def log(s):
 with open('activity.log','a') as f:f.write(datetime.datetime.now(datetime.timezone.utc).isoformat()+' '+s+'\n')
 print(s,flush=True)
def guard():
 if time.time()-START>390:raise TimeoutError('Save deadline reached; preserve completed checkpoints and report actual trial counts.')
PROPS=[{'id':'A','commune':'Etterbeek','n':20,'acq':1100000,'works':600000,'rent':520,'charges':90,'occ':.95,'debt':.5,'rate':.045,'ready':9,'area':14},{'id':'B','commune':'Ixelles','n':30,'acq':1500000,'works':800000,'rent':500,'charges':90,'occ':.95,'debt':.5,'rate':.045,'ready':12,'area':14},{'id':'C','commune':'Auderghem','n':40,'acq':1800000,'works':1000000,'rent':480,'charges':90,'occ':.95,'debt':.5,'rate':.045,'ready':18,'area':13}]
COMMON={'Project_Rate':.08,'Equity_Rate':.1,'Tax_Rate':.25,'Rent_Growth':.02,'Expense_Growth':.025,'Fixed_Opex':12000,'Variable_Opex':1200,'Management':.06,'Replacement':100,'Duty':.125,'VAT':.21,'Soft_Fee':.15,'Contingency':.1,'Furniture':1800,'Working_Capital':30000,'Legal_Fee':25000,'Exit_Cap':.055,'Selling_Fee':.02,'Land_Share':.2,'Building_Dep_Years':33,'Works_Dep_Years':20,'Debt_Term':20,'Rent_Elasticity':1.2,'Capital_Budget':3000000,'Risk_Penalty':.25,'Seed':29102026}
def calc(p,af=1,cf=1,rf=1,occ=None,rate=None,ready=None,cap=None,growth=None,of=1,horizon=20):
 a=p['acq']*af;works=p['works']*cf*(1+COMMON['Soft_Fee'])*(1+COMMON['Contingency'])*(1+COMMON['VAT']);invest=a*(1+COMMON['Duty'])+works+p['n']*COMMON['Furniture']+COMMON['Working_Capital']+COMMON['Legal_Fee'];debt=invest*p['debt'];rate=p['rate'] if rate is None else rate;ready=p['ready'] if ready is None else ready;opening=9+12*max(0,math.ceil((ready-9)/12));occ=p['occ'] if occ is None else occ;occ=min(1,max(0,occ*math.exp(-COMMON['Rent_Elasticity']*(rf-1))));cap=COMMON['Exit_Cap'] if cap is None else cap;growth=COMMON['Rent_Growth'] if growth is None else growth
 payment=debt*rate/(1-(1+rate)**(-COMMON['Debt_Term'])) if rate else debt/COMMON['Debt_Term'];project=[-invest];equity=[-invest+debt];balance=debt;dep=0;cfrows=[];cum=equity[0];funding=-cum
 depreciation=a*(1-COMMON['Land_Share'])/COMMON['Building_Dep_Years']+works/COMMON['Works_Dep_Years'];initialbook=a+works
 for y in range(1,horizon+1):
  frac=min(1,max(0,(12*y-opening)/12));revenue=p['n']*(p['rent']*rf+p['charges'])*12*occ*(1+growth)**(y-1)*frac;opex=(COMMON['Fixed_Opex']+p['n']*COMMON['Variable_Opex'])*of*(1+COMMON['Expense_Growth'])**(y-1)*frac;noi=revenue*(1-COMMON['Management'])-opex;dp=depreciation*frac;dep+=dp;interest=balance*rate;principal=min(balance,max(0,payment-interest)*frac);balance-=principal;replace=p['n']*COMMON['Replacement']*(1+COMMON['Expense_Growth'])**(y-1)*frac;taxu=max(0,noi-dp)*COMMON['Tax_Rate'];taxe=max(0,noi-dp-interest)*COMMON['Tax_Rate'];sale=0;gain=0
  if y==horizon:
   noif=p['n']*(p['rent']*rf+p['charges'])*12*occ*(1+growth)**horizon*(1-COMMON['Management'])-(COMMON['Fixed_Opex']+p['n']*COMMON['Variable_Opex'])*of*(1+COMMON['Expense_Growth'])**horizon;sale=noif/cap*(1-COMMON['Selling_Fee']);gain=max(0,sale-(initialbook-dep))*COMMON['Tax_Rate']
  u=noi-taxu-replace+sale-gain+(COMMON['Working_Capital'] if y==horizon else 0);eq=noi-taxe-replace-interest-principal+sale-gain+(COMMON['Working_Capital']-balance if y==horizon else 0);project.append(u);equity.append(eq);cum+=eq;funding=max(funding,-cum);cfrows.append([y,frac,revenue,opex,noi,dp,taxu,taxe,interest,principal,balance,replace,sale,gain,u,eq])
 return {'project_NPV':sum(v/(1+COMMON['Project_Rate'])**t for t,v in enumerate(project)),'equity_NPV':sum(v/(1+COMMON['Equity_Rate'])**t for t,v in enumerate(equity)),'profit':sum(equity),'funding':funding,'opening':opening,'investment':invest,'project':project,'equity':equity,'rows':cfrows}
def percentile(x,p):
 x=sorted(x);k=(len(x)-1)*p;i=int(k);return x[i]+(x[min(i+1,len(x)-1)]-x[i])*(k-i)
def irr(cf):
 def pv(r):return sum(c/(1+r)**t for t,c in enumerate(cf))
 lo=-.8;hi=1
 if pv(lo)*pv(hi)>0:return 'Undefined'
 for _ in range(60):
  mid=(lo+hi)/2
  if pv(mid)>0:lo=mid
  else:hi=mid
 return (lo+hi)/2
def workbook_core():
 w=openpyxl.Workbook();w.remove(w.active)
 def sh(n,h):s=w.create_sheet(n);s.append(h);return s
 read=sh('Read_me',['Subject','Disclosure'])
 for row in [('Missing source spreadsheet','The prompt references a spreadsheet, but none was present. A/B/C are synthetic screening typologies, not listings, appraisals or real investment opportunities. No actual property can be recommended.'),('Model','20-year acquisition-to-exit horizon, annual project and equity cash flows; readiness rounded to next September letting cycle. Initial capex assumed upfront, conservatively tying capital before opening.'),('Taxes','Illustrative taxable-company model:25% standard corporate rate; building33-year and works20-year straight-line depreciation are assumptions. No loss carryforwards, interest deductibility limits, reduced rates or tax structuring. Gain taxed on net sale less simplified tax book value.'),('Financing','50% capex debt, annual interest paid even before opening; annual amortization prorated by operating fraction, no negative balances. Closing debt deducted once at exit.'),('Simulation','50000 trials for each of A/B/C; all draws and outcomes in Sim_A/B/C and CSV files. Static simulation snapshots: rerun run_model.py after editing its input definitions; editing Excel does NOT rerun the static trials.'),('Correlation','Shared Gaussian factors induce assumed cost/delay latent correlation0.49 and rate/exit-cap correlation0.36 across properties. Economic rent-demand response adds inverse response. These are assumed, not estimated market correlations.'),('Legal','A/B/C are CONDITIONALLY FEASIBLE only on assumed area; permits, fire/sanitary/accessibility/PEB and university agreements unverified. ExampleD allocates only8m2 net/person with no ability to reconfigure, so is excluded before simulation.'),('Optimization','All8 portfolios searched under P95 funding<=EUR3m. Objective expected equity NPV minus0.25*mean worst5% NPV loss. No-build allowed. Actual legally cleared portfolio contains no verified asset.'),('Conventions','EUR nominal; project discount8%, equity10%. Negative outcomes retained. No claim of native Excel desktop execution. Main cash-flow formulas dynamic, simulation/sensitivity snapshots require regeneration.')]:read.append(row)
 c=sh('Common_inputs',['Input','Value','Status / unit'])
 for name,v in COMMON.items():c.append([name,v,'Assumption except verified general duty12.5% and standard corporate rate25%; transaction applicability unverified']);w.defined_names.add(DefinedName(name,attr_text=f"'Common_inputs'!$B${c._current_row}"))
 pi=sh('Property_inputs',['ID','Municipality','Rooms','Acquisition EUR','Works EUR','Base rent/month','Charges/month','Occupancy','Debt ratio','Debt rate','Readiness months','Allocated net area/person','Legal clearance'])
 for p in PROPS:pi.append([p['id'],p['commune'],p['n'],p['acq'],p['works'],p['rent'],p['charges'],p['occ'],p['debt'],p['rate'],p['ready'],p['area'],'Unverified'])
 pi.append(['D','City of Brussels',20,600000,300000,500,90,.95,.5,.045,12,8,'EXCLUDED: assumed total allocated net area8m2, no reconfiguration possible'])
 src=sh('Sources',['Topic','Verified evidence / limitation','URL','Checked'])
 for row in [('Market','Brukot ordinary room median500 base/590 incl charges; city asking listings, not achieved rents.','https://www.brukot.be/en/student-guide/rent-stats','2026-10-29'),('University','VUB standard449.30 all inclusive2026-27; restricted institutional supply.','https://www.vub.be/en/studying-vub/practical-info-for-students/student-housing/practical-housing-information','2026-10-29'),('Quality','2026 regional student housing minimum area framework12m2/person; collective allocation and permit rules need professional review.','https://www.slrb-bghm.brussels/sites/default/files/2026-02/arr%C3%AAt%C3%A9%20exigences%20%C3%A9l%C3%A9mentaires%20s%C3%A9curit%C3%A9%20%20texte%20coordonn%C3%A9%20fr%2001012026.pdf','2026-10-29'),('Housing safety','Regional housing conformity framework; no property inspection performed.','https://be.brussels/fr/logement/location/conformite-logement/normes-et-securite','2026-10-29'),('Planning Ixelles','Municipality identifies lawful use, permits and PEB services; no project legal approval.','https://www.ixelles.be/site/81-Service-de-l-Urbanisme-et-de-l-Environnement','2026-10-29'),('Corporate tax','Standard corporate rate25% verified; actual ownership regime and special provisions unverified.','https://finances.belgium.be/fr/entreprises/impot_des_societes/declaration','2026-10-29'),('Acquisition duty','General Brussels duty12.5%, transaction VAT/duty regime unverified.','https://be.brussels/en/taxes-finances/taxes/property-related-taxes/registration-duties','2026-10-29'),('Rent legality','Property-specific reference rent review outstanding; no statutory rent ceiling invented.','https://be.brussels/fr/logement/location/bail-dhabitation/loyers-de-reference','2026-10-29')]:src.append(row)
 summary=sh('Summary',['Illustrative property','Initial investment','Project NPV','Equity NPV','Project IRR','Equity IRR','Opening month','Nominal equity profit','Peak equity funding'])
 for idx,p in enumerate(PROPS,2):
  name=p['id'];s=sh('CF_'+name,['Year','Operating fraction','Revenue','Opex','NOI','Depreciation','Project tax','Equity tax','Interest','Principal','Closing debt','Replacement','Net sale','Gain tax','Project CF','Equity CF'])
  s['R1']='Investment';s['S1']=f'=Property_inputs!D{idx}*(1+Duty)+Property_inputs!E{idx}*(1+Soft_Fee)*(1+Contingency)*(1+VAT)+Property_inputs!C{idx}*Furniture+Working_Capital+Legal_Fee';s['R2']='Opening month';s['S2']=f'=9+12*MAX(0,INT((Property_inputs!K{idx}-9+12-0.00000001)/12))';s['R3']='Annual debt payment';s['S3']=f'=-PMT(Property_inputs!J{idx},Debt_Term,S1*Property_inputs!I{idx})';s['R4']='Annual depreciation';s['S4']=f'=Property_inputs!D{idx}*(1-Land_Share)/Building_Dep_Years+Property_inputs!E{idx}*(1+Soft_Fee)*(1+Contingency)*(1+VAT)/Works_Dep_Years';s._current_row=1
  s.append([0,0,0,0,0,0,0,0,0,0,f'=$S$1*Property_inputs!I{idx}',0,0,0,'=-$S$1',f'=-$S$1*(1-Property_inputs!I{idx})'])
  for y in range(1,21):
   rr=y+2;s.append([y,f'=MIN(1,MAX(0,(A{rr}*12-$S$2)/12))',f'=Property_inputs!C{idx}*(Property_inputs!F{idx}+Property_inputs!G{idx})*12*Property_inputs!H{idx}*(1+Rent_Growth)^(A{rr}-1)*B{rr}',f'=(Fixed_Opex+Property_inputs!C{idx}*Variable_Opex)*(1+Expense_Growth)^(A{rr}-1)*B{rr}',f'=C{rr}*(1-Management)-D{rr}',f'=$S$4*B{rr}',f'=MAX(0,E{rr}-F{rr})*Tax_Rate',f'=MAX(0,E{rr}-F{rr}-I{rr})*Tax_Rate',f'=K{rr-1}*Property_inputs!J{idx}',f'=MIN(K{rr-1},MAX(0,$S$3-I{rr})*B{rr})',f'=K{rr-1}-J{rr}',f'=Property_inputs!C{idx}*Replacement*(1+Expense_Growth)^(A{rr}-1)*B{rr}',0,0,f'=E{rr}-G{rr}-L{rr}+M{rr}-N{rr}'+('+Working_Capital' if y==20 else ''),f'=E{rr}-H{rr}-L{rr}-I{rr}-J{rr}+M{rr}-N{rr}'+(f'+Working_Capital-K{rr}' if y==20 else '')])
   if y==20:
    s.cell(rr,13,f'=(Property_inputs!C{idx}*(Property_inputs!F{idx}+Property_inputs!G{idx})*12*Property_inputs!H{idx}*(1+Rent_Growth)^20*(1-Management)-(Fixed_Opex+Property_inputs!C{idx}*Variable_Opex)*(1+Expense_Growth)^20)/Exit_Cap*(1-Selling_Fee)')
    s.cell(rr,14,f'=MAX(0,M{rr}-(Property_inputs!D{idx}+Property_inputs!E{idx}*(1+Soft_Fee)*(1+Contingency)*(1+VAT)-SUM(F3:F22)))*Tax_Rate')
  b=calc(p);summary.append([name,f'=CF_{name}!S1',f'=NPV(Project_Rate,CF_{name}!O3:O22)+CF_{name}!O2',f'=NPV(Equity_Rate,CF_{name}!P3:P22)+CF_{name}!P2',f'=IFERROR(IRR(CF_{name}!O2:O22),"Undefined")',f'=IFERROR(IRR(CF_{name}!P2:P22),"Undefined")',f'=CF_{name}!S2',f'=SUM(CF_{name}!P2:P22)',b['funding']])
 for p in PROPS:sh('Sim_'+p['id'],['Trial','Acquisition factor','Works factor','Rent factor','Occupancy raw','Debt rate','Readiness months','Exit cap','Annual rent growth','Opex factor','Opening month','Project NPV','Equity NPV','Nominal equity profit','Peak equity funding'])
 for s in w:
  s.freeze_panes='B2';s.column_dimensions['A'].width=34
  for c in s[1]:c.fill=PatternFill('solid',fgColor='18364D');c.font=Font(color='FFFFFF',bold=True)
  for col in range(2,min(s.max_column,20)+1):s.column_dimensions[L(col)].width=20
 w.calculation=CalcProperties(fullCalcOnLoad=True);w.save('model.xlsx');shutil.copy2('model.xlsx','intermediate/model_core.xlsx');return w
w=workbook_core();log('CORE SAVED: model.xlsx and intermediate/model_core.xlsx; synthetic inputs, official sources, tax/debt/year20 formulas and missing attachment disclosure.')
# Small core workbook can be fully evaluated before adding large static simulation datasets.
text=Path('../groupAube/recalculate.py').read_text();Path('recalculate.py').write_text(text)
import subprocess,sys
with open('intermediate/core_formula_validation.txt','w') as f:subprocess.run([sys.executable,'recalculate.py'],stdout=f,stderr=subprocess.STDOUT,check=True,timeout=30)
cached=openpyxl.load_workbook('model.xlsx',data_only=True)
for idx,p in enumerate(PROPS,2):
 b=calc(p);assert abs(cached['Summary'].cell(idx,3).value-b['project_NPV'])<.001;assert abs(cached['Summary'].cell(idx,4).value-b['equity_NPV'])<.001
log('VALIDATED base financial formulas against independent Python annual cash-flow engine, for all3 properties. Beginning50000 paired trials per conditionally feasible example.')
rng=random.Random(COMMON['Seed']);outputs={p['id']:{'eq':[],'pr':[],'profit':[],'fund':[]} for p in PROPS};handles={};writers={}
headers=['Trial','Acquisition factor','Works factor','Rent factor','Occupancy raw','Debt rate','Readiness months','Exit cap','Annual rent growth','Opex factor','Opening month','Project NPV','Equity NPV','Nominal equity profit','Peak equity funding']
for p in PROPS:
 f=open('intermediate/sim_'+p['id']+'.csv','w',newline='');handles[p['id']]=f;writers[p['id']]=csv.writer(f);writers[p['id']].writerow(headers)
def phi(z):return .5*(1+math.erf(z/math.sqrt(2)))
def tri(u,a,m,b):
 cut=(m-a)/(b-a);return a+math.sqrt(u*(b-a)*(m-a)) if u<cut else b-math.sqrt((1-u)*(b-a)*(b-m))
for trial in range(1,50001):
 zcost=rng.gauss(0,1);zcredit=rng.gauss(0,1);af=tri(rng.random(),.95,1,1.1)
 for p in PROPS:
  cf=tri(phi(.7*zcost+math.sqrt(.51)*rng.gauss(0,1)),.85,1,1.25);ready=tri(phi(.7*zcost+math.sqrt(.51)*rng.gauss(0,1)),6,12,24);rate=tri(phi(.6*zcredit+.8*rng.gauss(0,1)),.03,.045,.07);cap=tri(phi(.6*zcredit+.8*rng.gauss(0,1)),.045,.055,.075);rf=tri(rng.random(),.85,1,1.1);occ=tri(rng.random(),.75,.95,.99);growth=tri(rng.random(),.005,.02,.035);of=tri(rng.random(),.9,1,1.2)
  result=calc(p,af,cf,rf,occ,rate,ready,cap,growth,of);writers[p['id']].writerow([trial,af,cf,rf,occ,rate,ready,cap,growth,of,result['opening'],result['project_NPV'],result['equity_NPV'],result['profit'],result['funding']]);o=outputs[p['id']]
  for key,field in [('eq','equity_NPV'),('pr','project_NPV'),('profit','profit'),('fund','funding')]:o[key].append(result[field])
 if trial%10000==0:
  for f in handles.values():f.flush()
  log(f'CHECKPOINT: {trial} complete paired trials per property, {trial*3} total. Raw draws/outcomes saved in intermediate/sim_A.csv, sim_B.csv, sim_C.csv.');guard()
for f in handles.values():f.close()
log('SIMULATION COMPLETE: exactly50000 per property,150000 total; costs, rent, occupancy, debt pricing, readiness, exit cap, growth and opex vary together; negative outcomes retained.')
# Summaries, portfolio search, sensitivity and plots; raw trial worksheets streamed after small workbook calculation.
w=openpyxl.load_workbook('model.xlsx');st=w.create_sheet('Risk_summary');st.append(['Property','Trials','Mean equity NPV','P5','Median','P95','P(NPV<0)','P(target NPV>=0)','P(nominal loss)','Mean nominal profit','P95 funding','CVaR5 loss'])
metrics={}
for p in PROPS:
 o=outputs[p['id']];eq=o['eq'];tail=sorted(eq)[:2500];m={'trials':len(eq),'mean':statistics.mean(eq),'P5':percentile(eq,.05),'P50':percentile(eq,.5),'P95':percentile(eq,.95),'P_negative':sum(x<0 for x in eq)/50000,'P_target':sum(x>=0 for x in eq)/50000,'P_nominal_loss':sum(x<0 for x in o['profit'])/50000,'mean_profit':statistics.mean(o['profit']),'P95_funding':percentile(o['fund'],.95),'CVaR5':statistics.mean(max(0,-x) for x in tail)};metrics[p['id']]=m;st.append([p['id']]+list(m.values()))
port=w.create_sheet('Portfolio_optimization');port.append(['Portfolio','Mean equity NPV','CVaR5 loss','P95 funding','Budget pass','Risk adjusted score','P(loss)','Actual legal clearance'])
portfolio=[]
for mask in range(8):
 ids=[p['id'] for j,p in enumerate(PROPS) if mask&(1<<j)];eq=[sum(outputs[x]['eq'][t] for x in ids) for t in range(50000)];fund=[sum(outputs[x]['fund'][t] for x in ids) for t in range(50000)];mean=statistics.mean(eq);cvar=statistics.mean(max(0,-x) for x in sorted(eq)[:2500]);fund95=percentile(fund,.95);feasible=fund95<=COMMON['Capital_Budget'];score=mean-COMMON['Risk_Penalty']*cvar;row=['+'.join(ids) or 'No investment',mean,cvar,fund95,feasible,score,sum(x<0 for x in eq)/50000,'Unverified' if ids else 'No acquisition'];port.append(row);portfolio.append(row)
best=max((r for r in portfolio if r[4]),key=lambda r:r[5]);port['A12']='Exact enumeration of all8 subsets with paired trial outcomes. Risk penalty0.25 and P95 funding budgetEUR3m are assumed. Real legal-clearance constraint leaves only no investment.'
se=w.create_sheet('Sensitivity');se.append(['Property','Driver','Low','High','Equity NPV low','Equity NPV high','Swing'])
for p in PROPS:
 for label,lo,hi,kw in [('Acquisition factor',.8,1.2,'af'),('Works factor',.85,1.25,'cf'),('Rent factor',.85,1.1,'rf'),('Occupancy',.75,.99,'occ'),('Borrowing rate',.03,.07,'rate'),('Readiness months',6,24,'ready'),('Exit cap',.045,.075,'cap')]:
  low=calc(p,**{kw:lo})['equity_NPV'];high=calc(p,**{kw:hi})['equity_NPV'];se.append([p['id'],label,lo,hi,low,high,abs(high-low)])
be=w.create_sheet('Horizon_and_break_even');be.append(['Property','Horizon','Project NPV','Equity NPV','Max acquisition at project NPV0','Readiness9 equity NPV','Readiness12 equity NPV','Legal status'])
for p in PROPS:
 if calc(dict(p,acq=0))['project_NPV']<0:ceiling='No nonnegative purchase price supports target'
 else:
  lo=0;hi=p['acq']*3
  for _ in range(45):
   mid=(lo+hi)/2
   if calc(dict(p,acq=mid))['project_NPV']>0:lo=mid
   else:hi=mid
  ceiling=(lo+hi)/2
 for h in [10,15,20]:r=calc(p,horizon=h);be.append([p['id'],h,r['project_NPV'],r['equity_NPV'],ceiling,calc(p,ready=9)['equity_NPV'],calc(p,ready=12)['equity_NPV'],'Not a real listing; approval pending'])
hist=w.create_sheet('Distribution');hist.append(['Property','NPV upper bound','Count','CDF quantile','Equity NPV'])
for p in PROPS:
 eq=outputs[p['id']]['eq'];mn=min(eq);mx=max(eq);prev=0
 for j in range(1,31):upper=mn+(mx-mn)*j/30;count=sum(x<=upper for x in eq);hist.append([p['id'],upper,count-prev,j/30,percentile(eq,j/30)]);prev=count
for row1,row2,title in [(2,31,'A'),(32,61,'B'),(62,91,'C')]:
 ch=BarChart();ch.title=title+' equity NPV distribution (EUR)';ch.add_data(Reference(hist,min_col=3,min_row=row1,max_row=row2));ch.set_categories(Reference(hist,min_col=2,min_row=row1,max_row=row2));hist.add_chart(ch,'G'+str(row1))
 ch=LineChart();ch.title=title+' equity NPV percentile curve';ch.add_data(Reference(hist,min_col=5,min_row=row1,max_row=row2));hist.add_chart(ch,'Q'+str(row1))
ch=BarChart();ch.title='Expected equity NPV by synthetic property (EUR)';ch.add_data(Reference(st,min_col=3,min_row=1,max_row=4),titles_from_data=True);ch.set_categories(Reference(st,min_col=1,min_row=2,max_row=4));st.add_chart(ch,'N2')
ranking=w.create_sheet('Recommendation');ranking.append(['Finding','Result / condition'])
ranking.append(['Financial ranking',' > '.join(sorted(metrics,key=lambda k:metrics[k]['mean'],reverse=True))]);ranking.append(['Best assumed risk/budget portfolio',best[0]]);ranking.append(['Real investment recommendation','DO NOT COMMIT: missing source spreadsheet and no verified property, price, planning or compliance assessment. Screening ranks are conditional examples only.']);ranking.append(['Excluded propertyD','Total allocated net area8m2/person, no feasible reconfiguration by stated synthetic assumption; below12m2 minimum framework, excluded with zero simulations.']);ranking.append(['Required diligence','Supply actual spreadsheet, addresses, floor plans, lawful-use/permit files, sanitary/fire/PEB assessments, contractor and lender terms, tax structure and comparable achieved rents.']);ranking.append(['Probability interpretation','Subjective distribution assumptions, not statistically estimated market loss probabilities. P(target) defined as NPV>=0 at10% equity hurdle, not simulated IRR threshold.']);ranking.append(['Reinvestment limitation','Static simultaneous acquisition portfolio. Dynamic reinvestment option and exact transaction tax structuring were not implemented.'])
for s in [st,port,se,be,hist,ranking]:
 s.freeze_panes='B2';s.column_dimensions['A'].width=35
 for col in range(2,min(s.max_column,15)+1):s.column_dimensions[L(col)].width=22
w.save('.model.enhanced.xlsx');os.replace('.model.enhanced.xlsx','model.xlsx')
with open('intermediate/final_formula_validation.txt','w') as f:subprocess.run([sys.executable,'recalculate.py'],stdout=f,stderr=subprocess.STDOUT,check=True,timeout=30)
log('Summaries saved; exact8-portfolio enumeration and sensitivity/horizon/price ceilings completed. All main workbook formulas evaluated with zero errors; streaming150000 raw trial rows into simulation worksheets.')
# Stream raw numerical trial records into existing worksheet XML, preserving cached formulas and all package relationships.
wb=openpyxl.load_workbook('model.xlsx',read_only=True,data_only=True);sheetids={name:wb.sheetnames.index(name)+1 for name in ['Sim_A','Sim_B','Sim_C']};wb.close();ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
with zipfile.ZipFile('model.xlsx') as src,zipfile.ZipFile('.model.full.xlsx','w',zipfile.ZIP_DEFLATED) as dst:
 targets={f'xl/worksheets/sheet{v}.xml':k for k,v in sheetids.items()}
 for item in src.infolist():
  if item.filename not in targets:dst.writestr(item,src.read(item.filename));continue
  prop=targets[item.filename][-1]
  with dst.open(item.filename,'w') as out:
   out.write(f'<?xml version="1.0" encoding="UTF-8"?><worksheet xmlns="{ns}"><dimension ref="A1:O50001"/><sheetViews><sheetView workbookViewId="0"><pane xSplit="1" ySplit="1" topLeftCell="B2" state="frozen"/></sheetView></sheetViews><sheetData>'.encode())
   with open('intermediate/sim_'+prop+'.csv',newline='') as f:
    for rr,row in enumerate(csv.reader(f),1):
     cells=[]
     for col,v in enumerate(row,1):
      coord=L(col)+str(rr)
      if rr==1:cells.append(f'<c r="{coord}" t="inlineStr"><is><t>{v}</t></is></c>')
      else:cells.append(f'<c r="{coord}"><v>{float(v):.15g}</v></c>')
     out.write((f'<row r="{rr}">'+''.join(cells)+'</row>').encode())
   out.write(b'</sheetData><autoFilter ref="A1:O50001"/></worksheet>')
os.replace('.model.full.xlsx','model.xlsx');guard()
# Validate all150000 stored rows by streaming CSV; worksheet XML structure and package integrity checked without loading all cells.
for p in PROPS:
 with open('intermediate/sim_'+p['id']+'.csv',newline='') as f:
  rows=csv.reader(f);next(rows);count=0
  for row in rows:
   count+=1
   if count<=10:
    vals=list(map(float,row));z=calc(p,*vals[1:10]);assert abs(z['equity_NPV']-vals[12])<.001
  assert count==50000
with zipfile.ZipFile('model.xlsx') as z:
 assert z.testzip() is None
 for path in targets:
  with z.open(path) as f:
   count=sum(1 for event,node in ET.iterparse(f,events=['end']) if node.tag=='{'+ns+'}row')
  assert count==50001
Path('intermediate/metrics.json').write_text(json.dumps(metrics,indent=2));Path('intermediate/portfolio_results.json').write_text(json.dumps(portfolio,indent=2));Path('intermediate/input_snapshot.json').write_text(json.dumps({'properties':PROPS,'common':COMMON},indent=2))
report=[('Brussels student housing - conditional research screening',['Executive conclusion: no actual property purchase is recommended. The source spreadsheet is missing and none of the example sites has legal, pricing or condition verification. Synthetic A/B/C illustrate the decision-support method only.',f'Exactly150000 simulations completed:50000 per conditionally feasible example. An8m2 no-reconfiguration exampleD is excluded before simulation. Financial portfolio score selects {best[0]} under the assumed risk and funding constraints.','The analysis emphasizes expected economic value, downside cash losses, peak capital at risk, time to letting and compliance conditions. It does not manufacture live opportunities or market-calibrated success probabilities.']),('Comparative investment ranking',[k+': '+json.dumps(metrics[k]) for k in sorted(metrics,key=lambda k:metrics[k]['mean'],reverse=True)]+['Metrics are equity after simplified corporate/gain taxes. P(target) is nonnegative equity NPV at10% hurdle. Positive expected NPV alone does not imply acceptable downside or legal approval.']),('Financial projections and timing',['Acquisition, works, fees, assumed duty and VAT, furniture and liquidity are funded at time0. Annual rent, variable and fixed costs, taxes, replacement capex, loan amortization and terminal property sale are separate.','Readiness is rounded to the next September letting cycle. A project ready after September may miss almost a full academic year. Opening month21 means only one quarter of year2 generates rent; horizon still ends at year20.','Standard corporate rate25% and general Brussels acquisition duty12.5% were verified from official sources. Depreciation lives, tax basis, company regime and interest deduction treatment are simplified assumptions; no loss carryforwards.','Property value is capitalized year21 NOI at simulated exit cap. Income growth and exit cap capture appreciation; an additional appreciation uplift is not double-counted. Debt is repaid once at exit.']),('Simulation, dependence and risk',['Nine uncertain quantities vary simultaneously: acquisition cost, works, rent, occupancy, rate, readiness, exit cap, rent growth and operating cost. Seed29102026; all150000 input/output rows in Excel and CSV.','Triangular min/mode/max: acquisition0.95/1/1.1; works0.85/1/1.25; rent0.85/1/1.1; occupancy0.75/0.95/0.99; rate3/4.5/7%; readiness6/12/24 months; exit cap4.5/5.5/7.5%; growth0.5/2/3.5%; opex0.9/1/1.2. These are assumptions, not estimated distributions.','Shared cost/delay and credit/exit factors create assumed cross-property dependence. Portfolio outcomes use paired same-trial values rather than adding unrelated percentiles. All valid negative outcomes retained.','Main dynamic Excel cash-flow formulas are checked; simulation and sensitivity outputs are static and must be regenerated from the supplied source script after changing input definitions.']),('Optimization and managerial decision',[str(row) for row in portfolio]+['Objective: mean equity NPV minus0.25 times the mean worst5% NPV loss. Constraint:95th percentile funding<=EUR3m. Both are investor assumptions. All8 subsets including no investment were evaluated.','Legal approval is mandatory. Under actual evidence available, no candidate is legally cleared; this overrides financial ranking. Dynamic reinvestment was not implemented.','Compare maximum supportable acquisition prices, readiness9 versus12 month outcomes and10/15/20-year horizons in the workbook before negotiating a real asset.']),('Sources and outstanding diligence',['Brukot market asking rents: https://www.brukot.be/en/student-guide/rent-stats','2026 quality text: https://be.brussels/fr/logement/location/conformite-logement/normes-et-securite','Corporate tax: https://finances.belgium.be/fr/entreprises/impot_des_societes/declaration','Acquisition duty: https://be.brussels/en/taxes-finances/taxes/property-related-taxes/registration-duties','Ixelles planning: https://www.ixelles.be/site/81-Service-de-l-Urbanisme-et-de-l-Environnement','All checked29 October2026. Obtain actual spreadsheet, addresses, measured net areas, permitted use, fire/sanitary/PEB assessments, rent comparables, quotes, financing and tax advice.'])]
pdf('Investment_Report.pdf',report);Path('completion_summary.txt').write_text('\n'.join(report[0][1]+report[1][1]));shutil.copy2('model.xlsx','intermediate/model_validated.xlsx')
validation={'source_spreadsheet':'MISSING','trials_per_property':50000,'total_trials':150000,'excluded_D_trials':0,'main_formulas':'Evaluated zero errors; independently compared with Python DCF','trial_reproduction':'First10 per property reproduced from stored draws','stored_rows':'Exactly50000 simulation rows per worksheet verified by XML stream','optimization':'All8 portfolios enumerated with paired joint scenarios','limitations':['No actual property data or legal clearance','Synthetic costs, rents, demand and shock parameters are assumptions','Static simulation and sensitivity; regeneration required','Simplified tax basis/loss/interest rules','No dynamic reinvestment strategy','No native Excel desktop test','Report charts not reproduced; workbook charts provided']};Path('intermediate/validation.json').write_text(json.dumps(validation,indent=2));guard()
a=json.load(open('execution_started.json'));a.update(end_epoch=time.time(),elapsed_seconds=time.time()-START,completed=True,outputs=['model.xlsx','Investment_Report.pdf','completion_summary.txt'],trials_per_property=50000,total_trials=150000,limitations=validation['limitations']);Path('execution_summary.json').write_text(json.dumps(a,indent=2));log('COMPLETE '+json.dumps(a))
