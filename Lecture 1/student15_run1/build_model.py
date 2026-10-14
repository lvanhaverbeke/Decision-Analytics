from pathlib import Path
import os,json,math,zipfile,xml.etree.ElementTree as ET
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import LineChart,Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.properties import CalcProperties
R=Path(__file__).resolve().parent
if (R/'SAVE_NOW.txt').exists() and (R/'model.xlsx').exists():
 w=load_workbook(R/'model.xlsx');w.save(R/'model.tmp.xlsx');os.replace(R/'model.tmp.xlsx',R/'model.xlsx');print('SAVE_NOW: existing model saved; stopped');raise SystemExit
E=json.load(open(R/'evidence.json')); w=Workbook();w.remove(w.active); cache={}
for name in ['Executive report','Inputs','Market evidence','Development budget','Cash flows','Sensitivity','Regulations','Sources','Methods']:w.create_sheet(name)
def put(sh,cell,v,c=None):
 w[sh][cell]=v
 if isinstance(v,str) and v.startswith('=') and c is not None:cache[(sh,cell)]=c
S='Inputs'; data=[
('Rooms',100,'rooms','ASSUMPTION','Design brief required'),
('Market all-in monthly price',600,'EUR/room/month','OFFICIAL BENCHMARK','S1 p55; survey 2023–24, all typologies; not a 2026 quotation'),
('Monthly utility cost',75,'EUR/occupied room/month','ASSUMPTION','Obtain energy/water/service quotes'),
('Utilities included? (1=yes; 0=separate)',1,'binary','POLICY','All-in affordability is unchanged by billing mode'),
('Affordable room share',.2,'fraction','POLICY','20% illustrative project commitment'),
('Housing share of student budget',.3,'fraction','POLICY','30% illustrative; no official student affordability threshold found'),
('Student housing monthly budget',833.5,'EUR/month','OFFICIAL BENCHMARK','S1 p110; mean spending budget, not household income'),
('Paid months per year',12,'months','ASSUMPTION','S3 maximum lease 12 months does not ensure collection for 12 months'),
('Occupancy / collection factor',.95,'fraction','ASSUMPTION','Vacancy and nonpayment combined; not official occupancy'),
('Gross floor area per room',30,'m2/room','ASSUMPTION','Includes common space; architect must verify'),
('Construction unit cost ex VAT',2200,'EUR/m2','ASSUMPTION','No official Brussels PBSA EUR/m2 quotation verified'),
('Land and acquisition all-in',1500000,'EUR','ASSUMPTION','Include land purchase tax and fees; obtain site valuation'),
('Professional fees / hard cost',.12,'fraction','ASSUMPTION','Design, permits and project management'),
('Contingency / hard cost plus fees',.1,'fraction','ASSUMPTION','Cost overrun allowance'),
('Effective unrecoverable VAT on works',.21,'fraction','ASSUMPTION','Conservative budget allowance, not tax advice; verify entity and regime'),
('Furniture per room all-in',2000,'EUR/room','ASSUMPTION','Supplier quote required'),
('Annual maintenance / hard cost',.01,'fraction','ASSUMPTION','Cash maintenance; obtain lifecycle plan'),
('Annual other operating expenses',100000,'EUR/year','ASSUMPTION','Management, insurance, property tax, common areas; replace with quotes'),
('Nominal annual rent growth',.02,'fraction','ASSUMPTION','Forecast, not actual statutory indexation'),
('Nominal annual cost growth',.02,'fraction','ASSUMPTION','Forecast; future inflation is unknown'),
('Nominal required return',.07,'fraction','INVESTOR INPUT','Unlevered pre-income-tax hurdle'),
('Operating years',30,'years','INVESTOR INPUT','Fixed model horizon: 30 years'),
('Years to opening',1,'years','ASSUMPTION','Whole-year delay; all development spending at time zero'),
('Terminal capitalization rate',.06,'fraction','ASSUMPTION','Sale on next-year NOI; not official market yield'),
('Terminal selling cost',.02,'fraction','ASSUMPTION','Deducted from terminal gross sale'),
('Deposit months requested',2,'months','POLICY / LEGAL CAP','S2 maximum two months; average actual deposit unknown'),
('Major replacement year',15,'operating year','ASSUMPTION','Furniture/plant replacement'),
('Major replacement amount year-1 prices',300000,'EUR','ASSUMPTION','Escalated with cost growth; distinct from recurring maintenance'),
('Include terminal sale (1/0)',1,'binary','INVESTOR INPUT','Test zero for no-sale downside')]
w[S].append(['Input','Value','Unit','Classification','Source / interpretation'])
for row in data:w[S].append(row)
# row lookup, formulas and cached values
V={r[0]:r[1] for r in data}; idx={r[0]:i+2 for i,r in enumerate(data)}
def ref(k):return "Inputs!$B$"+str(idx[k])
n=100;area=3000;hard=6600000;fees=hard*.12;cont=(hard+fees)*.1;vat=(hard+fees+cont)*.21;capex=1500000+hard+fees+cont+vat+200000
b='Development budget';w[b].append(['Component','EUR','Method'])
for cell,f,val,label in [('B2',f'={ref("Rooms")}*{ref("Gross floor area per room")}',area,'Gross floor area (m2)'),('B3',f'=B2*{ref("Construction unit cost ex VAT")}',hard,'Hard construction'),('B4',f'=B3*{ref("Professional fees / hard cost")}',fees,'Professional fees'),('B5',f'=(B3+B4)*{ref("Contingency / hard cost plus fees")}',cont,'Contingency'),('B6',f'=(B3+B4+B5)*{ref("Effective unrecoverable VAT on works")}',vat,'Unrecoverable works VAT allowance'),('B7',f'={ref("Land and acquisition all-in")}',1500000,'Land and acquisition'),('B8',f'={ref("Rooms")}*{ref("Furniture per room all-in")}',200000,'Furniture'),('B9','=SUM(B3:B8)',capex,'Total initial development spend')]:
 put(b,cell,f,val);put(b,'A'+cell[1:],label)
w[b]['C9']='All paid at t=0; financing excluded. Replace budget assumptions before investment approval.'
m='Market evidence';w[m].append(['Measure','Value','Interpretation'])
market=[('All-in median rent',600,'S1 p55; 2023–24 survey'),('Shared housing median/person',525,'S1 p56; distinct product, not lower end of market range'),('Mean student-housing monthly budget',833.5,'S1 p110; includes rent and other spending'),('Estimated housing demand',37200,'S4 p26; approximation, excludes promotion sociale'),('Dedicated housing units',11305,'S4 p28; 2025 inventory; not all housing occupied by students'),('Dedicated coverage ratio','=B6/B5','Not vacancy or occupancy'),('Demand less dedicated inventory','=B5-B6','NOT a measured shortage: ordinary rental housing also serves students'),('Live available rooms','Unknown','Inventory does not provide current vacancies'),('Observed average deposit','Unknown','Legal maximum must not be represented as an average'),('Official affordable student threshold','Not identified','User policy threshold shown in Inputs'),('Current rent distribution / range','Not verified','525 and 600 are comparable benchmarks only, not minimum/maximum prices')]
for row in market:w[m].append(row)
cache[(m,'B7')]=11305/37200;cache[(m,'B8')]=37200-11305
# derived operating economics
s='Executive report';w[s].append(['Brussels student accommodation | decision model','Value / assessment'])
w[s].append(['Decision','DEFER capital commitment. Illustrative project does not meet the assumed required return; project costs remain unverified.'])
w[s].append(['Evidence and scope','Official regional benchmarks + explicitly marked project assumptions. Nominal EUR; unlevered, before income tax; 30 operating years.'])
aff=833.5*.3; blended=.8*600+.2*aff;base=blended-75
for row,label,f,val in [(5,'Policy affordable all-in threshold',f'={ref("Housing share of student budget")}*{ref("Student housing monthly budget")}',aff),(6,'Affordable rooms',f'={ref("Rooms")}*{ref("Affordable room share")}',20),(7,'Blended all-in monthly price',f'=(1-{ref("Affordable room share")})*{ref("Market all-in monthly price")}+{ref("Affordable room share")}*MIN(B5,{ref("Market all-in monthly price")})',blended),(8,'Monthly rent collection incl bundled charges',f'=B7-(1-{ref("Utilities included? (1=yes; 0=separate)")})*{ref("Monthly utility cost")}',blended),(9,'Separate monthly utility charge',f'=(1-{ref("Utilities included? (1=yes; 0=separate)")})*{ref("Monthly utility cost")}',0),(10,'Maximum deposit / market room',f'={ref("Deposit months requested")}*MAX(0,{ref("Market all-in monthly price")}-{ref("Monthly utility cost")})',1050),(11,'Initial investment',"='Development budget'!B9",capex)]:put(s,'A'+str(row),label);put(s,'B'+str(row),f,val)
c='Cash flows';headers=['Year t','Operating year','Rent collected','Utility recovery','Total receipts','Utility expense','Maintenance','Other operating','NOI','Replacement capex','Terminal net sale','Project cash flow','Discount factor','PV','Cumulative cash','Cumulative PV','Operating-only cumulative','PBP crossing','Discounted PBP crossing']
w[c].append(headers);cash=[];cum=-capex;cumop=-capex;cumdv=-capex;pbs=[];dpbs=[]
for t in range(32):
 r=t+2;op=max(0,t-1);active=t>1;g=1.02**max(0,op-1);rec=n*12*.95*blended*g if active else 0;utility=n*12*.95*75*g if active else 0;maint=hard*.01*g if active else 0;other=100000*g if active else 0;noi=rec-utility-maint-other;rep=300000*g if op==15 else 0;term=(noi*1.02/.06)*.98 if t==31 else 0;cf=-capex if t==0 else noi-rep+term;pv=cf/1.07**t;old=cum;olddv=cumdv
 if t: cum+=cf;cumdv+=pv;cumop+=noi-rep
 cash.append(cf);pb=t-1-old/cf if t>0 and old<0<=cum and cf>0 else 0;dpb=t-1-olddv/pv if t>0 and olddv<0<=cumdv and pv>0 else 0
 if pb:pbs.append(pb)
 if dpb:dpbs.append(dpb)
 vals={'A':t,'B':op,'C':rec,'D':0,'E':rec,'F':utility,'G':maint,'H':other,'I':noi,'J':rep,'K':term,'L':cf,'M':1/1.07**t,'N':pv,'O':cum,'P':cumdv,'Q':cumop,'R':pb,'S':dpb}
 forms={'B':f'=MAX(0,A{r}-{ref("Years to opening")})','C':f'=IF(B{r}>0,{ref("Rooms")}*{ref("Paid months per year")}*{ref("Occupancy / collection factor")}*\'Executive report\'!B8*(1+{ref("Nominal annual rent growth")})^(B{r}-1),0)','D':f'=IF(B{r}>0,{ref("Rooms")}*{ref("Paid months per year")}*{ref("Occupancy / collection factor")}*\'Executive report\'!B9*(1+{ref("Nominal annual rent growth")})^(B{r}-1),0)','E':f'=C{r}+D{r}','F':f'=IF(B{r}>0,{ref("Rooms")}*{ref("Paid months per year")}*{ref("Occupancy / collection factor")}*{ref("Monthly utility cost")}*(1+{ref("Nominal annual cost growth")})^(B{r}-1),0)','G':f'=IF(B{r}>0,\'Development budget\'!B3*{ref("Annual maintenance / hard cost")}*(1+{ref("Nominal annual cost growth")})^(B{r}-1),0)','H':f'=IF(B{r}>0,{ref("Annual other operating expenses")}*(1+{ref("Nominal annual cost growth")})^(B{r}-1),0)','I':f'=E{r}-SUM(F{r}:H{r})','J':f'=IF(B{r}={ref("Major replacement year")},{ref("Major replacement amount year-1 prices")}*(1+{ref("Nominal annual cost growth")})^(B{r}-1),0)','K':f'=IF(B{r}={ref("Operating years")},(E{r}*(1+{ref("Nominal annual rent growth")})-(F{r}+G{r}+H{r})*(1+{ref("Nominal annual cost growth")}))/{ref("Terminal capitalization rate")}*(1-{ref("Terminal selling cost")})*{ref("Include terminal sale (1/0)")},0)','L':f'=IF(A{r}=0,-\'Development budget\'!B9,I{r}-J{r}+K{r})','M':f'=1/(1+{ref("Nominal required return")})^A{r}','N':f'=L{r}*M{r}','O':f'=SUM($L$2:L{r})','P':f'=SUM($N$2:N{r})','Q':f'=-\'Development budget\'!B9+SUM($I$2:I{r})-SUM($J$2:J{r})'}
 if t:
  forms['R']=f'=IF(AND(O{r-1}<0,O{r}>=0,L{r}>0),A{r}-1-O{r-1}/L{r},0)';forms['S']=f'=IF(AND(P{r-1}<0,P{r}>=0,N{r}>0),A{r}-1-P{r-1}/N{r},0)'
 for col,val in vals.items():put(c,col+str(r),forms.get(col,val),val)
npv=sum(x/1.07**t for t,x in enumerate(cash))
def irr(flows):
 lo=-.9;hi=1
 for _ in range(150):
  mid=(lo+hi)/2;v=sum(x/(1+mid)**t for t,x in enumerate(flows))
  if v>0:lo=mid
  else:hi=mid
 return (lo+hi)/2
ir=irr(cash)
for row,label,f,val in [(13,'Net present value (NPV)',"=NPV(Inputs!B22,'Cash flows'!L3:L33)+'Cash flows'!L2",npv),(14,'Internal rate of return (IRR)',"=IFERROR(IRR('Cash flows'!L2:L33),\"No IRR\")",ir),(15,'Simple payback incl sale (years)',"=IF(SUM('Cash flows'!R3:R33)=0,\"Not recovered\",SUM('Cash flows'!R3:R33))",pbs[0] if pbs else 'Not recovered'),(16,'Discounted payback incl sale (years)',"=IF(SUM('Cash flows'!S3:S33)=0,\"Not recovered\",SUM('Cash flows'!S3:S33))",dpbs[0] if dpbs else 'Not recovered'),(17,'Operating-only recovery by horizon',"=IF('Cash flows'!Q33>=0,\"Recovered\",\"Not recovered\")",'Recovered' if cumop>=0 else 'Not recovered'),(18,'First operating-year NOI',"='Cash flows'!I4",cash[2]),(19,'NPV break-even initial spend',"=B11+B13",capex+npv)]:put(s,'A'+str(row),label);put(s,'B'+str(row),f,val)
put(s,'A12','Average base rent excluding utility charges');put(s,'B12',f'=B7-{ref("Monthly utility cost")}',base)
texts={21:('Interpretation','The default case mixes 80% market rooms with 20% policy affordable rooms. Results are illustrative, not an official forecast. Negative NPV and IRR below the hurdle argue against this budget.'),22:('Affordability','A 30% housing share of the reported student-housing budget gives a policy ceiling of EUR250.05 including utilities. Spending budget is not income; this is a chosen policy, not a legal threshold. Compare EUR525 shared median and EUR600 all-in median (S1).'),23:('Advantages','Adds purpose-built beds, may reduce pressure on ordinary housing and deliver efficient, accessible accommodation. Campus proximity and transparent all-in charges can support tenant appeal. These benefits do not establish profitable occupancy.'),24:('Drawbacks','Land and build cost uncertainty; permits and construction delays; annual turnover and summer vacancies; affordable rents reduce receipts; energy, repair and major replacement exposure; competition and terminal valuation risk.'),25:('Location recommendation','Shortlist sites near target campuses and transit, then compare acquisition price, zoning, walking time, safety and local competing pipeline. Ixelles/Etterbeek, central Brussels, Alma and Erasme/CERIA are candidate campus areas, not an official location ranking (S1 p8).'),26:('Investment gates','Obtain site valuation, quantity-surveyor budget, tax treatment, operator cost plan, current rent comparables, vacancies, planning/fire/energy approvals and campus-specific demand evidence. Proceed only if updated NPV >= 0, IRR >= hurdle and the affordable share remains deliverable.'),27:('Material limits','Income tax, debt, interest, construction payment phasing, working capital and grants are excluded. Property tax is an unverified component of other operating expenses. Deposits are refundable liabilities, excluded from revenue and project returns.'),28:('Timing and payback','All cost at t=0; opening after one full year; first receipts at t=2; sale at t=31. Payback interpolates the crossing year and can depend on sale; operating-only recovery is separately shown. Annual timing is a simplification.'),29:('Use','Blue inputs are editable assumptions; green inputs are official historical benchmarks. Workbook recalculates in Excel. Operating years and opening delay are fixed to this timeline and locked by validation; change code to extend it.')}
for r,(a,bv) in texts.items():put(s,'A'+str(r),a);put(s,'B'+str(r),bv)
# sensitivity as transparent formula adjustment to base cash flow / terminal
ss='Sensitivity';w[ss].append(['Case','All-in market price','Occupancy','Paid months','Capex multiplier','Terminal sale','NPV EUR','Interpretation'])
cases=[('Base',600,.95,12,1,1),('Shared benchmark price',525,.95,12,1,1),('Lower collection',600,.85,12,1,1),('10 paid months',600,.95,10,1,1),('20% development overrun',600,.95,12,1.2,1),('No sale proceeds',600,.95,12,1,0)]
# full scenario cash flow calculations linked to inputs, rather than scenario assumptions disguised as data
for rr,case in enumerate(cases,2):
 for cc,val in enumerate(case,1):w[ss].cell(rr,cc,val)
 p,oc,months,mul,sale=case[1:];blend=.8*p+.2*min(aff,p);noi1=n*months*oc*(blend-75)-166000
 fs=[-capex*mul,0]+[noi1*1.02**(y-1)-(300000*1.02**(y-1) if y==15 else 0)+(noi1*1.02**30/.06*.98*sale if y==30 else 0) for y in range(1,31)]
 val=sum(x/1.07**t for t,x in enumerate(fs))
 detail='Scenario cash flows'
 if detail not in w.sheetnames:
  w.create_sheet(detail);w[detail].append(['Year']+[x[0] for x in cases])
 col=rr;letter=w[detail].cell(1,col).column_letter
 # Scenario delta accounts for both receipts and variable utility costs.
 delta=f'({ref("Rooms")}*Sensitivity!D{rr}*Sensitivity!C{rr}*((1-{ref("Affordable room share")})*Sensitivity!B{rr}+{ref("Affordable room share")}*MIN(\'Executive report\'!B5,Sensitivity!B{rr}))-\'Cash flows\'!E{{row}})'
 for t,scf in enumerate(fs):
  row=t+2;put(detail,'A'+str(row),t)
  if t==0:f=f'=-\'Development budget\'!B9*Sensitivity!E{rr}'
  elif t==1:f='=0'
  else:
   d=f'({ref("Rooms")}*Sensitivity!D{rr}*Sensitivity!C{rr}*((1-{ref("Affordable room share")})*Sensitivity!B{rr}+{ref("Affordable room share")}*MIN(\'Executive report\'!B5,Sensitivity!B{rr}))*(1+{ref("Nominal annual rent growth")})^(\'Cash flows\'!B{row}-1)-\'Cash flows\'!E{row}-( {ref("Rooms")}*Sensitivity!D{rr}*Sensitivity!C{rr}*{ref("Monthly utility cost")}*(1+{ref("Nominal annual cost growth")})^(\'Cash flows\'!B{row}-1)-\'Cash flows\'!F{row}))'
   f=f'=\'Cash flows\'!I{row}-\'Cash flows\'!J{row}+{d}'
   if t==31:
    nextnoi=f'({ref("Rooms")}*Sensitivity!D{rr}*Sensitivity!C{rr}*((1-{ref("Affordable room share")})*Sensitivity!B{rr}+{ref("Affordable room share")}*MIN(\'Executive report\'!B5,Sensitivity!B{rr}))*(1+{ref("Nominal annual rent growth")})^30-({ref("Rooms")}*Sensitivity!D{rr}*Sensitivity!C{rr}*{ref("Monthly utility cost")}+\'Development budget\'!B3*{ref("Annual maintenance / hard cost")}+{ref("Annual other operating expenses")})*(1+{ref("Nominal annual cost growth")})^30)'
    f+=f'+{nextnoi}/{ref("Terminal capitalization rate")}*(1-{ref("Terminal selling cost")})*Sensitivity!F{rr}'
  put(detail,letter+str(row),f,scf)
 put(ss,'G'+str(rr),f'=NPV({ref("Nominal required return")},\'Scenario cash flows\'!{letter}3:{letter}33)+\'Scenario cash flows\'!{letter}2',val)
 w[ss].cell(rr,8,'Benchmark / explicit stress; no forecast probability assigned')
reg='Regulations';w[reg].append(['Topic','Official rule / project action','Source'])
for row in [('Student lease','Maximum 12 months; tenant may terminate with two months notice without penalty. Lease length is not average residence length.','S3'),('Deposit','Maximum two months rent for leases concluded/renewed from 1 Nov 2024. No average deposit identified. Held for tenant; not investor revenue.','S2'),('Quality label','Optional quality label lasts five years; application includes compliance declaration, student lease and valid PEB certificate.','S5'),('Site approvals','Before purchase verify zoning, building permit, fire safety, habitability, accessibility and energy performance with competent authorities. No site approved by this workbook.','S4 p7; site-specific review needed'),('Utilities','Show base rent and charges separately; actual billing and annual reconciliations require a compliant lease. Model assumes full cost recovery when separately billed.','Model assumption; obtain lease and meter plan')]:w[reg].append(row)
src='Sources';w[src].append(['ID','Official publisher / title','URL','Reference period / page','Access date','Use / limitation'])
for row in [('S1','Perspective — Panorama de la vie étudiante',E['survey_url'],'Survey 2023–24; Feb 2025; pp55,56,110,8','Historical rent / student budget; no future forecast'),('S2','Brussels Region — garantie locative','https://be.brussels/fr/logement/location/bail-dhabitation/garantie-locative','Effective 2024-11-01','Deposit legal cap'),('S3','Brussels Region — contrats de bail','https://be.brussels/fr/logement/location/bail-dhabitation/contrat-de-bail-modeles-et-annexe','Current page','Student lease term and termination'),('S4','Perspective — Monitoring student housing no.1',E['monitor_url'],'2025 inventory; published 2026; pp26,28,7','Dedicated stock and estimated demand; no live vacancy or measured total shortage'),('S5','Brussels Region — label logement étudiant','https://be.brussels/fr/logement/location/conformite-logement/label-logement-etudiant','Current page','Quality label'),('S6','Statbel — residential property price index','https://statbel.fgov.be/fr/themes/construction-logement/indice-des-prix-des-logements','Q2 2026; published 25 Sep 2026','Transaction price index is not EUR/m2 construction cost; not used as a construction quote')]:w[src].append([*row[:4],'2026-10-14',row[4]])
for r in range(2,w[src].max_row+1):w[src].cell(r,3).hyperlink=w[src].cell(r,3).value
mt='Methods';w[mt].append(['Calculation','Original formula / convention'])
for row in [('NPV','NPV = CF0 + sum(CFt/(1+r)^t), t=1..31. Excel NPV excludes t=0; add it explicitly.'),('IRR','IRR solves 0 = sum(CFt/(1+IRR)^t). Excel IRR on all cash flows including t=0; equally spaced annual periods.'),('PBP','First crossing k with cumulative CFk >= 0: (k-1) + abs(cumulative CF(k-1))/CFk. Includes terminal sale; separately show operating-only recovery.'),('Discounted PBP','Same crossing formula using discounted cash flows. If no crossing within 31 years, display Not recovered.'),('Revenue','Rooms × paid months × occupancy × blended monthly price; affordable units priced at min(policy budget threshold, market price).'),('Utility billing','Included rent = all-in price. Separate mode rent = all-in price minus utility allowance; utility recovery equals allowance. Model preserves total receipts.'),('NOI','Total receipts - utility expense - maintenance - other operating costs. Cash flow = NOI - replacement capex + terminal proceeds.'),('Terminal value','Next year NOI / terminal cap rate × (1-selling cost), received at the end of final operating year; optional zero-sale stress.'),('Threshold budget','Affordable threshold = official historical student-housing spending budget × policy share. This is not an official income-based definition.'),('Data gaps','No official project-specific land, build, maintenance, utility, occupancy, discount rate, exit yield, average stay/deposit or live vacancy data verified. Those cells remain explicit assumptions or Unknown.'),('Reproduction','Run python3 build_model.py; requires openpyxl. Atomic save and cached formula values are generated. Cached values reflect the default inputs; Excel recalculates when inputs change.')]:w[mt].append(row)
put(s,'B2','=IF(AND(B13>=0,B14>=Inputs!B22),"PASS numerical screen; verify all project inputs before commitment","DEFER: NPV / IRR fail the assumed investment hurdle")','DEFER: NPV / IRR fail the assumed investment hurdle')
# visual / usability
for sh in w:
 sh.freeze_panes='B2';sh.auto_filter.ref=sh.dimensions
 for ce in sh[1]:ce.fill=PatternFill('solid',fgColor='16324F');ce.font=Font(bold=True,color='FFFFFF')
 for row in sh.iter_rows(min_row=2):
  for ce in row:
   ce.alignment=Alignment(vertical='top',wrap_text=True)
   if ce.data_type=='f':ce.font=Font(color='172B4D');ce.number_format='#,##0.00;[Red](#,##0.00)'
 sh.column_dimensions['A'].width=39
 for col in range(2,sh.max_column+1):sh.column_dimensions[sh.cell(1,col).column_letter].width=23
 sh.sheet_properties.pageSetUpPr.fitToPage=True;sh.page_setup.orientation='landscape';sh.page_setup.paperSize=sh.PAPERSIZE_A4;sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0
 sh.print_title_rows='1:1'
for sh in [s,mt]:w[sh].column_dimensions['B'].width=115
w[s].row_dimensions[1].height=30
for r in list(range(21,30))+[2,3]:w[s].row_dimensions[r].height=62
for col in ['D','E']:w[S].column_dimensions[col].width=46
w[m].column_dimensions['C'].width=90;w[reg].column_dimensions['B'].width=110;w[src].column_dimensions['C'].width=90;w[src].column_dimensions['F'].width=65;w[ss].column_dimensions['H'].width=60
for r,item in enumerate(data,2):
 ce=w[S].cell(r,2);ce.font=Font(color='00804A' if item[3]=='OFFICIAL BENCHMARK' else '1565C0');ce.fill=PatternFill('solid',fgColor='E9F6EE' if item[3]=='OFFICIAL BENCHMARK' else 'EDF4FF')
 if item[2]=='fraction':ce.number_format='0.0%'
 dv=DataValidation(type='decimal',operator='between',formula1=0,formula2=1e9);dv.errorTitle='Invalid input';dv.error='Use a nonnegative value';dv.showErrorMessage=True
 if item[2]=='fraction':dv.formula2=1
 if item[2]=='binary':dv.type='whole';dv.formula2=1
 if item[0]=='Operating years':dv.type='whole';dv.formula1=dv.formula2=30
 if item[0]=='Years to opening':dv.type='whole';dv.formula1=dv.formula2=1
 if item[0]=='Paid months per year':dv.formula2=12
 if item[0]=='Deposit months requested':dv.formula2=2
 if item[0] in ['Nominal required return','Terminal capitalization rate']:dv.formula1=.0001
 w[S].add_data_validation(dv);dv.add(ce)
w[s]['B14'].number_format='0.00%';w[m]['B7'].number_format='0.0%'
chart=LineChart();chart.title='Cumulative project cash flows';chart.y_axis.title='EUR';chart.x_axis.title='Year';chart.add_data(Reference(w[c],min_col=15,max_col=16,min_row=1,max_row=33),titles_from_data=True);chart.set_categories(Reference(w[c],min_col=1,min_row=2,max_row=33));chart.width=25;chart.height=12;w[s].add_chart(chart,'A32')
w.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
tmp=R/'model.tmp.xlsx';w.save(tmp)
# Embed cached results so preview/readers show outputs without Excel calculation.
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};ET.register_namespace('',ns['m'])
with zipfile.ZipFile(tmp) as z:files={name:z.read(name) for name in z.namelist()}
for i,sh in enumerate(w.sheetnames,1):
 name=f'xl/worksheets/sheet{i}.xml';root=ET.fromstring(files[name])
 for ce in root.findall('.//m:c',ns):
  key=(sh,ce.attrib.get('r'))
  if key not in cache:continue
  val=cache[key];v=ce.find('m:v',ns)
  if v is None:v=ET.SubElement(ce,'{'+ns['m']+'}v')
  if isinstance(val,str):ce.set('t','str');v.text=val
  else:v.text=str(val)
 files[name]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as z:
 for name,content in files.items():z.writestr(name,content)
load_workbook(tmp);os.replace(tmp,R/'model.xlsx')
metrics={'initial_cost':capex,'all_in_blended_rent':blended,'affordable_threshold':aff,'NPV':npv,'IRR':ir,'payback':pbs[0] if pbs else None,'discounted_payback':dpbs[0] if dpbs else None,'operating_only_recovered':cumop>=0}
(R/'validation_metrics.json').write_text(json.dumps(metrics,indent=2))
report=f'''# Executive report — Brussels student accommodation\n\nRecommendation: defer capital commitment. The illustrative 100-room scheme, including 20 affordable rooms, has NPV EUR{npv:,.0f}, IRR {ir:.2%}, and simple payback {pbs[0]:.2f} years including sale. Discounted payback is {str(round(dpbs[0],2)) if dpbs else 'not achieved within 31 years'}. These outcomes depend on project assumptions, not official construction quotations.\n\nThe assumed initial spend is EUR{capex:,.0f}; required return 7%; opening delay one year; 30 operating years. The default terminal sale uses a 6% capitalization rate and 2% selling cost. Income tax and financing are excluded. The workbook shows operating-only recovery separately and tests zero sale proceeds.\n\nOfficial survey benchmarks are EUR600 monthly rent including charges and EUR833.50 monthly spending budget for students in student housing ([Perspective survey, pp55/110]({E['survey_url']})). The chosen 30% housing budget share gives an illustrative affordable ceiling of EUR{aff:.2f}; it is not an official affordability definition. Actual ability to pay requires student income/support data.\n\nAdvantages include adding dedicated accommodation, reducing pressure on ordinary housing, providing energy-efficient facilities and offering predictable charges. Drawbacks include uncertain acquisition and construction cost, approvals and delays, turnover and summer vacancies, energy and repair exposure, lower receipts from affordable units and reliance on terminal value.\n\nShortlist sites near target campuses and public transport. Compare zoning, acquisition price, walking time, safety and competing pipeline. Obtain a site-specific budget, current room prices and vacancies, operator cost plan, tax review and permit/fire/energy confirmations before proceeding. Approve only if updated NPV is nonnegative, IRR meets the hurdle and the affordable commitment is feasible.\n\nOfficial stock and demand figures do not measure current vacant rooms or a verified shortage; ordinary rental housing also serves students. No verified official project-specific land/build/maintenance prices or investment hurdle were available. The workbook identifies those assumptions explicitly rather than presenting invented statistics. All URLs, periods, formula definitions, sensitivities and cash flows are included in model.xlsx.\n'''
(R/'Executive_report.md').write_text(report)
print(json.dumps(metrics));print('Saved atomically. SAVE_NOW:',(R/'SAVE_NOW.txt').exists())
