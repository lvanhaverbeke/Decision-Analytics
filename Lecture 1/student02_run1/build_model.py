import os,json,math,zipfile,tempfile,xml.etree.ElementTree as ET
from pathlib import Path
from openpyxl import Workbook,load_workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.chart import LineChart,Reference
from openpyxl.workbook.properties import CalcProperties
ROOT=Path(__file__).resolve().parent
UPDATE=ROOT/'research.json'
data=json.loads(UPDATE.read_text()) if UPDATE.exists() else {}
wb=Workbook(); wb.remove(wb.active)
for name in ['Overview','Inputs','Cash flow','Rent sensitivity','Variables','Regulations','Sources','Student budget','Risk sensitivity']:wb.create_sheet(name)
cache={}
def put(s,cell,value,cached=None):
 wb[s][cell]=value
 if cached is not None:cache[(s,cell)]=cached
inputs=[
('Rooms',100,'rooms','Decision; discrete; ratio','Illustrative scale; no plot secured'),
('Gross floor area per room',28,'m²/room','Assumption; continuous; ratio','Includes share of circulation and common areas'),
('Construction cost',2600,'EUR/m²','Assumption; continuous; ratio','Concept estimate, gross of nonrecoverable VAT; obtain quotations'),
('Land / acquisition',1800000,'EUR','Assumption; continuous; ratio','No parcel or offer secured'),
('Professional / permit fees',.15,'% of construction','Assumption; continuous; ratio','Architect, studies, permit; gross cost assumption'),
('Contingency',.10,'% of construction + fees','Assumption; continuous; ratio','No double counting with operating reserve'),
('Monthly room rent',650,'EUR/room/month','Decision; continuous; ratio','Excludes utilities; market assumption'),
('Annual rent growth',.02,'%/year','Assumption; continuous; ratio','Subject to lawful indexation and lease provisions'),
('Stabilised occupancy',.95,'% of rooms','Assumption; continuous; ratio','Includes vacancy and replacement after departures'),
('Operating expense share',.25,'% of collected rent','Assumption; continuous; ratio','Management, repairs, insurance, property tax, owner utilities; excludes debt and income tax'),
('Annual capital reserve',150,'EUR/room/year','Assumption; continuous; ratio','Inflates at annual rent growth'),
('Discount rate',.07,'%/year','Assumption; continuous; ratio','Nominal unlevered required return'),
('Exit capitalisation rate',.055,'%','Assumption; continuous; ratio','Applied to next-year NOI before reserve'),
('Exit selling costs',.02,'% of terminal value','Assumption; continuous; ratio','Illustrative transaction costs'),
('Construction / approvals duration',2,'years','Assumption; discrete; ratio','Fixed schedule: first operating CF at end of year 2; rebuild code to change timing'),
('Operating horizon',30,'years','Assumption; discrete; ratio','Fixed 30-row operating model; changing this requires rebuilding schedule'),
('International student share',.23,'% of VUB students','Context; continuous; ratio','PROVISIONAL; verify definition, year and population'),
('Annual student resources',14000,'EUR/student/year','Scenario; continuous; ratio','Illustrative total resources incl. family support; not an observed VUB average or legal minimum'),
('Academic dropout proxy',.20,'% of entrant cohort','Context assumption; continuous; ratio','UNVERIFIED scenario; academic dropout is different from housing departure'),
('Baseline annual housing departures',.08,'% of residents/year','Assumption; continuous; ratio','Illustrative; no causal estimate from academic dropout'),
('Departure response to rent rise',.15,'fraction/fraction','Assumption; continuous; ratio','A 10% rent increase adds 1.5 percentage points of departures; unvalidated'),
]
for key,val in data.get('inputs',{}).items():
 i=int(key)-4; old=inputs[i]; inputs[i]=(old[0],val.get('value',old[1]),old[2],old[3],val.get('note',old[4]))
ws=wb['Inputs'];ws.append(['PROJECT INPUTS — blue cells editable']);ws.append(['Nominal EUR; unlevered, before income tax. Costs assumed inclusive of nonrecoverable taxes.']);ws.append(['Variable','Value','Unit','Variable classification','Definition / evidence'])
for row in inputs:ws.append(row)
v={i+4:r[1] for i,r in enumerate(inputs)}
area=v[4]*v[5];build=area*v[6];fee=build*v[8];cont=(build+fee)*v[9];total=v[7]+build+fee+cont
for cell,label,formula,value in [('A27','Gross area','=B4*B5',area),('A28','Construction subtotal','=B27*B6',build),('A29','Fees','=B28*B8',fee),('A30','Contingency','=(B28+B29)*B9',cont),('A31','Total investment','=B7+SUM(B28:B30)',total)]:
 ws[cell]=label;put('Inputs','B'+cell[1:],formula,value)
ws['A33']='Location';ws['B33']='Etterbeek / Ixelles, near VUB Main Campus, Pleinlaan / Boulevard de la Plaine';ws['A34']='Site status';ws['B34']='Search area only: no identified parcel, zoning clearance or buildability confirmation.'
cf=wb['Cash flow'];cf.append(['UNLEVERED ANNUAL CASH FLOW — end-of-year convention']);cf.append(['Year','Stage','Capex','Gross rent','Occupancy','Collected rent','Operating costs','NOI','Capital reserve','Net operating CF','Net sale proceeds','Total CF','Discount factor','PV','Cumulative CF'])
flows=[];pvs=[];cumulative=0;prev=0;payback=None
for year in range(32):
 r=year+3; op=year>=2; gross=v[4]*v[10]*12*(1+v[11])**(year-2) if op else 0; collected=gross*v[12]; costs=collected*v[13];noi=collected-costs;reserve=v[4]*v[14]*(1+v[11])**(year-2) if op else 0
 capex=-(v[7]+(total-v[7])/2) if year==0 else -(total-v[7])/2 if year==1 else 0
 terminal=noi*(1+v[11])/v[16]*(1-v[17]) if year==31 else 0
 net=noi-reserve;flow=capex+net+terminal;prev=cumulative;cumulative+=flow;pv=flow/(1+v[15])**year
 if payback is None and year>0 and cumulative>=0 and prev<0:payback=year-1+(-prev)/flow
 flows.append(flow);pvs.append(pv)
 vals=[year,'Construction' if not op else 'Operating',capex,gross,v[12] if op else 0,collected,costs,noi,reserve,net,terminal,flow,1/(1+v[15])**year,pv,cumulative]
 forms={2:f'=-Inputs!B7-(Inputs!B31-Inputs!B7)/2' if year==0 else '=-(Inputs!B31-Inputs!B7)/2' if year==1 else '=0',3:f'=Inputs!B4*Inputs!B10*12*(1+Inputs!B11)^(A{r}-Inputs!B18)' if op else '=0',4:'=Inputs!B12' if op else '=0',5:f'=D{r}*E{r}',6:f'=F{r}*Inputs!B13',7:f'=F{r}-G{r}',8:f'=Inputs!B4*Inputs!B14*(1+Inputs!B11)^(A{r}-Inputs!B18)' if op else '=0',9:f'=H{r}-I{r}',10:f'=H{r}*(1+Inputs!B11)/Inputs!B16*(1-Inputs!B17)' if year==31 else '=0',11:f'=C{r}+J{r}+K{r}',12:f'=1/(1+Inputs!B15)^A{r}',13:f'=L{r}*M{r}',14:f'=SUM(L$3:L{r})'}
 for c,x in enumerate(vals):put('Cash flow',f'{chr(65+c)}{r}',forms.get(c,x),x if c in forms else None)
def npv(rate):return sum(f/(1+rate)**t for t,f in enumerate(flows))
lo,hi=-.5,1
for _ in range(120):
 mid=(lo+hi)/2
 if npv(mid)>0:lo=mid
 else:hi=mid
irr=(lo+hi)/2
ov=wb['Overview'];ov.append(['VUB STUDENT HOUSING — investment screening']);ov.append([data.get('status','BASIC VERSION — provisional assumptions; research pending')]);ov.append(['Metric','Result','Classification / interpretation'])
metrics=[('Total investment','=Inputs!B31',total,'Derived; continuous; ratio; nominal EUR'),('NPV at required return','=NPV(Inputs!B15,\'Cash flow\'!L4:L34)+\'Cash flow\'!L3',sum(pvs),'Output; continuous; interval; EUR; t=0'),('Project IRR',"=IRR('Cash flow'!L3:L34)",irr,'Output; continuous; rate; unlevered'),('Simple payback incl. exit',"=IF(COUNTIF('Cash flow'!O3:O34,\">=0\")=0,\"Not reached\",MATCH(TRUE,INDEX('Cash flow'!O3:O34>=0,0),0)-2-INDEX('Cash flow'!O3:O34,MATCH(TRUE,INDEX('Cash flow'!O3:O34>=0,0),0)-1)/INDEX('Cash flow'!L3:L34,MATCH(TRUE,INDEX('Cash flow'!O3:O34>=0,0),0)))",payback if payback else 'Not reached','Output; continuous; ratio; years from t=0; interpolation'),('Time until operations','=Inputs!B18',v[18],'Assumption; discrete; ratio; years incl. approvals'),('Operating horizon','=Inputs!B19',v[19],'Assumption; discrete; ratio; years'),('International share','=Inputs!B20',v[20],'Context; continuous; ratio; see Inputs evidence'),('Annual student resources','=Inputs!B21',v[21],'Scenario; continuous; ratio; EUR/year; no VUB average established'),('Rent / month','=Inputs!B10',v[10],'Decision; continuous; ratio; EUR excl. utilities'),('Academic dropout proxy','=Inputs!B22',v[22],'Context; continuous; ratio; Flemish proxy, not VUB rate; S9'),('Annual rent / resources','=Inputs!B10*12/Inputs!B21',v[10]*12/v[21],'Derived; continuous; ratio; excludes utilities'),('First operating year NOI',"='Cash flow'!H5",(v[4]*v[10]*12*v[12])*(1-v[13]),'Output; continuous; ratio; EUR/year')]
for r,(lab,formula,value,desc) in enumerate(metrics,4):ov.cell(r,1,lab);put('Overview',f'B{r}',formula,value);ov.cell(r,3,desc)
ov['A18']='Investment interpretation';ov['B18']='Negative NPV at 7% implies the assumed rent, costs and exit value do not meet the required return.' if sum(pvs)<0 else 'Positive NPV under these assumptions; sensitivity and site feasibility remain essential.'
ov['A20']='Scope';ov['B20']='100-room illustrative development. No secured land or construction quote. 30 operating years; terminal sale in year 31.'
ov['A21']='Student income';ov['B21']='There is no universal minimum annual income for a student. Wages apply only to hours worked; family support, grants and earnings vary.'
ov['A22']='Dropout versus departures';ov['B22']='16% dropout is a Flemish first-year cohort proxy, not VUB-specific. Price-driven departures below are hypothetical, not measured VUB behaviour; baseline vacancy already includes turnover.'
ov['A23']='Payback convention';ov['B23']='Simple PBP includes sale if needed; the base case pays back before sale. Fractional years assume uniform within-year operating cash flow. Discounted payback uses the 7% required return.'
opcum=-total;oppb=None
for year in range(2,32):
 net=(v[4]*v[10]*12*v[12]*(1-v[13])-v[4]*v[14])*(1+v[11])**(year-2);prev=opcum;opcum+=net
 if oppb is None and opcum>=0:oppb=year-1-prev/net
ov['A25']='Operating-only payback (years)';ov['B25']=oppb if oppb else 'Not reached within 31 years';ov['C25']='Derived; continuous; ratio; excludes terminal sale'
ov['A26']='Discounted payback (years)';ov['B26']='Not reached within 31 years' if sum(pvs)<0 else 'See cumulative discounted flows';ov['C26']='Derived; continuous; ratio; includes sale'
rs=wb['Rent sensitivity'];rs.append(['ILLUSTRATIVE PRICE RESPONSE — NOT EMPIRICAL VUB DATA']);rs.append(['Rent increase','Monthly rent EUR','Annual departures','Incremental departures','Occupancy proxy','NPV EUR'])
for r,inc in enumerate([0,.025,.05,.075,.1,.15,.2,.25,.3],3):
 rent=v[10]*(1+inc);dep=min(1,max(0,v[23]+inc*v[24]));occ=max(0,v[12]-(dep-v[23]));snpv=0
 for t,base in enumerate(flows):
  if t<2:sf=base
  else:
   noi=v[4]*rent*12*(1+v[11])**(t-2)*occ*(1-v[13]);sf=noi-v[4]*v[14]*(1+v[11])**(t-2)
   if t==31:sf+=noi*(1+v[11])/v[16]*(1-v[17])
  snpv+=sf/(1+v[15])**t
 for c,f,val in [('A',inc,None),('B',f'=Inputs!B10*(1+A{r})',rent),('C',f'=MIN(1,MAX(0,Inputs!B23+A{r}*Inputs!B24))',dep),('D',f'=C{r}-Inputs!B23',dep-v[23]),('E',f'=MAX(0,Inputs!B12-D{r})',occ),('F',f"=Overview!B5+(B{r}*E{r}/(Inputs!B10*Inputs!B12)-1)*(SUM('Cash flow'!N5:N34)+SUMPRODUCT('Cash flow'!I5:I34,'Cash flow'!M5:M34))",snpv)]:put('Rent sensitivity',f'{c}{r}',f,val)
rs['A14']='Interpretation';rs['B14']='One-time rent uplift at opening; normal annual growth thereafter. Every incremental departure assumed to cause one full year of vacancy, without replacement: conservative scenario only.'
rs['A15']='Academic dropout';rs['B15']='Not used as a housing-exit probability; cohorts and denominators differ.'
chart=LineChart();chart.title='Rent increase and students leaving (scenario)';chart.y_axis.title='Monthly rent (EUR)';chart.x_axis.title='Rent increase';chart.add_data(Reference(rs,min_col=2,min_row=2,max_row=11),titles_from_data=True);chart.set_categories(Reference(rs,min_col=1,min_row=3,max_row=11));chart.height=10;chart.width=22
c2=LineChart();c2.add_data(Reference(rs,min_col=3,min_row=2,max_row=11),titles_from_data=True);c2.set_categories(Reference(rs,min_col=1,min_row=3,max_row=11));c2.y_axis.axId=200;c2.y_axis.title='Annual residents leaving (%)';c2.y_axis.numFmt='0%';c2.y_axis.crosses='max';chart+=c2;rs.add_chart(chart,'H2')
var=wb['Variables'];var.append(['VARIABLE DICTIONARY']);var.append(['Variable','Role / scale','Meaning'])
for row in inputs:var.append([row[0],row[3],row[4]])
for m in metrics:var.append([m[0],m[3],'See Overview and Cash flow'])
var.append(['Construction location','Decision; categorical; nominal','Etterbeek / Ixelles search area, parcel not selected']);var.append(['Project stage','Categorical; ordinal','Planning → construction → operations']);var.append(['Year','Index; discrete; interval','Relative annual dates from investment start'])
reg=wb['Regulations'];reg.append(['BRUSSELS REGULATORY DUE DILIGENCE']);reg.append(['Topic','Requirement / action','Model implication','Source'])
default_regs=[['Planning / site','Confirm zoning, lawful use, planning permit and municipal conditions before land commitment.','Two-year total development duration is an assumption, not an approval estimate.','Research pending'],['Housing quality','Verify applicable Brussels minimum safety, sanitation, equipment and room area rules.','28 m² gross per room is a budget allocation, not proof of compliant private living space.','Research pending'],['Lease / rent','Use compliant student tenancy documentation and verify indexation, deposit and rent rules.','2% rent growth is a scenario, not a legal entitlement.','Research pending'],['Fire / energy / accessibility','Obtain fire-service advice and validate EPB, accessibility and building requirements with architect.','Costs and timetable may rise after design review.','Research pending']]
for row in data.get('regulations',default_regs):reg.append(row)
src=wb['Sources'];src.append(['SOURCES AND RESEARCH LOG']);src.append(['ID','Organisation / title','URL','Period / access','Evidence and limits'])
for row in data.get('sources',[]):src.append(row)
if not data.get('sources'):src.append(['A0','Author assumptions','','2026-10-14','All inputs provisional pending research.'])
# Supplementary cash-flow controls and formula-based payback calculations.
cf.cell(2,16,'Cumulative CF excl. exit');cf.cell(2,17,'Cumulative discounted CF');cf.cell(2,18,'PBP crossing');cf.cell(2,19,'Operating PBP crossing');cf.cell(2,20,'Discounted PBP crossing')
noexit=0;dpv=0;discount_pb=None
for r in range(3,35):
 t=r-3;previous_noexit=noexit;previous_dpv=dpv
 terminal_value=cache[('Cash flow',f'K{r}')];noexit+=flows[t]-terminal_value;dpv+=pvs[t]
 put('Cash flow',f'P{r}',f'=SUM(L$3:L{r})-SUM(K$3:K{r})',noexit)
 put('Cash flow',f'Q{r}',f'=SUM(N$3:N{r})',dpv)
 for col,prevcol,amountcol,cur,prev in [('R','O','L',sum(flows[:t+1]),sum(flows[:t])),('S','P','J',noexit,previous_noexit),('T','Q','N',dpv,previous_dpv)]:
  crossing=t-1-prev/(flows[t] if col=='R' else flows[t]-terminal_value if col=='S' else pvs[t]) if t>0 and cur>=0 and prev<0 else 0
  formula=f'=IF(AND({prevcol}{r}>=0,{prevcol}{r-1}<0),A{r}-1-{prevcol}{r-1}/{amountcol}{r},0)' if t>0 else '=0'
  put('Cash flow',f'{col}{r}',formula,crossing)
  if col=='T' and crossing:discount_pb=crossing
put('Overview','B7', '=IF(COUNTIF(\'Cash flow\'!R3:R34,">0")=0,"Not reached",SUM(\'Cash flow\'!R3:R34))',payback or 'Not reached')
put('Overview','B25','=IF(COUNTIF(\'Cash flow\'!S3:S34,">0")=0,"Not reached",SUM(\'Cash flow\'!S3:S34))',oppb or 'Not reached')
put('Overview','B26','=IF(COUNTIF(\'Cash flow\'!T3:T34,">0")=0,"Not reached",SUM(\'Cash flow\'!T3:T34))',discount_pb or 'Not reached')
ov['A28']='Minimum proof of means / year';put('Overview','B28',"='Student budget'!B5*12",12744);ov['C28']='Regulatory benchmark; continuous; ratio; EUR; certain non-EU students (S3)'
ov['A29']='Minimum student wage, age 21';put('Overview','B29',"='Student budget'!B12",13.5644);ov['C29']='Regulatory benchmark; continuous; ratio; EUR/hour; fallback absent sector CBA (S11)'
ov['A30']='Illustrative 650h gross pay, age 21';put('Overview','B30',"='Student budget'!C12",13.5644*650);ov['C30']='Derived scenario; continuous; ratio; EUR/year; not VUB average'
ov['A32']='Model controls';ov['B32']='Blue cells are assumptions. Development duration and horizon are fixed to 2 and 30 years; rebuild the saved code for another schedule. No debt, grants, income tax or financing fees; costs gross of nonrecoverable tax.'
budget=wb['Student budget'];budget.append(['STUDENT RESOURCES AND AFFORDABILITY']);budget.append(['Item / age','Value / hourly pay EUR','Annual gross EUR','After solidarity contribution EUR','Classification / evidence'])
budget['A4']='Illustrative annual total resources';put('Student budget','B4','=Inputs!B21',v[21]);budget['E4']='Scenario; continuous; ratio. Family support + grants + work; no observed VUB average.'
budget['A5']='Monthly proof of means';budget['B5']=1062;budget['E5']='S3: 2026–2027 benchmark for relevant non-EU student residence applicants; not wages.'
budget['A6']='Reduced-contribution annual hours';budget['B6']=650;budget['E6']='S4: count, discrete, ratio; scenario assumes entire quota worked, not guaranteed employment.'
budget['A7']='Student solidarity contribution';budget['B7']=.0271;budget['E7']='S4: continuous ratio; 2.71%; tax treatment depends on personal circumstances.'
for rr,age,rate in [(9,18,10.7158),(10,19,11.5297),(11,20,12.2079),(12,21,13.5644)]:
 budget.cell(rr,1,age);budget.cell(rr,2,rate);put('Student budget',f'C{rr}',f'=B{rr}*$B$6',rate*650);put('Student budget',f'D{rr}',f'=C{rr}*(1-$B$7)',rate*650*(1-.0271));budget.cell(rr,5,'S11: fallback hourly rate effective July 2026, indicative 38h-week basis; sector CBA may apply. Gross annual pay assumes rate unchanged for all 650h; D is before any income tax.')
budget['A15']='VUB standard rent incl. listed costs';budget['B15']=449.30;budget['E15']='S10: EUR/month; continuous ratio; 2026–2027; subsidised university offering with eligibility conditions.'
budget['A16']='Project rent excl. utilities';put('Student budget','B16','=Inputs!B10',v[10]);budget['E16']='Decision; continuous ratio; EUR/month.'
budget['A17']='Rent share of proof-of-means amount';put('Student budget','B17','=B16/B5',v[10]/1062);budget['E17']='Derived; continuous ratio; excludes utilities, tuition and other living expenses.'
budget['A19']='Minimum annual earnings';budget['B19']='No guaranteed amount: a student with no paid work may earn EUR0. Actual average VUB earnings were not established.'
risk=wb['Risk sensitivity'];risk.append(['NPV SENSITIVITY — required return versus exit capitalisation']);risk.append(['Discount rate / exit cap',.045,.055,.065])
for rr,discount in enumerate([.05,.06,.07,.08,.09],3):
 risk.cell(rr,1,discount)
 for cc,cap in enumerate([.045,.055,.065],2):
  terminal=cache[('Cash flow','H34')]*(1+v[11])/cap*(1-v[17]);value=sum(f/(1+discount)**t for t,f in enumerate(flows))+(terminal-cache[('Cash flow','K34')])/(1+discount)**31
  letter=chr(64+cc);formula=f"=NPV($A{rr},'Cash flow'!L4:L34)+'Cash flow'!L3+('Cash flow'!H34*(1+Inputs!B11)/{letter}$2*(1-Inputs!B17)-'Cash flow'!K34)/(1+$A{rr})^31"
  put('Risk sensitivity',f'{letter}{rr}',formula,value)
risk['A10']='Break-even opening rent, no extra departures';rent_pv=sum(pvs[2:])+sum(cache[('Cash flow',f'I{r}')]*cache[('Cash flow',f'M{r}')] for r in range(5,35));breakeven=v[10]*(1-sum(pvs)/rent_pv)
put('Risk sensitivity','B10',"=Inputs!B10*(1-Overview!B5/(SUM('Cash flow'!N5:N34)+SUMPRODUCT('Cash flow'!I5:I34,'Cash flow'!M5:M34)))",breakeven)
risk['A12']='Interpretation';risk['B12']='Discount and exit cap are continuous ratio assumptions; NPV is a derived continuous EUR output. Break-even rent holds occupancy fixed; verify affordability, lawfulness and departures separately.'
from openpyxl.worksheet.datavalidation import DataValidation
for rr,num in [(18,2),(19,30)]:
 fixed=DataValidation(type='whole',operator='equal',formula1=str(num));fixed.errorTitle='Fixed schedule';fixed.error='Rebuild build_model.py to change this schedule assumption.';fixed.showErrorMessage=True;fixed.errorStyle='stop';wb['Inputs'].add_data_validation(fixed);fixed.add(f'B{rr}')
for sh in wb:
 sh.freeze_panes='B4' if sh.title=='Inputs' else 'B3';sh.sheet_view.showGridLines=False
 sh.row_dimensions[1].height=28
 for cell in sh[1]:cell.fill=PatternFill('solid',fgColor='16324F');cell.font=Font(color='FFFFFF',bold=True,size=13)
 for row in sh.iter_rows(min_row=2):
  for cell in row:
   cell.alignment=Alignment(vertical='top',wrap_text=True)
   if isinstance(cell.value,(int,float)) or (isinstance(cell.value,str) and cell.value.startswith('=')):cell.number_format='#,##0.00;[Red](#,##0.00)'
 for col in sh.columns:
  letter=col[0].column_letter;sh.column_dimensions[letter].width=22 if sh.title=='Cash flow' else 29
 if sh.title not in ['Cash flow','Rent sensitivity']:
  sh.column_dimensions['A'].width=35;sh.column_dimensions['B'].width=65;sh.column_dimensions['C'].width=48;sh.column_dimensions['D'].width=40;sh.column_dimensions['E'].width=95
 sh.auto_filter.ref=sh.dimensions
 for r in range(2,sh.max_row+1):sh.row_dimensions[r].height=42 if sh.title not in ['Cash flow','Rent sensitivity'] else 23
for r in range(4,25):
 wb['Inputs'].cell(r,2).font=Font(color='0070C0');wb['Inputs'].cell(r,2).fill=PatternFill('solid',fgColor='E8F2FB')
for r in [8,9,11,12,13,15,16,17,20,22,23,24]:wb['Inputs'].cell(r,2).number_format='0.0%'
for r in [6,10,13,14]:ov.cell(r,2).number_format='0.0%'
for row in rs.iter_rows(min_row=3,max_row=11):
 for c in [0,2,3,4]:row[c].number_format='0.0%'
for r in range(3,35):cf.cell(r,5).number_format='0.0%'
for row in src.iter_rows(min_row=3):
 if row[2].value:row[2].hyperlink=row[2].value;row[2].style='Hyperlink'
for rr in [18,19]:
 wb['Inputs'].cell(rr,2).font=Font(color='595959');wb['Inputs'].cell(rr,2).fill=PatternFill('solid',fgColor='EEEEEE')
wb['Inputs']['E18']='Fixed 2-year schedule incl. approvals; change code and rebuild for alternative timing.'
wb['Student budget']['B7'].number_format='0.00%';wb['Student budget']['B17'].number_format='0.0%'
for rr in range(3,8):wb['Risk sensitivity'].cell(rr,1).number_format='0.0%'
for cc in range(2,5):wb['Risk sensitivity'].cell(2,cc).number_format='0.0%'
# Final presentation and input controls.
from openpyxl.formatting.rule import ColorScaleRule
for sh in wb:
 last='T' if sh.title=='Cash flow' else 'F' if sh.title=='Rent sensitivity' else 'E'
 sh.merge_cells(f'A1:{last}1')
 sh.row_dimensions[1].height=31
 header=3 if sh.title in ['Overview','Inputs'] else 2
 for c in sh[header]:
  c.fill=PatternFill('solid',fgColor='DCE6F1');c.font=Font(bold=True,color='16324F')
 sh.sheet_properties.pageSetUpPr.fitToPage=True
 sh.page_setup.orientation='landscape';sh.page_setup.paperSize=sh.PAPERSIZE_A3
 sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0
 if sh.title in ['Sources','Regulations']:
  for rr in range(3,sh.max_row+1):sh.row_dimensions[rr].height=78
 if sh.title=='Overview':
  sh.merge_cells('A2:C2');sh.row_dimensions[2].height=40
  for rr in [18,20,21,22,23,32]:sh.row_dimensions[rr].height=65
wb['Student budget'].row_dimensions[19].height=70
for rr in range(9,13):wb['Student budget'].row_dimensions[rr].height=75
wb['Risk sensitivity'].conditional_formatting.add('B3:D7',ColorScaleRule(start_type='min',start_color='F4CCCC',mid_type='num',mid_value=0,mid_color='FFF2CC',end_type='max',end_color='D9EAD3'))
positive=DataValidation(type='decimal',operator='greaterThan',formula1='0');positive.showErrorMessage=True;positive.error='Enter a positive number.'
wb['Inputs'].add_data_validation(positive)
for rr in [4,5,6,7,10,14,15,16,21]:positive.add(f'B{rr}')
percent=DataValidation(type='decimal',operator='between',formula1='0',formula2='1');percent.showErrorMessage=True;percent.error='Enter a fraction from 0 to 1.'
wb['Inputs'].add_data_validation(percent)
for rr in [8,9,11,12,13,17,20,22,23,24]:percent.add(f'B{rr}')
wb['Inputs']['B31'].fill=PatternFill('solid',fgColor='D9EAD3');wb['Inputs']['B31'].font=Font(bold=True)
wb['Overview'].print_area='A1:C32';wb['Rent sensitivity'].print_area='A1:S22'
wb.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
tmp=ROOT/'model.tmp.xlsx';wb.save(tmp)
# Add calculated caches so the workbook is immediately readable without an Excel recalculation.
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
patched=ROOT/'model.cached.tmp.xlsx'
with zipfile.ZipFile(tmp) as zin,zipfile.ZipFile(patched,'w',zipfile.ZIP_DEFLATED) as zout:
 for item in zin.infolist():
  content=zin.read(item.filename)
  if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
   num=int(item.filename.split('sheet')[-1].split('.')[0]);name=wb.sheetnames[num-1];root=ET.fromstring(content)
   for c in root.findall('.//s:c',ns):
    key=(name,c.attrib.get('r'))
    if key in cache:
     value=cache[key];vv=c.find('s:v',ns)
     if vv is None:vv=ET.SubElement(c,'{'+ns['s']+'}v')
     if isinstance(value,str):c.set('t','str');vv.text=value
     else:vv.text=str(value)
   content=ET.tostring(root,encoding='utf-8',xml_declaration=True)
  zout.writestr(item,content)
# Validate zip and workbook before atomic replace.
with zipfile.ZipFile(patched) as z:assert z.testzip() is None
check=load_workbook(patched,data_only=True);assert check['Overview']['B5'].value is not None
os.replace(patched,ROOT/'model.xlsx');tmp.unlink(missing_ok=True)
print(json.dumps({'saved':'model.xlsx','investment':total,'NPV':sum(pvs),'IRR':irr,'payback':payback,'SAVE_NOW':(ROOT/'SAVE_NOW.txt').exists()}))
