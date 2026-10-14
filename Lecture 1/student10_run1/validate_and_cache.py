from pathlib import Path
import math, json, os, zipfile, xml.etree.ElementTree as ET
from openpyxl import load_workbook
P=Path(__file__).resolve().parent
if (P/'SAVE_NOW.txt').exists():print('SAVE_NOW: saved workbook retained');raise SystemExit
w=load_workbook(P/'model.xlsx'); cache={}; results={}
def put(sh,cell,val):cache.setdefault(sh,{})[cell]=val
def irr(cf):
 def npv(r):return sum(v/(1+r)**t for t,v in enumerate(cf))
 lo,hi=-.95,5
 if npv(lo)*npv(hi)>0:return 'No IRR'
 for _ in range(160):
  mid=(lo+hi)/2
  if npv(lo)*npv(mid)<=0:hi=mid
  else:lo=mid
 return (lo+hi)/2
for col,sc in enumerate(['Bad','Intermediate','Optimal'],2):
 v={r:w['Inputs'].cell(r,col).value for r in range(2,28)}
 rooms=v[2]; dev=v[8]; total=v[3]*(1+v[4])+v[5]*(1+v[6])*(1+v[7]);loan=total*v[22]
 payment=loan*v[23]/(1-(1+v[23])**(-v[24])) if v[23] else loan/v[24]
 flows=[];cum=0;dcum=0;balance=0;cum_no_sale=0
 for y in range(21):
  r=y+2; k=y-dev-1
  spend=v[3]*(1+v[4]) if y==0 else v[5]*(1+v[6])*(1+v[7])/dev if y<=dev else 0
  rent=rooms*12*min(v[9]*(1+v[11])**k,v[10]*(1+v[12])**k)*v[13]*(1-v[14]) if y>dev else 0
  cost=v[15]*(1+v[17])**k+rent*v[16] if y>dev else 0
  noi=rent-cost;reserve=rooms*v[18]*(1+v[17])**k if y>dev else 0
  rent21=rooms*12*min(v[9]*(1+v[11])**(20-dev),v[10]*(1+v[12])**(20-dev))*v[13]*(1-v[14])
  noi21=rent21*(1-v[16])-v[15]*(1+v[17])**(20-dev)
  sale=max(0,noi21)/v[19]*(1-v[20]) if y==20 else 0
  cf=noi-reserve-spend+sale;dcf=cf/(1+v[21])**y;cum+=cf;dcum+=dcf
  draw=loan if y==dev else 0;service=payment if dev<y<=dev+v[24] else 0
  balance=max(0,balance*(1+(v[23] if service else 0))+draw-service)
  equity=cf+draw-service-(balance if y==20 else 0);dscr=(noi-reserve)/service if service else 0
  cum_no_sale+=cf-sale
  vals=[y,spend,rent,cost,noi,reserve,sale,cf,dcf,cum,dcum,draw,service,balance,equity,dscr,cum_no_sale,0]
  flows.append(vals)
  for c,val in enumerate(vals,1):
   if c>1:put(sc,f'{chr(64+c)}{r}',val)
 assert abs(sum(x[1] for x in flows)-total)<1e-6
 assert all(x[2]==0 for x in flows[:dev+1])
 assert all(abs(x[7]-(x[4]-x[5]-x[1]+x[6]))<1e-7 for x in flows)
 assert abs(sum(x[11] for x in flows)-loan)<1e-7
 def pb(k):return next((x[0] for x in flows[1:] if x[k]>=0),'Beyond 20')
 initnoi=flows[dev+1][4]
 pv_in=sum((x[4]-x[5]+x[6])/(1+v[21])**x[0] for x in flows)
 share=(min(v[9],v[10])+v[26])/v[25]
 metrics=[dcum,irr([x[7] for x in flows]),pb(9),pb(10),total,total/rooms,initnoi,initnoi/total,min(x[15] for x in flows if x[12]>0),irr([x[14] for x in flows]),flows[-1][6],flows[-1][6]/(1+v[21])**20/pv_in,share,'Pass' if share<=v[27] else 'Fail',(v[15]+rooms*v[18])/(rooms*12*min(v[9],v[10])*(1-v[14])*(1-v[16])),f'{dev} / {20-dev}']
 metrics.append('Financial screen passes' if metrics[0]>0 and isinstance(metrics[1],float) and metrics[1]>v[21] and metrics[8]>=1.2 and metrics[13]=='Pass' else 'Review / reject')
 metrics.append(pb(16))
 for r,val in enumerate(metrics,2):put('Dashboard',f'{chr(64+col)}{r}',val)
 results[sc]={'npv':dcum,'irr':metrics[1],'payback':metrics[2],'discounted_payback':metrics[3],'cost':total,'min_dscr':metrics[8],'operating_only_payback':metrics[-1]}
 if sc=='Intermediate':baseflows=flows;basev=v;basenpv=dcum
for r,shift in [(2,-.01),(3,0),(4,.01)]:
 for c,change in [(2,-.2),(3,0),(4,.2)]:
  val=basenpv+change*sum(x[8] for x in baseflows[1:3])+baseflows[-1][6]*(basev[19]/(basev[19]+shift)-1)/(1+basev[21])**20
  put('Sensitivity',f'{chr(64+c)}{r}',val)
assert abs(cache['Sensitivity']['C3']-results['Intermediate']['npv'])<1e-6
# Check all formula cells receive caches; workbook preserves Excel formulas for edits.
formula_count=0
for ws in w:
 for row in ws:
  for cell in row:
   if cell.data_type=='f':
    formula_count+=1
    assert cell.coordinate in cache.get(ws.title,{}),(ws.title,cell.coordinate)
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ET.register_namespace('',ns['s'])
target=P/'model.cached.tmp.xlsx'
with zipfile.ZipFile(P/'model.xlsx') as src,zipfile.ZipFile(target,'w',zipfile.ZIP_DEFLATED) as out:
 for item in src.infolist():
  data=src.read(item.filename)
  if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
   idx=int(item.filename.rsplit('sheet',1)[1].split('.')[0])-1; sheet=w.sheetnames[idx]
   if sheet in cache:
    tree=ET.fromstring(data)
    for cell in tree.findall('.//s:c',ns):
     coord=cell.attrib['r']
     if coord in cache[sheet]:
      value=cache[sheet][coord]; old=cell.find('s:v',ns)
      if old is not None:cell.remove(old)
      if isinstance(value,str):cell.set('t','str')
      else:cell.pop('t',None) if False else cell.attrib.pop('t',None)
      el=ET.SubElement(cell,'{'+ns['s']+'}v');el.text=str(value)
    data=ET.tostring(tree,encoding='utf-8',xml_declaration=True)
  out.writestr(item,data)
os.replace(target,P/'model.xlsx')
# Reopen both ways to verify formulas and cached numeric values survive.
f=load_workbook(P/'model.xlsx',data_only=False);d=load_workbook(P/'model.xlsx',data_only=True)
for sc in results:
 assert abs(d['Dashboard'].cell(2,['Bad','Intermediate','Optimal'].index(sc)+2).value-results[sc]['npv'])<1e-6
assert f['Dashboard']['B2'].data_type=='f'
report={'formula_cells_checked':formula_count,'checks':['total costs reconcile','no revenue before opening','project cash flows reconcile','debt draw reconciles','sensitivity base equals intermediate NPV','all formula cells cached','workbook reopens with formulas and cached values'],'results':results,'limitations':'Independent Python calculations verify default assumptions; Excel recalculates on editing. No property-specific legal clearance or tax advice.'}
(P/'validation_results.json').write_text(json.dumps(report,indent=2))
print(json.dumps(report,indent=2));print('SAVE_NOW=',(P/'SAVE_NOW.txt').exists())
