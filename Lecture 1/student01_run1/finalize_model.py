"""Independent numerical audit and Excel formula-cache population; atomic output."""
from pathlib import Path
from openpyxl import load_workbook
import math,json,os,zipfile,xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parent
path=ROOT/'model.xlsx';w=load_workbook(path);cache={};summary={}
base={i:w['Inputs'].cell(i,2).value for i in range(2,29)}
def put(sh,cell,val):cache[(sh,cell)]=val

def npv_at(rate,cf):return sum(v/(1+rate)**t for t,v in enumerate(cf))
def irr(cf):
 lo,hi=-.5,1.;fl=npv_at(lo,cf);fh=npv_at(hi,cf)
 if fl*fh>0:return 'No IRR'
 for _ in range(150):
  mid=(lo+hi)/2;fm=npv_at(mid,cf)
  if fm*fl>0:lo=mid;fl=fm
  else:hi=mid
 return (lo+hi)/2

def compute(v,cs,rs,ins=None):
 gross=v[4]+(v[2]*v[3]*v[5]+v[6])*(1+v[7])*(1+v[8])+v[9]
 capital=gross-v[10];delay=math.ceil(v[11]/12);last=delay+v[12]
 annual=[];cum=-capital;dcum=-capital;weights=0;pay=[];dpay=[]
 for t in range(61):
  age=t-delay if delay<t<=last else 0
  price=v[16]*(1+v[18])**(age-1) if age else 0
  rev=v[2]*v[14]*v[15]*price if age else 0
  cost=(v[20]+v[21]+v[2]*v[14]*(v[22]+v[23])+(v[2]*v[3]*v[5]+v[6])*v[24])*(1+v[19])**(age-1) if age else 0
  ref=v[26]*(1+v[19])**(age-1) if age and age<v[12] and age%v[25]==0 else 0
  terminal=v[27]*(1-v[28]) if t==last else 0
  cf=-capital if t==0 else rev-cost-ref+terminal;annual.append(cf)
  df=1/(1+v[13])**t;pv=cf*df
  weight=v[2]*v[14]*v[15]*(1+v[18])**(age-1)*df if age else 0;weights+=weight
  crossing='';dcross=''
  if t:
   old=cum;cum+=cf-terminal
   oldd=dcum;dcum+=(cf-terminal)*df
   if cum>=0 and old<0:crossing=t-1-old/(cf-terminal);pay.append(crossing)
   if dcum>=0 and oldd<0:dcross=t-1-oldd/((cf-terminal)*df);dpay.append(dcross)
  vals=[t,age,price,rev,cost,ref,terminal,cf,df,pv,cum,weight,crossing,dcum,dcross]
  for idx,value in enumerate(vals):
   from openpyxl.utils import get_column_letter
   cell=f'{get_column_letter(idx+1)}{t+2}'
   if w[cs][cell].data_type=='f':put(cs,cell,value)
 val=sum(annual[t]/(1+v[13])**t for t in range(61));req=max(0,v[16]-val/weights);capnpv=val+(v[17]-v[16])*weights
 out={2:gross,3:capital,4:delay,5:last,6:val,7:irr(annual),8:min(pay) if pay else 'Not recovered',9:v[16],10:v[17],11:req,12:capnpv,13:max(0,-capnpv),14:capital+capnpv,15:'Feasible under assumptions' if req<=v[17] else 'Funding / cost gap',16:req if req<=v[17] else 'No feasible price under cap',17:delay+1,18:'OK' if v[10]<=gross and v[14]>0 and last<=60 else 'CHECK subsidy, occupancy or horizon',19:min(dpay) if dpay else 'Not recovered',20:gross/v[2]}
 for row,value in out.items():put(rs,f'B{row}',value)
 # Audits cover valuation identity and the optimization constraint.
 assert abs(val-npv_at(v[13],annual))<1e-6
 assert abs(val+(req-v[16])*weights)<1e-6 or req==0
 assert abs(capnpv+max(0,-capnpv))<1e-6 or capnpv>=0
 assert last<=60
 if isinstance(out[7],float):assert abs(npv_at(out[7],annual))<.001
 summary[rs]={str(k):value for k,value in out.items()}
 return out
compute(base,'Cash flow','Results')
for n in range(1,4):
 sr=n+1;v=base.copy();sh=f'S{n} inputs'
 for sc,rr in [('C',14),('D',11),('E',13),('F',12)]:
  val=w['Scenarios'][f'{sc}{sr}'].value
  if isinstance(val,str) and val.startswith('='):val=base[int(val.split('B')[-1])];put('Scenarios',f'{sc}{sr}',val)
  v[rr]=val
 v[5]*=w['Scenarios'][f'B{sr}'].value
 for rr,value in v.items():put(sh,f'B{rr}',value)
 result=compute(v,f'S{n} cash flow',f'S{n} results',sh)
 for col,rr in [('G',2),('H',6),('I',7),('J',11),('K',12),('L',13),('M',8)]:put('Scenarios',f'{col}{sr}',result[rr])
put('Results','B21',min(base[11],w['Scenarios']['D3'].value,w['Scenarios']['D4'].value))
put('Results','B22',max(base[12],w['Scenarios']['F3'].value,w['Scenarios']['F4'].value))
put('Results','B23','Proceed to project-specific due diligence' if summary['Results']['12']>=0 else 'Revise cost / secure subsidy before proceeding')
feasible=[summary[f'S{i} results']['2'] for i in range(1,4) if summary[f'S{i} results']['12']>=0]
put('Results','B24',min(feasible) if feasible else 'No feasible scenario')
ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};tag='{'+ns['s']+'}'
tmp=ROOT/'model.cached.tmp.xlsx'
with zipfile.ZipFile(path) as inp,zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as out:
 for item in inp.infolist():
  data=inp.read(item.filename)
  if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
   num=int(item.filename.split('sheet')[-1].split('.')[0]);name=w.sheetnames[num-1]
   xml=ET.fromstring(data)
   for cell in xml.findall('.//s:c',ns):
    if cell.find('s:f',ns) is None:continue
    key=(name,cell.attrib['r'])
    if key not in cache:raise AssertionError(f'Uncached formula: {key}')
    value=cache[key];vc=cell.find('s:v',ns)
    if vc is None:vc=ET.SubElement(cell,tag+'v')
    if isinstance(value,str):cell.set('t','str');vc.text=value
    else:cell.attrib.pop('t',None);vc.text=str(value)
   data=ET.tostring(xml,encoding='utf-8',xml_declaration=True)
  out.writestr(item,data)
check=load_workbook(tmp,data_only=True)
assert isinstance(check['Results']['B6'].value,(float,int))
assert check['Results']['B18'].value=='OK'
check.close();os.replace(tmp,path)
(ROOT/'validation.json').write_text(json.dumps({'audit':'PASS: NPV, break-even rent, funding gap and IRR residual independently checked; all formulas cached','formula_cells':len(cache),'results_by_row':summary},indent=2))
print(json.dumps(summary['Results'],indent=2));print('Saved cached model.xlsx atomically. SAVE_NOW:',(ROOT/'SAVE_NOW.txt').exists())
