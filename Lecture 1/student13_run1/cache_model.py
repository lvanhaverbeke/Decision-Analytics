"""Calculate initial workbook values independently and cache them for spreadsheet previews.
Excel formulas remain intact and recalculate when inputs change.
"""
import math, zipfile, xml.etree.ElementTree as ET, os
from pathlib import Path

def cache_workbook(wb,path):
 vals={};O=wb['Opportunities'];A=wb['Assumptions'];C=wb['Compliance'];b=lambda r:A.cell(r,2).value
 def put(s,cell,v):vals[(s,cell)]=v
 def irr(cf):
  if not cf[0] or not any(x>0 for x in cf[1:]):return 0
  lo,hi=-.95,10
  for _ in range(180):
   mid=(lo+hi)/2;v=sum(x/(1+mid)**y for y,x in enumerate(cf))
   if v>0:lo=mid
   else:hi=mid
  return (lo+hi)/2
 projects=[]
 for i in range(8):
  r=8+i;id,name,u,rent,p,c,other,d,occ,oc,f=[O.cell(r,j).value for j in range(1,12)]
  cap=p*(1+b(13))+c*(1+b(14))+other if name else 0
  statuses=[C.cell(5+12*i+j,4).value for j in range(12)];evidence=[C.cell(5+12*i+j,5).value for j in range(12)]
  gate='EMPTY' if not name else ('FAIL' if 'Fail' in statuses else ('PASS' if all(x=='Verified' for x in statuses) and all(evidence) else ('DEMO PASS' if b(5)=='YES' and all(x=='Demo pass' for x in statuses) else 'UNVERIFIED')))
  flows=[]
  for stress,fr in [(False,5+i),(True,18+i)]:
   h=b(15) if stress else 0;delay=d+(b(16) if stress else 0)
   noi=u*rent*12*occ*(1-oc)*(1-h)-f
   cf=[-cap-(c*(1+b(14))*b(17) if stress else 0)]
   for y in range(1,11):cf.append(noi*(1+b(10))**(y-1)*max(0,min(1,(12*y-delay)/12)))
   cf[-1]+=max(0,noi*(1+b(10))**10/b(11))*(1-b(12))
   put('Cash flows',f'A{fr}',name or '')
   for j,x in enumerate(cf,2):put('Cash flows',f'{chr(64+j)}{fr}',x)
   flows.append(cf)
  npvs=[sum(x/(1+b(8))**y for y,x in enumerate(cf)) for cf in flows]
  eligible=bool(name and gate in ['PASS','DEMO PASS'] and 0<=d<=b(18) and u>0 and rent>0 and p>0 and 0<occ<=1 and 0<=oc<1 and min(c,other,f)>=0 and npvs[0]>0)
  result={'name':name or '', 'cap':cap,'gate':gate,'npv':npvs[0],'stress':npvs[1],'irr':irr(flows[0]),'delay':d,'eligible':eligible,'profit':sum(flows[0])}
  projects.append(result)
  for col,v in {'L':cap,'M':gate,'N':flows[0][1],'O':(u*rent*12*occ*(1-oc)-f)*(1+b(10))**10,'P':npvs[0],'Q':result['irr'],'R':npvs[1],'S':'EMPTY' if not name else ('ELIGIBLE' if eligible else 'EXCLUDE')}.items():put('Opportunities',f'{col}{r}',v)
 ranked=sorted([i for i,p in enumerate(projects) if p['eligible']],key=lambda i:(-projects[i]['npv'],i))
 for i in range(8):put('Opportunities',f'T{8+i}',ranked.index(i)+1 if i in ranked else '')
 subsets=[]
 for mask in range(256):
  bits=[(mask>>i)&1 for i in range(8)];selected=[p for i,p in enumerate(projects) if bits[i]]
  capital=sum(p['cap'] for p in selected);npv=sum(p['npv'] for p in selected);feasible=capital<=b(6) and len(selected)<=b(7) and all(p['eligible'] for p in selected);obj=npv if feasible else -1e99
  subsets.append((obj,capital,bits,len(selected)))
  for col,v in {'J':capital,'K':len(selected),'L':npv,'M':int(feasible),'N':obj}.items():put('Portfolio',f'{col}{5+mask}',v)
 winner=max(range(256),key=lambda m:subsets[m][0]);objective,capital,bits,count=subsets[winner]
 for i,bit in enumerate(bits):put('Opportunities',f'U{8+i}',bit)
 for cell,v in {'B4':'DEMONSTRATION — hypothetical opportunities' if b(5)=='YES' else 'LIVE INPUTS — evidence required','B6':objective,'B7':winner+1,'B8':capital,'B9':b(6)-capital,'B10':count,'B11':'No eligible positive-NPV portfolio — retain capital' if not count else ('Illustration only; obtain actual opportunities and evidence before investment' if b(5)=='YES' else 'Selected projects require final due diligence and transaction approval'),'B34':sum(bool(p['name']) and not p['eligible'] for p in projects),'B35':sum(p['stress']*bits[i] for i,p in enumerate(projects)),'B36':sum(p['profit']*bits[i] for i,p in enumerate(projects))}.items():put('Report',cell,v)
 for j in range(8):
  row=15+j
  if j>=len(ranked):
   for col in 'ABCDEFGHI':put('Report',f'{col}{row}','')
  else:
   i=ranked[j];p=projects[i]
   for col,v in zip('ABCDEFGHI',[j+1,p['name'],p['npv'],p['irr'],p['cap'],p['delay'],p['stress'],p['gate'],'SELECTED' if bits[i] else 'WATCHLIST']):put('Report',f'{col}{row}',v)
 assert capital<=b(6) and count<=b(7)
 assert all(not bits[i] or p['eligible'] for i,p in enumerate(projects))
 assert projects[4]['eligible']==False and projects[5]['eligible']==False
 assert all(p['stress']<=p['npv'] for p in projects)
 assert objective==max(s[0] for s in subsets)
 ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'};ET.register_namespace('',ns['m']);tmp=Path(path).with_name('cached.pending.xlsx')
 with zipfile.ZipFile(path) as zin,zipfile.ZipFile(tmp,'w',zipfile.ZIP_DEFLATED) as zout:
  for item in zin.infolist():
   content=zin.read(item.filename)
   if item.filename.startswith('xl/worksheets/sheet') and item.filename.endswith('.xml'):
    n=int(item.filename.split('sheet')[-1].split('.')[0]);sheet=wb.sheetnames[n-1];root=ET.fromstring(content)
    for cell in root.findall('.//m:c',ns):
     key=(sheet,cell.attrib['r'])
     if key not in vals:continue
     val=vals[key];v=cell.find('m:v',ns)
     if v is None:v=ET.SubElement(cell,'{'+ns['m']+'}v')
     if isinstance(val,str):cell.set('t','str');v.text=val
     else:cell.attrib.pop('t',None);v.text=str(val)
    content=ET.tostring(root,encoding='utf-8',xml_declaration=True)
   zout.writestr(item,content)
 os.replace(tmp,path)
 Path(path).parent.joinpath('validation.txt').write_text(f'Validated all 256 portfolios; selected mask {winner}; capital EUR {capital:,.0f}; base NPV EUR {objective:,.0f}.\nDelay and compliance exclusions passed. Downside no better than base. Formula cells retain recalculation.\n')
