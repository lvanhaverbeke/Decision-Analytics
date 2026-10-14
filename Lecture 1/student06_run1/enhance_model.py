from pathlib import Path
import os,re,math,zipfile,xml.etree.ElementTree as ET,json
from openpyxl import load_workbook
from openpyxl.styles import Font,PatternFill,Alignment
from openpyxl.worksheet.datavalidation import DataValidation
P=Path(__file__).resolve().parent
if (P/'SAVE_NOW.txt').exists():
 print('SAVE_NOW: retained saved workbook; stopping.');raise SystemExit
w=load_workbook(P/'model.xlsx');A=w['Assumptions'];S=w['Sources & Uses']
A.append(['Mortgage secured amount / debt',1.10,'%','Illustrative security uplift; confirm mortgage versus mandate split'])
A.append(['Mortgage duty',0.01,'%','Verified regional rate; applied to secured amount'])
A.append(['Other loan deed / security costs',10000,'EUR','Illustrative notary, inscription and security costs; obtain quote'])
S.append(['Mortgage taxes & other loan costs',"=B12*'Assumptions'!B32*'Assumptions'!B33+'Assumptions'!B34"])
S['B10']='=SUM(B2:B9)+B16'
A['D4']='Verified 12.5% ordinary existing-property acquisition duty; no abatement assumed'
A['D5']='Notary and acquisition diligence allowance; mortgage costs separately included'
A['D27']='Fixed 10-year model; do not edit (extend formulas to change horizon)'
A['B27'].font=Font(color='111111')
for rr in [32,33]:A.cell(rr,2).number_format='0.0%'
for r in range(32,35):A.cell(r,2).font=Font(color='0066CC')
for start,end,lo,hi in [(11,11,0,1),(23,23,0,0.9),(31,31,0,12),(2,2,1,10000),(12,12,1,12)]:
 v=DataValidation(type='decimal',operator='between',formula1=lo,formula2=hi);v.error='Enter a value within the stated range.';v.showErrorMessage=True;A.add_data_validation(v);v.add(f'B{start}:B{end}')
# Calculate every formula for initial cached display, using independent Excel-compatible basic functions.
cache={};active=set()
def flatten(args):
 out=[]
 for x in args:
  out.extend(flatten(x)) if isinstance(x,(list,tuple)) else out.append(x)
 return out

def irr(v):
 lo=-0.9999;hi=10.0
 for _ in range(160):
  mid=(lo+hi)/2; val=sum(x/(1+mid)**i for i,x in enumerate(v))
  if val>0:lo=mid
  else:hi=mid
 return (lo+hi)/2

def pmt(r,n,p):return -p/n if r==0 else -p*r/(1-(1+r)**(-n))
def calc(sh,ref):
 ref=ref.replace('$','');key=(sh,ref)
 if key in cache:return cache[key]
 val=w[sh][ref].value
 if not isinstance(val,str) or not val.startswith('='):return val if isinstance(val,(int,float)) else 0
 if key in active:raise ValueError('Circular reference '+str(key))
 active.add(key);e=val[1:]
 # Replace ranges before individual cell references.
 def ran(m):
  sn=m.group(1) or sh;a=m.group(2).replace('$','');b=m.group(3).replace('$','')
  from openpyxl.utils.cell import range_boundaries
  x1,y1,x2,y2=range_boundaries(a+':'+b)
  from openpyxl.utils import get_column_letter
  return repr([calc(sn,f'{get_column_letter(x)}{y}') for y in range(y1,y2+1) for x in range(x1,x2+1)])
 e=re.sub(r"(?:'([^']+)'!)?(\$?[A-Z]+\$?\d+):(\$?[A-Z]+\$?\d+)",ran,e)
 def one(m):return repr(calc(m.group(1) or sh,m.group(2)))
 e=re.sub(r"(?:'([^']+)'!)?(\$?[A-Z]+\$?\d+)",one,e)
 e=e.replace('^','**')
 funcs={'SUM':lambda *x:sum(flatten(x)),'MIN':min,'MAX':max,'PMT':pmt,'IRR':irr,'NPV':lambda r,v:sum(x/(1+r)**(i+1) for i,x in enumerate(v))}
 result=eval(e,{'__builtins__':{}},funcs);cache[key]=result;active.remove(key);return result
for ws in w:
 for row in ws:
  for c in row:
   if c.data_type=='f':calc(ws.title,c.coordinate)
# Executive dashboard with formulas and investment flags.
D=w.create_sheet('Executive Summary',0)
for row in [['Student housing: investment screening','Illustrative case — 14 October 2026'],['Metric','Value'],['Total project cost',"='Sources & Uses'!B10"],['Initial equity',"='Sources & Uses'!B13"],['Year 1 NOI',"='Operations'!B13"],['Year 1 cash flow after debt',"='Returns'!C3"],['Equity IRR',"='Returns'!B9"],['NPV at 9% required return',"='Returns'!B10"],['Year 1 DSCR',"='Returns'!B14"],['Break-even economic occupancy',"='Returns'!B15"],['Investment view','Do not proceed at illustrative terms: return is below target. Confirm lawful capacity and rent before bidding.'],['Required property data','Address / commune; price; approved room count and plans; floor areas; EPC; surveys; renovation quotes; utility allocation; historic rents / occupancy; tax bill; lender terms; investor tax structure.'],['Legal gate','Open: property-specific permit, fire, quality, rent and tax verification. Regional rules checked; compliance not certified.']]:D.append(row)
D.column_dimensions['A'].width=40;D.column_dimensions['B'].width=100
for c in D[1]:c.fill=PatternFill('solid',fgColor='183B56');c.font=Font(color='FFFFFF',bold=True)
for row in D:
 for c in row:c.alignment=Alignment(wrap_text=True,vertical='top')
for r in range(3,11):D.cell(r,2).number_format='#,##0.00';calc(D.title,f'B{r}')
D['B7'].number_format='0.0%';D['B9'].number_format='0.00x';D['B10'].number_format='0.0%'
D.freeze_panes='B3'
# Store formula caches in workbook XML. Excel recalculates on edits and on load.
t=P/'model.tmp.xlsx';w.save(t)
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile(t) as z:parts={n:z.read(n) for n in z.namelist()}
for i,ws in enumerate(w,1):
 fn=f'xl/worksheets/sheet{i}.xml';root=ET.fromstring(parts[fn])
 for c in root.findall('.//m:c',ns):
  if c.find('m:f',ns) is None:continue
  v=c.find('m:v',ns)
  if v is None:v=ET.SubElement(c,'{'+ns['m']+'}v')
  v.text=str(cache[(ws.title,c.attrib['r'])])
 parts[fn]=ET.tostring(root,encoding='utf-8',xml_declaration=True)
with zipfile.ZipFile(t,'w',zipfile.ZIP_DEFLATED) as z:
 for n,b in parts.items():z.writestr(n,b)
os.replace(t,P/'model.xlsx')
metrics={k:calc('Returns',r) for k,r in [('irr','B9'),('npv','B10'),('multiple','B11'),('yield','B12'),('dscr','B14'),('break_even','B15')]}
metrics.update(cost=calc('Sources & Uses','B10'),equity=calc('Sources & Uses','B13'),noi=calc('Operations','B13'),cash=calc('Returns','C3'))
(P/'metrics.json').write_text(json.dumps(metrics,indent=2))
print(json.dumps(metrics,indent=2));print('Saved enhanced model atomically. SAVE_NOW:',(P/'SAVE_NOW.txt').exists())
