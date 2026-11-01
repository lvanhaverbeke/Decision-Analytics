import openpyxl,math,re,statistics,zipfile,xml.etree.ElementTree as ET,os,json,time
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import range_boundaries
w=openpyxl.load_workbook('model.xlsx');cache={};vis=set();errors=[]
def flat(a):
 out=[]
 for v in a:
  if isinstance(v,list):out+=flat(v)
  else:out.append(v)
 return out
def nums(a):return [float(v) for v in flat(a) if isinstance(v,(int,float))]
def irr(a):
 a=nums(a)
 def f(r):return sum(v/(1+r)**i for i,v in enumerate(a))
 lo=-.95;hi=1.
 try:
  while f(lo)*f(hi)>0 and hi<100:hi*=2
  if f(lo)*f(hi)>0:return 'No IRR'
  for _ in range(80):
   mid=(lo+hi)/2
   if f(lo)*f(mid)<=0:hi=mid
   else:lo=mid
  return (lo+hi)/2
 except:return 'No IRR'
def countif(arr,criterion):
 arr=flat([arr]);m=re.match(r'(<=|>=|<>|<|>|=)?(.*)',str(criterion));op=m[1] or '=';value=m[2]
 try:value=float(value)
 except:pass
 def check(x):
  try:return {'<':lambda:x<value,'>':lambda:x>value,'<=':lambda:x<=value,'>=':lambda:x>=value,'=':lambda:x==value,'<>':lambda:x!=value}[op]()
  except:return False
 return sum(check(x) for x in arr)
class Parser:
 def __init__(self,f):self.ts=[t for t in Tokenizer(f).items if t.type!='WHITE-SPACE'];self.i=0
 def expr(self,p=0):
  t=self.ts[self.i];self.i+=1
  if t.type=='OPERATOR-PREFIX':a=('un',t.value,self.expr(6))
  elif t.type in ('FUNC','PAREN') and t.subtype=='OPEN':
   if t.type=='PAREN':a=self.expr();self.i+=1
   else:
    args=[]
    while not(self.ts[self.i].subtype=='CLOSE'):
     args.append(self.expr())
     if self.ts[self.i].type=='SEP':self.i+=1
     else:break
    self.i+=1;a=('fn',t.value[:-1].upper(),args)
  elif t.type=='OPERAND':
   if t.subtype=='NUMBER':a=('lit',float(t.value))
   elif t.subtype=='TEXT':a=('lit',t.value[1:-1].replace('""','"'))
   elif t.subtype=='LOGICAL':a=('lit',t.value=='TRUE')
   else:a=('ref',t.value)
  else:raise ValueError(('token',t))
  prec={'=':1,'<>':1,'<':1,'>':1,'<=':1,'>=':1,'&':2,'+':3,'-':3,'*':4,'/':4,'^':5}
  while self.i<len(self.ts):
   t=self.ts[self.i]
   if t.type=='OPERATOR-POSTFIX':self.i+=1;a=('bin','/',a,('lit',100));continue
   if t.type!='OPERATOR-INFIX' or prec[t.value]<=p:break
   self.i+=1;a=('bin',t.value,a,self.expr(prec[t.value]-(1 if t.value=='^' else 0)))
  return a
asts={}
def ev(a,sh):
 kind=a[0]
 if kind=='lit':return a[1]
 if kind=='ref':
  ref=a[1].replace('$','')
  if '!' in ref:ss,ref=ref.rsplit('!',1);ss=ss.strip("'").replace("''", "'")
  else:ss=sh
  if ss not in w.sheetnames:raise ValueError('sheet '+ss)
  if ref in w.defined_names:
   target=list(w.defined_names[ref].destinations)[0]
   return get(target[0],target[1].replace('$',''))
  if ':' in ref:
   c1,r1,c2,r2=range_boundaries(ref)
   return [get(ss,w[ss].cell(r,c).coordinate) for r in range(r1,r2+1) for c in range(c1,c2+1)]
  return get(ss,ref)
 if kind=='un':return ev(a[2],sh)*(-1 if a[1]=='-' else 1)
 if kind=='bin':
  x=ev(a[2],sh);y=ev(a[3],sh);op=a[1]
  if op=='&':return str(x)+str(y)
  if op in ['=','<>','<','>','<=','>=']:
   return {'=':lambda:x==y,'<>':lambda:x!=y,'<':lambda:x<y,'>':lambda:x>y,'<=':lambda:x<=y,'>=':lambda:x>=y}[op]()
  return {'+':lambda:x+y,'-':lambda:x-y,'*':lambda:x*y,'/':lambda:x/y,'^':lambda:x**y}[op]()
 fn,args=a[1],a[2]
 if fn=='IF':return ev(args[1] if ev(args[0],sh) else args[2],sh)
 if fn=='IFERROR':
  try:return ev(args[0],sh)
  except:return ev(args[1],sh)
 vals=[ev(x,sh) for x in args];ns=nums(vals)
 if fn=='SUM':return sum(ns)
 if fn=='MIN':return min(ns)
 if fn=='MAX':return max(ns)
 if fn=='ABS':return abs(vals[0])
 if fn=='INT':return math.floor(vals[0])
 if fn=='EXP':return math.exp(vals[0])
 if fn=='SQRT':return math.sqrt(vals[0])
 if fn=='AND':return all(flat(vals))
 if fn=='OR':return any(flat(vals))
 if fn=='COUNT':return len(ns)
 if fn=='COUNTIF':return countif(*vals)
 if fn=='STDEV':return statistics.stdev(ns)
 if fn=='AVERAGE':return statistics.mean(ns)
 if fn=='MEDIAN':return statistics.median(ns)
 if fn=='PERCENTILE':
  vs=sorted(nums([vals[0]]));pos=(len(vs)-1)*vals[1];i=int(pos);return vs[i]+(vs[min(i+1,len(vs)-1)]-vs[i])*(pos-i)
 if fn=='NPV':return sum(v/(1+vals[0])**(i+1) for i,v in enumerate(nums(vals[1:])))
 if fn=='IRR':return irr(vals)
 if fn=='PMT':
  rate,n,pv=vals[:3];return -pv/n if rate==0 else -pv*rate/(1-(1+rate)**(-n))
 if fn=='INDEX':return flat([vals[0]])[int(vals[1])-1]
 if fn=='MATCH':
  for i,v in enumerate(flat([vals[1]])):
   if v==vals[0]:return i+1
  raise ValueError('No match')
 if fn=='RAND':return .5
 raise ValueError('Function '+fn)
def get(sh,ref):
 key=(sh,ref)
 if key in cache:return cache[key]
 if key in vis:raise ValueError('Circular '+str(key))
 vis.add(key)
 try:
  v=w[sh][ref].value
  if isinstance(v,str) and v.startswith('='):
   if v not in asts:asts[v]=Parser(v).expr()
   v=ev(asts[v],sh)
  elif v is None:v=0
  cache[key]=v;return v
 finally:vis.remove(key)
for sh in w:
 for row in sh:
  for c in row:
   if c.data_type=='f':
    try:get(sh.title,c.coordinate)
    except Exception as e:errors.append((sh.title,c.coordinate,str(e),c.value))
print('FORMULA ERRORS',len(errors));print(errors[:10])
if errors:json.dump(errors,open('intermediate/formula_errors.json','w'));raise SystemExit(1)
# cache all formula results without changing formulas or source inputs
ns={'m':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
with zipfile.ZipFile('model.xlsx') as src,zipfile.ZipFile('.model.cached.xlsx','w',zipfile.ZIP_DEFLATED) as dst:
 for item in src.infolist():
  data=src.read(item.filename)
  match=re.fullmatch(r'xl/worksheets/sheet(\d+)\.xml',item.filename)
  if match:
   sh=w.worksheets[int(match[1])-1];root=ET.fromstring(data)
   for c in root.findall('.//m:c',ns):
    if c.find('m:f',ns) is None:continue
    val=cache.get((sh.title,c.attrib['r']));v=c.find('m:v',ns)
    if v is None:v=ET.SubElement(c,'{'+ns['m']+'}v')
    if isinstance(val,str):c.set('t','str');v.text=val
    elif isinstance(val,bool):c.set('t','b');v.text='1' if val else '0'
    else:c.pop('t',None) if False else None;v.text=str(val if val is not None else 0)
   data=ET.tostring(root,encoding='utf-8',xml_declaration=True)
  dst.writestr(item,data)
os.replace('.model.cached.xlsx','model.xlsx');print('CACHED ATOMIC SAVE COMPLETE',len(cache))
json.dump({f'{s}!{c}':v for (s,c),v in cache.items()},open('intermediate/calculated_values.json','w'))
print('SAVE_NOW',os.path.exists('.deadline_runs/20261029T053252Z-5d40d38f/SAVE_NOW.txt'))
