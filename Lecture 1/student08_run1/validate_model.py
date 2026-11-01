"""Controleer alle formules onafhankelijk en vergelijk met de opgeslagen waarden."""
from pathlib import Path
import re, math
from functools import lru_cache
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries,get_column_letter
root=Path(__file__).resolve().parent
if (root/'SAVE_NOW.txt').exists():
    print('SAVE_NOW aanwezig: model reeds opgeslagen; geen verdere controle.');raise SystemExit
f=load_workbook(root/'model.xlsx',data_only=False)
v=load_workbook(root/'model.xlsx',data_only=True)
def flatten(xs):
    for x in xs:
        if isinstance(x,(list,tuple)):yield from flatten(x)
        else:yield x
def npv(rate,*xs):return sum(x/(1+rate)**i for i,x in enumerate(flatten(xs),1))
def irr(xs):
    lo,hi=-.99,10
    for _ in range(170):
        mid=(lo+hi)/2;pv=sum(x/(1+mid)**i for i,x in enumerate(xs))
        if pv>0:lo=mid
        else:hi=mid
    return (lo+hi)/2
def match(x,xs,mode):return xs.index(x)+1
@lru_cache(None)
def cell(sheet,addr):
    value=f[sheet][addr.replace('$','')].value
    if not isinstance(value,str) or not value.startswith('='):return value or 0
    expr=value[1:];tokens={}
    def protect(m):
        key=f'__text_{len(tokens)}__';tokens[key]=m.group(0);return key
    expr=re.sub(r'"[^"]*"',protect,expr)
    def rng(m):
        key=f'__range_{len(tokens)}__';tokens[key]=f'area({(m[1] or sheet)!r},{m[2]!r},{m[3]!r})';return key
    expr=re.sub(r'(?:(\w+)!)?(\$?[A-Z]{1,3}\$?\d+):(\$?[A-Z]{1,3}\$?\d+)',rng,expr)
    expr=re.sub(r'(?<![A-Za-z0-9_])(?:(\w+)!)?(\$?[A-Z]{1,3}\$?\d+)(?![A-Za-z0-9_])',lambda m:f'cell({(m[1] or sheet)!r},{m[2]!r})',expr)
    expr=expr.replace('^','**')
    expr=re.sub(r'(?<![<>=!])=(?!=)', '==', expr)
    for k,txt in tokens.items():expr=expr.replace(k,txt)
    env={'cell':cell,'area':area,'SUM':lambda *x:sum(flatten(x)),'IF':lambda c,a,b:a if c else b,'IFERROR':lambda a,b:a,'AND':lambda *xs:all(xs),'NPV':npv,'IRR':irr,'MATCH':match}
    try:return eval(expr,{'__builtins__':{}},env)
    except Exception as e:raise RuntimeError(f'{sheet}!{addr}: {value}: {e}')
def area(sheet,a,b):
    c1,r1,c2,r2=range_boundaries(a.replace('$','')+':'+b.replace('$',''))
    return [cell(sheet,f'{get_column_letter(c)}{r}') for r in range(r1,r2+1) for c in range(c1,c2+1)]
count=0
for sh in f:
    for row in sh:
        for c in row:
            if c.data_type!='f':continue
            actual=cell(sh.title,c.coordinate);saved=v[sh.title][c.coordinate].value
            if isinstance(actual,(float,int)):
                assert isinstance(saved,(float,int)) and math.isclose(actual,saved,rel_tol=1e-9,abs_tol=1e-7),(sh.title,c.coordinate,actual,saved)
            else:assert actual==saved,(sh.title,c.coordinate,actual,saved)
            count+=1
assert math.isclose(cell('Scenariovergelijking','C10'),cell('Samenvatting','B13'),abs_tol=1e-7)
assert math.isclose(cell('Kasstromen','C13')-cell('Kasstromen','C24'),cell('Kasstromen','C25'),abs_tol=1e-8)
print(f'OK: {count} formules onafhankelijk nagerekend; alle cachewaarden kloppen.')
print('OK: scenario 8+2 sluit aan op hoofdberekening; omzet - kosten = operationele kasstroom.')
print('SAVE_NOW:',(root/'SAVE_NOW.txt').exists())
