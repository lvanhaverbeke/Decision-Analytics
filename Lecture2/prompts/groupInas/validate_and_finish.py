import math,json,time,zipfile,shutil,statistics,random,os
from pathlib import Path
import openpyxl
from openpyxl.chart import BarChart,Reference
w=openpyxl.load_workbook('model.xlsx',data_only=True)
p={r[0].value:r[1].value for r in w['Inputs'].iter_rows(min_row=2)}
def calc(n,rent,cost=1,demand=1,opex=1,cap=.055,aff=None):
 a=math.ceil(n/10) if aff is None else aff
 invest=p['Acquisition_EUR']*(1+p['Transaction_Rate'])+n*p['Gross_Area_Per_Room']*p['Hard_Cost_Per_m2']*cost*(1+p['Soft_Cost_Rate'])*(1+p['Contingency_Rate'])*(1+p['Nonrecoverable_VAT'])+n*p['Furniture_Per_Room']+p['Working_Capital']
 occ=min(n-a,p['Reference_Demand']*demand*math.exp(-p['Demand_Elasticity']*(rent/p['Reference_Market_Rent']-1)))
 rev=12*(occ*rent+a*min(1,demand)*p['Affordable_Rent_EUR_Month'])
 cf=rev*(1-p['Management_Fee'])-(p['Fixed_Opex']+n*(p['Variable_Opex']+p['Energy_kWh_Per_Room']*p['Energy_Price']))*opex-n*p['Replacement_Reserve']
 cfs=[-invest]+[cf*(1+p['Annual_Growth'])**t for t in range(20)]
 cfs[-1]+=cf*(1+p['Annual_Growth'])**20/cap*(1-p['Selling_Cost_Rate'])
 npv=sum(c/(1+p['Project_Discount_Rate'])**t for t,c in enumerate(cfs))
 feasible=n>=p['Minimum_Rooms'] and n*28<=p['Area_Capacity'] and invest<=p['Budget_EUR'] and rent<=p['Rent_Policy_Max'] and a>=n*.1
 return npv,invest,occ,feasible
base=calc(120,590,aff=12);assert abs(base[0]-w['KPI_Dashboard']['B2'].value)<.001
best=max((calc(n,r)[0],n,r) for n in range(40,141) for r in range(400,651,5) if calc(n,r)[3]);assert abs(best[0]-w['Optimization']['I2'].value)<.001
assert calc(120,590,cost=1.1)[0]<base[0]
assert calc(120,730)[2]<base[2]
assert calc(120,590,demand=.8)[0]<base[0]
for row in list(w['Monte_Carlo'].iter_rows(min_row=2,max_row=1001))[:40]:
 vals=[c.value for c in row[:6]];assert abs(calc(120,590,*vals[1:5],aff=12)[0]-vals[5])<.001
for r in range(2,7):assert w['Model_Checks'].cell(r,2).value==True
assert w['Model_Checks']['B7'].value==1000
assert w['Model_Checks']['B8'].value==1000
result={'independent_annual_NPV':base[0],'best_grid_NPV':best[0],'best_rooms':best[1],'best_rent':best[2],'checks':'Base and optimal NPV reconciled to independent annual cash flow; price/demand and cost shocks tested; first 40 MC rows reconciled; affordability/debt/capacity/histogram passed','limitations':['Native Excel execution not performed','Triangular endpoints and independence subjective','All costs and site demand require validation','No floor plan or studio mix optimization']}
Path('intermediate/independent_validation.json').write_text(json.dumps(result,indent=2))
# Add interest sensitivity on equity: project NPV correctly excludes financing.
f=openpyxl.load_workbook('model.xlsx');s=f.create_sheet('Interest_Sensitivity');s.append(['Interest rate','Annual debt service','Equity NPV EUR'])
for r,rate in enumerate([.03,.04,.045,.05,.06,.07],2):
 s.append([rate,f'=-PMT(A{r},Loan_Term,Initial_Investment_EUR*Debt_Ratio)',f'=-Initial_Investment_EUR*(1-Debt_Ratio)'+''.join(f'+(Cash_Flows!E{t+2}-B{r})/(1+Equity_Discount_Rate)^{t}' for t in range(1,21))])
s['E1']='Project cash flow is unlevered, so interest affects equity NPV only. Loan term is fixed 20 years here.'
v=f.create_sheet('Verified_Tests');v.append(['Performed test','Result'])
for k,val in result.items():v.append([k,str(val)])
for sh in [s,v]:
 sh.column_dimensions['A'].width=38;sh.column_dimensions['B'].width=36;sh.column_dimensions['C'].width=30
f.save('.model.tmp.xlsx');os.replace('.model.tmp.xlsx','model.xlsx')
print(json.dumps(result))
