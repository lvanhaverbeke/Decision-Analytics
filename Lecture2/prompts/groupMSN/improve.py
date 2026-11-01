import openpyxl,os,shutil
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.workbook.properties import CalcProperties
w=openpyxl.load_workbook('model.xlsx');s=w['Inputs']
for row in [('Price response parameter',1.2,'Illustrative, uncalibrated; test 0.6 to 2'),('Reference weighted rent',480,'EUR/person/month frozen calibration'),('Reference occupancy',.95,'Management assumption'),('Opening occupancy ratio',.85/.95,'Management assumption')]:s.append(row)
s['B55']='=(60*480+40*390+20*650)/120'
s['B9']="=MIN(1,B56*EXP(-B54*(('Room mix'!H5/'Room mix'!D5/B8)/B55-1)))"
s['B29']='=MIN(1,B9*B57)'
for shname in ['Scenarios','Sensitivity','Monte Carlo']:
 sh=w[shname]
 for row in sh:
  r=row[0].row
  if sh.cell(r,6).data_type!='f' or 'NPV(' not in sh.cell(r,6).value:continue
  for c in range(8,31):
   cell=sh.cell(r,c)
   if cell.data_type!='f':continue
   f=cell.value
   # Replace occupancy in revenue; keep construction and discounting unchanged.
   f=f.replace(f'MIN(1,Inputs!B29*C{r}/Inputs!B9)',f'MIN(1,Inputs!B29*C{r}/Inputs!B9*EXP(-Inputs!B54*(D{r}-1)))')
   f=f.replace(f'*D{r}*C{r}*',f'*D{r}*MIN(1,C{r}*EXP(-Inputs!B54*(D{r}-1)))*')
   cell.value=f
w['Sensitivity']['B26']='No closed-form threshold with rent-demand response; inspect sensitivity cases'
w['Sensitivity']['B27']='Linear estimate invalid with occupancy caps; inspect sensitivity cases'
w['Sensitivity']['A28']='Rent-demand response is illustrative exponential; ranges are management stress assumptions, not statistically estimated confidence intervals.'
w['Read me']['A19']='Revised this execution: price-demand feedback in base model, scenarios, sensitivity and Monte Carlo. Elasticity is an explicit unverified assumption.'
w['Decision Summary']['B33']='Pre-tax, no grants; illustrative uncalibrated common rent-demand response. Mix optimization exact for affine per-unit economics and fixed envelope; not a city demand equilibrium.'
w['Class methodology']['B7']='Applied exponential rent-demand response. Elasticity requires estimation; occupancy shocks and rent shocks combine simultaneously.'
w['Decision Summary']['B21']='See sensitivity; higher rent can lower occupancy and NPV.'
# Remove stale claims and replace with checks performed by this run.
del w['Validation'];v=w.create_sheet('Validation');v.append(['Check','Status'])
for a,b in [('Legacy class XLS','Read OLE/BIFF structure, labels, cached numbers and formula tokens; see intermediate/class_inspection.txt'),('Model change','Added price response to base cash flows, scenario and simulation paths'),('Formula evaluation','Run recalculate.py; results retained in intermediate/formula_validation.txt'),('Excel desktop','Not opened interactively; custom evaluator and structural tests only'),('Sensitivity tables','Dynamic equivalent formulas, not native Excel Data Tables'),('Optimization','Existing exact bounded integer enumeration retained; includes zero no-build option'),('Input evidence','Legacy assumptions retained; not reverified market data'),('Uncertainty','300 reproducible triangular scenarios; assumed parameters and independence'),('Limitations','Tax/legal/site/credit diligence required. Elasticity uncalibrated.')]:v.append([a,b])
for n,cell in [('Demand_Elasticity','B54'),('Reference_Occupancy','B56')]:w.defined_names.add(DefinedName(n,attr_text="'Inputs'!"+cell))
w.calculation=CalcProperties(calcId=191029,fullCalcOnLoad=True)
w.save('intermediate/model_price_response.xlsx');shutil.copy2('intermediate/model_price_response.xlsx','.model.tmp.xlsx');os.replace('.model.tmp.xlsx','model.xlsx')
p='recalculate.py';t=open(p).read();t=t.replace("if fn=='SQRT':", "if fn=='EXP':return math.exp(vals[0])\n if fn=='SQRT':")
t=t.replace("'/private/tmp/da_errors.json'","'intermediate/formula_errors.json'").replace("'/private/tmp/results.json'","'intermediate/calculated_values.json'")
open(p,'w').write(t)
print('Updated workbook saved')
