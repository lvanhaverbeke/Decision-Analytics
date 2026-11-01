import openpyxl,json,math,zipfile,shutil,time,datetime,csv
from pathlib import Path
from document_tools import pdf
w=openpyxl.load_workbook('model.xlsx',data_only=True);p={r[0].value:r[1].value for r in w['Inputs'].iter_rows(min_row=2)}
def log(s):
 with open('activity.log','a') as f:f.write(datetime.datetime.now(datetime.timezone.utc).isoformat()+' '+s+'\n')
def model(nm=None,ns=None,rent=None,occ=None,occs=None,acq=None,cost=1,discount=None,exitmul=1):
 nm=p['Market_Rooms'] if nm is None else nm;ns=p['Social_Rooms'] if ns is None else ns;rent=p['Market_Base_Rent'] if rent is None else rent;occ=p['Market_Occupancy'] if occ is None else occ;occs=p['Social_Occupancy'] if occs is None else occs;acq=p['Acquisition_Price'] if acq is None else acq;discount=p['Project_Discount'] if discount is None else discount;n=nm+ns
 invest=acq*(1+p['Acquisition_Duty'])+p['Works_Cost']*cost*(1+p['Soft_Fee_Rate'])*(1+p['Contingency_Rate'])*(1+p['Works_VAT'])+n*p['Furniture_Per_Room']+p['Legal_Fees']+p['Working_Capital']
 receipt=(nm*occ*(rent+p['Monthly_Charges'])+ns*occs*(p['Social_Base_Rent']+p['Monthly_Charges']))*p['Months_Per_Year'];op=n*(p['Monthly_Charges']*12+p['Maintenance_Per_Room']+p['Property_Tax_Per_Room']+p['Insurance_Per_Room'])+p['Fixed_Admin'];cfs=[-invest]
 for t in range(1,21):cfs.append(receipt*(1-p['Management_Rate'])*(1+p['Rent_Growth'])**(t-1)-(op+n*p['Replacement_Per_Room'])*(1+p['Expense_Growth'])**(t-1))
 terminal=(receipt*(1-p['Management_Rate'])*(1+p['Rent_Growth'])**20-op*(1+p['Expense_Growth'])**20)/p['Exit_Cap_Rate']*(1-p['Selling_Fee'])*exitmul
 cfs[-1]+=terminal+p['Working_Capital'];npv=sum(c/(1+discount)**t for t,c in enumerate(cfs))
 return npv,invest,cfs,terminal
base=model();assert abs(base[0]-w['Results']['B2'].value)<.001
no_sale=base[0]-base[3]/(1+p['Project_Discount'])**20;assert abs(no_sale-w['Results']['C2'].value)<.001
assert model(cost=1.2)[0]<base[0] and model(occ=.8)[0]<base[0] and model(rent=550)[0]>base[0]
maxacq=p['Acquisition_Price']+base[0]/(1+p['Acquisition_Duty']);assert abs(model(acq=maxacq)[0])<.001
best=max((model(nm=120-s,ns=s)[0],s) for s in range(121) if s>=p['Required_Social_Quota']*120 and model(nm=120-s,ns=s)[1]<=p['Budget_Limit']);assert abs(best[0]-w['Optimization']['I1'].value)<.001
for row in list(w['Monte_Carlo'].iter_rows(min_row=2,max_row=1001))[:50]:
 vals=[c.value for c in row[:6]];assert abs(model(rent=p['Market_Base_Rent']*vals[3],occ=vals[2],cost=vals[1],exitmul=vals[4])[0]-vals[5])<.001
for r in range(2,13):assert w['Model_checks'].cell(r,2).value==True,('check',r,w['Model_checks'].cell(r,2).value)
# Debt independently reconciles to opening + draws - principal each year; initial sources equal uses.
for col in range(3,23):assert abs(w['Financing'].cell(2,col).value-w['Financing'].cell(5,col).value-w['Financing'].cell(7,col).value)<.001
assert w['Model_checks']['B13'].value==False
result={'base_project_NPV':base[0],'no_sale_NPV':no_sale,'project_IRR':w['Results']['B3'].value,'equity_NPV':w['Results']['B4'].value,'simple_payback':w['Results']['B6'].value,'discounted_payback':w['Results']['B7'].value,'best_mix_NPV':best[0],'best_social_rooms':best[1],'maximum_acquisition':maxacq,'MC_negative_probability':w['Monte_Carlo']['I4'].value,'independent_checks':'Annual NPV with/without sale; all 121 mix alternatives; 50 MC trials; zero-NPV purchase ceiling; input direction and debt reconciliation passed','limits':['Quota applicability unverified; no site legal clearance','No native Excel desktop or Solver execution','Pre-tax: corporate income taxes excluded','Commute times and land budgets unverified','Occupancy independent of rent in this basic course model; no demand elasticity calibration']}
Path('intermediate/independent_validation.json').write_text(json.dumps(result,indent=2))
log('VALIDATION: formula evaluator reported zero errors. Independently reconciled annual DCF, terminal exclusion, integer optimization, first 50 MC trials, acquisition ceiling and debt. Quota approval correctly remains FALSE.')
with open('intermediate/Monte_Carlo_trials.csv','w',newline='') as f:
 writer=csv.writer(f);writer.writerow([c.value for c in w['Monte_Carlo'][1][:6]])
 for row in w['Monte_Carlo'].iter_rows(min_row=2,max_row=1001):writer.writerow([c.value for c in row[:6]])
brief=f'''Illustrative Brussels student residence investment\nBase project NPV: EUR {base[0]:,.0f}; project IRR {w['Results']['B3'].value:.2%}.\nNPV excluding terminal sale: EUR {no_sale:,.0f}. Excluding sale omits a residual asset that remains owned.\nMaximum zero-NPV acquisition price: EUR {maxacq:,.0f}.\nBest social allocation: {best[1]} of 120 rooms; NPV EUR {best[0]:,.0f}.\nProbability negative NPV under subjective independent triangular inputs: {result['MC_negative_probability']:.1%}.\nRecommendation: HOLD pending legal/site clearance; financial recommendation is conditional on the model assumptions and DSCR.\nLegal quota reference: City of Brussels 2021 guidance 20% from 30 units; applicability to selected municipality City of Brussels remains unverified.\nLimitations: pre-tax, no construction delay schedule; no native Solver/Data Table/Excel test. Travel times and land prices are illustrative.\n'''
Path('completion_summary.txt').write_text(brief)
pages=[('Student housing decision - executive summary',brief.splitlines()[1:]),('Financial model and audit',['Year 0 contains acquisition, duties, works, professional fees, contingency, works VAT, furniture and working capital. Operations run years 1-20. Nominal rent growth and expense inflation are separate inputs.','Project cash flows exclude debt financing. Equity cash flows deduct interest and principal; remaining debt repaid on sale. Working capital is released at exit.','Simple and discounted payback are presented separately with/without terminal sale. Negative financial outcomes are preserved.','Monte Carlo uses 1000 seeded independent triangular draws for works, rent, occupancy and exit value. Parameters are management assumptions, not calibrated empirical probabilities.']),('Sources and outstanding legal checks',[str(row[1].value)+' '+str(row[2].value) for row in w['Legal_and_sources'].iter_rows(min_row=2,max_row=8)]+['Selected building planning permit, university agreement, fire/accessibility review, tax assessment and measured layout remain outstanding.']),('Validation and next actions',[result['independent_checks']]+result['limits']+['Obtain actual offers and an address, verify municipal permit conditions, replace assumed travel/land budgets, and secure finance before committing.'])]
pdf('Executive_Report.pdf',pages)
for file in ['model.xlsx']:
 with zipfile.ZipFile(file) as z:assert z.testzip() is None
shutil.copy2('model.xlsx','intermediate/model_validated.xlsx')
log('OUTPUTS: model.xlsx, Executive_Report.pdf, completion_summary.txt, intermediate/model_core.xlsx, model_analysis.xlsx, model_validated.xlsx, Monte_Carlo_trials.csv, calculated_values.json, independent_validation.json and formula_validation.txt. No source inputs overwritten.')
a=json.load(open('execution_started.json'));a.update(end_epoch=time.time(),elapsed_seconds=time.time()-a['start_epoch'],completed=True,outputs=['model.xlsx','Executive_Report.pdf','completion_summary.txt'],limitations=result['limits']);assert a['elapsed_seconds']<420;Path('execution_summary.json').write_text(json.dumps(a,indent=2));log('COMPLETE in %.2f seconds (limit 420).'%a['elapsed_seconds']);print(json.dumps(a));print(json.dumps(result))
