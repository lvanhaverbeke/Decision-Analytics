from pathlib import Path
import openpyxl,shutil,json,time,datetime,subprocess,sys,math,os
start=json.load(open('execution_started.json'))['start_epoch']
def log(s):
 with open('activity.log','a') as f:f.write(datetime.datetime.now(datetime.timezone.utc).isoformat()+' '+s+'\n')
def recalc(label):
 with open('intermediate/'+label+'_formula_validation.txt','w') as f:subprocess.run([sys.executable,'recalculate.py'],stdout=f,stderr=subprocess.STDOUT,check=True,timeout=max(1,400-(time.time()-start)))
def observe(label):
 w=openpyxl.load_workbook('model.xlsx',data_only=True);vals=sorted(r[5].value for r in w['Monte_Carlo'].iter_rows(min_row=2,max_row=1001));pos=.1*(len(vals)-1);j=int(pos);p10=vals[j]+(vals[j+1]-vals[j])*(pos-j)
 p={r[0].value:r[1].value for r in w['Inputs'].iter_rows(min_row=2)}
 result={'iteration':label,'social_rooms':p['Social_Rooms'],'required_quota':p['Required_Social_Quota'],'acquisition':p['Acquisition_Price'],'works':p['Works_Cost'],'NPV':w['Results']['B2'].value,'loss_probability':w['Monte_Carlo']['I4'].value,'NPV_P10':p10,'equity_funded':w['Cost_summary']['B12'].value*(1-p['Debt_Ratio'])}
 shutil.copy2('model.xlsx','intermediate/model_'+label+'.xlsx');log('EVALUATED '+json.dumps(result));return result
# Keep the early prototype as a separate intermediate; final concept is one office conversion, not mixed project types.
shutil.copy2('model.xlsx','intermediate/model_preliminary_prototype.xlsx')
w=openpyxl.load_workbook('model.xlsx');ref={r[0].value:r[1] for r in w['Inputs'].iter_rows(min_row=2)};ref['Acquisition_Price'].value=3000000;ref['Works_Cost'].value=2200000
w['README']['B2']='Illustrative conversion of an office property into 120 student rooms in City of Brussels; new student units subject to planning and social allocation review. No actual property secured.'
w.save('.iter.xlsx');os.replace('.iter.xlsx','model.xlsx');recalc('iteration_1');a=observe('iteration_1')
w=openpyxl.load_workbook('model.xlsx');ref={r[0].value:r[1] for r in w['Inputs'].iter_rows(min_row=2)}
ref['Market_Rooms'].value=96;ref['Social_Rooms'].value=24;ref['Required_Social_Quota'].value=.2
ref['Required_Social_Quota'].offset(column=2).value='Constant';ref['Required_Social_Quota'].offset(column=3).value='Assumption pending permit';ref['Quota_Verified'].offset(column=4).value='City guidance verified historically; actual permit/university allocation terms pending.'
w['README']['B6']='City of Brussels 2021 municipal guidance: 20% social student units from 30 units. Apply 20% as a conservative design constraint to this hypothetical 120-unit office conversion. Exact permit and agreement remain unverified; not a universal Brussels law.'
w.save('.iter.xlsx');os.replace('.iter.xlsx','model.xlsx');recalc('iteration_2');b=observe('iteration_2')
# Risk policy: 90% non-loss probability in the specified subjective simulation. NPV shifts exactly by -(1+duty)*purchase change.
newprice=max(0,min(3000000,3000000+b['NPV_P10']/1.125))
w=openpyxl.load_workbook('model.xlsx');ref={r[0].value:r[1] for r in w['Inputs'].iter_rows(min_row=2)};ref['Acquisition_Price'].value=newprice
s=w.create_sheet('Agile_iterations');s.append(['Iteration','Social rooms','Quota','Acquisition EUR','Works EUR','NPV EUR','P(loss)','NPV P10','Equity EUR','Decision passed to next iteration'])
for result,action in [(a,'Preliminary 10% social design; not approved. Apply published municipal policy in next iteration.'),(b,'Adopt 20% social rooms; negotiate price down to meet assumed 90% non-loss risk policy. No relaxation of social rent ceiling.')]:s.append([result[k] for k in ['iteration','social_rooms','required_quota','acquisition','works','NPV','loss_probability','NPV_P10','equity_funded']]+[action])
s.append(['iteration_3',24,.2,'=Acquisition_Price','=Works_Cost','=Results!B2','=Monte_Carlo!I4','=PERCENTILE(Monte_Carlo!F2:F1001,0.1)','=Initial_Investment*(1-Debt_Ratio)','Conditional negotiated purchase ceiling. Not an actual vendor offer or investment.'])
s['A7']='The 90% non-loss target is an investor assumption. Same 1000 shock draws used across iterations for comparability; it is not an empirical confidence guarantee.'
s['A8']='Equity EUR is hypothetical funding if approved, not money spent or profits reinvested. No external investment occurred.'
s['A9']='If the risk-adjusted acquisition ceiling is zero, test feasibility at zero before any commitment. Bid and engineering validation may make the ceiling unattainable.'
cl=w.create_sheet('Input_classification');cl.append(['Input','Top-level category','Uncertainty subtype','Explanation'])
for row in w['Inputs'].iter_rows(min_row=2):
 name,cell,unit,category,status,why=[c.value for c in row]
 if 'uncertainty' in str(category).lower():top='Uncertain';sub='Approximate' if 'Approximate' in category else 'Inherently uncertain'
 else:top='Control / decision' if category=='Control' else 'Constant';sub=''
 cl.append([name,top,sub,why])
from openpyxl.chart import BarChart,Reference
ch=BarChart();ch.title='Model iteration NPV (EUR)';ch.add_data(Reference(s,min_col=6,min_row=1,max_row=4),titles_from_data=True);ch.set_categories(Reference(s,min_col=1,min_row=2,max_row=4));s.add_chart(ch,'L2')
for sh in [s,cl]:sh.freeze_panes='B2';sh.column_dimensions['A'].width=32;sh.column_dimensions['B'].width=28
w.save('.iter.xlsx');os.replace('.iter.xlsx','model.xlsx');recalc('formula_validation');c=observe('iteration_3');Path('intermediate/iteration_evaluations.json').write_text(json.dumps([a,b,c],indent=2))
subprocess.run([sys.executable,'validate_finish.py'],check=True,timeout=max(1,410-(time.time()-start)))
from document_tools import pdf
report=[('Student housing in City of Brussels - iterative decision',['Illustrative office conversion: 120 rooms, including 24 social rooms in final design. All prices and project-specific feasibility are assumptions.','Municipal distinction: City of Brussels publishes 2021 guidance for 20% social student housing from 30 created units. Current guidance listing checked; exact permit conditions and university agreement remain outstanding. No corresponding universal regional quota or different Ixelles threshold has been verified.','The statement that Brussels is Belgium\'s most expensive student city was not independently established; it is not used as a calculation input.']),('Three evaluated model iterations',[json.dumps(r) for r in [a,b,c]]+['Iteration 1 is preliminary and not compliant with the adopted 20% design requirement. Iteration 2 reserves 20%. Iteration 3 changes only acquisition price to improve downside risk, preserving housing cost ceilings.','The 90% probability target and triangular shock distributions are investor assumptions. Reusing the same seeded draws makes model changes comparable, not empirically certain.']),('Decision and limitations',[f'Conditional negotiated acquisition ceiling: EUR {newprice:,.0f}. Final NPV EUR {c["NPV"]:,.0f}; simulated loss probability {c["loss_probability"]:.1%}.', 'Do not commit without a vendor agreeing to the price, verified conversion costs, planning permission, university/social allocation terms and loan underwriting.','The equity amount in the iteration ledger is a prospective funding requirement. No real capital was invested or reinvested.','Inputs are classified as constants, controls and uncertain inputs, with approximate versus inherent uncertainty. Sensitivity, risk distributions and iteration charts support decisions.','Pre-tax financial model; unknown exact transaction VAT/tax treatment, static rent-occupancy relation, and unverified physical capacity limit confidence.']),('References',[str(r[2].value) for r in w['Legal_and_sources'].iter_rows(min_row=2,max_row=8)]+['Current municipal guidance listing: https://www.bruxelles.be/plans-reglements-urbanisme'])]
pdf('Executive_Report.pdf',report);log('REPORT: final Executive_Report.pdf explains all three iterations, source status, uncertainty classification, capital decision ledger, risk-adjusted negotiation ceiling and unresolved permit conditions.')
audit=json.load(open('execution_summary.json'));audit.update(end_epoch=time.time(),elapsed_seconds=time.time()-start,iteration_count=3);assert audit['elapsed_seconds']<420;Path('execution_summary.json').write_text(json.dumps(audit,indent=2));log('FINAL COMPLETE %.2fs'%audit['elapsed_seconds']);print(audit)
