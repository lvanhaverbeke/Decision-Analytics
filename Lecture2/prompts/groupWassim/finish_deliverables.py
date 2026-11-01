import json,math,time,zipfile,shutil,os,statistics
from pathlib import Path
from xml.etree import ElementTree as ET
from openpyxl import load_workbook
from reporting import docx,pdf,pptx,result_pages
w=load_workbook('model.xlsx',data_only=True)
p={r[0].value:r[1].value for r in w['Inputs'].iter_rows(min_row=2)}
strategies=[(0,2400,.95,.1),(.06,1800,1,.1),(.12,1100,1.05,.2)]
def calc(n,rent,strategy=2,cost=1,demand=1,opex=1,cap=.055,growth=.02):
 premium,energy,appeal,affshare=strategies[strategy-1];aff=math.ceil(n*affshare-1e-8)
 investment=1500000*1.15+n*28*2200*cost*(1+premium)*1.15*1.1*1.21+n*1800+50000
 occupied=min(n-aff,125*demand*appeal*math.exp(-1.2*(rent/590-1)))
 rev=12*(occupied*rent+aff*min(1,demand)*449.3)
 cf=rev*.94-(60000+n*(700+energy*.25))*opex-n*150
 cfs=[-investment]+[cf*(1+growth)**t for t in range(20)];cfs[-1]+=cf*(1+growth)**20/cap*.98
 npv=sum(c/(1.08)**t for t,c in enumerate(cfs));feasible=investment<=14000000 and n*28<=4000
 return npv,investment,occupied,feasible,cfs
base=calc(120,590);assert abs(base[0]-w['KPI_Dashboard']['B2'].value)<.001
best=max((calc(n,r,s)[0],n,r,s) for s in [1,2,3] for n in range(40,141) for r in range(400,651,10) if calc(n,r,s)[3]);assert abs(best[0]-w['Optimization']['I2'].value)<.001
for row in list(w['Monte_Carlo'].iter_rows(min_row=2,max_row=5001))[:60]:
 vals=[c.value for c in row[:7]];assert abs(calc(120,590,cost=vals[1],demand=vals[2],opex=vals[3],cap=vals[4],growth=vals[6])[0]-vals[5])<.001
assert calc(120,730)[2]<base[2];assert calc(120,590,cost=1.1)[0]<base[0]
for r in range(2,7):assert w['Model_Checks'].cell(r,2).value==True
assert w['Model_Checks']['B7'].value==5000 and w['Model_Checks']['B8'].value==5000
metrics={'base_NPV':base[0],'best_constructed':[best[0],best[1],best[2],best[3]],'negative_NPV_probability':w['Monte_Carlo']['I8'].value,'joint_success_probability':w['Monte_Carlo']['I15'].value,'formula_tests':'Zero supported-function formula errors','independent_tests':'Base annual cash flow; all finite grid candidates; first 60 Monte Carlo draws; rent demand and cost monotonicity; debt/affordability/histogram reconciled','limits':['Tax and delays excluded','No native Excel/Solver/Data Tables test','Finite EUR10 rent grid','Scenario weights and shock distributions subjective','Charts are in workbook; reports/slides text-based','Finance DSCR and sustainability treated as approval gates; finite project-NPV search does not enforce these gates']}
Path('intermediate/independent_validation.json').write_text(json.dumps(metrics,indent=2))
# Actual IRR distribution retained as static, reproducible appendix; rerun this script after model changes.
irrs=[];rows=[]
for row in w['Monte_Carlo'].iter_rows(min_row=2,max_row=5001):
 vals=[c.value for c in row[:7]];cf=calc(120,590,cost=vals[1],demand=vals[2],opex=vals[3],cap=vals[4],growth=vals[6])[4]
 def pv(rate):return sum(c/(1+rate)**t for t,c in enumerate(cf))
 low=-.5;high=1
 if pv(low)*pv(high)>0:irr=None
 else:
  for _ in range(42):
   mid=(low+high)/2
   if pv(mid)>0:low=mid
   else:high=mid
  irr=(low+high)/2;irrs.append(irr)
 rows.append([int(vals[0]),irr])
import csv
with open('intermediate/Monte_Carlo_IRR.csv','w',newline='') as out:writer=csv.writer(out);writer.writerow(['Trial','Project IRR - static appendix; rerun finish_deliverables.py after changing inputs']);writer.writerows(rows)
metrics['IRR_mean']=statistics.mean(irrs);metrics['IRR_undefined_trials']=5000-len(irrs);metrics['IRR_meeting_8pct']=sum(x>=.08 for x in irrs)/5000
Path('intermediate/independent_validation.json').write_text(json.dumps(metrics,indent=2))
pages=result_pages(w)
pages[0][1].append(f'Best built strategy index {best[3]} (1=A,2=B,3=C); gates for DSCR and sustainability still require approval. No-build avoids negative investment NPV.')
pages[4][1][1]=pages[4][1][1].replace('1000','5000')
pages[4][1].append(f'Fifth simultaneous uncertainty: common annual rent/opex growth, triangular 0.5%/2%/3.5%. Mean project IRR {metrics["IRR_mean"]:.2%}; undefined IRRs {metrics["IRR_undefined_trials"]}; fraction meeting 8% hurdle {metrics["IRR_meeting_8pct"]:.1%}. Full static trial IRRs in CSV appendix; rerun after changes.')
pages[3][1][0]=pages[3][1][0].replace('EUR 5','EUR 10')
pages[3][1][-1]=pages[3][1][-1].replace('EUR 5','EUR 10')
pages[5][1].append('A/B/C modelled energy = 2400/1800/1100 kWh per room; construction premiums 0%/6%/12%; affordable shares 10%/10%/20%. Illustrative engineering alternatives, not certifications.')
pages[6][1].append('Full requested rubric is not complete: financing choice, delays, native tables/Solver and empirical calibration remain. Source benchmarks are verified; group review and exact tax/legal conclusions are outstanding.')
pdf('Student_Housing_Brussels_Executive_Report.pdf',pages)
slides=[pages[0],('Five managerial questions',['1. Which rent strategy maximizes NPV?','2. Which development concept should be selected?','3. How robust is the investment?','4. Which inputs drive risk?','5. Which alternative best handles conflicting objectives?']),pages[1],pages[2],('Base financial results',pages[0][1][1:3]+[f'Equity NPV EUR {w["KPI_Dashboard"]["B4"].value:,.0f}; year 1 DSCR {w["KPI_Dashboard"]["B11"].value:.2f}.']),('Optimization and investment gate',pages[3][1]),('Sensitivity findings',[pages[4][1][0],'See both dynamic two-way heatmaps in workbook. Native Excel Data Table instructions provided.', 'Negative NPV outcomes retained; no-build explicitly zero.']),('Monte Carlo results',pages[4][1][1:]),('Scenario decision criteria',['Four coherent scenarios; illustrative equal weights, not market event probabilities.']+[f'{w["Decision_Criteria"].cell(r,1).value}: {w["Decision_Criteria"].cell(r,2).value}' for r in range(9,14)]),('Affordability and sustainability',pages[5][1][:2]+[pages[5][1][-1]]),('Recommendation and due diligence',pages[5][1][2:4]),('Limits, AI and oral defense',['Native Solver and Data Tables require manual setup. Distributions and elasticity are assumptions. No desktop Office verification.','AI created analysis; student checks and team contributions remain to be completed. Roles: evidence/model; sensitivity/optimization; risk/report.','Defense: explain 8% nominal hurdle, demand response, subjective triangles, independence risks and no-build comparison. All three students should review every result.'])]
pptx('Student_Housing_Brussels_Presentation.pptx',slides)
shutil.copy2('model.xlsx','Student_Housing_Brussels_Model.xlsx');shutil.copy2('model.xlsx','intermediate/model_validated.xlsx')
Path('intermediate/report_content.json').write_text(json.dumps(pages,indent=2));Path('intermediate/presentation_content.json').write_text(json.dumps(slides,indent=2))
# Structural validation, with no claim of Office application execution.
for file in ['Student_Housing_Brussels_Model.xlsx','Student_Housing_Brussels_Presentation.pptx']:
 with zipfile.ZipFile(file) as z:
  assert z.testzip() is None
  for name in z.namelist():
   if name.endswith('.xml') or name.endswith('.rels'):ET.fromstring(z.read(name))
assert Path('Student_Housing_Brussels_Executive_Report.pdf').read_bytes().startswith(b'%PDF-1.4')
with zipfile.ZipFile('Student_Housing_Brussels_Submission.zip','w',zipfile.ZIP_DEFLATED) as z:
 for name in ['Student_Housing_Brussels_Model.xlsx','Student_Housing_Brussels_Executive_Report.pdf','Student_Housing_Brussels_Presentation.pptx','intermediate/independent_validation.json','intermediate/Monte_Carlo_IRR.csv']:z.write(name,name)
a=json.load(open('execution_started.json'));a.update(end_epoch=time.time(),elapsed_seconds=time.time()-a['start_epoch'],completed=True,limitations=metrics['limits'],outputs=['Student_Housing_Brussels_Model.xlsx','Student_Housing_Brussels_Executive_Report.pdf','Student_Housing_Brussels_Presentation.pptx','Student_Housing_Brussels_Submission.zip']);Path('execution_summary.json').write_text(json.dumps(a,indent=2));print(json.dumps(a));print(json.dumps(metrics))
