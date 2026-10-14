from pathlib import Path
import json
p=Path('research.json');d=json.loads(p.read_text())
d['inputs']['22']={'value':.16,'note':'S9: contextual proxy ONLY, not VUB-specific: 16% first-year bachelor dropout among Flemish 2023–2024 starters; published 2026-03-31, provisional. No direct link to housing vacancy.'}
for row in d['sources']:
 if row[0]=='S9':row[3]='Published 2026-03-31; cohort 2023–2024; accessed 2026-10-14';row[4]='16% first-year bachelor dropout (8,881 entrants); figure provisional. Flemish academic and professional bachelors, not VUB-specific. Obtain VUB cohort data before treating as a VUB estimate.'
d['sources'].append(['S12','Brussels Housing: comparison of minimum requirements','https://admin.be.brussels/sites/default/files/2025-06/Tableau%20comparatif%20AGRBC%202003-2023%20FR.pdf','2026 rules, right-hand column; pp.13–14; accessed 2026-10-14','Student housing surface: 12 m² for one person, 18 m² for two. Habitable rooms at least 70% of dwelling area. Collective common rooms allocated by maximum residents using them. These are housing-code minima; planning/fire rules can be stricter.'])
d['regulations'].append(['Student housing minimum surface','2026 housing-code rules: student dwelling minimum 12 m² for one person, 18 m² for two. Habitable rooms must cover at least 70% of dwelling area. Collective common rooms counted pro rata by maximum residents using them.','Do not equate 12 m² with a universal private-bedroom minimum. Architect must calculate qualifying area and apply stricter planning/fire rules; 28 m² gross allocation is not compliance proof.','S12, pp.13–14, right-hand column'])
p.write_text(json.dumps(d,indent=2,ensure_ascii=False))
p=Path('build_model.py');s=p.read_text();s=s.replace("('Academic dropout rate',.20", "('Academic dropout proxy',.20")
s=s.replace("('Academic dropout','=Inputs!B22',v[22],'Context assumption; continuous; ratio; no observed VUB rate established')", "('Academic dropout proxy','=Inputs!B22',v[22],'Context; continuous; ratio; Flemish proxy, not VUB rate; S9')")
s=s.replace("ov['A22']='Dropout versus departures';", "ov['A22']='Dropout versus departures';")
s=s.replace("Academic dropout is contextual only. Price-driven departures below are hypothetical", "16% dropout is a Flemish first-year cohort proxy, not VUB-specific. Price-driven departures below are hypothetical")
s=s.replace("ov['A23']='Payback limitation';ov['B23']='Simple PBP includes sale; interpolation is indicative because exit occurs at year end. Operating-only and discounted payback shown below.'", "ov['A23']='Payback convention';ov['B23']='Simple PBP includes sale if needed; the base case pays back before sale. Fractional years assume uniform within-year operating cash flow. Discounted payback uses the 7% required return.'")
marker='wb.calculation=CalcProperties'
s=s.replace(marker,'''# Final presentation and input controls.
from openpyxl.formatting.rule import ColorScaleRule
for sh in wb:
 last='T' if sh.title=='Cash flow' else 'F' if sh.title=='Rent sensitivity' else 'E'
 sh.merge_cells(f'A1:{last}1')
 sh.row_dimensions[1].height=31
 header=3 if sh.title in ['Overview','Inputs'] else 2
 for c in sh[header]:
  c.fill=PatternFill('solid',fgColor='DCE6F1');c.font=Font(bold=True,color='16324F')
 sh.sheet_properties.pageSetUpPr.fitToPage=True
 sh.page_setup.orientation='landscape';sh.page_setup.paperSize=sh.PAPERSIZE_A3
 sh.page_setup.fitToWidth=1;sh.page_setup.fitToHeight=0
 if sh.title in ['Sources','Regulations']:
  for rr in range(3,sh.max_row+1):sh.row_dimensions[rr].height=78
 if sh.title=='Overview':
  sh.merge_cells('A2:C2');sh.row_dimensions[2].height=40
  for rr in [18,20,21,22,23,32]:sh.row_dimensions[rr].height=65
wb['Student budget'].row_dimensions[19].height=70
for rr in range(9,13):wb['Student budget'].row_dimensions[rr].height=75
wb['Risk sensitivity'].conditional_formatting.add('B3:D7',ColorScaleRule(start_type='min',start_color='F4CCCC',mid_type='num',mid_value=0,mid_color='FFF2CC',end_type='max',end_color='D9EAD3'))
positive=DataValidation(type='decimal',operator='greaterThan',formula1='0');positive.showErrorMessage=True;positive.error='Enter a positive number.'
wb['Inputs'].add_data_validation(positive)
for rr in [4,5,6,7,10,14,15,16,21]:positive.add(f'B{rr}')
percent=DataValidation(type='decimal',operator='between',formula1='0',formula2='1');percent.showErrorMessage=True;percent.error='Enter a fraction from 0 to 1.'
wb['Inputs'].add_data_validation(percent)
for rr in [8,9,11,12,13,17,20,22,23,24]:percent.add(f'B{rr}')
wb['Inputs']['B31'].fill=PatternFill('solid',fgColor='D9EAD3');wb['Inputs']['B31'].font=Font(bold=True)
wb['Overview'].print_area='A1:C32';wb['Rent sensitivity'].print_area='A1:S22'
wb.calculation=CalcProperties''')
p.write_text(s)
