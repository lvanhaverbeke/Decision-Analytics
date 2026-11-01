from pathlib import Path
p=Path('build_model.py');t=p.read_text()
t=t.replace("i=sh('Inputs'", "inputs += [('Selected_Location',1,'1 Ixelles,2 Etterbeek,3 Auderghem,4 City','Control','Assumption','Select hypothetical commune land budget'),('Building_Value_Ex_Land',2700000,'EUR','Approximate uncertainty','Assumption','Illustrative building value; not a valuation')]\ni=sh('Inputs'")
t=t.replace("w.defined_names.add(DefinedName('Total_Rooms'", "i.cell(ref['Acquisition_Price'],2,'=Building_Value_Ex_Land+INDEX(Locations!E2:E5,Selected_Location)')\ni.cell(ref['Acquisition_Price'],4,'Calculated from controls');i.cell(ref['Acquisition_Price'],5,'Assumption')\nw.defined_names.add(DefinedName('Total_Rooms'")
t=t.replace('Cash_flow!{col}12/{col}6,0)', 'Cash_flow!{col}12/{col}6,1E99)')
t=t.replace("for name in ['Planning", "legal.append(['S8 Current City guidance page','City still lists 17 June 2021 student housing recommendations; guidance is not itself a regulation.','https://www.bruxelles.be/plans-reglements-urbanisme','Live page 29 Oct 2026','2026-10-29','Verified publication; selected permit applicability remains to confirm.'])\nfor name in ['Planning")
needle="check=sh('Model_checks'"
block="""lp=sh('Location_finance',['Municipality','Assumed land budget EUR','Assumed building value EUR','Acquisition EUR','Project NPV EUR','Interpretation'])
for rr in range(2,6):
 er=case('Location land comparison',acq=f'=Building_Value_Ex_Land+Locations!E{rr}')
 lp.append([f'=Locations!A{rr}',f'=Locations!E{rr}','=Building_Value_Ex_Land',f'=B{rr}+C{rr}',f'=Analysis_engine!L{er}','All parcels, land budgets and constant building value are UNVERIFIED. Not a market ranking; same room plan and operating economics held fixed.'])
lp['A8']='No verified current comparable buildable land sales were found within the time budget. Statbel / IBSA dwelling medians are not land EUR/m2 and were not substituted.'
ch=BarChart();ch.title='Illustrative land-cost effect on NPV (EUR)';ch.add_data(Reference(lp,min_col=5,min_row=1,max_row=5),titles_from_data=True);ch.set_categories(Reference(lp,min_col=1,min_row=2,max_row=5));lp.add_chart(ch,'H2')
"""
t=t.replace(needle,block+needle)
p.write_text(t)
