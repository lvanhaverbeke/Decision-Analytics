from pathlib import Path
import os
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.chart import LineChart, Reference
ROOT=Path(__file__).resolve().parent

def save_atomic(wb):
    tmp=ROOT/'model.tmp.xlsx'
    wb.save(tmp)
    os.replace(tmp, ROOT/'model.xlsx')

def baseline():
    w=Workbook(); r=w.active; r.title='Read Me'
    for row in [['BRUSSELS STUDENT HOUSING | INVESTMENT SCREEN'],['Usable initial version; illustrative assumptions, not quotations.'],['100 rooms; 10% affordable rooms; 20-year investment horizon.'],['Inputs are blue. Cash flows are annual, nominal EUR.'],['See Inputs, Cash Flow and Summary. Detailed research and simulations follow.']]: r.append(row)
    a=w.create_sheet('Inputs')
    data=[['Assumption','Value','Unit'],['Rooms',100,'rooms'],['Affordable share',.10,'%'],['Standard monthly rent',650,'EUR/month'],['Affordable monthly rent',450,'EUR/month'],['Occupancy',.96,'%'],['Rent growth',.02,'%/year'],['Operating costs',.30,'% revenue'],['Total development cost',12500000,'EUR'],['Discount rate',.07,'%'],['Exit cap rate',.055,'%'],['Selling costs',.02,'%']]
    for row in data:a.append(row)
    c=w.create_sheet('Cash Flow');c.append(['Year','Gross rent','Effective rent','Operating costs','NOI','Exit proceeds','Project cash flow'])
    c.append([0,0,0,0,0,0,'=-Inputs!B9'])
    for y in range(1,21):
        n=y+2;c.append([y,f'=Inputs!B2*((1-Inputs!B3)*Inputs!B4+Inputs!B3*Inputs!B5)*12*(1+Inputs!B7)^(A{n}-1)',f'=B{n}*Inputs!B6',f'=C{n}*Inputs!B8',f'=C{n}-D{n}',f'=IF(A{n}=20,E{n}*(1+Inputs!B7)/Inputs!B11*(1-Inputs!B12),0)',f'=E{n}+F{n}'])
    s=w.create_sheet('Summary');s.append(['Metric','Value']);s.append(['Project NPV','=NPV(Inputs!B10,\'Cash Flow\'!G3:G22)+\'Cash Flow\'!G2']);s.append(['Project IRR','=IRR(\'Cash Flow\'!G2:G22)']);s.append(['Year 1 NOI',"='Cash Flow'!E3"]);s.append(['Affordable rooms','=Inputs!B2*Inputs!B3'])
    for sh in w:
        sh.freeze_panes='B2';sh.column_dimensions['A'].width=45
        for col in ['B','C','D','E','F','G']:sh.column_dimensions[col].width=21
        for cell in sh[1]:cell.fill=PatternFill('solid',fgColor='123D45');cell.font=Font(color='FFFFFF',bold=True)
    for cell in a['B'][1:]:cell.font=Font(color='1565C0')
    save_atomic(w)
    print('Saved usable baseline model.xlsx; SAVE_NOW=',(ROOT/'SAVE_NOW.txt').exists())
if __name__=='__main__':baseline()
