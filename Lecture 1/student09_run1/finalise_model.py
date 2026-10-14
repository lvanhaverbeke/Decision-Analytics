from pathlib import Path
p=Path('build_model.py');s=p.read_text()
s=s.replace("'Payback flag'])", "'Payback flag','Cumulative equity CF','Equity payback flag'])")
s=s.replace("c.cell(r,10,f'=IF(I{r}>=0,1,0)')", "c.cell(r,10,f'=IF(I{r}>=0,1,0)')\n c.cell(r,11,f'=H{r}' if year==0 else f'=K{r-1}+H{r}')\n c.cell(r,12,f'=IF(K{r}>=0,1,0)')")
s=s.replace("B19>=Inputs!B24),", "B19>=Inputs!B24,Inputs!B2<=Inputs!B29),")
s=s.replace("i['A31']='Legal", "i['A32']='Dynamic equity payback';i['B32']='=IFERROR(MATCH(1,\'Annual cash flow\'!L2:L42,0)-1,\"Beyond model horizon\")'\ni['A31']='Legal")
# Use double quotes to preserve the quoted worksheet name in the builder source.
s=s.replace("i['B32']='=IFERROR(MATCH(1,'Annual cash flow'!L2:L42,0)-1,\"Beyond model horizon\")'", "i['B32']=\"=IFERROR(MATCH(1,'Annual cash flow'!L2:L42,0)-1,\\\"Beyond model horizon\\\")\"")
s=s.replace("('Operating horizon',30,'years','No sale proceeds modelled')", "('Operating horizon',30,'years','No sale proceeds modelled; supported range 1–38 years')")
s=s.replace("wb.calculation=", "from openpyxl.worksheet.datavalidation import DataValidation\ndv=DataValidation(type='whole',operator='between',formula1=1,formula2=38);dv.errorTitle='Horizon range';dv.error='Use 1 to 38 operating years';dv.showErrorMessage=True;a.add_data_validation(dv);dv.add(a['B27'])\nwb.calculation=")
p.write_text(s)
