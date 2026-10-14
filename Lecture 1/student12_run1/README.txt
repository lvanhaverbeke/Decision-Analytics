Student Housing Investment — VUB Decision Analytics

Open model.xlsx in Excel. Blue cells in Inputs are editable; Excel recalculates green formulas.
Financial inputs and travel times are illustrative. Official sources and unresolved project requirements are in Sources and Read me.

Rebuild in order (Python 3 with openpyxl):
python3 build_model.py
python3 improve_model.py
python3 finalize_model.py

The workbook is saved via a temporary .xlsx and atomic os.replace. Numeric caches allow immediate baseline viewing. The Results tab is last.
LLM_prompt.txt contains the reusable prompt also included in the workbook.
