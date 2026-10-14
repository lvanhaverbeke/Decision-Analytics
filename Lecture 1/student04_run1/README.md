Student housing decision model

Deliverables: model.xlsx and executive_report.md.

Rebuild with Python 3 and openpyxl:
1. python3 enhance_model.py
2. python3 finalise_model.py

build_model.py saves a basic workbook first; enhance_model.py adds the research and analysis. finalise_model.py independently evaluates formulas, writes the report and adds cached spreadsheet results. Every workbook save writes model.tmp.xlsx, validates it and atomically replaces model.xlsx. SAVE_NOW.txt stops further research/enhancement.

Blue cells are assumptions. Sources carry dates and distinguish official evidence from estimates. Excel recalculates edited formulas on opening. Investor taxation and debt are excluded.
