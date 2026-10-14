# Brussels student housing model

Open model.xlsx in Excel. Blue Inputs cells and scenario drivers are editable. Results contains the decision metrics; Prompt contains the reusable research/analysis prompt. Financial assumptions are illustrative, not validated quotes.

Reproduce with Python 3 and openpyxl:

    python3 build_model.py
    python3 finalize_model.py

Both scripts live here. Saves validate a temporary XLSX and atomically replace model.xlsx. build_model.py checks SAVE_NOW.txt between improvements and stops if present. finalize_model.py performs the final numerical audit and caches current formula results without altering Excel formulas. Excel recalculates after input changes; rerun caching if an external preview needs updated values.

validation.json records independent checks of the initial configuration (NPV identity, zero-NPV rent, funding gap, IRR residual). These checks establish arithmetic consistency, not the accuracy of commercial or engineering assumptions. Hidden S1/S2/S3 sheets contain auditable scenario inputs, results and cash flows.

Model convention: nominal EUR, unlevered pre-tax, all capital at time zero, construction duration conservatively rounded up to whole years, 60-year cash-flow horizon, no terminal value by default. Maintenance is an annual cash expense and major refurbishment is separate. Simple and discounted payback exclude terminal sale. Initial affordability does not guarantee affordability after annual rent escalation.
