# Excel Filter & Export Tool

A fast Streamlit app to filter Excel files by zone → export one sheet per ward + summary pivot.

## Features
- Upload `.xlsx`, `.xls`, or `.csv`
- Auto-detects zone / ward / amount columns
- Multi-select zone filter
- Per-column visibility toggle (show/hide columns in export)
- Strips trailing `-` from E-Holding column (checkbox toggle)
- Exports one sheet per ward (named Zone_Ward)
- Summary sheet with pivot table: zone totals, row counts, amount sums, grand total row

## Setup

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Usage
1. Upload your file
2. Set zone/ward/amount columns in sidebar (auto-detected)
3. Select zones to include
4. Toggle column visibility
5. Enable dash-removal if needed
6. Click **Generate Excel Export** → Download

## Output Structure
```
filtered_export.xlsx
├── 📊 Summary        ← first sheet: detail table + pivot
├── Zone1_Ward1
├── Zone1_Ward2
├── Zone2_Ward1
└── ...
```
