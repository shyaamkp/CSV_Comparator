# CSV/XLSX Comparator

Standalone PyQt5 desktop app to compare two CSV / XLS / XLSX files.

## Features
- Select two CSV or single-sheet Excel files
- Compare **by row position** or **by key columns** (composite keys supported)
- Ignore selected columns
- Sort results by any column
- Color-coded comparison view (Match / Different / Only in File 1 / Only in File 2)
- Differing cells highlighted with bold orange
- Preview tab (first 100 rows of each file)
- Summary tab (file stats, comparison stats)
- Export full report to a multi-sheet Excel workbook (colors preserved)

## Install (any OS)
```bash
pip install -r requirements.txt
python csv_comparator.py
```

## Build a Windows `.exe`
On a Windows machine with Python installed:
```bash
pip install -r requirements.txt
pip install pyinstaller
pyinstaller --onefile --windowed --name "CSV_Comparator" csv_comparator.py
```
The single-file executable will be at `dist\CSV_Comparator.exe`.

## Notes
- Cells are read as strings to preserve formatting (leading zeros, etc.).
- For `.xls` files, `xlrd==2.0.1` is pinned because newer versions dropped xls support.
- `.xlsx` uses `openpyxl`.
