# Excel Customer Segregator (CLI + Streamlit)

Split a workbook into multiple sheets based on a customer code column while preserving your template/header formatting.

- Reads an input `.xlsx`
- Groups rows by the value in a chosen "Customer Code" column (e.g., column F)
- Writes a new workbook with one sheet per unique code
- Preserves header area formatting (top N rows), column widths, merges, row heights, cell styles, and freeze panes
- Use via command line or a Streamlit web UI (upload and download)

## Features
- Preserve formatting for the top header rows (default 8) exactly as in the source sheet
- Keep column widths, merged header cells, row heights, cell styles, and number formats
- Freeze panes right below the header
- Flexible column selector: Excel letter (F), 1-based index (6), or header title ("Customer Code")
- Safe Excel sheet names (invalid characters removed, 31-char limit handled)

## Requirements
- Python 3.9+
- Packages: `pandas`, `openpyxl`, `streamlit`

Install dependencies:
```bash
pip install -r requirements.txt
```

## Command Line Usage
Run from the project directory:
```bash
python3 segregate_by_customer_code.py --input test.xlsx --column F --header-rows 8
```

Options:
- `--input, -i` Path to the source `.xlsx` (default: `./test.xlsx`)
- `--sheet, -s` Sheet name or 0-based index to read (default: first sheet)
- `--column, -c` Customer code column. Accepts:
  - Letter: `F`
  - 1-based index: `6`
  - Header label: `"Customer Code"`
- `--header-rows` Number of top rows to copy verbatim (formatting + merges). Default: `8`
- `--output, -o` Output path (default: `<input>_segregated.xlsx`)

Examples:
```bash
# Use a header label instead of letter
python3 segregate_by_customer_code.py -i mydata.xlsx -c "Customer Code"

# Choose a specific sheet and 7 header rows
python3 segregate_by_customer_code.py -i mydata.xlsx -s "Data" -c F --header-rows 7

# Write to a custom output path
python3 segregate_by_customer_code.py -i mydata.xlsx -o output.xlsx
```

## Streamlit Web App
Launch the UI:
```bash
streamlit run streamlit_app.py
```
Then:
1. Upload your Excel file
2. Pick the worksheet, customer code column, and header-row count
3. Click "Segregate"
4. Download the resulting workbook

## What Formatting Is Preserved?
- Top N header rows (values, cell styles, merges, row heights)
- Column widths
- Cell formatting for data rows (font, fill, borders, alignment, protection, number formats)
- Freeze panes set just below the header (e.g., row 9 if header rows = 8)

Not copied (by default):
- AutoFilters, conditional formatting, data validation, page layout/margins/print titles.
If you need these, open an issue or request and they can be added.

## Project Structure
- `segregate_by_customer_code.py` — Core logic and CLI. Main function: `segregate()`
- `streamlit_app.py` — Web UI for upload/segregate/download
- `requirements.txt` — Python dependencies

## Tips & Troubleshooting
- Codes like `C0005` are preserved as text.
- Blank code cells are skipped.
- Excel sheet names are limited to 31 characters; duplicates are suffixed automatically.
- If the header block in your file is not 8 rows, pass `--header-rows <N>` or change it in the app.
- If you get an error about the code column, try specifying the header title exactly as it appears in your sheet, or use the column letter.

## License
This project is provided as-is for internal/automation use. Add your preferred license if needed.
