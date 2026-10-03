---
name: pdf-to-excel
description: Convert PDF to Excel. Use when the user says "convert pdf to excel", "pdf to excel", "pdf to xlsx", "pdf to csv", "extract tables from this pdf", "put this bank statement in excel", "turn this invoice pdf into a spreadsheet", or "copy the table from this pdf into excel". Extracts every table from one or more PDFs (bank statements, invoices, price lists, reports) into a clean .xlsx with real numbers and dates, merged multi-page tables, bold headers, filters and frozen header row.
---

# PDF to Excel

Turn tables inside PDFs into a spreadsheet the user can sort, filter and sum right away.

## Steps

1. **Find the input.** Use the PDF path(s) the user gave. If they attached a file, use its path. If none was given, ask which PDF.
2. **Check dependencies** once: `python -c "import pdfplumber, openpyxl"`. If it fails, run `pip install pdfplumber openpyxl`.
3. **Run the converter** (from this skill's folder):
   ```bash
   python scripts/pdf_to_excel.py "input.pdf" -o "input.xlsx"
   ```
   Options:
   - several PDFs at once: list them all; each gets its own sheet(s)
   - `--pages 2-4` only some pages
   - `--decimal comma` for European numbers like `1.234,50` (Norwegian, German, French statements); `--decimal dot` for `1,234.50`. Default `auto` guesses per cell.
   - `--no-merge` keep each table on its own sheet (default merges tables that continue across pages)
   - `--csv` also write CSV files
4. **Read the JSON summary** the script prints and check it:
   - `header` looks like real column names, `rows` count is plausible for the document.
   - If the header looks like data or columns are shifted, re-run with `--no-merge` or a page range and compare.
   - If there is a warning about a **scanned PDF**, say so plainly: the file is an image and needs OCR. Offer to read the pages visually and type the table into Excel instead (fine for a page or two).
   - If **no tables** were detected, read the PDF text and build the sheet yourself with openpyxl, using the same formatting (bold header, frozen first row, filter, numbers as numbers).
5. **Spot-check** 2–3 values: open the workbook with openpyxl and compare a few cells (first row, last row, a total) against the PDF text. Fix the decimal option if amounts are off by a factor of 100 or 1000.
6. **Report back** briefly: output path, sheets, row counts, and anything the user should check (e.g. a column that stayed text). For bank statements, offer a totals row or a pivot by category.

## Rules

- Never invent numbers. If a cell cannot be read, leave it empty and mention it.
- Keep the original PDF unchanged; write the .xlsx next to it unless the user named another location.
- Account numbers and IDs with leading zeros stay text so the zeros are kept.
