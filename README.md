# PDF to Excel for Claude

Convert PDF to Excel inside Claude. Bank statements, invoices, price lists and reports become a clean `.xlsx`: real numbers and dates (not text), tables that span pages merged into one sheet, bold headers, filters and a frozen header row.

Just ask:

> convert this pdf to excel
> put my bank statement in a spreadsheet
> extract the tables from report.pdf to xlsx

## Install

**Claude Code**
```
/plugin marketplace add Cavada76/pdf-to-excel
/plugin install pdf-to-excel@pdf-to-excel
```

**Claude.ai / Claude desktop:** download `pdf-to-excel-skill.zip` from Releases, then Settings → Capabilities → Skills → Upload.

## What it handles

- Multi-page tables (repeated headers removed)
- European and US number formats: `1.234,50`, `1,234.50`, `1 234,50 kr`, `(99.10)`, `250,00-`
- Dates as real Excel dates; account numbers keep their leading zeros
- Several PDFs at once, page ranges, optional CSV output
- Scanned PDFs are detected and reported (they need OCR)

Requires Python with `pdfplumber` and `openpyxl` (Claude installs them if missing).

## Privacy

Runs locally, makes no network calls, and writes only to your folder. See the [Privacy Policy](PRIVACY.md).

MIT licensed.
