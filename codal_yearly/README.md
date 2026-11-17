# CODAL Yearly / Semiannual Financial Statements

This project collects audited 6-month and 12-month financial statements for all symbols in `companies.xlsx`.

Current output contains exactly two sheets:

1. `صورت سود و زیان`
2. `صورت وضعیت مالی`

The future `خلاصه اطلاعات گزارش تفسیری - صفحه 1` extraction is intentionally not enabled yet.

## Folder structure

```text
codal_yearly/
├── companies.xlsx
├── codal_yearly.py
├── requirements.txt
├── README.md
├── config/
│   └── line_item_aliases.json
├── cache/
│   ├── search/
│   └── reports/
├── audit/
└── output/
```

## Important behavior

- Only `LetterType=6` financial statements are searched.
- Only audited reports are kept.
- Only 6-month and 12-month reports are kept.
- Corrected disclosures (`اصلاحیه`) supersede the older version of the same period.
- A later report column explicitly marked `تجدید ارائه شده` / `اصلاح شده` is applied back to that historical period.
- If an orphan correction cannot expose its target period in the title, it is attached to the latest earlier matching 6/12-month report.
- If both consolidated and standalone filings are found for one company/period, the conflict is written to `audit/report_scope_conflicts.csv`. The default preference is consolidated; change `SCOPE_PREFERENCE` in the script if needed.
- Final workbook has one row per company / period and a common set of line-item columns. Missing line items remain blank.
- Source monetary units are normalized to **million rials** when a unit marker can be detected.

## First setup on macOS

From Terminal:

```bash
cd /path/to/Codal/codal_yearly
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python3 codal_yearly.py
```

## Main settings

At the top of `codal_yearly.py`:

- `MIN_JALALI_YEAR = None` and `MAX_JALALI_YEAR = None` -> all available audited 6/12 month reports.
- Set them, for example, to `1402` and `1404` to limit output.
- `REFRESH_SEARCH_CACHE = False` and `REFRESH_REPORT_CACHE = False` make reruns faster and safer against CODAL rate limits.
- `SCOPE_PREFERENCE = "consolidated"` controls what happens if both consolidated and standalone reports exist for the same period.

## Audit files

After a run, check `audit/`:

- `report_selection.csv`: reports selected / skipped and why
- `report_scope_conflicts.csv`: periods where both consolidated and standalone candidates exist
- `processing_errors.csv`: report download / parse errors
- `line_items_income.csv`: discovered income-statement row names
- `line_items_financial_position.csv`: discovered financial-position row names
- `restatements_applied.csv`: historical periods overwritten by explicit restated/corrected comparative columns

The two `line_items_*.csv` files are specifically for the standardization review. If CODAL uses a genuinely different label for the same concept, add a mapping to `config/line_item_aliases.json` and rerun.
