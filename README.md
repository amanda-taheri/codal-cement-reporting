# CODAL Cement Reporting & Financial Analysis

A Python-based collection of scripts and Excel workbooks for working with publicly available disclosures from the Iranian CODAL reporting system. The original project structure, input workbooks, sample reports, downloaded caches, code, and historical output files are intentionally preserved in the provided archive.

## What the project does

- **Monthly company reporting:** searches and processes monthly cement-sector disclosures and builds a consolidated Excel report.
- **Materials and energy comparison:** extracts and compares material and energy line items across companies.
- **Semiannual and annual financial statements:** selects audited 6-month and 12-month reports, handles corrections and restated comparisons, and generates Excel output plus audit records.
- **Earlier/experimental scripts:** includes additional report builders and tests that were part of the original project.

This repository is a working-code archive, not a packaged Python library. Files are kept in their original locations to avoid changing script-relative paths.

## Repository layout

```text
Codal/
├── codal/
│   ├── codal_monthly_all_companies.py   # Monthly reporting
│   ├── codal_material_energy_compare.py # Materials / energy comparison
│   ├── codal_report_builder_5.py        # Additional report builder
│   ├── codal_monthly_builder_test.py    # Development version
│   ├── codal_test_5.py                  # Testing / exploratory script
│   ├── companies.xlsx                   # Monthly company list
│   ├── companies_2.xlsx                 # Additional company list
│   ├── template_v1.xlsx                 # Excel template
│   ├── Output/                          # Existing generated workbooks
│   └── Codal_*Cache*/                   # Download / search cache (not for Git)
├── codal_yearly/
│   ├── codal_yearly.py                  # Audited 6-/12-month statements
│   ├── companies.xlsx                   # Yearly company list
│   ├── requirements.txt
│   ├── README.md                         # Original module documentation
│   ├── config/line_item_aliases.json
│   ├── output/                           # Existing financial workbook
│   └── audit/                            # Existing audit CSVs
├── Reports/                              # Original company Excel reports
├── Codal_Cement_40_Vocabulary_*.xlsx     # Vocabulary/standardization workbooks
├── template_v1.xlsx                     # Original top-level template
├── README.md
├── LICENSE
├── requirements.txt
├── .gitignore
└── .gitattributes
```

## Requirements

- Python 3.10+ recommended (the yearly script uses modern type annotations).
- Python packages: `pandas`, `requests`, `lxml`, `openpyxl`, `xlrd`.
- Network access for live CODAL requests. CODAL availability and response formats may change.

### Install

From the `Codal` directory:

```bash
python3 -m venv .venv
source .venv/bin/activate   # macOS / Linux
python -m pip install -r requirements.txt
```

On Windows, activate the environment using `.venv\Scripts\activate` instead.

### Run the main scripts

```bash
python codal/codal_monthly_all_companies.py
python codal/codal_material_energy_compare.py
python codal_yearly/codal_yearly.py
```

The scripts read `companies.xlsx` from their own folder. Those original workbooks are included without modification; the expected columns include `نماد` and `نام شرکت`. Review the symbol list and period settings before running. The scripts make live external requests and may overwrite generated output/cache files.

For yearly settings and audit details, see [`codal_yearly/README.md`](codal_yearly/README.md). **Documentation caveat:** that existing README describes a two-sheet workbook, but the currently included yearly script reports a four-sheet output (two comparison and two detail sheets). Historical output files are preserved as-is and might reflect an earlier code version. This package does not edit either the code or the historical workbooks.

## Data and reproducibility

- Source: publicly accessible CODAL disclosures; source data ownership and reuse terms remain with their respective rights holders.
- `Reports/`, the root spreadsheets, `codal/Output/`, `codal_yearly/output/`, and `codal_yearly/audit/` contain the original workbooks and audit results. They are intentionally **not** excluded by `.gitignore`.
- Python environments, interpreter caches, macOS metadata, download/search caches, and temporary Excel lock files are excluded by `.gitignore` to keep commits manageable. They **remain inside the original-preserved ZIP**.
- The raw archived contents are not an assertion that all scripts were freshly executed or validated in the present environment.

## Publishing to GitHub

**Use Git (CLI or GitHub Desktop), not a bulk drag-and-drop upload of the extracted archive.** Git honors `.gitignore`; GitHub's web upload dialog may still let you select ignored files. Do not upload `__MACOSX/` from the ZIP.

1. Extract this archive and enter the **`Codal`** directory (not the parent directory that also contains `__MACOSX`).
2. Review `LICENSE` and confirm that code developed for an organization may be published.
3. Inspect the workbook contents, source and configuration for secrets, private data, or third-party redistribution restrictions. Public CODAL data does not automatically grant publication rights to an employer's code.
4. In your terminal run:

```bash
git init
git add .
git status --short
# Check the staged files carefully before continuing.
git commit -m "Add CODAL reporting project and documentation"
```

5. Create a GitHub repository (recommended: **Private** while permissions are being reviewed) and follow GitHub's instructions for linking it with a remote and pushing.

## License status

See [`LICENSE`](LICENSE). This archive does **not** apply an MIT or other open-source license without confirming the rights to publish and relicense the work.
