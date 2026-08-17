# UseHirable USA Sponsor Preprocessor

This repository builds a compact U.S. visa sponsor JSON file for UseHirable.

The processor downloads the latest cumulative U.S. Department of Labor OFLC LCA disclosure workbook, reads the large XLSX file outside Google Apps Script, aggregates case-level filings into employer/location sponsor records, and writes `dist/usa-sponsors.json`.

## Data Source

Source page:

`https://www.dol.gov/agencies/eta/foreign-labor/performance`

The script looks for the latest cumulative workbook matching:

`LCA_Disclosure_Data_FY{year}_Q{quarter}.xlsx`

Appendix and worksite files are ignored. The selected workbook is processed with `openpyxl` in read-only mode so large files can be handled more safely than in Apps Script.

## Local Setup

```bash
python -m venv .venv
.venv\Scripts\activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

On macOS/Linux, activate with:

```bash
source .venv/bin/activate
```

## Manual Run

Download the latest DOL workbook and build the JSON:

```bash
python src/build_usa_sponsors.py
```

Use an existing local workbook:

```bash
python src/build_usa_sponsors.py --workbook downloads/LCA_Disclosure_Data_FY2026_Q3.xlsx --source-url https://example.com/source.xlsx
```

The output is written to:

```text
dist/usa-sponsors.json
```

## Output Structure

The JSON payload uses this high-level shape:

```json
{
  "version": "YYYY-MM-DD",
  "generatedAt": "ISO timestamp",
  "source": "U.S. DOL LCA Disclosure",
  "sourceUrl": "https://...",
  "fiscalYear": 2026,
  "quarter": 3,
  "totalCompanies": 0,
  "companies": [
    {
      "company_name": "",
      "country": "United States",
      "industry": "",
      "location": "City, ST",
      "tags": "",
      "sponsor_type": "LCA Employer",
      "visa_program": "H-1B / H-1B1 / E-3",
      "source": "U.S. DOL LCA Disclosure",
      "source_url": "",
      "filing_count": 0
    }
  ]
}
```

The processor also includes helper fields such as `company_slug`, `country_code`, `country_slug`, `location_slug`, and `top_job_titles` so downstream import scripts can avoid recomputing common values.

## Processing Notes

- Headers are inspected dynamically instead of hardcoding exact column names.
- Employer records are grouped by normalized employer name, city, and state.
- Company names are normalized conservatively by trimming and collapsing whitespace only.
- Relevant visa programs are `H-1B`, `H-1B1`, and `E-3`.
- Certified LCA cases are kept when a case status column is present.
- Output records are sorted predictably by company name and location.
- The script fails if no workbook is found, required employer columns are missing, or no company records are produced.

## GitHub Actions

The workflow at `.github/workflows/update-usa-sponsors.yml` can be run manually with `workflow_dispatch`.

It:

1. Checks out the repository.
2. Sets up Python 3.12.
3. Installs `requirements.txt`.
4. Runs `python src/build_usa_sponsors.py`.
5. Uploads `dist/usa-sponsors.json` as a workflow artifact.

Automatic commits and weekly cron scheduling are intentionally not enabled yet.

## Future Google Apps Script Integration

Google Apps Script should consume the compact `dist/usa-sponsors.json` output instead of reading the raw DOL workbook. That keeps heavy XLSX parsing, aggregation, and normalization in Python, while Apps Script can focus on importing or publishing already-prepared sponsor records into the wider visa sponsor directory pipeline.
