# UseHirable USA Sponsor Preprocessor

This repository builds a compact U.S. visa sponsor JSON file for UseHirable.

The processor downloads the latest cumulative U.S. Department of Labor OFLC LCA disclosure workbook, reads the large XLSX file outside Google Apps Script, aggregates case-level filings into employer/location sponsor records, and writes both a full JSON file and chunked JSON files for safe resumable imports.

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

The primary outputs are written to:

```text
dist/usa-sponsors.json
dist/manifest.json
dist/usa-001.json
dist/usa-002.json
...
```

## Output Structure

### Full Dataset

`dist/usa-sponsors.json` contains the complete generated USA sponsor dataset. It remains useful for debugging, inspection, archival checks, and comparing the full dataset against the chunked files.

The full JSON payload uses this high-level shape:

```json
{
  "version": "YYYY-MM-DD",
  "generatedAt": "ISO timestamp",
  "source": "U.S. DOL LCA Disclosure",
  "sourceUrl": "https://...",
  "fiscalYear": 2026,
  "quarter": 3,
  "totalCompanies": 0,
  "totalFilings": 0,
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

### Manifest

`dist/manifest.json` is the entry point for consumers such as Google Apps Script. It describes the dataset, source workbook, totals, and chunk files to import.

Example:

```json
{
  "version": "2026-08-17",
  "generatedAt": "2026-08-17T18:00:00Z",
  "country": "United States",
  "countryCode": "US",
  "source": "U.S. DOL LCA Disclosure",
  "sourceUrl": "https://www.dol.gov/media/LCA_Disclosure_Data_FY2026_Q3.xlsx",
  "fiscalYear": 2026,
  "quarter": 3,
  "totalCompanies": 60846,
  "totalFilings": 427628,
  "chunkSize": 5000,
  "chunkCount": 13,
  "chunks": [
    {
      "index": 1,
      "file": "usa-001.json",
      "recordCount": 5000
    },
    {
      "index": 2,
      "file": "usa-002.json",
      "recordCount": 5000
    }
  ]
}
```

### Chunks

`dist/usa-001.json`, `dist/usa-002.json`, and the remaining `usa-###.json` files contain 5,000 sponsor records per file, except the final chunk, which may contain fewer records.

Each chunk is a JSON object:

```json
{
  "version": "2026-08-17",
  "country": "United States",
  "countryCode": "US",
  "fiscalYear": 2026,
  "quarter": 3,
  "chunk": 1,
  "chunkCount": 13,
  "chunkSize": 5000,
  "recordCount": 5000,
  "companies": []
}
```

The chunked output is generated from the same ordered `companies` list as `usa-sponsors.json`. Chunking does not re-sort or re-deduplicate the dataset.

## Processing Notes

- Headers are inspected dynamically instead of hardcoding exact column names.
- Employer records are grouped by normalized employer name, city, and state.
- Company names are normalized conservatively by trimming and collapsing whitespace only.
- Relevant visa programs are `H-1B`, `H-1B1`, and `E-3`.
- Certified LCA cases are kept when a case status column is present.
- Output records are sorted predictably by company name and location.
- Generated chunk files matching `usa-###.json` are cleaned before new chunks are written so stale chunk files do not remain.
- The script fails if no workbook is found, required employer columns are missing, or no company records are produced.
- The script validates that manifest totals, chunk totals, chunk filenames, and concatenated chunk records match the full dataset before completing successfully.

## GitHub Actions

The workflow at `.github/workflows/update-usa-sponsors.yml` can be run manually with `workflow_dispatch`.

It:

1. Checks out the repository.
2. Sets up Python 3.12.
3. Installs `requirements.txt`.
4. Runs `python src/build_usa_sponsors.py`.
5. Uploads `dist/usa-sponsors.json`, `dist/manifest.json`, and `dist/usa-*.json` as workflow artifacts.

Automatic commits and weekly cron scheduling are intentionally not enabled yet.

## Future Google Apps Script Integration

Google Apps Script should consume the compact JSON outputs instead of reading the raw DOL workbook. That keeps heavy XLSX parsing, aggregation, and normalization in Python, while Apps Script can focus on importing or publishing already-prepared sponsor records into the wider visa sponsor directory pipeline.

Recommended consumer flow:

1. Fetch `manifest.json`.
2. Read the `chunks` list.
3. Import one chunk file at a time.
4. Track and validate the total imported records against `manifest.totalCompanies`.
5. Only replace the previous USA staging snapshot after every chunk imports successfully.

Apps Script implementation is intentionally not part of this repository.
