"""Build employer-level U.S. visa sponsor data from DOL LCA disclosure files."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from dataclasses import dataclass, field
from datetime import date, datetime, timezone
from pathlib import Path
from tempfile import TemporaryDirectory
from typing import Iterable
from urllib.parse import urljoin

import requests
from bs4 import BeautifulSoup
from openpyxl import load_workbook


DOL_PERFORMANCE_URL = "https://www.dol.gov/agencies/eta/foreign-labor/performance"
SOURCE_NAME = "U.S. DOL LCA Disclosure"
COUNTRY = "United States"
COUNTRY_CODE = "US"
DISCLOSURE_RE = re.compile(
    r"LCA_Disclosure_Data_FY(?P<year>\d{4})_Q(?P<quarter>[1-4])\.xlsx",
    re.IGNORECASE,
)
EXCLUDED_LINK_TEXT = ("appendix", "worksite")
RELEVANT_VISA_CLASSES = {"H-1B", "H-1B1", "E-3"}
RELEVANT_STATUS_PREFIXES = ("CERTIFIED",)

HEADER_ALIASES = {
    "employer_name": {
        "employername",
        "employerlegalbusinessname",
        "employerlegalname",
        "petitionername",
        "companyname",
    },
    "employer_city": {
        "employercity",
        "employercityname",
        "employerlocationcity",
        "petitionercity",
    },
    "employer_state": {
        "employerstate",
        "employerstateprovince",
        "employerprovince",
        "employerlocationstate",
        "petitionerstate",
    },
    "visa_class": {
        "visaclass",
        "classofadmission",
        "visa",
        "program",
    },
    "soc_title": {
        "soctitle",
        "socname",
        "sococcupationtitle",
        "socjobtitle",
        "jobtitle",
        "jobtitletext",
    },
    "case_status": {
        "casestatus",
        "status",
        "decisionstatus",
        "casefinalstatus",
    },
}

REQUIRED_COLUMNS = ("employer_name", "employer_city", "employer_state")


@dataclass(frozen=True)
class DisclosureFile:
    url: str
    filename: str
    fiscal_year: int
    quarter: int


@dataclass
class EmployerAggregate:
    company_name: str
    city: str
    state: str
    visa_programs: Counter[str] = field(default_factory=Counter)
    job_titles: Counter[str] = field(default_factory=Counter)
    filing_count: int = 0


def normalize_header(value: object) -> str:
    text = normalize_whitespace(value)
    return re.sub(r"[^a-z0-9]", "", text.lower())


def normalize_whitespace(value: object) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value)).strip()


def normalize_company_name(value: object) -> str:
    return normalize_whitespace(value)


def normalize_group_key(*parts: str) -> tuple[str, ...]:
    return tuple(normalize_whitespace(part).casefold() for part in parts)


def slugify(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug or "unknown"


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Build compact U.S. LCA employer sponsor JSON from DOL disclosure data."
    )
    parser.add_argument(
        "--source-page",
        default=DOL_PERFORMANCE_URL,
        help="DOL performance page used for workbook discovery.",
    )
    parser.add_argument(
        "--workbook",
        type=Path,
        help="Optional local XLSX workbook. Skips DOL discovery and download.",
    )
    parser.add_argument(
        "--source-url",
        help="Source URL to write into output when --workbook is used.",
    )
    parser.add_argument(
        "--download-dir",
        type=Path,
        help="Optional directory for downloaded workbooks.",
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("dist/usa-sponsors.json"),
        help="Output JSON path.",
    )
    return parser.parse_args()


def discover_latest_disclosure(source_page: str) -> DisclosureFile:
    response = requests.get(source_page, timeout=45)
    response.raise_for_status()
    soup = BeautifulSoup(response.text, "html.parser")
    candidates: list[DisclosureFile] = []

    for link in soup.find_all("a", href=True):
        href = str(link["href"])
        text = " ".join(link.stripped_strings)
        combined = f"{href} {text}".lower()
        if any(excluded in combined for excluded in EXCLUDED_LINK_TEXT):
            continue

        match = DISCLOSURE_RE.search(href) or DISCLOSURE_RE.search(text)
        if not match:
            continue

        filename = match.group(0)
        candidates.append(
            DisclosureFile(
                url=urljoin(source_page, href),
                filename=filename,
                fiscal_year=int(match.group("year")),
                quarter=int(match.group("quarter")),
            )
        )

    if not candidates:
        raise RuntimeError(f"No cumulative LCA disclosure workbook found at {source_page}")

    return max(candidates, key=lambda item: (item.fiscal_year, item.quarter))


def download_file(disclosure: DisclosureFile, destination_dir: Path) -> Path:
    destination_dir.mkdir(parents=True, exist_ok=True)
    destination = destination_dir / disclosure.filename

    with requests.get(disclosure.url, stream=True, timeout=(15, 180)) as response:
        response.raise_for_status()
        with destination.open("wb") as output:
            for chunk in response.iter_content(chunk_size=1024 * 1024):
                if chunk:
                    output.write(chunk)

    if destination.stat().st_size == 0:
        raise RuntimeError(f"Downloaded workbook is empty: {destination}")
    return destination


def find_header_row(rows: Iterable[tuple[object, ...]]) -> tuple[int, list[str]]:
    best_row_number = 0
    best_headers: list[str] = []
    best_score = 0

    for row_number, row in enumerate(rows, start=1):
        headers = [normalize_header(cell) for cell in row]
        header_set = set(headers)
        score = sum(
            1
            for aliases in HEADER_ALIASES.values()
            if aliases.intersection(header_set)
        )
        if score > best_score:
            best_row_number = row_number
            best_headers = headers
            best_score = score
        if score >= len(REQUIRED_COLUMNS):
            return row_number, headers
        if row_number >= 30:
            break

    if best_score:
        return best_row_number, best_headers
    raise RuntimeError("Could not locate the workbook header row.")


def resolve_columns(headers: list[str]) -> dict[str, int | None]:
    columns: dict[str, int | None] = {}
    for column_name, aliases in HEADER_ALIASES.items():
        columns[column_name] = next(
            (index for index, header in enumerate(headers) if header in aliases),
            None,
        )

    missing = [name for name in REQUIRED_COLUMNS if columns[name] is None]
    if missing:
        raise RuntimeError(f"Missing required employer columns: {', '.join(missing)}")

    return columns


def get_cell(row: tuple[object, ...], index: int | None) -> str:
    if index is None or index >= len(row):
        return ""
    return normalize_whitespace(row[index])


def is_relevant_case(visa_class: str, case_status: str) -> bool:
    visa_class_upper = visa_class.upper()
    case_status_upper = case_status.upper()

    if visa_class_upper and visa_class_upper not in RELEVANT_VISA_CLASSES:
        return False
    if case_status_upper and not case_status_upper.startswith(RELEVANT_STATUS_PREFIXES):
        return False
    return True


def get_metadata_from_filename(filename: str) -> tuple[int, int]:
    match = DISCLOSURE_RE.search(filename)
    if not match:
        raise RuntimeError(f"Could not determine fiscal year and quarter from {filename}")
    return int(match.group("year")), int(match.group("quarter"))


def process_workbook(workbook_path: Path) -> dict[tuple[str, str, str], EmployerAggregate]:
    workbook = load_workbook(filename=workbook_path, read_only=True, data_only=True)
    try:
        worksheet = workbook[workbook.sheetnames[0]]
        row_iterator = worksheet.iter_rows(values_only=True)
        header_row_number, headers = find_header_row(row_iterator)
        columns = resolve_columns(headers)

        aggregates: dict[tuple[str, str, str], EmployerAggregate] = {}
        for row_number, row in enumerate(row_iterator, start=header_row_number + 1):
            company_name = normalize_company_name(get_cell(row, columns["employer_name"]))
            city = normalize_whitespace(get_cell(row, columns["employer_city"])).title()
            state = normalize_whitespace(get_cell(row, columns["employer_state"])).upper()
            if not company_name or not city or not state:
                continue

            visa_class = get_cell(row, columns["visa_class"]).upper()
            case_status = get_cell(row, columns["case_status"])
            if not is_relevant_case(visa_class, case_status):
                continue

            key = normalize_group_key(company_name, city, state)
            aggregate = aggregates.setdefault(
                key,
                EmployerAggregate(company_name=company_name, city=city, state=state),
            )
            aggregate.filing_count += 1
            if visa_class:
                aggregate.visa_programs[visa_class] += 1

            soc_title = get_cell(row, columns["soc_title"])
            if soc_title:
                aggregate.job_titles[normalize_whitespace(soc_title).title()] += 1

            if row_number % 25000 == 0:
                print(f"Processed {row_number:,} workbook rows...", file=sys.stderr)
    finally:
        workbook.close()

    if not aggregates:
        raise RuntimeError("No employer-level sponsor records were produced.")
    return aggregates


def build_company_records(
    aggregates: dict[tuple[str, str, str], EmployerAggregate],
    source_url: str,
) -> list[dict[str, object]]:
    records: list[dict[str, object]] = []
    for aggregate in aggregates.values():
        location = f"{aggregate.city}, {aggregate.state}"
        visa_programs = (
            " / ".join(sorted(aggregate.visa_programs))
            if aggregate.visa_programs
            else "H-1B / H-1B1 / E-3"
        )
        top_titles = [title for title, _count in aggregate.job_titles.most_common(5)]

        records.append(
            {
                "company_name": aggregate.company_name,
                "company_slug": slugify(aggregate.company_name),
                "country": COUNTRY,
                "country_code": COUNTRY_CODE,
                "country_slug": "united-states",
                "industry": "",
                "location": location,
                "location_slug": slugify(location),
                "tags": ", ".join(top_titles),
                "sponsor_type": "LCA Employer",
                "visa_program": visa_programs,
                "top_job_titles": top_titles,
                "source": SOURCE_NAME,
                "source_url": source_url,
                "filing_count": aggregate.filing_count,
            }
        )

    return sorted(
        records,
        key=lambda record: (
            str(record["company_name"]).casefold(),
            str(record["location"]).casefold(),
        ),
    )


def write_output(
    output_path: Path,
    companies: list[dict[str, object]],
    source_url: str,
    fiscal_year: int,
    quarter: int,
) -> None:
    payload = {
        "version": date.today().isoformat(),
        "generatedAt": datetime.now(timezone.utc).isoformat(),
        "source": SOURCE_NAME,
        "sourceUrl": source_url,
        "fiscalYear": fiscal_year,
        "quarter": quarter,
        "totalCompanies": len(companies),
        "companies": companies,
    }

    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
        encoding="utf-8",
    )


def run() -> None:
    args = parse_args()

    if args.workbook:
        workbook_path = args.workbook
        source_url = args.source_url or str(workbook_path)
        fiscal_year, quarter = get_metadata_from_filename(workbook_path.name)
    else:
        disclosure = discover_latest_disclosure(args.source_page)
        source_url = disclosure.url
        fiscal_year = disclosure.fiscal_year
        quarter = disclosure.quarter
        if args.download_dir:
            workbook_path = download_file(disclosure, args.download_dir)
        else:
            with TemporaryDirectory() as temp_dir:
                workbook_path = download_file(disclosure, Path(temp_dir))
                aggregates = process_workbook(workbook_path)
                companies = build_company_records(aggregates, source_url)
                write_output(args.output, companies, source_url, fiscal_year, quarter)
                print(f"Wrote {len(companies):,} companies to {args.output}")
                return

    aggregates = process_workbook(workbook_path)
    companies = build_company_records(aggregates, source_url)
    write_output(args.output, companies, source_url, fiscal_year, quarter)
    print(f"Wrote {len(companies):,} companies to {args.output}")


def main() -> int:
    try:
        run()
    except Exception as error:
        print(f"Error: {error}", file=sys.stderr)
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
