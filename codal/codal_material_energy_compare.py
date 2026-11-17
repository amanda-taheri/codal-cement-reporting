from __future__ import annotations

import io
import re
import time
from collections import defaultdict
from pathlib import Path

import pandas as pd
import requests

from lxml import html, etree
from openpyxl import Workbook
from openpyxl.styles import (
    Font,
    PatternFill,
    Alignment,
    Border,
    Side,
)


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

COMPANIES_FILE = BASE_DIR / "companies.xlsx"

CACHE_DIR = BASE_DIR / "Codal_Cache_Latest"
OUTPUT_DIR = BASE_DIR / "Output"

CACHE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

REFRESH_CACHE = False

CODAL_SEARCH_URL = (
    "https://search.codal.ir/api/search/v2/q"
)

COMPANY_DELAY = 3
INITIAL_429_WAIT = 10
MAX_RETRIES = 6


# ============================================================
# HTTP
# ============================================================

session = requests.Session()

session.headers.update(
    {
        "User-Agent": (
            "Mozilla/5.0 "
            "(Macintosh; Intel Mac OS X 10_15_7) "
            "AppleWebKit/537.36 "
            "(KHTML, like Gecko) "
            "Chrome/140.0 Safari/537.36"
        ),
        "Accept": "*/*",
        "Accept-Language": "fa-IR,fa;q=0.9,en;q=0.8",
    }
)


# ============================================================
# NORMALIZATION
# ============================================================

FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
EN_DIGITS = "0123456789"

DIGIT_MAP = str.maketrans(
    FA_DIGITS,
    EN_DIGITS,
)


def normalize_text(value) -> str:

    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value)

    text = text.translate(
        DIGIT_MAP
    )

    text = (
        text
        .replace("ي", "ی")
        .replace("ك", "ک")
        .replace("\u200c", " ")
        .replace("\xa0", " ")
        .replace("–", "-")
        .replace("—", "-")
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def safe_get(row, index):

    if index is None:
        return None

    if index < 0:
        return None

    if index >= len(row):
        return None

    return row[index]


# ============================================================
# LOAD COMPANIES
# ============================================================

def load_companies():

    df = pd.read_excel(
        COMPANIES_FILE
    )

    symbol_col = None
    company_col = None

    for col in df.columns:

        normalized = normalize_text(
            col
        )

        if normalized == "نماد":
            symbol_col = col

        elif normalized == "نام شرکت":
            company_col = col

    if symbol_col is None:
        raise RuntimeError(
            "Column 'نماد' not found."
        )

    if company_col is None:
        raise RuntimeError(
            "Column 'نام شرکت' not found."
        )

    companies = []

    for _, row in df.iterrows():

        symbol = normalize_text(
            row[symbol_col]
        )

        company = normalize_text(
            row[company_col]
        )

        if not symbol:
            continue

        if not company:
            company = symbol

        companies.append(
            (
                symbol,
                company,
            )
        )

    return companies


# ============================================================
# SEARCH CODAL WITH RETRY
# ============================================================

def search_reports(
    symbol,
):

    params = {
        "Symbol": symbol,
        "LetterType": 58,
        "Category": 3,
        "PageNumber": 1,
        "search": "true",
    }

    wait_seconds = (
        INITIAL_429_WAIT
    )

    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):

        try:

            response = session.get(
                CODAL_SEARCH_URL,
                params=params,
                timeout=60,
            )

            if response.status_code == 429:

                print(
                    f"    HTTP 429 "
                    f"(attempt {attempt}/{MAX_RETRIES})"
                )

                print(
                    f"    Waiting {wait_seconds}s..."
                )

                time.sleep(
                    wait_seconds
                )

                wait_seconds = min(
                    wait_seconds * 2,
                    120,
                )

                continue

            response.raise_for_status()

            data = response.json()

            return data.get(
                "Letters",
                [],
            )

        except Exception as error:

            if attempt >= MAX_RETRIES:
                raise

            print(
                f"    Search error: "
                f"{type(error).__name__}"
            )

            print(
                f"    Retrying in "
                f"{wait_seconds}s..."
            )

            time.sleep(
                wait_seconds
            )

            wait_seconds = min(
                wait_seconds * 2,
                120,
            )

    return []


# ============================================================
# FIND LATEST MONTHLY REPORT
# ============================================================

def find_latest_monthly_report(
    reports,
):

    matches = []

    for report in reports:

        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )

        if (
            "گزارش فعالیت ماهانه"
            not in title
        ):
            continue

        matches.append(
            report
        )

    if not matches:
        return None

    matches.sort(
        key=lambda x: (
            x.get(
                "PublishDateTime",
                "",
            )
            or x.get(
                "SentDateTime",
                "",
            )
        ),
        reverse=True,
    )

    return matches[0]


# ============================================================
# EXCEL URL
# ============================================================

def get_excel_url(report):

    excel_url = (
        report.get(
            "ExcelUrl"
        )
        or ""
    )

    if excel_url:

        if excel_url.startswith("//"):
            return "https:" + excel_url

        if excel_url.startswith("/"):
            return (
                "https://excel.codal.ir"
                + excel_url
            )

        return excel_url

    report_url = (
        report.get(
            "Url"
        )
        or ""
    )

    match = re.search(
        r"LetterSerial=([^&]+)",
        report_url,
    )

    if match:

        serial = match.group(1)

        return (
            "https://excel.codal.ir/"
            "service/Excel/GetAll/"
            f"{serial}/0"
        )

    return None


# ============================================================
# DOWNLOAD LATEST REPORT
# ============================================================

def download_report(
    symbol,
    report,
):

    raw_path = (
        CACHE_DIR
        / f"{symbol}_latest_raw.xls"
    )

    if (
        not REFRESH_CACHE
        and raw_path.exists()
        and raw_path.stat().st_size > 500
    ):

        print(
            "    Using cached file"
        )

        return raw_path

    url = get_excel_url(
        report
    )

    if not url:

        raise RuntimeError(
            "Excel URL not found."
        )

    wait_seconds = (
        INITIAL_429_WAIT
    )

    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):

        response = session.get(
            url,
            timeout=90,
        )

        if response.status_code == 429:

            print(
                f"    Download 429 "
                f"(attempt {attempt}/{MAX_RETRIES})"
            )

            print(
                f"    Waiting {wait_seconds}s..."
            )

            time.sleep(
                wait_seconds
            )

            wait_seconds = min(
                wait_seconds * 2,
                120,
            )

            continue

        response.raise_for_status()

        raw_path.write_bytes(
            response.content
        )

        print(
            f"    Downloaded "
            f"{len(response.content):,} bytes"
        )

        return raw_path

    raise RuntimeError(
        "Download failed."
    )


# ============================================================
# FILE TYPE
# ============================================================

def detect_file_type(path):

    data = path.read_bytes()

    if data.startswith(
        b"\xD0\xCF\x11\xE0"
    ):
        return "xls"

    if data.startswith(b"PK"):
        return "xlsx"

    sample = (
        data[:15000]
        .decode(
            "utf-8",
            errors="ignore",
        )
        .lower()
    )

    if (
        "<html" in sample
        or "<table" in sample
        or "<!doctype" in sample
    ):
        return "html"

    return "unknown"


# ============================================================
# HTML READER
# ============================================================

def decode_html(data):

    for encoding in [
        "utf-8-sig",
        "utf-8",
        "windows-1256",
        "cp1256",
    ]:

        try:
            return data.decode(
                encoding
            )

        except UnicodeDecodeError:
            pass

    return data.decode(
        "utf-8",
        errors="ignore",
    )


def read_html_tables_robust(
    decoded,
):

    try:

        tables = pd.read_html(
            io.StringIO(decoded),
            header=None,
            flavor="lxml",
        )

        if tables:
            return tables

    except Exception:
        pass

    document = html.fromstring(
        decoded
    )

    html_tables = document.xpath(
        "//table"
    )

    result = []

    for table_element in html_tables:

        table_html = etree.tostring(
            table_element,
            encoding="unicode",
            method="html",
        )

        try:

            parsed = pd.read_html(
                io.StringIO(
                    table_html
                ),
                header=None,
                flavor="lxml",
            )

            for df in parsed:

                if (
                    df is not None
                    and not df.empty
                ):

                    result.append(
                        df
                    )

        except Exception:
            continue

    if not result:

        raise RuntimeError(
            "No readable tables found."
        )

    return result


def read_codal_file(path):

    file_type = detect_file_type(
        path
    )

    data = path.read_bytes()

    if file_type == "xls":

        sheets = pd.read_excel(
            io.BytesIO(data),
            sheet_name=None,
            header=None,
            engine="xlrd",
        )

        return list(
            sheets.values()
        )

    if file_type == "xlsx":

        sheets = pd.read_excel(
            io.BytesIO(data),
            sheet_name=None,
            header=None,
            engine="openpyxl",
        )

        return list(
            sheets.values()
        )

    if file_type == "html":

        decoded = decode_html(
            data
        )

        return read_html_tables_robust(
            decoded
        )

    raise RuntimeError(
        "Unsupported file type."
    )


# ============================================================
# NORMALIZE HTML TABLE
# ============================================================

def normalize_html_table(df):

    df = df.copy()

    header_rows = []

    if isinstance(
        df.columns,
        pd.MultiIndex,
    ):

        for level in range(
            df.columns.nlevels
        ):

            row = []

            for col in df.columns:

                value = normalize_text(
                    col[level]
                )

                if (
                    value.startswith(
                        "Unnamed:"
                    )
                    or value.lower()
                    == "nan"
                ):
                    value = ""

                row.append(
                    value
                )

            if any(
                normalize_text(x)
                for x in row
            ):

                header_rows.append(
                    row
                )

    else:

        row = []

        has_header = False

        for col in df.columns:

            if isinstance(
                col,
                int,
            ):

                value = ""

            else:

                value = normalize_text(
                    col
                )

                if (
                    value.startswith(
                        "Unnamed:"
                    )
                    or value.lower()
                    == "nan"
                ):

                    value = ""

                elif value:

                    has_header = True

            row.append(
                value
            )

        if has_header:

            header_rows.append(
                row
            )

    body_rows = []

    for row in df.itertuples(
        index=False,
        name=None,
    ):

        body_rows.append(
            [
                normalize_text(value)
                for value in row
            ]
        )

    all_rows = (
        header_rows
        + body_rows
    )

    if not all_rows:
        return pd.DataFrame()

    result = pd.DataFrame(
        all_rows
    )

    result = result[
        ~result.apply(
            lambda row:
                all(
                    normalize_text(v) == ""
                    for v in row
                ),
            axis=1,
        )
    ]

    result = result.loc[
        :,
        ~result.apply(
            lambda col:
                all(
                    normalize_text(v) == ""
                    for v in col
                ),
            axis=0,
        )
    ]

    result = result.reset_index(
        drop=True
    )

    result.columns = range(
        result.shape[1]
    )

    return result


# ============================================================
# TABLE TEXT
# ============================================================

def table_text(df):

    df = normalize_html_table(
        df
    )

    items = []

    for value in (
        df.fillna("")
        .values
        .flatten()
    ):

        text = normalize_text(
            value
        )

        if text:
            items.append(text)

    return " ".join(
        items
    )


# ============================================================
# SCORES
# ============================================================

def material_score(df):

    text = table_text(
        df
    )

    score = 0

    checks = [
        ("خرید مواد اولیه", 6),
        ("مواد اولیه داخلی", 4),
        ("مواد اولیه وارداتی", 3),
        ("خرید دوره یک ماهه", 4),
    ]

    for keyword, points in checks:

        if keyword in text:
            score += points

    return score


def energy_score(df):

    text = table_text(
        df
    )

    score = 0

    checks = [
        ("نوع انرژی", 6),
        ("مقدار مصرف", 5),
        ("واحد اندازه گیری", 3),
        ("برق", 1),
        ("گاز", 1),
        ("مازوت", 1),
        ("گازوئیل", 1),
    ]

    for keyword, points in checks:

        if keyword in text:
            score += points

    return score


def find_best_table(
    tables,
    scorer,
):

    best_table = None
    best_score = -1

    for table in tables:

        df = normalize_html_table(
            table
        )

        if df.empty:
            continue

        score = scorer(
            df
        )

        if score > best_score:

            best_score = score
            best_table = df

    return (
        best_table,
        best_score,
    )


# ============================================================
# EXTRACT RAW MATERIALS
# ============================================================

def extract_material_names(df):

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    items = []

    purchase_type = None

    for row in values:

        name = normalize_text(
            safe_get(
                row,
                0,
            )
        )

        unit = normalize_text(
            safe_get(
                row,
                1,
            )
        )

        if (
            "مواد اولیه داخلی"
            in name
        ):

            purchase_type = "داخلی"
            continue

        if (
            "مواد اولیه وارداتی"
            in name
        ):

            purchase_type = "وارداتی"
            continue

        if purchase_type is None:
            continue

        if not name:
            continue

        if name.startswith(
            "جمع"
        ):
            continue

        if not unit:
            continue

        items.append(
            name
        )

    return sorted(
        set(items)
    )


# ============================================================
# EXTRACT RAW ENERGY
# ============================================================

def extract_energy_names(df):

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    items = []

    for row in values:

        name = normalize_text(
            safe_get(
                row,
                2,
            )
        )

        unit = normalize_text(
            safe_get(
                row,
                3,
            )
        )

        if not name:
            continue

        if not unit:
            continue

        if name == "نوع انرژی":
            continue

        if (
            "مقدار مصرف" in name
            or "دوره یک ماهه" in name
            or "دوره 1 ماهه" in name
            or name == "نرخ"
            or name == "مبلغ"
        ):

            continue

        items.append(
            name
        )

    return sorted(
        set(items)
    )


# ============================================================
# CREATE EXCEL
# ============================================================

def create_excel(
    companies,
    company_materials,
    company_energy,
):

    output_file = (
        OUTPUT_DIR
        / "Codal_Materials_Energy_Comparison.xlsx"
    )

    wb = Workbook()

    wb.remove(
        wb.active
    )


    # ========================================================
    # STYLES
    # ========================================================

    yellow_fill = PatternFill(
        fill_type="solid",
        fgColor="FFF2CC",
    )

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7",
    )

    header_font = Font(
        bold=True,
    )

    thin = Side(
        style="thin",
        color="D9D9D9",
    )

    border = Border(
        left=thin,
        right=thin,
        top=thin,
        bottom=thin,
    )

    center = Alignment(
        horizontal="center",
        vertical="center",
        wrap_text=True,
    )

    right = Alignment(
        horizontal="right",
        vertical="center",
        wrap_text=True,
    )


    # ========================================================
    # MATERIAL COUNTS
    # ========================================================

    material_count = (
        defaultdict(int)
    )

    for symbol, company in companies:

        for item in set(
            company_materials.get(
                symbol,
                [],
            )
        ):

            material_count[
                item
            ] += 1


    # ========================================================
    # ENERGY COUNTS
    # ========================================================

    energy_count = (
        defaultdict(int)
    )

    for symbol, company in companies:

        for item in set(
            company_energy.get(
                symbol,
                [],
            )
        ):

            energy_count[
                item
            ] += 1


    # ========================================================
    # MATERIALS SHEET
    # ========================================================

    ws = wb.create_sheet(
        "مواد اولیه"
    )

    ws.sheet_view.rightToLeft = True

    for col_index, (
        symbol,
        company,
    ) in enumerate(
        companies,
        start=1,
    ):

        cell = ws.cell(
            1,
            col_index,
            company,
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

        items = company_materials.get(
            symbol,
            [],
        )

        for row_index, item in enumerate(
            items,
            start=2,
        ):

            cell = ws.cell(
                row_index,
                col_index,
                item,
            )

            cell.alignment = right
            cell.border = border

            if (
                material_count[
                    item
                ]
                == 1
            ):

                cell.fill = yellow_fill

    for col_index in range(
        1,
        len(companies) + 1,
    ):

        letter = ws.cell(
            1,
            col_index,
        ).column_letter

        ws.column_dimensions[
            letter
        ].width = 26

    ws.freeze_panes = "A2"


    # ========================================================
    # ENERGY SHEET
    # ========================================================

    ws = wb.create_sheet(
        "انرژی"
    )

    ws.sheet_view.rightToLeft = True

    for col_index, (
        symbol,
        company,
    ) in enumerate(
        companies,
        start=1,
    ):

        cell = ws.cell(
            1,
            col_index,
            company,
        )

        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center
        cell.border = border

        items = company_energy.get(
            symbol,
            [],
        )

        for row_index, item in enumerate(
            items,
            start=2,
        ):

            cell = ws.cell(
                row_index,
                col_index,
                item,
            )

            cell.alignment = right
            cell.border = border

            if (
                energy_count[
                    item
                ]
                == 1
            ):

                cell.fill = yellow_fill

    for col_index in range(
        1,
        len(companies) + 1,
    ):

        letter = ws.cell(
            1,
            col_index,
        ).column_letter

        ws.column_dimensions[
            letter
        ].width = 26

    ws.freeze_panes = "A2"


    wb.save(
        output_file
    )

    return output_file


# ============================================================
# MAIN
# ============================================================

def main():

    companies = load_companies()

    print()
    print(
        "=" * 70
    )

    print(
        "CODAL LATEST MATERIAL / ENERGY COMPARISON"
    )

    print(
        f"Companies: {len(companies)}"
    )

    print(
        "=" * 70
    )

    company_materials = {}
    company_energy = {}


    for index, (
        symbol,
        company,
    ) in enumerate(
        companies,
        start=1,
    ):

        print()
        print(
            f"[{index}/{len(companies)}] "
            f"{symbol} - {company}"
        )

        try:

            # =================================================
            # ONE SEARCH ONLY
            # =================================================

            reports = search_reports(
                symbol
            )

            print(
                f"    Search results: "
                f"{len(reports)}"
            )


            # =================================================
            # LATEST MONTHLY REPORT ONLY
            # =================================================

            report = (
                find_latest_monthly_report(
                    reports
                )
            )

            if report is None:

                print(
                    "    Monthly report not found"
                )

                company_materials[
                    symbol
                ] = []

                company_energy[
                    symbol
                ] = []

                time.sleep(
                    COMPANY_DELAY
                )

                continue


            title = normalize_text(
                report.get(
                    "Title",
                    "",
                )
            )

            print(
                f"    Latest report: "
                f"{title}"
            )


            # =================================================
            # DOWNLOAD / CACHE
            # =================================================

            raw_file = download_report(
                symbol,
                report,
            )


            # =================================================
            # TABLES
            # =================================================

            tables = read_codal_file(
                raw_file
            )

            print(
                f"    Tables detected: "
                f"{len(tables)}"
            )


            # =================================================
            # MATERIALS
            # =================================================

            material_table, m_score = (
                find_best_table(
                    tables,
                    material_score,
                )
            )

            materials = []

            if (
                material_table
                is not None
                and m_score > 0
            ):

                materials = (
                    extract_material_names(
                        material_table
                    )
                )


            # =================================================
            # ENERGY
            # =================================================

            energy_table, e_score = (
                find_best_table(
                    tables,
                    energy_score,
                )
            )

            energies = []

            if (
                energy_table
                is not None
                and e_score > 0
            ):

                energies = (
                    extract_energy_names(
                        energy_table
                    )
                )


            company_materials[
                symbol
            ] = materials

            company_energy[
                symbol
            ] = energies


            print(
                f"    Materials: "
                f"{len(materials)}"
            )

            print(
                f"    Energy: "
                f"{len(energies)}"
            )

            print(
                "    RESULT: OK"
            )


        except Exception as error:

            print(
                f"    ERROR: "
                f"{type(error).__name__}: "
                f"{error}"
            )

            company_materials[
                symbol
            ] = []

            company_energy[
                symbol
            ] = []


        # Gentle delay before next company search
        time.sleep(
            COMPANY_DELAY
        )


    # ========================================================
    # OUTPUT
    # ========================================================

    output_file = create_excel(
        companies,
        company_materials,
        company_energy,
    )


    # ========================================================
    # SUMMARY
    # ========================================================

    all_materials = set()

    for items in (
        company_materials.values()
    ):

        all_materials.update(
            items
        )


    all_energy = set()

    for items in (
        company_energy.values()
    ):

        all_energy.update(
            items
        )


    print()
    print(
        "=" * 70
    )

    print(
        "DONE"
    )

    print(
        f"Unique material titles: "
        f"{len(all_materials)}"
    )

    print(
        f"Unique energy titles: "
        f"{len(all_energy)}"
    )

    print(
        f"Output created: "
        f"{output_file.name}"
    )

    print(
        "=" * 70
    )


if __name__ == "__main__":
    main()