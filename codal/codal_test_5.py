from __future__ import annotations

import io
import json
import re
import time
from pathlib import Path

import pandas as pd
import requests


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

CACHE_DIR = BASE_DIR / "Codal_Cache"
DEBUG_DIR = BASE_DIR / "Debug"

CACHE_DIR.mkdir(exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)


COMPANIES = [
    ("سکارون", "سیمان کارون"),
    ("سلار", "سیمان لارستان"),
    ("ستران", "سیمان تهران"),
    ("سآبیک", "سیمان آبیک"),
    ("سصوفی", "سیمان صوفیان"),
]


PERIODS = [
    "1405/04/31",
    "1405/05/31",
]


# True:
# always download again from Codal
#
# False:
# use already downloaded file if it exists
REFRESH_CACHE = False


SEARCH_URL = (
    "https://search.codal.ir/api/search/v2/q"
)


# ============================================================
# HTTP SESSION
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
# TEXT NORMALIZATION
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

    if pd.isna(value):
        return ""

    text = str(value)

    text = text.translate(DIGIT_MAP)

    text = (
        text
        .replace("ي", "ی")
        .replace("ك", "ک")
        .replace("\u200c", " ")
        .replace("\xa0", " ")
    )

    text = re.sub(
        r"\s+",
        " ",
        text,
    )

    return text.strip()


def normalize_date(value) -> str:

    return (
        normalize_text(value)
        .replace("-", "/")
    )


# ============================================================
# NUMBER
# ============================================================

def to_number(value):

    text = normalize_text(value)

    if text in {
        "",
        "-",
        "—",
        "nan",
        "None",
    }:
        return None

    text = (
        text
        .replace(",", "")
        .replace("٬", "")
    )

    try:

        number = float(text)

        if number.is_integer():
            return int(number)

        return number

    except Exception:

        return None


# ============================================================
# CODAL SEARCH
# ============================================================

def search_reports(symbol):

    params = {
        "Symbol": symbol,
        "LetterType": 58,
        "Category": 3,
        "PageNumber": 1,
        "search": "true",
    }

    response = session.get(
        SEARCH_URL,
        params=params,
        timeout=45,
    )

    response.raise_for_status()

    content_type = (
        response.headers
        .get("Content-Type", "")
        .lower()
    )

    # Normally JSON
    if "json" in content_type:

        data = response.json()

        return data.get(
            "Letters",
            []
        )

    # fallback
    try:

        data = response.json()

        return data.get(
            "Letters",
            []
        )

    except Exception:

        raise RuntimeError(
            "Codal search response is not JSON."
        )


# ============================================================
# SELECT REPORT
# ============================================================

def find_report(symbol, period):

    target_period = normalize_date(period)

    reports = search_reports(symbol)

    matches = []

    for report in reports:

        title = normalize_text(
            report.get(
                "Title",
                ""
            )
        )

        if target_period not in title:
            continue

        if "گزارش فعالیت ماهانه" not in title:
            continue

        matches.append(report)

    if not matches:
        return None

    # Latest PublishDateTime wins.
    # Therefore correction automatically wins.
    matches.sort(
        key=lambda x: (
            x.get("PublishDateTime", "")
            or x.get("SentDateTime", "")
        ),
        reverse=True,
    )

    return matches[0]


# ============================================================
# EXCEL URL
# ============================================================

def get_excel_url(report):

    excel_url = (
        report.get("ExcelUrl")
        or ""
    )

    if excel_url:

        if excel_url.startswith("//"):

            return (
                "https:"
                + excel_url
            )

        if excel_url.startswith("/"):

            return (
                "https://excel.codal.ir"
                + excel_url
            )

        return excel_url

    # fallback using URL
    report_url = (
        report.get("Url")
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
# DOWNLOAD RAW FILE
# ============================================================

def download_report(
    symbol,
    period,
    report,
):

    period_dir = (
        CACHE_DIR
        / period.replace("/", "-")
    )

    period_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    raw_path = (
        period_dir
        / f"{symbol}_raw.xls"
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

    url = get_excel_url(report)

    if not url:

        raise RuntimeError(
            "Excel URL not found."
        )

    print(
        "    Downloading report..."
    )

    response = session.get(
        url,
        timeout=90,
    )

    response.raise_for_status()

    raw_path.write_bytes(
        response.content
    )

    print(
        f"    Downloaded bytes: "
        f"{len(response.content):,}"
    )

    return raw_path


# ============================================================
# DETECT REAL FILE TYPE
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
        data[:10000]
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
# READ CODAL FILE
# ============================================================

def read_codal_file(path):

    file_type = detect_file_type(path)

    print(
        f"    Detected file type: "
        f"{file_type}"
    )

    data = path.read_bytes()

    # --------------------------------------------------------
    # Real old XLS
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # Real XLSX
    # --------------------------------------------------------

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

    # --------------------------------------------------------
    # HTML disguised as XLS
    # --------------------------------------------------------

    if file_type == "html":

        decoded = None

        encodings = [
            "utf-8-sig",
            "utf-8",
            "windows-1256",
            "cp1256",
        ]

        for encoding in encodings:

            try:

                decoded = data.decode(
                    encoding
                )

                break

            except UnicodeDecodeError:
                pass

        if decoded is None:

            decoded = data.decode(
                "utf-8",
                errors="ignore",
            )

        tables = pd.read_html(
            io.StringIO(decoded),
            header=None,
            flavor="lxml",
        )

        return tables

    raise RuntimeError(
        "Unsupported file type."
    )


# ============================================================
# CLEAN TABLE
# ============================================================

def clean_dataframe(df):

    return normalize_html_table(
        df
    )

def normalize_html_table(df):
    """
    Convert Codal HTML tables with MultiIndex headers
    into a normal DataFrame where header levels become
    actual rows.

    This is important because Codal stores many meaningful
    labels inside <th> cells, which pandas may convert into
    MultiIndex columns.
    """

    df = df.copy()

    header_rows = []

    # ========================================================
    # MultiIndex columns
    # ========================================================

    if isinstance(df.columns, pd.MultiIndex):

        for level in range(df.columns.nlevels):

            row = []

            for col in df.columns:

                value = normalize_text(
                    col[level]
                )

                if (
                    value.startswith("Unnamed:")
                    or value.lower() == "nan"
                ):
                    value = ""

                row.append(value)

            # Avoid completely empty header rows
            if any(
                normalize_text(x)
                for x in row
            ):
                header_rows.append(row)

    # ========================================================
    # Normal named columns
    # ========================================================

    else:

        meaningful_headers = []

        has_real_header = False

        for col in df.columns:

            value = normalize_text(col)

            # Numeric RangeIndex is not a real header
            if isinstance(col, int):
                value = ""

            elif (
                value.startswith("Unnamed:")
                or value.lower() == "nan"
            ):
                value = ""

            else:
                has_real_header = True

            meaningful_headers.append(
                value
            )

        if has_real_header:
            header_rows.append(
                meaningful_headers
            )

    # ========================================================
    # Convert all cells to normalized text
    # ========================================================

    data_rows = []

    for row in df.itertuples(
        index=False,
        name=None,
    ):

        data_rows.append(
            [
                normalize_text(value)
                for value in row
            ]
        )

    # ========================================================
    # Combine header + body
    # ========================================================

    all_rows = (
        header_rows
        + data_rows
    )

    if not all_rows:
        return pd.DataFrame()

    result = pd.DataFrame(
        all_rows
    )

    # Remove completely empty rows
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

    # Remove completely empty columns
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

    # CRITICAL:
    # Make columns plain integers.
    # No MultiIndex remains.
    result.columns = range(
        result.shape[1]
    )

    return result

# ============================================================
# MAKE MAC-READABLE XLSX
# ============================================================

def save_readable_xlsx(
    symbol,
    period,
    tables,
):

    period_dir = (
        CACHE_DIR
        / period.replace("/", "-")
    )

    output = (
        period_dir
        / f"{symbol}.xlsx"
    )

    written_tables = 0

    with pd.ExcelWriter(
        output,
        engine="openpyxl",
    ) as writer:

        for i, table in enumerate(
            tables,
            start=1,
        ):

            df = normalize_html_table(
                table
            )

            if df.empty:
                continue

            # Just in case:
            # force simple RangeIndex columns
            df.columns = range(
                df.shape[1]
            )

            df.to_excel(
                writer,
                sheet_name=f"Table_{i}",
                index=False,
                header=False,
            )

            written_tables += 1

    print(
        f"    XLSX created: "
        f"{output.name}"
    )

    print(
        f"    XLSX tables written: "
        f"{written_tables}"
    )

    return output

# ============================================================
# SAVE DEBUG TEXT
# ============================================================

def save_debug(
    symbol,
    period,
    tables,
):

    folder = (
        DEBUG_DIR
        / period.replace("/", "-")
    )

    folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    path = (
        folder
        / f"{symbol}.txt"
    )

    with open(
        path,
        "w",
        encoding="utf-8",
    ) as file:

        for i, table in enumerate(
            tables,
            start=1,
        ):

            df = clean_dataframe(
                table
            )

            file.write(
                "\n"
                + "=" * 120
                + "\n"
            )

            file.write(
                f"TABLE {i}\n"
            )

            file.write(
                f"ROWS={df.shape[0]} "
                f"COLS={df.shape[1]}\n"
            )

            file.write(
                "=" * 120
                + "\n\n"
            )

            file.write(
                df.to_string(
                    index=False,
                    header=False,
                )
            )

            file.write(
                "\n\n"
            )

    print(
        f"    Debug saved: "
        f"{path.name}"
    )


# ============================================================
# GET ALL TABLE TEXT
# ============================================================

def table_text(df):

    df = clean_dataframe(df)

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

    return " ".join(items)


# ============================================================
# TABLE SCORING
# ============================================================

def production_score(df):

    text = table_text(df)

    score = 0

    tests = [
        ("مقدار تولید", 5),
        ("مقدار فروش", 5),
        ("نرخ فروش", 3),
        ("مبلغ فروش", 3),
        ("سیمان", 2),
        ("کلینکر", 2),
        ("دوره یک ماهه", 2),
    ]

    for keyword, points in tests:

        if keyword in text:
            score += points

    return score


def material_score(df):

    text = table_text(df)

    score = 0

    tests = [
        ("خرید مواد اولیه", 6),
        ("مواد اولیه داخلی", 4),
        ("مواد اولیه وارداتی", 3),
        ("خرید دوره یک ماهه", 4),
        ("نرخ", 1),
        ("مبلغ", 1),
    ]

    for keyword, points in tests:

        if keyword in text:
            score += points

    return score


def energy_score(df):

    text = table_text(df)

    score = 0

    tests = [
        ("نوع انرژی", 6),
        ("مقدار مصرف", 5),
        ("واحد اندازه گیری", 3),
        ("برق", 1),
        ("گاز", 1),
        ("مازوت", 1),
    ]

    for keyword, points in tests:

        if keyword in text:
            score += points

    return score


# ============================================================
# FIND BEST TABLE
# ============================================================

def find_best_table(
    tables,
    scorer,
    label,
):

    best_index = None
    best_score = -1
    best_table = None

    for index, table in enumerate(
        tables,
        start=1,
    ):

        df = clean_dataframe(
            table
        )

        score = scorer(df)

        print(
            f"        Table {index}: "
            f"score={score}, "
            f"shape={df.shape}"
        )

        if score > best_score:

            best_score = score
            best_index = index
            best_table = df

    print(
        f"    Best {label} table: "
        f"{best_index} "
        f"(score={best_score})"
    )

    return (
        best_table,
        best_score,
    )


# ============================================================
# PROCESS ONE REPORT
# ============================================================

def process_report(
    symbol,
    company,
    period,
):

    print()
    print(
        f"Processing: "
        f"{symbol} / {period}"
    )

    # --------------------------------------------------------
    # Search
    # --------------------------------------------------------

    report = find_report(
        symbol,
        period,
    )

    if report is None:

        print(
            "    ERROR: "
            "Monthly report not found."
        )

        return False

    title = normalize_text(
        report.get(
            "Title",
            ""
        )
    )

    print(
        f"    Report: {title}"
    )

    # --------------------------------------------------------
    # Download
    # --------------------------------------------------------

    raw_file = download_report(
        symbol,
        period,
        report,
    )

    print(
        f"    Raw file: "
        f"{raw_file.name}"
    )

    # --------------------------------------------------------
    # Read
    # --------------------------------------------------------

    tables = read_codal_file(
        raw_file
    )

    print(
        f"    Tables detected: "
        f"{len(tables)}"
    )

    # --------------------------------------------------------
    # Convert to real XLSX
    # --------------------------------------------------------

    save_readable_xlsx(
        symbol,
        period,
        tables,
    )

    # --------------------------------------------------------
    # Debug
    # --------------------------------------------------------

    save_debug(
        symbol,
        period,
        tables,
    )

    # --------------------------------------------------------
    # Detect tables
    # --------------------------------------------------------

    print(
        "\n    Production table candidates:"
    )

    prod_table, prod_score = (
        find_best_table(
            tables,
            production_score,
            "production",
        )
    )

    print(
        "\n    Material table candidates:"
    )

    material_table, material_score_value = (
        find_best_table(
            tables,
            material_score,
            "materials",
        )
    )

    print(
        "\n    Energy table candidates:"
    )

    energy_table, energy_score_value = (
        find_best_table(
            tables,
            energy_score,
            "energy",
        )
    )

    # --------------------------------------------------------
    # Summary
    # --------------------------------------------------------

    print()
    print(
        "    Detection summary:"
    )

    print(
        f"        Production score: "
        f"{prod_score}"
    )

    print(
        f"        Materials score: "
        f"{material_score_value}"
    )

    print(
        f"        Energy score: "
        f"{energy_score_value}"
    )

    if (
        prod_score <= 0
        or material_score_value <= 0
        or energy_score_value <= 0
    ):

        print(
            "    RESULT: CHECK DEBUG"
        )

        return False

    print(
        "    RESULT: OK"
    )

    return True


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print(
        "=" * 60
    )

    print(
        "CODAL 5-COMPANY DIAGNOSTIC"
    )

    print(
        "=" * 60
    )

    success = 0
    total = 0

    for period in PERIODS:

        print()
        print(
            "=" * 60
        )

        print(
            f"PERIOD: {period}"
        )

        print(
            "=" * 60
        )

        for symbol, company in COMPANIES:

            total += 1

            try:

                ok = process_report(
                    symbol,
                    company,
                    period,
                )

                if ok:
                    success += 1

            except Exception as error:

                print(
                    f"    ERROR: "
                    f"{type(error).__name__}: "
                    f"{error}"
                )

            time.sleep(2)

    print()
    print(
        "=" * 60
    )

    print(
        f"SUCCESS: {success}/{total}"
    )

    print(
        "=" * 60
    )


if __name__ == "__main__":
    main()