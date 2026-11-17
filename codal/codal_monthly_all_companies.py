from __future__ import annotations

import io
import json
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

CACHE_DIR = BASE_DIR / "Codal_Cache"
SEARCH_CACHE_DIR = BASE_DIR / "Codal_Search_Cache"
OUTPUT_DIR = BASE_DIR / "Output"

CACHE_DIR.mkdir(exist_ok=True)
SEARCH_CACHE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)


OUTPUT_FILE = (
    OUTPUT_DIR
    / "گزارش_عملکرد_ماهانه_شرکتهای_سیمانی.xlsx"
)


PERIODS = [
    "1405/01/31",
    "1405/02/31",
    "1405/03/31",
    "1405/04/31",
    "1405/05/31",
]


# ------------------------------------------------------------
# CACHE SETTINGS
# ------------------------------------------------------------

# False = use existing downloaded Codal files
REFRESH_REPORT_CACHE = False

# False = use existing Codal search JSON
REFRESH_SEARCH_CACHE = False


# ------------------------------------------------------------
# RATE LIMIT SETTINGS
# ------------------------------------------------------------

MAX_RETRIES = 6

INITIAL_WAIT = 10

COMPANY_DELAY = 3

DOWNLOAD_DELAY = 1


CODAL_SEARCH_URL = (
    "https://search.codal.ir/api/search/v2/q"
)


NUMBER_FORMAT = "#,##0"


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


def normalize_date(value) -> str:

    return (
        normalize_text(value)
        .replace("-", "/")
    )


def safe_get(row, index):

    if index is None:
        return None

    if index < 0:
        return None

    if index >= len(row):
        return None

    return row[index]


# ============================================================
# NUMBERS
# ============================================================

def to_number(value):

    text = normalize_text(
        value
    )

    if text in {
        "",
        "-",
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

        return float(
            text
        )

    except Exception:

        return None


# ============================================================
# LOAD COMPANIES.XLSX
# ============================================================

def load_companies():

    df = pd.read_excel(
        COMPANIES_FILE
    )

    symbol_col = None
    company_col = None

    for col in df.columns:

        value = normalize_text(
            col
        )

        if value == "نماد":
            symbol_col = col

        elif value == "نام شرکت":
            company_col = col

    if symbol_col is None:

        raise RuntimeError(
            "Column 'نماد' not found in companies.xlsx"
        )

    if company_col is None:

        raise RuntimeError(
            "Column 'نام شرکت' not found in companies.xlsx"
        )

    companies = []

    seen_symbols = set()

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

        # avoid accidental duplicate symbol
        if symbol in seen_symbols:
            continue

        seen_symbols.add(
            symbol
        )

        companies.append(
            (
                symbol,
                company,
            )
        )

    return companies


# ============================================================
# ENERGY STANDARDIZATION
# ============================================================

ENERGY_MASTER = [
    "آب",
    "برق",
    "گاز طبیعی",
    "مازوت",
    "گازوئیل",
]


ENERGY_UNITS = {
    "آب": "مترمکعب",
    "برق": "مگاوات ساعت",
    "گاز طبیعی": "مترمکعب",
    "مازوت": "لیتر",
    "گازوئیل": "لیتر",
}


ENERGY_ALIASES = {

    # WATER
    "آب": "آب",
    "آب مصرفی": "آب",
    "آب صنعتی": "آب",
    "آب خریداری": "آب",
    "آب خریداری شده": "آب",
    "آب خریداری‌شده": "آب",

    # ELECTRICITY
    "برق": "برق",
    "برق مصرفی": "برق",
    "برق خریداری": "برق",
    "برق خریداری شده": "برق",
    "برق خریداری‌شده": "برق",

    # NATURAL GAS
    "گاز": "گاز طبیعی",
    "گاز طبیعی": "گاز طبیعی",
    "گاز مصرفی": "گاز طبیعی",
    "گاز سوخت": "گاز طبیعی",
    "گاز صنعتی": "گاز طبیعی",
    "گاز خریداری شده": "گاز طبیعی",
    "گاز خریداری‌شده": "گاز طبیعی",

    # MAZUT
    "مازوت": "مازوت",
    "نفت کوره": "مازوت",
    "نفت کوره (مازوت)": "مازوت",

    # GASOIL
    "گازوییل": "گازوئیل",
    "گازوئیل": "گازوئیل",
    "نفت گاز": "گازوئیل",
}


def canonical_energy(
    name
):

    name = normalize_text(
        name
    )

    return ENERGY_ALIASES.get(
        name,
        name,
    )


# ============================================================
# SEARCH CACHE
# ============================================================

def get_search_cache_path(
    symbol
):

    safe_symbol = (
        symbol
        .replace("/", "_")
        .replace("\\", "_")
    )

    return (
        SEARCH_CACHE_DIR
        / f"{safe_symbol}.json"
    )


# ============================================================
# CODAL SEARCH
# ============================================================

def search_reports(
    symbol
):

    cache_file = (
        get_search_cache_path(
            symbol
        )
    )

    # --------------------------------------------------------
    # CACHE
    # --------------------------------------------------------

    if (
        not REFRESH_SEARCH_CACHE
        and cache_file.exists()
    ):

        try:

            with open(
                cache_file,
                "r",
                encoding="utf-8",
            ) as file:

                data = json.load(
                    file
                )

            print(
                "    Using cached search results"
            )

            return data

        except Exception:
            pass


    # --------------------------------------------------------
    # SEARCH CODAL
    # --------------------------------------------------------

    params = {
        "Symbol": symbol,
        "LetterType": 58,
        "Category": 3,
        "PageNumber": 1,
        "search": "true",
    }

    wait_seconds = (
        INITIAL_WAIT
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

            if (
                response.status_code
                == 429
            ):

                print(
                    f"    HTTP 429 "
                    f"(attempt "
                    f"{attempt}/{MAX_RETRIES})"
                )

                print(
                    f"    Waiting "
                    f"{wait_seconds}s..."
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

            letters = data.get(
                "Letters",
                [],
            )

            # save cache
            with open(
                cache_file,
                "w",
                encoding="utf-8",
            ) as file:

                json.dump(
                    letters,
                    file,
                    ensure_ascii=False,
                    indent=2,
                )

            return letters


        except Exception as error:

            if (
                attempt
                >= MAX_RETRIES
            ):

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
# FIND EXACT PERIOD
# ============================================================

def find_report_from_list(
    reports,
    period,
):

    target_period = (
        normalize_date(
            period
        )
    )

    matches = []

    for report in reports:

        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )

        if (
            target_period
            not in title
        ):
            continue

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


    # --------------------------------------------------------
    # IMPORTANT:
    # latest version of exact month wins.
    #
    # Original:
    # 1405/05/31
    #
    # Correction:
    # 1405/05/31 (اصلاحیه)
    #
    # Latest PublishDateTime wins.
    # --------------------------------------------------------

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

def get_excel_url(
    report
):

    excel_url = (
        report.get(
            "ExcelUrl"
        )
        or ""
    )

    if excel_url:

        if excel_url.startswith(
            "//"
        ):

            return (
                "https:"
                + excel_url
            )

        if excel_url.startswith(
            "/"
        ):

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

        serial = (
            match.group(1)
        )

        return (
            "https://excel.codal.ir/"
            "service/Excel/GetAll/"
            f"{serial}/0"
        )

    return None


# ============================================================
# DOWNLOAD REPORT
# ============================================================

def download_report(
    symbol,
    period,
    report,
):

    period_dir = (
        CACHE_DIR
        / period.replace(
            "/",
            "-",
        )
    )

    period_dir.mkdir(
        parents=True,
        exist_ok=True,
    )

    raw_path = (
        period_dir
        / f"{symbol}_raw.xls"
    )


    # --------------------------------------------------------
    # FILE CACHE
    # --------------------------------------------------------

    if (
        not REFRESH_REPORT_CACHE
        and raw_path.exists()
        and raw_path.stat().st_size
        > 500
    ):

        print(
            "        Using cached report"
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
        INITIAL_WAIT
    )


    for attempt in range(
        1,
        MAX_RETRIES + 1,
    ):

        try:

            response = session.get(
                url,
                timeout=90,
            )

            if (
                response.status_code
                == 429
            ):

                print(
                    f"        Download 429 "
                    f"(attempt "
                    f"{attempt}/{MAX_RETRIES})"
                )

                print(
                    f"        Waiting "
                    f"{wait_seconds}s..."
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
                f"        Downloaded "
                f"{len(response.content):,} "
                f"bytes"
            )

            time.sleep(
                DOWNLOAD_DELAY
            )

            return raw_path


        except Exception:

            if (
                attempt
                >= MAX_RETRIES
            ):

                raise

            time.sleep(
                wait_seconds
            )

            wait_seconds = min(
                wait_seconds * 2,
                120,
            )


    raise RuntimeError(
        "Report download failed."
    )


# ============================================================
# FILE TYPE
# ============================================================

def detect_file_type(
    path
):

    data = path.read_bytes()

    if data.startswith(
        b"\xD0\xCF\x11\xE0"
    ):

        return "xls"

    if data.startswith(
        b"PK"
    ):

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
        "<html"
        in sample
        or "<table"
        in sample
        or "<!doctype"
        in sample
    ):

        return "html"


    return "unknown"


# ============================================================
# READ HTML
# ============================================================

def decode_html(
    data
):

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


def read_html_tables(
    decoded
):

    try:

        tables = pd.read_html(
            io.StringIO(
                decoded
            ),
            header=None,
            flavor="lxml",
        )

        if tables:

            return tables

    except Exception:

        pass


    # --------------------------------------------------------
    # FALLBACK
    # --------------------------------------------------------

    document = html.fromstring(
        decoded
    )

    html_tables = (
        document.xpath(
            "//table"
        )
    )

    result = []

    for element in (
        html_tables
    ):

        table_html = (
            etree.tostring(
                element,
                encoding="unicode",
                method="html",
            )
        )

        try:

            parsed = pd.read_html(
                io.StringIO(
                    table_html
                ),
                header=None,
                flavor="lxml",
            )

            result.extend(
                parsed
            )

        except Exception:

            continue


    return result


# ============================================================
# READ REPORT
# ============================================================

def read_codal_file(
    path
):

    file_type = (
        detect_file_type(
            path
        )
    )

    data = path.read_bytes()


    if file_type == "html":

        return read_html_tables(
            decode_html(
                data
            )
        )


    if file_type == "xlsx":

        sheets = pd.read_excel(
            io.BytesIO(
                data
            ),
            sheet_name=None,
            header=None,
            engine="openpyxl",
        )

        return list(
            sheets.values()
        )


    if file_type == "xls":

        sheets = pd.read_excel(
            io.BytesIO(
                data
            ),
            sheet_name=None,
            header=None,
            engine="xlrd",
        )

        return list(
            sheets.values()
        )


    raise RuntimeError(
        "Unsupported Codal file."
    )


# ============================================================
# NORMALIZE TABLE
# ============================================================

def normalize_table(
    df
):

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

            for column in (
                df.columns
            ):

                value = normalize_text(
                    column[level]
                )

                if (
                    value.startswith(
                        "Unnamed:"
                    )
                    or
                    value.lower()
                    == "nan"
                ):

                    value = ""

                row.append(
                    value
                )

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
                normalize_text(
                    value
                )
                for value
                in row
            ]
        )


    if not (
        header_rows
        or body_rows
    ):

        return pd.DataFrame()


    result = pd.DataFrame(
        header_rows
        + body_rows
    )


    # remove fully empty rows
    result = result[
        ~result.apply(
            lambda row:
                all(
                    normalize_text(v)
                    == ""
                    for v in row
                ),
            axis=1,
        )
    ]


    # remove fully empty columns
    result = result.loc[
        :,
        ~result.apply(
            lambda col:
                all(
                    normalize_text(v)
                    == ""
                    for v in col
                ),
            axis=0,
        )
    ]


    result = (
        result.reset_index(
            drop=True
        )
    )

    result.columns = range(
        result.shape[1]
    )


    return result


# ============================================================
# TABLE DETECTION
# ============================================================

def table_text(
    df
):

    return " ".join(
        normalize_text(
            value
        )
        for value
        in df.fillna("")
        .values
        .flatten()
    )


def production_score(
    df
):

    text = table_text(
        df
    )

    score = 0

    for keyword, points in [

        ("مقدار تولید", 5),
        ("مقدار فروش", 5),
        ("نرخ فروش", 3),
        ("مبلغ فروش", 3),
        ("سیمان", 2),
        ("کلینکر", 2),

    ]:

        if keyword in text:

            score += points


    return score


def energy_score(
    df
):

    text = table_text(
        df
    )

    score = 0

    for keyword, points in [

        ("نوع انرژی", 6),
        ("مقدار مصرف", 5),
        ("واحد اندازه گیری", 3),
        ("برق", 1),
        ("گاز", 1),
        ("آب", 1),

    ]:

        if keyword in text:

            score += points


    return score


def find_best_table(
    tables,
    scorer,
):

    best = None

    best_score = -1


    for table in tables:

        df = normalize_table(
            table
        )

        if df.empty:
            continue

        score = scorer(
            df
        )

        if score > best_score:

            best = df
            best_score = score


    return (
        best,
        best_score,
    )


# ============================================================
# HEADER SEARCH
# ============================================================

def find_header_column(
    df,
    condition,
):

    if df is None:
        return None

    rows = (
        df.fillna("")
        .values
        .tolist()
    )


    for row in rows[:15]:

        for index, value in enumerate(
            row
        ):

            text = normalize_text(
                value
            )

            if condition(
                text
            ):

                return index


    return None


# ============================================================
# CURRENT MONTH
# ============================================================

def find_current_month_block(
    df,
    period,
):

    period = (
        normalize_date(
            period
        )
    )


    result = (
        find_header_column(
            df,
            lambda text:
                (
                    period in text
                    and (
                        "یک ماهه"
                        in text
                        or
                        "1 ماهه"
                        in text
                    )
                )
        )
    )


    # --------------------------------------------------------
    # FARVARDIN FALLBACK
    # --------------------------------------------------------

    if (
        result is None
        and "/01/"
        in period
    ):

        result = (
            find_header_column(
                df,
                lambda text:
                    (
                        "دوره یک ماهه"
                        in text
                        or
                        "دوره 1 ماهه"
                        in text
                    )
            )
        )


    return result


# ============================================================
# YTD ORIGINAL / CORRECTED
# ============================================================

def find_original_ytd_block(
    df,
    previous_period,
):

    previous_period = (
        normalize_date(
            previous_period
        )
    )


    return find_header_column(
        df,
        lambda text:
            (
                previous_period
                in text
                and
                "از ابتدای سال"
                in text
                and
                "اصلاح شده"
                not in text
            )
    )


def find_corrected_ytd_block(
    df,
    previous_period,
):

    previous_period = (
        normalize_date(
            previous_period
        )
    )


    return find_header_column(
        df,
        lambda text:
            (
                previous_period
                in text
                and
                "اصلاح شده"
                in text
            )
    )


# ============================================================
# PRODUCTION EXTRACTOR
# ============================================================

def empty_production():

    return {
        ("داخلی", "سیمان"):
            [None, None, None],

        ("داخلی", "کلینکر"):
            [None, None, None],

        ("خارجی", "سیمان"):
            [None, None, None],

        ("خارجی", "کلینکر"):
            [None, None, None],
    }


def extract_production_block(
    df,
    start_col,
):

    result = empty_production()

    if (
        df is None
        or start_col is None
    ):

        return result


    current_market = None


    rows = (
        df.fillna("")
        .values
        .tolist()
    )


    for row in rows:

        row_text = " ".join(
            normalize_text(
                value
            )
            for value in row
        )


        # ----------------------------------------------------
        # DOMESTIC SECTION
        # ----------------------------------------------------

        if (
            "فروش داخلی"
            in row_text
        ):

            current_market = (
                "داخلی"
            )

            continue


        # ----------------------------------------------------
        # EXPORT SECTION
        # ----------------------------------------------------

        if (
            "فروش صادراتی"
            in row_text
            or
            "فروش خارجی"
            in row_text
        ):

            current_market = (
                "خارجی"
            )

            continue


        if current_market is None:

            continue


        left_text = " ".join(
            normalize_text(
                value
            )
            for value
            in row[:8]
        )


        # skip totals
        if (
            left_text.startswith(
                "جمع"
            )
            or
            "جمع فروش داخلی"
            in left_text
            or
            "جمع فروش صادراتی"
            in left_text
        ):

            continue


        # ----------------------------------------------------
        # PRODUCT FAMILY
        # ----------------------------------------------------

        if (
            "کلینکر"
            in left_text
        ):

            product = (
                "کلینکر"
            )

        elif (
            "سیمان"
            in left_text
        ):

            product = (
                "سیمان"
            )

        else:

            continue


        production = (
            to_number(
                safe_get(
                    row,
                    start_col,
                )
            )
        )


        sales = (
            to_number(
                safe_get(
                    row,
                    start_col + 1,
                )
            )
        )


        amount = (
            to_number(
                safe_get(
                    row,
                    start_col + 3,
                )
            )
        )


        if (
            production is None
            and
            sales is None
            and
            amount is None
        ):

            continue


        key = (
            current_market,
            product,
        )


        values = [
            production,
            sales,
            amount,
        ]


        for i, value in enumerate(
            values
        ):

            if value is None:
                continue

            if (
                result[key][i]
                is None
            ):

                result[key][i] = 0

            result[key][i] += (
                value
            )


    return result


# ============================================================
# ENERGY EXTRACTOR
# ============================================================

def empty_energy():

    return {
        item:
            [None, None]
        for item
        in ENERGY_MASTER
    }


def extract_energy_block(
    df,
    start_col,
):

    result = empty_energy()

    if (
        df is None
        or start_col is None
    ):

        return result


    rows = (
        df.fillna("")
        .values
        .tolist()
    )


    for row in rows:

        raw_name = (
            normalize_text(
                safe_get(
                    row,
                    2,
                )
            )
        )


        if not raw_name:

            continue


        energy_name = (
            canonical_energy(
                raw_name
            )
        )


        # Prevent garbage entering output
        if (
            energy_name
            not in ENERGY_MASTER
        ):

            continue


        quantity = (
            to_number(
                safe_get(
                    row,
                    start_col,
                )
            )
        )


        amount = (
            to_number(
                safe_get(
                    row,
                    start_col + 2,
                )
            )
        )


        if (
            quantity is None
            and
            amount is None
        ):

            continue


        current = result[
            energy_name
        ]


        if quantity is not None:

            if current[0] is None:

                current[0] = 0

            current[0] += (
                quantity
            )


        if amount is not None:

            if current[1] is None:

                current[1] = 0

            current[1] += (
                amount
            )


    return result


# ============================================================
# APPLY CORRECTION DELTA
# ============================================================

def apply_production_delta(
    target_month,
    original_ytd,
    corrected_ytd,
):

    for key in (
        target_month.keys()
    ):

        for i in range(
            3
        ):

            old = (
                original_ytd[
                    key
                ][i]
            )

            new = (
                corrected_ytd[
                    key
                ][i]
            )


            if (
                old is None
                or
                new is None
            ):

                continue


            delta = (
                new
                - old
            )


            if abs(
                delta
            ) < 0.000001:

                continue


            if (
                target_month[
                    key
                ][i]
                is None
            ):

                target_month[
                    key
                ][i] = 0


            target_month[
                key
            ][i] += delta


def apply_energy_delta(
    target_month,
    original_ytd,
    corrected_ytd,
):

    for key in (
        ENERGY_MASTER
    ):

        for i in range(
            2
        ):

            old = (
                original_ytd[
                    key
                ][i]
            )

            new = (
                corrected_ytd[
                    key
                ][i]
            )


            if (
                old is None
                or
                new is None
            ):

                continue


            delta = (
                new
                - old
            )


            if abs(
                delta
            ) < 0.000001:

                continue


            if (
                target_month[
                    key
                ][i]
                is None
            ):

                target_month[
                    key
                ][i] = 0


            target_month[
                key
            ][i] += delta


# ============================================================
# PROCESS ONE COMPANY
# ============================================================

def process_company(
    symbol,
    company,
):

    print()
    print(
        f"Processing: "
        f"{symbol} - {company}"
    )


    reports = search_reports(
        symbol
    )


    monthly_production = {}

    monthly_energy = {}


    for index, period in enumerate(
        PERIODS
    ):

        print(
            f"    Period: "
            f"{period}"
        )


        report = (
            find_report_from_list(
                reports,
                period,
            )
        )


        if report is None:

            print(
                "        Report not found"
            )

            continue


        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )


        print(
            f"        {title}"
        )


        raw_file = download_report(
            symbol,
            period,
            report,
        )


        tables = read_codal_file(
            raw_file
        )


        # ====================================================
        # PRODUCTION TABLE
        # ====================================================

        (
            production_table,
            production_score_value,
        ) = find_best_table(
            tables,
            production_score,
        )


        current_col = (
            find_current_month_block(
                production_table,
                period,
            )
        )


        monthly_production[
            period
        ] = (
            extract_production_block(
                production_table,
                current_col,
            )
        )


        # ====================================================
        # ENERGY TABLE
        # ====================================================

        (
            energy_table,
            energy_score_value,
        ) = find_best_table(
            tables,
            energy_score,
        )


        current_energy_col = (
            find_current_month_block(
                energy_table,
                period,
            )
        )


        monthly_energy[
            period
        ] = (
            extract_energy_block(
                energy_table,
                current_energy_col,
            )
        )


        # ====================================================
        # CORRECTION INSIDE CURRENT REPORT
        #
        # Example:
        #
        # We are processing Mordad.
        # Codal contains:
        #
        # YTD to Tir
        # YTD to Tir (corrected)
        #
        # Delta is applied to Tir.
        # ====================================================

        if index > 0:

            previous_period = (
                PERIODS[
                    index - 1
                ]
            )


            # ------------------------------------------------
            # PRODUCTION
            # ------------------------------------------------

            original_col = (
                find_original_ytd_block(
                    production_table,
                    previous_period,
                )
            )


            corrected_col = (
                find_corrected_ytd_block(
                    production_table,
                    previous_period,
                )
            )


            if (
                original_col
                is not None
                and
                corrected_col
                is not None
                and
                previous_period
                in monthly_production
            ):

                original_ytd = (
                    extract_production_block(
                        production_table,
                        original_col,
                    )
                )


                corrected_ytd = (
                    extract_production_block(
                        production_table,
                        corrected_col,
                    )
                )


                apply_production_delta(
                    monthly_production[
                        previous_period
                    ],
                    original_ytd,
                    corrected_ytd,
                )


                print(
                    f"        Production "
                    f"correction applied "
                    f"to {previous_period}"
                )


            # ------------------------------------------------
            # ENERGY
            # ------------------------------------------------

            original_energy_col = (
                find_original_ytd_block(
                    energy_table,
                    previous_period,
                )
            )


            corrected_energy_col = (
                find_corrected_ytd_block(
                    energy_table,
                    previous_period,
                )
            )


            if (
                original_energy_col
                is not None
                and
                corrected_energy_col
                is not None
                and
                previous_period
                in monthly_energy
            ):

                original_energy = (
                    extract_energy_block(
                        energy_table,
                        original_energy_col,
                    )
                )


                corrected_energy = (
                    extract_energy_block(
                        energy_table,
                        corrected_energy_col,
                    )
                )


                apply_energy_delta(
                    monthly_energy[
                        previous_period
                    ],
                    original_energy,
                    corrected_energy,
                )


                print(
                    f"        Energy "
                    f"correction applied "
                    f"to {previous_period}"
                )


    return {
        "symbol": symbol,
        "company": company,
        "production":
            monthly_production,
        "energy":
            monthly_energy,
    }


# ============================================================
# OUTPUT
# ============================================================

def create_output(
    results
):

    wb = Workbook()

    wb.remove(
        wb.active
    )


    # ========================================================
    # STYLE
    # ========================================================

    header_fill = (
        PatternFill(
            fill_type="solid",
            fgColor="D9EAF7",
        )
    )


    header_font = Font(
        bold=True
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


    # ========================================================
    # PRODUCTION SHEET
    # ========================================================

    ws = wb.create_sheet(
        "تولید و فروش"
    )

    ws.sheet_view.rightToLeft = (
        True
    )


    headers = [
        "نام شرکت",
        "نماد",
        "دوره",
        "نوع فروش",
        "نام کالا",
        "واحد",
        "مقدار تولید",
        "مقدار فروش",
        "نرخ فروش (ریال)",
        "مبلغ فروش (میلیون ریال)",
    ]


    for col, header in enumerate(
        headers,
        start=1,
    ):

        cell = ws.cell(
            1,
            col,
            header,
        )

        cell.fill = (
            header_fill
        )

        cell.font = (
            header_font
        )

        cell.border = border

        cell.alignment = center


    row = 2


    # --------------------------------------------------------
    # PERIOD FIRST
    #
    # Farvardin all companies,
    # Ordibehesht all companies,
    # ...
    # --------------------------------------------------------

    for period in PERIODS:

        for result in results:

            data = (
                result[
                    "production"
                ]
                .get(
                    period
                )
            )


            if data is None:

                continue


            for (
                market,
                product,
            ) in [

                ("داخلی", "سیمان"),
                ("داخلی", "کلینکر"),
                ("خارجی", "سیمان"),
                ("خارجی", "کلینکر"),

            ]:

                values = data[
                    (
                        market,
                        product,
                    )
                ]


                production = (
                    values[0]
                )

                sales = (
                    values[1]
                )

                amount = (
                    values[2]
                )


                ws.cell(
                    row,
                    1,
                    result[
                        "company"
                    ],
                )


                ws.cell(
                    row,
                    2,
                    result[
                        "symbol"
                    ],
                )


                ws.cell(
                    row,
                    3,
                    period,
                )


                ws.cell(
                    row,
                    4,
                    market,
                )


                ws.cell(
                    row,
                    5,
                    product,
                )


                ws.cell(
                    row,
                    6,
                    "تن",
                )


                ws.cell(
                    row,
                    7,
                    production,
                )


                ws.cell(
                    row,
                    8,
                    sales,
                )


                ws.cell(
                    row,
                    10,
                    amount,
                )


                if (
                    sales is not None
                    and
                    sales != 0
                    and
                    amount is not None
                ):

                    rate = (
                        amount
                        * 1_000_000
                        / sales
                    )


                    ws.cell(
                        row,
                        9,
                        rate,
                    )


                for col in range(
                    1,
                    11,
                ):

                    ws.cell(
                        row,
                        col,
                    ).border = border


                for col in [
                    7,
                    8,
                    9,
                    10,
                ]:

                    ws.cell(
                        row,
                        col,
                    ).number_format = (
                        NUMBER_FORMAT
                    )


                row += 1


    # ========================================================
    # ENERGY SHEET
    # ========================================================

    ws = wb.create_sheet(
        "انرژی"
    )


    ws.sheet_view.rightToLeft = (
        True
    )


    headers = [
        "نام شرکت",
        "نماد",
        "دوره",
        "نوع انرژی",
        "واحد",
        "مقدار مصرف",
        "نرخ",
        "مبلغ (میلیون ریال)",
    ]


    for col, header in enumerate(
        headers,
        start=1,
    ):

        cell = ws.cell(
            1,
            col,
            header,
        )

        cell.fill = (
            header_fill
        )

        cell.font = (
            header_font
        )

        cell.border = (
            border
        )

        cell.alignment = (
            center
        )


    row = 2


    for period in PERIODS:

        for result in results:

            data = (
                result[
                    "energy"
                ]
                .get(
                    period
                )
            )


            if data is None:

                continue


            for energy_name in (
                ENERGY_MASTER
            ):

                quantity = (
                    data[
                        energy_name
                    ][0]
                )

                amount = (
                    data[
                        energy_name
                    ][1]
                )


                ws.cell(
                    row,
                    1,
                    result[
                        "company"
                    ],
                )


                ws.cell(
                    row,
                    2,
                    result[
                        "symbol"
                    ],
                )


                ws.cell(
                    row,
                    3,
                    period,
                )


                ws.cell(
                    row,
                    4,
                    energy_name,
                )


                ws.cell(
                    row,
                    5,
                    ENERGY_UNITS[
                        energy_name
                    ],
                )


                ws.cell(
                    row,
                    6,
                    quantity,
                )


                ws.cell(
                    row,
                    8,
                    amount,
                )


                if (
                    quantity is not None
                    and
                    quantity != 0
                    and
                    amount is not None
                ):

                    rate = (
                        amount
                        * 1_000_000
                        / quantity
                    )


                    ws.cell(
                        row,
                        7,
                        rate,
                    )


                for col in range(
                    1,
                    9,
                ):

                    ws.cell(
                        row,
                        col,
                    ).border = (
                        border
                    )


                for col in [
                    6,
                    7,
                    8,
                ]:

                    ws.cell(
                        row,
                        col,
                    ).number_format = (
                        NUMBER_FORMAT
                    )


                row += 1


    # ========================================================
    # FORMATTING
    # ========================================================

    for ws in wb.worksheets:

        ws.freeze_panes = (
            "A2"
        )


        for column_cells in (
            ws.columns
        ):

            max_length = 0


            column_letter = (
                column_cells[
                    0
                ].column_letter
            )


            for cell in (
                column_cells
            ):

                if (
                    cell.value
                    is None
                ):

                    continue


                max_length = max(
                    max_length,
                    len(
                        str(
                            cell.value
                        )
                    ),
                )


            ws.column_dimensions[
                column_letter
            ].width = min(
                max_length + 3,
                35,
            )


    wb.save(
        OUTPUT_FILE
    )


# ============================================================
# MAIN
# ============================================================

def main():

    companies = (
        load_companies()
    )


    print()
    print(
        "=" * 70
    )

    print(
        "CODAL MONTHLY PERFORMANCE BUILDER"
    )

    print(
        f"Companies loaded: "
        f"{len(companies)}"
    )

    print(
        f"Periods: "
        f"{len(PERIODS)}"
    )

    print(
        "=" * 70
    )


    results = []


    for index, (
        symbol,
        company,
    ) in enumerate(
        companies,
        start=1,
    ):

        print()
        print(
            f"[{index}/"
            f"{len(companies)}]"
        )


        try:

            result = (
                process_company(
                    symbol,
                    company,
                )
            )


            results.append(
                result
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


        time.sleep(
            COMPANY_DELAY
        )


    # ========================================================
    # OUTPUT
    # ========================================================

    create_output(
        results
    )


    print()
    print(
        "=" * 70
    )

    print(
        f"Successful companies: "
        f"{len(results)}/"
        f"{len(companies)}"
    )

    print()

    print(
        "Output created:"
    )

    print(
        OUTPUT_FILE
    )

    print(
        "=" * 70
    )


if __name__ == "__main__":

    main()