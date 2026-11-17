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

CACHE_DIR = BASE_DIR / "Codal_Cache"
OUTPUT_DIR = BASE_DIR / "Output"

CACHE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

OUTPUT_FILE = (
    OUTPUT_DIR
    / "گزارش_عملکرد_ماهانه_تجمیعی.xlsx"
)


COMPANIES = [
    ("سکارون", "سیمان کارون"),
    ("سلار", "سیمان لارستان"),
    ("ستران", "سیمان تهران"),
    ("سآبیک", "سیمان آبیک"),
    ("سصوفی", "سیمان صوفیان"),
]


PERIODS = [
    "1405/01/31",
    "1405/02/31",
    "1405/03/31",
    "1405/04/31",
    "1405/05/31",
]


# Existing files are reused
REFRESH_CACHE = False


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
# NUMBER
# ============================================================

def to_number(value):

    text = normalize_text(value)

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
        return float(text)

    except Exception:
        return None


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


ENERGY_ALIASES = {

    # --------------------------
    # Water
    # --------------------------

    "آب": "آب",
    "آب مصرفی": "آب",
    "آب خریداری شده": "آب",
    "آب خریداری‌شده": "آب",
    "آب صنعتی": "آب",

    # --------------------------
    # Electricity
    # --------------------------

    "برق": "برق",
    "برق مصرفی": "برق",
    "برق خریداری شده": "برق",
    "برق خریداری‌شده": "برق",
    "برق خریداری": "برق",

    # --------------------------
    # Natural gas
    # --------------------------

    "گاز": "گاز طبیعی",
    "گاز طبیعی": "گاز طبیعی",
    "گاز مصرفی": "گاز طبیعی",
    "گاز خریداری شده": "گاز طبیعی",
    "گاز خریداری‌شده": "گاز طبیعی",

    # --------------------------
    # Mazut
    # --------------------------

    "مازوت": "مازوت",
    "نفت کوره": "مازوت",
    "نفت کوره (مازوت)": "مازوت",

    # --------------------------
    # Gasoil
    # --------------------------

    "گازوییل": "گازوئیل",
    "گازوئیل": "گازوئیل",
    "نفت گاز": "گازوئیل",
}


def canonical_energy(name):

    name = normalize_text(
        name
    )

    return ENERGY_ALIASES.get(
        name,
        name,
    )


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
        CODAL_SEARCH_URL,
        params=params,
        timeout=60,
    )

    response.raise_for_status()

    return response.json().get(
        "Letters",
        [],
    )


# ============================================================
# FIND LATEST VERSION OF EXACT PERIOD
# ============================================================

def find_report_from_list(
    reports,
    period,
):

    target_period = normalize_date(
        period
    )

    matches = []

    for report in reports:

        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )

        if target_period not in title:
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
    #
    # If original + correction exist,
    # latest publication wins.
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

def get_excel_url(report):

    excel_url = (
        report.get("ExcelUrl")
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
# DOWNLOAD
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

    if (
        not REFRESH_CACHE
        and raw_path.exists()
        and raw_path.stat().st_size > 500
    ):

        print(
            "        Using cached file"
        )

        return raw_path

    url = get_excel_url(
        report
    )

    if not url:

        raise RuntimeError(
            "Excel URL not found."
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
        f"        Downloaded "
        f"{len(response.content):,} bytes"
    )

    return raw_path


# ============================================================
# FILE READER
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

    for element in html_tables:

        table_html = etree.tostring(
            element,
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

            result.extend(
                parsed
            )

        except Exception:
            continue

    return result


def read_codal_file(path):

    file_type = detect_file_type(
        path
    )

    data = path.read_bytes()

    if file_type == "html":

        return read_html_tables_robust(
            decode_html(data)
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

    raise RuntimeError(
        "Unsupported Codal file."
    )


# ============================================================
# NORMALIZE TABLE
# ============================================================

def normalize_table(df):

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
                    or value.lower() == "nan"
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
                normalize_text(v)
                for v in row
            ]
        )

    result = pd.DataFrame(
        header_rows
        + body_rows
    )

    result = result.dropna(
        how="all"
    )

    result.columns = range(
        result.shape[1]
    )

    return result


# ============================================================
# TABLE DETECTION
# ============================================================

def table_text(df):

    return " ".join(
        normalize_text(v)
        for v in df.fillna("")
        .values
        .flatten()
    )


def production_score(df):

    text = table_text(
        df
    )

    score = 0

    for keyword, points in [
        ("مقدار تولید", 5),
        ("مقدار فروش", 5),
        ("مبلغ فروش", 3),
        ("سیمان", 2),
        ("کلینکر", 2),
    ]:

        if keyword in text:
            score += points

    return score


def energy_score(df):

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

        score = scorer(
            df
        )

        if score > best_score:

            best = df
            best_score = score

    return best


# ============================================================
# FIND HEADER BLOCK
# ============================================================

def find_header_column(
    df,
    text_condition,
):

    for row in (
        df.fillna("")
        .values
        .tolist()[:15]
    ):

        for index, value in enumerate(
            row
        ):

            text = normalize_text(
                value
            )

            if text_condition(text):

                return index

    return None


# ============================================================
# CURRENT MONTH BLOCK
# ============================================================

def find_current_month_block(
    df,
    period,
):

    period = normalize_date(
        period
    )

    return find_header_column(
        df,
        lambda text:
            period in text
            and (
                "یک ماهه" in text
                or "1 ماهه" in text
            )
    )


# ============================================================
# CORRECTED YTD BLOCK
# ============================================================

def find_corrected_ytd_block(
    df,
    previous_period,
):

    previous_period = normalize_date(
        previous_period
    )

    return find_header_column(
        df,
        lambda text:
            previous_period in text
            and "اصلاح شده" in text
    )


# ============================================================
# PRODUCTION BLOCK EXTRACTOR
# ============================================================

def extract_production_block(
    df,
    start_col,
):

    result = {
        ("داخلی", "سیمان"): [None, None, None],
        ("داخلی", "کلینکر"): [None, None, None],
        ("خارجی", "سیمان"): [None, None, None],
        ("خارجی", "کلینکر"): [None, None, None],
    }

    if start_col is None:
        return result

    market = None

    for row in (
        df.fillna("")
        .values
        .tolist()
    ):

        row_text = " ".join(
            normalize_text(v)
            for v in row
        )

        if "فروش داخلی" in row_text:

            market = "داخلی"
            continue

        if (
            "فروش صادراتی" in row_text
            or "فروش خارجی" in row_text
        ):

            market = "خارجی"
            continue

        if market is None:
            continue

        left = " ".join(
            normalize_text(v)
            for v in row[:8]
        )

        if left.startswith(
            "جمع"
        ):
            continue

        if "کلینکر" in left:

            product = "کلینکر"

        elif "سیمان" in left:

            product = "سیمان"

        else:

            continue

        production = to_number(
            safe_get(
                row,
                start_col,
            )
        )

        sales = to_number(
            safe_get(
                row,
                start_col + 1,
            )
        )

        amount = to_number(
            safe_get(
                row,
                start_col + 3,
            )
        )

        key = (
            market,
            product,
        )

        current = result[key]

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

            if current[i] is None:
                current[i] = 0

            current[i] += value

    return result


# ============================================================
# ENERGY BLOCK EXTRACTOR
# ============================================================

def extract_energy_block(
    df,
    start_col,
):

    result = {
        item: [None, None]
        for item in ENERGY_MASTER
    }

    if start_col is None:
        return result

    for row in (
        df.fillna("")
        .values
        .tolist()
    ):

        raw_name = normalize_text(
            safe_get(
                row,
                2,
            )
        )

        if not raw_name:
            continue

        energy = canonical_energy(
            raw_name
        )

        if energy not in ENERGY_MASTER:
            continue

        quantity = to_number(
            safe_get(
                row,
                start_col,
            )
        )

        amount = to_number(
            safe_get(
                row,
                start_col + 2,
            )
        )

        if (
            quantity is None
            and amount is None
        ):
            continue

        result[
            energy
        ] = [
            quantity,
            amount,
        ]

    return result


# ============================================================
# CORRECTION LOGIC
# ============================================================

def sum_previous_production(
    monthly_data,
    periods,
    end_index,
):

    totals = defaultdict(
        lambda: [0, 0, 0]
    )

    for period in periods[:end_index]:

        data = monthly_data.get(
            period,
            {}
        )

        for key, values in data.items():

            for i, value in enumerate(
                values
            ):

                if value is not None:
                    totals[key][i] += value

    return totals


def sum_previous_energy(
    monthly_data,
    periods,
    end_index,
):

    totals = defaultdict(
        lambda: [0, 0]
    )

    for period in periods[:end_index]:

        data = monthly_data.get(
            period,
            {}
        )

        for key, values in data.items():

            for i, value in enumerate(
                values
            ):

                if value is not None:
                    totals[key][i] += value

    return totals


def apply_production_correction(
    monthly_data,
    corrected_ytd,
    previous_period,
    period_index,
):

    previous_totals = (
        sum_previous_production(
            monthly_data,
            PERIODS,
            period_index,
        )
    )

    target = monthly_data[
        previous_period
    ]

    for key, corrected_values in (
        corrected_ytd.items()
    ):

        for i, corrected_value in enumerate(
            corrected_values
        ):

            if corrected_value is None:
                continue

            old_total = (
                previous_totals[
                    key
                ][i]
            )

            delta = (
                corrected_value
                - old_total
            )

            if abs(delta) < 0.000001:
                continue

            if target[key][i] is None:
                target[key][i] = 0

            target[key][i] += delta


def apply_energy_correction(
    monthly_data,
    corrected_ytd,
    previous_period,
    period_index,
):

    previous_totals = (
        sum_previous_energy(
            monthly_data,
            PERIODS,
            period_index,
        )
    )

    target = monthly_data[
        previous_period
    ]

    for key, corrected_values in (
        corrected_ytd.items()
    ):

        for i, corrected_value in enumerate(
            corrected_values
        ):

            if corrected_value is None:
                continue

            old_total = (
                previous_totals[
                    key
                ][i]
            )

            delta = (
                corrected_value
                - old_total
            )

            if abs(delta) < 0.000001:
                continue

            if target[key][i] is None:
                target[key][i] = 0

            target[key][i] += delta


# ============================================================
# PROCESS COMPANY
# ============================================================

def process_company(
    symbol,
    company,
):

    print()
    print(
        f"Processing: {symbol} - {company}"
    )

    reports = search_reports(
        symbol
    )

    monthly_production = {}
    monthly_energy = {}


    for index, period in enumerate(
        PERIODS
    ):

        report = (
            find_report_from_list(
                reports,
                period,
            )
        )

        if report is None:

            print(
                f"    {period}: report not found"
            )

            continue


        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )

        print(
            f"    {period}: {title}"
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
        # PRODUCTION
        # ====================================================

        production_table = (
            find_best_table(
                tables,
                production_score,
            )
        )

        current_col = (
            find_current_month_block(
                production_table,
                period,
            )
        )

        current_production = (
            extract_production_block(
                production_table,
                current_col,
            )
        )

        monthly_production[
            period
        ] = current_production


        # ====================================================
        # ENERGY
        # ====================================================

        energy_table = (
            find_best_table(
                tables,
                energy_score,
            )
        )

        current_energy_col = (
            find_current_month_block(
                energy_table,
                period,
            )
        )

        current_energy = (
            extract_energy_block(
                energy_table,
                current_energy_col,
            )
        )

        monthly_energy[
            period
        ] = current_energy


        # ====================================================
        # INTERNAL CORRECTION OF PREVIOUS MONTH
        # ====================================================

        if index > 0:

            previous_period = (
                PERIODS[
                    index - 1
                ]
            )


            # ------------------------------------------------
            # PRODUCTION CORRECTION
            # ------------------------------------------------

            corrected_col = (
                find_corrected_ytd_block(
                    production_table,
                    previous_period,
                )
            )

            if corrected_col is not None:

                corrected_ytd = (
                    extract_production_block(
                        production_table,
                        corrected_col,
                    )
                )

                apply_production_correction(
                    monthly_production,
                    corrected_ytd,
                    previous_period,
                    index,
                )

                print(
                    f"        Production correction "
                    f"applied to {previous_period}"
                )


            # ------------------------------------------------
            # ENERGY CORRECTION
            # ------------------------------------------------

            corrected_energy_col = (
                find_corrected_ytd_block(
                    energy_table,
                    previous_period,
                )
            )

            if corrected_energy_col is not None:

                corrected_energy = (
                    extract_energy_block(
                        energy_table,
                        corrected_energy_col,
                    )
                )

                apply_energy_correction(
                    monthly_energy,
                    corrected_energy,
                    previous_period,
                    index,
                )

                print(
                    f"        Energy correction "
                    f"applied to {previous_period}"
                )


    return {
        "company": company,
        "symbol": symbol,
        "production": monthly_production,
        "energy": monthly_energy,
    }


# ============================================================
# EXCEL OUTPUT
# ============================================================

def create_output(
    results,
):

    wb = Workbook()

    default_sheet = wb.active

    wb.remove(
        default_sheet
    )


    # ========================================================
    # STYLE
    # ========================================================

    header_fill = PatternFill(
        fill_type="solid",
        fgColor="D9EAF7",
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
    )


    # ========================================================
    # PRODUCTION SHEET
    # ========================================================

    ws = wb.create_sheet(
        "تولید و فروش"
    )

    ws.sheet_view.rightToLeft = True


    headers = [
        "نام شرکت",
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

        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center


    row = 2


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


            for market, product in [
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

                production = values[0]
                sales = values[1]
                amount = values[2]


                ws.cell(
                    row,
                    1,
                    result["company"],
                )

                ws.cell(
                    row,
                    2,
                    period,
                )

                ws.cell(
                    row,
                    3,
                    market,
                )

                ws.cell(
                    row,
                    4,
                    product,
                )

                ws.cell(
                    row,
                    5,
                    "تن",
                )

                ws.cell(
                    row,
                    6,
                    production,
                )

                ws.cell(
                    row,
                    7,
                    sales,
                )

                ws.cell(
                    row,
                    9,
                    amount,
                )


                if (
                    sales is not None
                    and sales != 0
                    and amount is not None
                ):

                    rate = (
                        amount
                        * 1_000_000
                        / sales
                    )

                    ws.cell(
                        row,
                        8,
                        rate,
                    )


                for col in range(
                    1,
                    10,
                ):

                    ws.cell(
                        row,
                        col,
                    ).border = border


                for col in [
                    6,
                    7,
                    8,
                    9,
                ]:

                    ws.cell(
                        row,
                        col,
                    ).number_format = NUMBER_FORMAT


                row += 1


    # ========================================================
    # ENERGY SHEET
    # ========================================================

    ws = wb.create_sheet(
        "انرژی"
    )

    ws.sheet_view.rightToLeft = True


    headers = [
        "نام شرکت",
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

        cell.fill = header_fill
        cell.font = header_font
        cell.border = border
        cell.alignment = center


    row = 2


    ENERGY_UNITS = {
        "آب": "مترمکعب",
        "برق": "مگاوات ساعت",
        "گاز طبیعی": "مترمکعب",
        "مازوت": "لیتر",
        "گازوئیل": "لیتر",
    }


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

                quantity, amount = (
                    data[
                        energy_name
                    ]
                )


                ws.cell(
                    row,
                    1,
                    result["company"],
                )

                ws.cell(
                    row,
                    2,
                    period,
                )

                ws.cell(
                    row,
                    3,
                    energy_name,
                )

                ws.cell(
                    row,
                    4,
                    ENERGY_UNITS[
                        energy_name
                    ],
                )

                ws.cell(
                    row,
                    5,
                    quantity,
                )

                ws.cell(
                    row,
                    7,
                    amount,
                )


                if (
                    quantity is not None
                    and quantity != 0
                    and amount is not None
                ):

                    rate = (
                        amount
                        * 1_000_000
                        / quantity
                    )

                    ws.cell(
                        row,
                        6,
                        rate,
                    )


                for col in range(
                    1,
                    8,
                ):

                    ws.cell(
                        row,
                        col,
                    ).border = border


                for col in [
                    5,
                    6,
                    7,
                ]:

                    ws.cell(
                        row,
                        col,
                    ).number_format = NUMBER_FORMAT


                row += 1


    # ========================================================
    # COLUMN WIDTHS
    # ========================================================

    for ws in wb.worksheets:

        for column_cells in (
            ws.columns
        ):

            max_length = 0

            column_letter = (
                column_cells[0]
                .column_letter
            )

            for cell in column_cells:

                if cell.value is None:
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

        ws.freeze_panes = "A2"


    wb.save(
        OUTPUT_FILE
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print()
    print(
        "=" * 70
    )

    print(
        "CODAL MONTHLY PERFORMANCE BUILDER"
    )

    print(
        "=" * 70
    )


    results = []


    for index, (
        symbol,
        company,
    ) in enumerate(
        COMPANIES,
        start=1,
    ):

        print()
        print(
            f"[{index}/{len(COMPANIES)}]"
        )

        try:

            result = process_company(
                symbol,
                company,
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


        time.sleep(2)


    create_output(
        results
    )


    print()
    print(
        "=" * 70
    )

    print(
        f"Output created:"
    )

    print(
        OUTPUT_FILE.name
    )

    print(
        "=" * 70
    )


if __name__ == "__main__":
    main()