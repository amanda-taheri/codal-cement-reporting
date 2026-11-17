from __future__ import annotations

import io
import re
import shutil
import time
from copy import copy
from pathlib import Path
from typing import Any

import pandas as pd
import requests
from openpyxl import load_workbook


# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

TEMPLATE_FILE = BASE_DIR / "template_v1.xlsx"

CACHE_DIR = BASE_DIR / "Codal_Cache"
OUTPUT_DIR = BASE_DIR / "Output"
DEBUG_DIR = BASE_DIR / "Debug"

CACHE_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)
DEBUG_DIR.mkdir(exist_ok=True)


# True = download again
# False = use existing cached files
REFRESH_CACHE = True


PERIODS = [
    "1405/04/31",
    "1405/05/31",
]


TEST_COMPANIES = [
    {
        "symbol": "سکارون",
        "company": "سیمان کارون",
    },
    {
        "symbol": "سلار",
        "company": "سیمان لارستان",
    },
    {
        "symbol": "ستران",
        "company": "سیمان تهران",
    },
    {
        "symbol": "سآبیک",
        "company": "سیمان آبیک",
    },
    {
        "symbol": "سصوفی",
        "company": "سیمان صوفیان",
    },
]


SHEET_PRODUCTION = "تولید و فروش"
SHEET_MATERIALS = "خرید مواد اولیه"
SHEET_ENERGY = "انرژی"


CODAL_SEARCH_URL = (
    "https://search.codal.ir/api/search/v2/q"
)


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


def normalize_text(value: Any) -> str:

    if value is None:
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


def normalize_date(value: str) -> str:

    return (
        normalize_text(value)
        .replace("-", "/")
    )


def to_number(value):

    if value is None:
        return None

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
# MATERIAL NORMALIZATION
# ============================================================

MATERIAL_ALIASES = {

    "سنگ آهک": "سنگ آهک",
    "سنگ آهک انفجاری": "سنگ آهک",

    "کائولن": "کائولن",
    "سنگ کائولن": "کائولن",

    "بوکسیت": "بوکسیت",
    "سنگ بوکسیت": "بوکسیت",

    "گچ": "سنگ گچ",
    "سنگ گچ": "سنگ گچ",

    "پوزولان": "پوزولان",
    "سنگ پوزولان": "پوزولان",
    "سنگ تراس پوزولان": "پوزولان",

    "پاکت": "پاکت سیمان",
    "کیسه": "پاکت سیمان",
    "پاکت سیمان": "پاکت سیمان",

    "حاک رس": "خاک رس",
    "خاک رس": "خاک رس",

    "خاک": "خاک",
    "خاک صنعتی": "خاک صنعتی",

    "سنگ آهن": "سنگ آهن",

    "سیلیس": "سیلیس",
    "سنگ سیلیس": "سیلیس",

    "آلوویوم": "آلوویوم",

    # intentionally separate:
    "آلوویوم و سنگ آهک":
        "آلوویوم و سنگ آهک",

    "مارل": "مارل",

    "سرباره آهن": "سرباره آهن",

    "دولومیت": "دولومیت",

    "سنگ دانه رنگی":
        "سنگ دانه رنگی",

    "کلینکر خریداری شده":
        "کلینکر خریداری‌شده",

    "ادتیو سیمانهای حفاری":
        "ادتیو سیمان‌های حفاری",

    "شیل": "شیل",
}


def canonical_material(name):

    name = normalize_text(name)

    return MATERIAL_ALIASES.get(
        name,
        name,
    )


# ============================================================
# ENERGY NORMALIZATION
# ============================================================

ENERGY_ALIASES = {

    "گاز": "گاز طبیعی",
    "گاز مصرفی": "گاز طبیعی",
    "گاز طبیعی": "گاز طبیعی",

    "برق": "برق",

    "مازوت": "مازوت",
    "نفت کوره": "مازوت",

    "گازوییل": "گازوئیل",
    "گازوئیل": "گازوئیل",
}


def canonical_energy(name):

    name = normalize_text(name)

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
        timeout=45,
    )

    response.raise_for_status()

    data = response.json()

    return data.get(
        "Letters",
        [],
    )


def find_report(
    symbol,
    period,
):

    period = normalize_date(period)

    reports = search_reports(symbol)

    matches = []

    for report in reports:

        title = normalize_text(
            report.get(
                "Title",
                "",
            )
        )

        if period not in title:
            continue

        if "گزارش فعالیت ماهانه" not in title:
            continue

        matches.append(report)

    if not matches:
        return None

    # Latest published report wins
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
# DOWNLOAD
# ============================================================

def get_excel_url(report):

    url = report.get(
        "ExcelUrl",
        "",
    )

    if url:

        if url.startswith("/"):
            url = (
                "https://excel.codal.ir"
                + url
            )

        return url

    return ""


def download_report(
    symbol,
    period,
    report,
):

    period_folder = (
        CACHE_DIR
        / period.replace("/", "-")
    )

    period_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    raw_file = (
        period_folder
        / f"{symbol}_raw.xls"
    )

    if (
        raw_file.exists()
        and raw_file.stat().st_size > 1000
        and not REFRESH_CACHE
    ):

        print(
            "    Using cached raw file"
        )

        return raw_file

    url = get_excel_url(report)

    if not url:

        raise RuntimeError(
            "Excel URL not found"
        )

    print(
        "    Downloading Codal file..."
    )

    response = session.get(
        url,
        timeout=90,
    )

    response.raise_for_status()

    raw_file.write_bytes(
        response.content
    )

    print(
        f"    Raw file saved: "
        f"{raw_file.name}"
    )

    return raw_file


# ============================================================
# READ CODAL XLS
# ============================================================

def read_codal_file(file_path):

    data = file_path.read_bytes()

    # --------------------------------------------------------
    # real XLS
    # --------------------------------------------------------

    if data.startswith(
        b"\xD0\xCF\x11\xE0"
    ):

        print(
            "    File type: real XLS"
        )

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
    # XLSX
    # --------------------------------------------------------

    if data.startswith(b"PK"):

        print(
            "    File type: XLSX"
        )

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

    print(
        "    File type: HTML disguised as XLS"
    )

    decoded = None

    for encoding in [
        "utf-8-sig",
        "utf-8",
        "windows-1256",
        "cp1256",
    ]:

        try:

            decoded = data.decode(
                encoding
            )

            break

        except Exception:

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

    print(
        f"    HTML tables detected: "
        f"{len(tables)}"
    )

    return tables


# ============================================================
# CONVERT RAW CODAL FILE TO XLSX
# ============================================================

def convert_tables_to_xlsx(
    tables,
    output_file,
):

    with pd.ExcelWriter(
        output_file,
        engine="openpyxl",
    ) as writer:

        for i, df in enumerate(
            tables,
            start=1,
        ):

            temp = df.copy()

            # flatten MultiIndex
            if isinstance(
                temp.columns,
                pd.MultiIndex,
            ):

                temp.columns = [
                    " | ".join(
                        [
                            normalize_text(x)
                            for x in col
                            if normalize_text(x)
                            and not normalize_text(x)
                            .startswith("Unnamed")
                        ]
                    )
                    or f"Column_{n}"
                    for n, col in enumerate(
                        temp.columns,
                        start=1,
                    )
                ]

            temp.to_excel(
                writer,
                sheet_name=f"Table_{i}",
                index=False,
            )

    print(
        f"    Converted XLSX saved: "
        f"{output_file.name}"
    )


# ============================================================
# CLEAN DATAFRAME
# ============================================================

def clean_dataframe(df):

    result = df.copy()

    for col in result.columns:

        result[col] = result[col].map(
            lambda x:
                ""
                if pd.isna(x)
                else normalize_text(x)
        )

    return result


def dataframe_text(df):

    values = (
        df.fillna("")
        .astype(str)
        .values
        .flatten()
    )

    return " ".join(
        normalize_text(x)
        for x in values
        if normalize_text(x)
    )


# ============================================================
# DEBUG
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

    output_file = (
        folder
        / f"{symbol}_tables.txt"
    )

    with open(
        output_file,
        "w",
        encoding="utf-8",
    ) as f:

        for i, table in enumerate(
            tables,
            start=1,
        ):

            df = clean_dataframe(table)

            f.write(
                "\n"
                + "=" * 80
                + "\n"
            )

            f.write(
                f"TABLE {i}\n"
            )

            f.write(
                f"Shape: {df.shape}\n"
            )

            f.write(
                "=" * 80
                + "\n\n"
            )

            f.write(
                df.to_string(
                    index=False,
                    header=False,
                )
            )

            f.write("\n\n")

    print(
        f"    Debug file saved: "
        f"{output_file.name}"
    )


# ============================================================
# TABLE DETECTION
# ============================================================

def find_production_table(tables):

    best = None
    best_score = 0

    for i, table in enumerate(
        tables,
        start=1,
    ):

        df = clean_dataframe(table)

        text = dataframe_text(df)

        score = 0

        if "مقدار تولید" in text:
            score += 3

        if "مقدار فروش" in text:
            score += 3

        if "نرخ فروش" in text:
            score += 2

        if "مبلغ فروش" in text:
            score += 2

        if "سیمان" in text:
            score += 2

        if "کلینکر" in text:
            score += 1

        if score > best_score:

            best_score = score
            best = df

    if (
        best is not None
        and best_score >= 5
    ):

        print(
            f"    Production table: OK "
            f"(score={best_score})"
        )

        return best

    print(
        f"    Production table: NOT FOUND "
        f"(best score={best_score})"
    )

    return None


def find_material_table(tables):

    best = None
    best_score = 0

    for table in tables:

        df = clean_dataframe(table)

        text = dataframe_text(df)

        score = 0

        if "خرید دوره یک ماهه" in text:
            score += 4

        if "مواد اولیه" in text:
            score += 3

        if "مبلغ" in text:
            score += 1

        if "نرخ" in text:
            score += 1

        if score > best_score:

            best_score = score
            best = df

    if (
        best is not None
        and best_score >= 5
    ):

        print(
            f"    Material table: OK "
            f"(score={best_score})"
        )

        return best

    print(
        f"    Material table: NOT FOUND "
        f"(best score={best_score})"
    )

    return None


def find_energy_table(tables):

    best = None
    best_score = 0

    for table in tables:

        df = clean_dataframe(table)

        text = dataframe_text(df)

        score = 0

        if "نوع انرژی" in text:
            score += 4

        if "مقدار مصرف" in text:
            score += 3

        if "واحد اندازه گیری" in text:
            score += 2

        if "برق" in text:
            score += 1

        if (
            "گاز" in text
            or "مازوت" in text
        ):
            score += 1

        if score > best_score:

            best_score = score
            best = df

    if (
        best is not None
        and best_score >= 5
    ):

        print(
            f"    Energy table: OK "
            f"(score={best_score})"
        )

        return best

    print(
        f"    Energy table: NOT FOUND "
        f"(best score={best_score})"
    )

    return None


# ============================================================
# FIND MONTH COLUMN BLOCK
# ============================================================

def find_month_block(
    df,
    period,
):

    target = normalize_date(
        period
    )

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    for row in values[:15]:

        for col_index, value in enumerate(
            row
        ):

            text = normalize_text(
                value
            )

            if target not in text:
                continue

            if (
                "یک ماهه" in text
                or "1 ماهه" in text
                or "خرید دوره" in text
            ):

                return col_index

    return None


# ============================================================
# PRODUCTION EXTRACTION
# ============================================================

def extract_production(
    df,
    period,
):

    start_col = find_month_block(
        df,
        period,
    )

    if start_col is None:

        print(
            "    ERROR: Monthly production "
            "column not found"
        )

        return {}

    print(
        f"    Production month column: "
        f"{start_col}"
    )

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    result = {}

    combinations = [
        ("داخلی", "سیمان"),
        ("داخلی", "کلینکر"),
        ("خارجی", "سیمان"),
        ("خارجی", "کلینکر"),
    ]

    for key in combinations:

        result[key] = {
            "exists": False,
            "production": 0,
            "sales": 0,
            "amount": 0,
        }

    for row in values:

        row_text = " ".join(
            normalize_text(x)
            for x in row[:8]
        )

        if (
            "سیمان" not in row_text
            and "کلینکر" not in row_text
        ):
            continue

        if any(
            x in row_text
            for x in [
                "جمع",
                "تخفیف",
                "برگشت",
            ]
        ):
            continue

        if "کلینکر" in row_text:

            product = "کلینکر"

        else:

            product = "سیمان"

        if (
            "صادرات" in row_text
            or "خارجی" in row_text
        ):

            market = "خارجی"

        else:

            market = "داخلی"

        production = (
            to_number(
                row[start_col]
            )
            if start_col < len(row)
            else None
        )

        sales = (
            to_number(
                row[start_col + 1]
            )
            if start_col + 1 < len(row)
            else None
        )

        amount = (
            to_number(
                row[start_col + 3]
            )
            if start_col + 3 < len(row)
            else None
        )

        key = (
            market,
            product,
        )

        item = result[key]

        item["exists"] = True

        if production is not None:
            item["production"] += production

        if sales is not None:
            item["sales"] += sales

        if amount is not None:
            item["amount"] += amount

    return result


# ============================================================
# MATERIAL EXTRACTION
# ============================================================

def extract_materials(
    df,
    period,
):

    start_col = find_month_block(
        df,
        period,
    )

    if start_col is None:

        print(
            "    ERROR: Monthly material "
            "column not found"
        )

        return {}

    print(
        f"    Material month column: "
        f"{start_col}"
    )

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    output = {}

    purchase_type = None

    for row in values:

        name = normalize_text(
            row[0]
            if len(row) > 0
            else ""
        )

        unit = normalize_text(
            row[1]
            if len(row) > 1
            else ""
        )

        if "مواد اولیه داخلی" in name:

            purchase_type = "داخلی"
            continue

        if "مواد اولیه وارداتی" in name:

            purchase_type = "وارداتی"
            continue

        if not purchase_type:
            continue

        if not name:
            continue

        if name.startswith("جمع"):
            continue

        if not unit:
            continue

        canonical = canonical_material(
            name
        )

        quantity = (
            to_number(
                row[start_col]
            )
            if start_col < len(row)
            else None
        )

        amount = (
            to_number(
                row[start_col + 2]
            )
            if start_col + 2 < len(row)
            else None
        )

        key = (
            purchase_type,
            canonical,
            unit,
        )

        if key not in output:

            output[key] = {
                "quantity": None,
                "amount": None,
            }

        if quantity is not None:

            if output[key]["quantity"] is None:
                output[key]["quantity"] = 0

            output[key]["quantity"] += quantity

        if amount is not None:

            if output[key]["amount"] is None:
                output[key]["amount"] = 0

            output[key]["amount"] += amount

    return output


# ============================================================
# ENERGY EXTRACTION
# ============================================================

def extract_energy(
    df,
    period,
):

    start_col = find_month_block(
        df,
        period,
    )

    if start_col is None:

        print(
            "    ERROR: Monthly energy "
            "column not found"
        )

        return {}

    print(
        f"    Energy month column: "
        f"{start_col}"
    )

    values = (
        df.fillna("")
        .values
        .tolist()
    )

    output = {}

    for row in values:

        if len(row) < 4:
            continue

        name = normalize_text(
            row[2]
        )

        unit = normalize_text(
            row[3]
        )

        if not name:
            continue

        if not unit:
            continue

        if name in {
            "نوع انرژی",
        }:
            continue

        canonical = canonical_energy(
            name
        )

        quantity = (
            to_number(
                row[start_col]
            )
            if start_col < len(row)
            else None
        )

        amount = (
            to_number(
                row[start_col + 2]
            )
            if start_col + 2 < len(row)
            else None
        )

        key = (
            canonical,
            unit,
        )

        output[key] = {
            "quantity": quantity,
            "amount": amount,
        }

    return output


# ============================================================
# PROCESS ONE COMPANY
# ============================================================

def process_company(
    symbol,
    company,
    period,
):

    print(
        f"\nProcessing {symbol}"
    )

    report = find_report(
        symbol,
        period,
    )

    if report is None:

        print(
            "    ERROR: Report not found"
        )

        return None

    print(
        "    Report found"
    )

    print(
        f"    Title: "
        f"{normalize_text(report.get('Title'))}"
    )

    raw_file = download_report(
        symbol,
        period,
        report,
    )

    tables = read_codal_file(
        raw_file
    )

    # --------------------------------------------------------
    # Save readable XLSX for Mac
    # --------------------------------------------------------

    xlsx_file = (
        raw_file.parent
        / f"{symbol}.xlsx"
    )

    convert_tables_to_xlsx(
        tables,
        xlsx_file,
    )

    save_debug(
        symbol,
        period,
        tables,
    )

    production_table = (
        find_production_table(
            tables
        )
    )

    material_table = (
        find_material_table(
            tables
        )
    )

    energy_table = (
        find_energy_table(
            tables
        )
    )

    if production_table is None:

        return None

    production = extract_production(
        production_table,
        period,
    )

    materials = (
        extract_materials(
            material_table,
            period,
        )
        if material_table is not None
        else {}
    )

    energy = (
        extract_energy(
            energy_table,
            period,
        )
        if energy_table is not None
        else {}
    )

    print(
        f"    Production groups: "
        f"{sum(1 for x in production.values() if x['exists'])}"
    )

    print(
        f"    Material rows: "
        f"{len(materials)}"
    )

    print(
        f"    Energy rows: "
        f"{len(energy)}"
    )

    return {
        "symbol": symbol,
        "company": company,
        "period": period,
        "production": production,
        "materials": materials,
        "energy": energy,
    }


# ============================================================
# COPY EXCEL STYLE
# ============================================================

def copy_row_style(
    ws,
    source_row,
    target_row,
    max_col,
):

    for col in range(
        1,
        max_col + 1,
    ):

        src = ws.cell(
            source_row,
            col,
        )

        dst = ws.cell(
            target_row,
            col,
        )

        if src.has_style:
            dst._style = copy(
                src._style
            )

        dst.font = copy(
            src.font
        )

        dst.fill = copy(
            src.fill
        )

        dst.border = copy(
            src.border
        )

        dst.alignment = copy(
            src.alignment
        )

        dst.number_format = (
            src.number_format
        )


# ============================================================
# CREATE OUTPUT
# ============================================================

def create_output(
    period,
    results,
    material_master,
    energy_master,
):

    filename = (
        OUTPUT_DIR
        / (
            "گزارش_عملکرد_"
            + period.replace("/", "-")
            + ".xlsx"
        )
    )

    shutil.copy2(
        TEMPLATE_FILE,
        filename,
    )

    wb = load_workbook(
        filename
    )

    # ========================================================
    # PRODUCTION
    # ========================================================

    ws = wb[
        SHEET_PRODUCTION
    ]

    if ws.max_row > 1:

        ws.delete_rows(
            2,
            ws.max_row - 1,
        )

    row = 2

    combinations = [
        ("داخلی", "سیمان"),
        ("داخلی", "کلینکر"),
        ("خارجی", "سیمان"),
        ("خارجی", "کلینکر"),
    ]

    for result in results:

        first = True

        for market, product in combinations:

            item = (
                result["production"]
                .get(
                    (market, product)
                )
            )

            ws.cell(
                row,
                1,
                result["company"]
                if first
                else None,
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

            if (
                item
                and item["exists"]
            ):

                ws.cell(
                    row,
                    6,
                    item["production"],
                )

                ws.cell(
                    row,
                    7,
                    item["sales"],
                )

                ws.cell(
                    row,
                    9,
                    item["amount"],
                )

            ws.cell(
                row,
                8,
                (
                    f'=IF(OR('
                    f'G{row}="",'
                    f'G{row}=0,'
                    f'I{row}=""'
                    f'),"",'
                    f'I{row}*1000000/G{row})'
                ),
            )

            copy_row_style(
                ws,
                2,
                row,
                9,
            )

            first = False
            row += 1

    # ========================================================
    # MATERIALS
    # ========================================================

    ws = wb[
        SHEET_MATERIALS
    ]

    if ws.max_row > 1:

        ws.delete_rows(
            2,
            ws.max_row - 1,
        )

    row = 2

    for result in results:

        first = True

        for (
            purchase_type,
            name,
            unit,
        ) in material_master:

            item = (
                result["materials"]
                .get(
                    (
                        purchase_type,
                        name,
                        unit,
                    )
                )
            )

            ws.cell(
                row,
                1,
                result["company"]
                if first
                else None,
            )

            ws.cell(
                row,
                2,
                period,
            )

            ws.cell(
                row,
                3,
                purchase_type,
            )

            ws.cell(
                row,
                4,
                name,
            )

            ws.cell(
                row,
                5,
                unit,
            )

            if item:

                ws.cell(
                    row,
                    6,
                    item["quantity"],
                )

                ws.cell(
                    row,
                    8,
                    item["amount"],
                )

            ws.cell(
                row,
                7,
                (
                    f'=IF(OR('
                    f'F{row}="",'
                    f'F{row}=0,'
                    f'H{row}=""'
                    f'),"",'
                    f'H{row}*1000000/F{row})'
                ),
            )

            copy_row_style(
                ws,
                2,
                row,
                8,
            )

            first = False
            row += 1

    # ========================================================
    # ENERGY
    # ========================================================

    ws = wb[
        SHEET_ENERGY
    ]

    if ws.max_row > 1:

        ws.delete_rows(
            2,
            ws.max_row - 1,
        )

    row = 2

    for result in results:

        first = True

        for name, unit in energy_master:

            item = (
                result["energy"]
                .get(
                    (
                        name,
                        unit,
                    )
                )
            )

            ws.cell(
                row,
                1,
                result["company"]
                if first
                else None,
            )

            ws.cell(
                row,
                2,
                period,
            )

            ws.cell(
                row,
                3,
                name,
            )

            ws.cell(
                row,
                4,
                unit,
            )

            if item:

                ws.cell(
                    row,
                    5,
                    item["quantity"],
                )

                ws.cell(
                    row,
                    7,
                    item["amount"],
                )

            ws.cell(
                row,
                6,
                (
                    f'=IF(OR('
                    f'E{row}="",'
                    f'E{row}=0,'
                    f'G{row}=""'
                    f'),"",'
                    f'G{row}*1000000/E{row})'
                ),
            )

            copy_row_style(
                ws,
                2,
                row,
                7,
            )

            first = False
            row += 1

    wb.save(
        filename
    )

    print(
        f"\nOutput created: "
        f"{filename.name}"
    )


# ============================================================
# MAIN
# ============================================================

def main():

    print(
        "\n======================================"
    )

    print(
        " CODAL 5-COMPANY TEST"
    )

    print(
        "======================================"
    )

    all_period_results = {}

    material_master = set()
    energy_master = set()

    for period in PERIODS:

        print(
            "\n======================================"
        )

        print(
            f"Period: {period}"
        )

        print(
            "======================================"
        )

        period_results = []

        for index, item in enumerate(
            TEST_COMPANIES,
            start=1,
        ):

            print(
                f"\n[{index}/"
                f"{len(TEST_COMPANIES)}]"
            )

            try:

                result = process_company(
                    item["symbol"],
                    item["company"],
                    period,
                )

                if result is None:

                    print(
                        "    RESULT: FAILED"
                    )

                    continue

                period_results.append(
                    result
                )

                material_master.update(
                    result[
                        "materials"
                    ].keys()
                )

                energy_master.update(
                    result[
                        "energy"
                    ].keys()
                )

                print(
                    "    RESULT: OK"
                )

            except Exception as e:

                print(
                    f"    ERROR: "
                    f"{type(e).__name__}: "
                    f"{e}"
                )

            time.sleep(2)

        all_period_results[
            period
        ] = period_results

    # --------------------------------------------------------
    # Fixed master lists across both months
    # --------------------------------------------------------

    material_master = sorted(
        material_master,
        key=lambda x: (
            x[0],
            x[1],
            x[2],
        ),
    )

    preferred_energy = {
        "گاز طبیعی": 1,
        "برق": 2,
        "مازوت": 3,
        "گازوئیل": 4,
    }

    energy_master = sorted(
        energy_master,
        key=lambda x: (
            preferred_energy.get(
                x[0],
                999,
            ),
            x[0],
        ),
    )

    print(
        "\n======================================"
    )

    print(
        f"Master material rows: "
        f"{len(material_master)}"
    )

    print(
        f"Master energy rows: "
        f"{len(energy_master)}"
    )

    print(
        "======================================"
    )

    for period in PERIODS:

        results = (
            all_period_results[
                period
            ]
        )

        print(
            f"\nCreating output "
            f"for {period}..."
        )

        print(
            f"Successful companies: "
            f"{len(results)}/"
            f"{len(TEST_COMPANIES)}"
        )

        create_output(
            period,
            results,
            material_master,
            energy_master,
        )

    print(
        "\n======================================"
    )

    print(
        "DONE"
    )

    print(
        "======================================"
    )


if __name__ == "__main__":
    main()