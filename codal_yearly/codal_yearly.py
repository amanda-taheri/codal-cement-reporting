from __future__ import annotations

import io
import json
import random
import re
import time
from collections import defaultdict
from dataclasses import dataclass
from pathlib import Path
from statistics import median
from typing import Any, Iterable

import pandas as pd
import requests
from lxml import etree, html
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# ============================================================
# CONFIG
# ============================================================

BASE_DIR = Path(__file__).resolve().parent

COMPANIES_FILE = BASE_DIR / "companies.xlsx"
ALIAS_FILE = BASE_DIR / "config" / "line_item_aliases.json"

SEARCH_CACHE_DIR = BASE_DIR / "cache" / "search"
REPORT_CACHE_DIR = BASE_DIR / "cache" / "reports"
OUTPUT_DIR = BASE_DIR / "output"
AUDIT_DIR = BASE_DIR / "audit"

for folder in [
    SEARCH_CACHE_DIR,
    REPORT_CACHE_DIR,
    OUTPUT_DIR,
    AUDIT_DIR,
]:
    folder.mkdir(parents=True, exist_ok=True)

OUTPUT_FILE = OUTPUT_DIR / "صورتهای_مالی_شش_ماهه_و_سالانه_شرکتهای_سیمانی.xlsx"

# Only include financial periods from Jalali year 1404 onward.
# MAX_JALALI_YEAR=None means keep everything from 1404 up to the latest available period.
MIN_JALALI_YEAR: int | None = 1404
MAX_JALALI_YEAR: int | None = None

# If both consolidated and standalone filings exist for the same period:
#   "consolidated" -> prefer consolidated
#   "standalone"   -> prefer standalone
#   "latest"       -> latest published candidate regardless of scope
SCOPE_PREFERENCE = "consolidated"

# Cache controls
REFRESH_SEARCH_CACHE = False
REFRESH_REPORT_CACHE = False

# Search / network limits
MAX_SEARCH_PAGES = 30
MAX_RETRIES = 6
INITIAL_WAIT = 10
PAGE_DELAY = 1.5
COMPANY_DELAY = 3.0
DOWNLOAD_DELAY = 1.0

# CODAL API filters.
# LetterType 6 = interim / annual financial statements.
CODAL_SEARCH_URL = "https://search.codal.ir/api/search/v2/q"
CODAL_CATEGORY = 1
CODAL_LETTER_TYPE = 6

NUMBER_FORMAT = "#,##0"

# Final workbook sheets: comparison (wide) + detailed (long) for each statement.
SHEET_INCOME_COMPARISON = "سود و زیان - مقایسه ای"
SHEET_POSITION_COMPARISON = "وضعیت مالی - مقایسه ای"
SHEET_INCOME_DETAIL = "سود و زیان - تفصیلی"
SHEET_POSITION_DETAIL = "وضعیت مالی - تفصیلی"


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
# TEXT / NUMBER NORMALIZATION
# ============================================================

FA_DIGITS = "۰۱۲۳۴۵۶۷۸۹"
AR_DIGITS = "٠١٢٣٤٥٦٧٨٩"
EN_DIGITS = "0123456789"
DIGIT_MAP = str.maketrans(FA_DIGITS + AR_DIGITS, EN_DIGITS + EN_DIGITS)

DATE_RE = re.compile(r"(1[34]\d{2}/\d{2}/\d{2})")


def normalize_text(value: Any) -> str:
    if value is None:
        return ""

    try:
        if pd.isna(value):
            return ""
    except Exception:
        pass

    text = str(value).translate(DIGIT_MAP)
    text = (
        text.replace("ي", "ی")
        .replace("ى", "ی")
        .replace("ك", "ک")
        .replace("ة", "ه")
        .replace("ۀ", "ه")
        .replace("\u200c", " ")
        .replace("\u200f", " ")
        .replace("\u200e", " ")
        .replace("\xa0", " ")
        .replace("–", "-")
        .replace("—", "-")
    )
    text = re.sub(r"\s+", " ", text)
    return text.strip()


def normalize_date(value: Any) -> str:
    return normalize_text(value).replace("-", "/")


def normalize_line_item(value: Any) -> str:
    """Conservative orthographic normalization for financial line labels."""
    text = normalize_text(value)
    if not text:
        return ""

    # Keep semantic words; normalize only punctuation/spacing variants.
    text = text.replace("‌", " ")
    text = re.sub(r"\s*\(\s*", " (", text)
    text = re.sub(r"\s*\)\s*", ") ", text)
    text = re.sub(r"\s*،\s*", "، ", text)
    text = re.sub(r"\s*/\s*", "/", text)
    text = re.sub(r"\s+", " ", text).strip()

    # Safe spelling variants.
    text = text.replace("داراییها", "دارایی ها")
    text = text.replace("بدهیها", "بدهی ها")
    text = text.replace("فعالیتهای", "فعالیت های")
    text = text.replace("هزینه های", "هزینه های")
    return text


def to_number(value: Any) -> float | None:
    text = normalize_text(value)
    if text in {"", "-", "--", "---", "nan", "None", "null"}:
        return None

    negative = False

    # Accounting parentheses: (1,234) means -1234.
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()

    text = (
        text.replace(",", "")
        .replace("٬", "")
        .replace(" ", "")
        .replace("+", "")
        .replace("−", "-")
    )

    # Ignore percentages in statement value columns.
    if "%" in text or "٪" in text:
        return None

    try:
        number = float(text)
    except Exception:
        return None

    if negative:
        number = -abs(number)

    return number


def safe_get(row: list[Any], index: int | None) -> Any:
    if index is None or index < 0 or index >= len(row):
        return None
    return row[index]


def jalali_year(date_text: str) -> int | None:
    match = re.match(r"^(1[34]\d{2})/", normalize_date(date_text))
    return int(match.group(1)) if match else None


def in_year_range(period: str) -> bool:
    year = jalali_year(period)
    if year is None:
        return False
    if MIN_JALALI_YEAR is not None and year < MIN_JALALI_YEAR:
        return False
    if MAX_JALALI_YEAR is not None and year > MAX_JALALI_YEAR:
        return False
    return True


# ============================================================
# DATA CLASSES
# ============================================================


@dataclass
class ReportCandidate:
    symbol: str
    company: str
    title: str
    report_period: str
    length_months: int | None
    scope: str
    is_correction: bool
    publish_dt: str
    sent_dt: str
    tracing_no: str
    raw: dict[str, Any]


@dataclass
class TableEntry:
    name: str
    df: pd.DataFrame


@dataclass
class StatementExtraction:
    kind: str
    current_period: str
    values: dict[str, float | None]
    raw_labels: dict[str, set[str]]
    positions: dict[str, float]
    restated_columns: list[tuple[str, dict[str, float | None]]]
    unit_text: str
    unit_factor_to_million: float
    source_table_name: str


# ============================================================
# COMPANIES
# ============================================================


def load_companies() -> list[tuple[str, str]]:
    # Prefer the first sheet. In the user's current file this is "شرکت‌ها".
    df = pd.read_excel(COMPANIES_FILE, sheet_name=0)

    symbol_col = None
    company_col = None

    for col in df.columns:
        col_name = normalize_text(col)
        if col_name == "نماد":
            symbol_col = col
        elif col_name == "نام شرکت":
            company_col = col

    if symbol_col is None:
        raise RuntimeError("Column 'نماد' not found in companies.xlsx")
    if company_col is None:
        raise RuntimeError("Column 'نام شرکت' not found in companies.xlsx")

    companies: list[tuple[str, str]] = []
    seen: set[str] = set()

    for _, row in df.iterrows():
        symbol = normalize_text(row[symbol_col])
        company = normalize_text(row[company_col])
        if not symbol:
            continue
        if symbol in seen:
            continue
        seen.add(symbol)
        companies.append((symbol, company or symbol))

    return companies


# ============================================================
# ALIASES
# ============================================================


def load_aliases() -> dict[str, dict[str, str]]:
    if not ALIAS_FILE.exists():
        return {"income_statement": {}, "financial_position": {}}

    try:
        data = json.loads(ALIAS_FILE.read_text(encoding="utf-8"))
    except Exception:
        return {"income_statement": {}, "financial_position": {}}

    result: dict[str, dict[str, str]] = {
        "income_statement": {},
        "financial_position": {},
    }

    for kind in result:
        raw_map = data.get(kind, {}) if isinstance(data, dict) else {}
        if not isinstance(raw_map, dict):
            continue
        for source, target in raw_map.items():
            source_n = normalize_line_item(source)
            target_n = normalize_line_item(target)
            if source_n and target_n:
                result[kind][source_n] = target_n

    return result


ALIASES = load_aliases()


def canonical_line_item(kind: str, raw_label: Any) -> str:
    label = normalize_line_item(raw_label)
    if not label:
        return ""
    return ALIASES.get(kind, {}).get(label, label)


# ============================================================
# SEARCH CACHE / CODAL SEARCH
# ============================================================


def safe_filename(text: str) -> str:
    text = normalize_text(text)
    return re.sub(r"[^0-9A-Za-zآ-ی_.-]+", "_", text).strip("_") or "item"


def search_cache_path(symbol: str, page: int) -> Path:
    return SEARCH_CACHE_DIR / f"{safe_filename(symbol)}_page_{page}.json"


def request_json_with_retry(url: str, params: dict[str, Any]) -> dict[str, Any]:
    """Request CODAL JSON without ever converting a network failure into an empty result."""
    wait_seconds = INITIAL_WAIT
    last_error: Exception | None = None

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, params=params, timeout=60)

            if response.status_code == 429:
                last_error = RuntimeError("CODAL returned HTTP 429 Too Many Requests")
                print(f"    HTTP 429 (attempt {attempt}/{MAX_RETRIES})")
                if attempt >= MAX_RETRIES:
                    break
                print(f"    Waiting {wait_seconds}s...")
                time.sleep(wait_seconds + random.uniform(0.0, 1.5))
                wait_seconds = min(wait_seconds * 2, 120)
                continue

            response.raise_for_status()
            data = response.json()
            if not isinstance(data, dict):
                raise RuntimeError("CODAL search response is not a JSON object")
            return data

        except Exception as error:
            last_error = error
            if attempt >= MAX_RETRIES:
                break
            print(f"    Search error: {type(error).__name__}")
            print(f"    Retrying in {wait_seconds}s...")
            time.sleep(wait_seconds + random.uniform(0.0, 1.5))
            wait_seconds = min(wait_seconds * 2, 120)

    raise RuntimeError(
        f"CODAL search failed after {MAX_RETRIES} attempts; "
        f"last error: {last_error}"
    )


def search_financial_reports(symbol: str) -> list[dict[str, Any]]:
    all_letters: list[dict[str, Any]] = []
    seen_tracing: set[str] = set()
    expected_pages: int | None = None

    for page in range(1, MAX_SEARCH_PAGES + 1):
        cache_file = search_cache_path(symbol, page)
        data: dict[str, Any] | None = None

        if not REFRESH_SEARCH_CACHE and cache_file.exists():
            try:
                data = json.loads(cache_file.read_text(encoding="utf-8"))
                print(f"    Using cached search page {page}")
            except Exception:
                data = None

        if data is None:
            params = {
                "Symbol": symbol,
                "LetterType": CODAL_LETTER_TYPE,
                "Category": CODAL_CATEGORY,
                "Audited": "true",
                "NotAudited": "false",
                "Consolidatable": "true",
                "NotConsolidatable": "true",
                "Childs": "false",
                "Mains": "true",
                "Publisher": "false",
                "PageNumber": page,
                "search": "true",
            }
            data = request_json_with_retry(CODAL_SEARCH_URL, params)
            # Only successful HTTP/JSON responses reach this point, so failed
            # searches (including exhausted 429 retries) are never cached.
            cache_file.write_text(
                json.dumps(data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
            time.sleep(PAGE_DELAY)

        letters = data.get("Letters", []) if isinstance(data, dict) else []
        if not isinstance(letters, list):
            letters = []

        if expected_pages is None:
            try:
                expected_pages = int(data.get("Page", 0))
            except Exception:
                expected_pages = None

        new_count = 0
        for letter in letters:
            tracing = normalize_text(letter.get("TracingNo", ""))
            key = tracing or normalize_text(letter.get("Url", ""))
            if key and key in seen_tracing:
                continue
            if key:
                seen_tracing.add(key)
            all_letters.append(letter)
            new_count += 1

        if not letters or new_count == 0:
            break
        if expected_pages and page >= expected_pages:
            break

    return all_letters


# ============================================================
# REPORT FILTERING / SELECTION
# ============================================================


def parse_report_period(title: str) -> str:
    title_n = normalize_text(title)

    match = re.search(r"منتهی\s*به\s*:?[ ]*(1[34]\d{2}/\d{2}/\d{2})", title_n)
    if match:
        return normalize_date(match.group(1))

    # Fallback: first Jalali date in title.
    dates = DATE_RE.findall(title_n)
    return normalize_date(dates[0]) if dates else ""


def parse_length_months(title: str) -> int | None:
    title_n = normalize_text(title)

    match = re.search(r"دوره\s*(\d{1,2})\s*ماهه", title_n)
    if match:
        value = int(match.group(1))
        return value

    if "سال مالی منتهی" in title_n:
        return 12

    return None


def is_audited_title(title: str) -> bool:
    title_n = normalize_text(title)
    return "حسابرسی شده" in title_n and "حسابرسی نشده" not in title_n


def detect_scope(title: str) -> str:
    return "consolidated" if "تلفیقی" in normalize_text(title) else "standalone"


def is_correction_title(title: str) -> bool:
    return "اصلاحیه" in normalize_text(title)


def candidate_from_letter(
    symbol: str,
    company: str,
    letter: dict[str, Any],
) -> ReportCandidate | None:
    title = normalize_text(letter.get("Title", ""))

    if not title:
        return None
    if not is_audited_title(title):
        return None

    length = parse_length_months(title)
    if length not in {6, 12}:
        return None

    period = parse_report_period(title)
    if not period:
        # Keep only correction orphans for fallback matching.
        if not is_correction_title(title):
            return None
    elif not in_year_range(period):
        return None

    return ReportCandidate(
        symbol=symbol,
        company=company,
        title=title,
        report_period=period,
        length_months=length,
        scope=detect_scope(title),
        is_correction=is_correction_title(title),
        publish_dt=normalize_text(letter.get("PublishDateTime", "")),
        sent_dt=normalize_text(letter.get("SentDateTime", "")),
        tracing_no=normalize_text(letter.get("TracingNo", "")),
        raw=letter,
    )


def candidate_sort_key(candidate: ReportCandidate) -> tuple[str, str]:
    return (candidate.publish_dt or candidate.sent_dt, candidate.tracing_no)


def attach_orphan_corrections(
    candidates: list[ReportCandidate],
    selection_log: list[dict[str, Any]],
) -> list[ReportCandidate]:
    normal = [c for c in candidates if c.report_period]
    orphans = [c for c in candidates if not c.report_period and c.is_correction]

    for orphan in sorted(orphans, key=candidate_sort_key):
        eligible = [
            c
            for c in normal
            if c.scope == orphan.scope
            and c.length_months == orphan.length_months
            and candidate_sort_key(c) <= candidate_sort_key(orphan)
        ]

        if not eligible:
            # Wider fallback: same length regardless of scope.
            eligible = [
                c
                for c in normal
                if c.length_months == orphan.length_months
                and candidate_sort_key(c) <= candidate_sort_key(orphan)
            ]

        if not eligible:
            selection_log.append(
                {
                    "symbol": orphan.symbol,
                    "company": orphan.company,
                    "period": "",
                    "title": orphan.title,
                    "action": "skip_orphan_correction",
                    "reason": "No earlier matching 6/12-month report found",
                }
            )
            continue

        target = max(eligible, key=candidate_sort_key)
        orphan.report_period = target.report_period
        orphan.length_months = target.length_months
        orphan.scope = target.scope
        normal.append(orphan)
        selection_log.append(
            {
                "symbol": orphan.symbol,
                "company": orphan.company,
                "period": orphan.report_period,
                "title": orphan.title,
                "action": "attach_orphan_correction",
                "reason": f"Attached to latest earlier {target.length_months}-month {target.scope} report",
            }
        )

    return normal


def choose_scope_candidate(
    period_candidates: list[ReportCandidate],
    scope_conflicts: list[dict[str, Any]],
    selection_log: list[dict[str, Any]],
) -> ReportCandidate:
    by_scope: dict[str, list[ReportCandidate]] = defaultdict(list)
    for c in period_candidates:
        by_scope[c.scope].append(c)

    # First choose latest version inside each scope. This makes explicit
    # correction filings supersede the original report automatically.
    latest_by_scope = {
        scope: max(items, key=candidate_sort_key)
        for scope, items in by_scope.items()
    }

    if len(latest_by_scope) == 1:
        chosen = next(iter(latest_by_scope.values()))
    else:
        sample = period_candidates[0]
        scope_conflicts.append(
            {
                "symbol": sample.symbol,
                "company": sample.company,
                "period": sample.report_period,
                "scopes_found": ", ".join(sorted(latest_by_scope)),
                "consolidated_title": latest_by_scope.get("consolidated").title
                if latest_by_scope.get("consolidated")
                else "",
                "standalone_title": latest_by_scope.get("standalone").title
                if latest_by_scope.get("standalone")
                else "",
                "preference": SCOPE_PREFERENCE,
            }
        )

        if SCOPE_PREFERENCE == "consolidated" and "consolidated" in latest_by_scope:
            chosen = latest_by_scope["consolidated"]
        elif SCOPE_PREFERENCE == "standalone" and "standalone" in latest_by_scope:
            chosen = latest_by_scope["standalone"]
        else:
            chosen = max(latest_by_scope.values(), key=candidate_sort_key)

    for candidate in period_candidates:
        selection_log.append(
            {
                "symbol": candidate.symbol,
                "company": candidate.company,
                "period": candidate.report_period,
                "title": candidate.title,
                "scope": candidate.scope,
                "length_months": candidate.length_months,
                "is_correction": candidate.is_correction,
                "publish_dt": candidate.publish_dt,
                "action": "selected" if candidate is chosen else "skipped_version_or_scope",
                "reason": "latest correction/version + scope preference",
            }
        )

    return chosen


def select_reports_for_company(
    symbol: str,
    company: str,
    letters: list[dict[str, Any]],
    scope_conflicts: list[dict[str, Any]],
    selection_log: list[dict[str, Any]],
) -> list[ReportCandidate]:
    candidates = [
        c
        for letter in letters
        if (c := candidate_from_letter(symbol, company, letter)) is not None
    ]

    candidates = attach_orphan_corrections(candidates, selection_log)

    grouped: dict[str, list[ReportCandidate]] = defaultdict(list)
    for c in candidates:
        if c.report_period:
            grouped[c.report_period].append(c)

    selected = [
        choose_scope_candidate(items, scope_conflicts, selection_log)
        for _, items in sorted(grouped.items())
    ]

    return sorted(selected, key=lambda c: c.report_period)


# ============================================================
# REPORT DOWNLOAD / FILE READING
# ============================================================


def get_excel_url(report: dict[str, Any]) -> str | None:
    excel_url = normalize_text(report.get("ExcelUrl", ""))
    if excel_url:
        if excel_url.startswith("//"):
            return "https:" + excel_url
        if excel_url.startswith("/"):
            return "https://excel.codal.ir" + excel_url
        return excel_url

    report_url = normalize_text(report.get("Url", ""))
    match = re.search(r"LetterSerial=([^&]+)", report_url)
    if match:
        serial = match.group(1)
        return f"https://excel.codal.ir/service/Excel/GetAll/{serial}/0"

    return None


def report_cache_path(candidate: ReportCandidate) -> Path:
    company_dir = REPORT_CACHE_DIR / safe_filename(candidate.symbol)
    company_dir.mkdir(parents=True, exist_ok=True)

    identity = candidate.tracing_no or re.sub(
        r"[^A-Za-z0-9]+", "_", normalize_text(candidate.raw.get("Url", ""))
    )[:80]
    if not identity:
        identity = f"{candidate.report_period}_{candidate.scope}"

    return company_dir / f"{safe_filename(identity)}.xls"


def download_report(candidate: ReportCandidate) -> Path:
    raw_path = report_cache_path(candidate)

    if (
        not REFRESH_REPORT_CACHE
        and raw_path.exists()
        and raw_path.stat().st_size > 500
    ):
        print("        Using cached report")
        return raw_path

    url = get_excel_url(candidate.raw)
    if not url:
        raise RuntimeError("Excel URL not found")

    wait_seconds = INITIAL_WAIT

    for attempt in range(1, MAX_RETRIES + 1):
        try:
            response = session.get(url, timeout=90)

            if response.status_code == 429:
                print(f"        Download 429 (attempt {attempt}/{MAX_RETRIES})")
                print(f"        Waiting {wait_seconds}s...")
                time.sleep(wait_seconds + random.uniform(0.0, 1.5))
                wait_seconds = min(wait_seconds * 2, 120)
                continue

            response.raise_for_status()
            raw_path.write_bytes(response.content)
            print(f"        Downloaded {len(response.content):,} bytes")
            time.sleep(DOWNLOAD_DELAY)
            return raw_path

        except Exception as error:
            if attempt >= MAX_RETRIES:
                raise
            print(f"        Download error: {type(error).__name__}")
            print(f"        Retrying in {wait_seconds}s...")
            time.sleep(wait_seconds + random.uniform(0.0, 1.5))
            wait_seconds = min(wait_seconds * 2, 120)

    raise RuntimeError("Report download failed")


def detect_file_type(path: Path) -> str:
    data = path.read_bytes()

    if data.startswith(b"\xD0\xCF\x11\xE0"):
        return "xls"
    if data.startswith(b"PK"):
        return "xlsx"

    sample = data[:15000].decode("utf-8", errors="ignore").lower()
    if "<html" in sample or "<table" in sample or "<!doctype" in sample:
        return "html"

    return "unknown"


def decode_html(data: bytes) -> str:
    for encoding in ["utf-8-sig", "utf-8", "windows-1256", "cp1256"]:
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            pass
    return data.decode("utf-8", errors="ignore")


def read_html_tables(decoded: str) -> list[TableEntry]:
    result: list[TableEntry] = []

    try:
        tables = pd.read_html(io.StringIO(decoded), header=None, flavor="lxml")
        for i, table in enumerate(tables, start=1):
            result.append(TableEntry(name=f"html_table_{i}", df=table))
        if result:
            return result
    except Exception:
        pass

    document = html.fromstring(decoded)
    html_tables = document.xpath("//table")

    for i, element in enumerate(html_tables, start=1):
        table_html = etree.tostring(
            element,
            encoding="unicode",
            method="html",
        )
        try:
            parsed = pd.read_html(
                io.StringIO(table_html),
                header=None,
                flavor="lxml",
            )
            for j, table in enumerate(parsed, start=1):
                result.append(TableEntry(name=f"html_table_{i}_{j}", df=table))
        except Exception:
            continue

    return result


def read_codal_file(path: Path) -> list[TableEntry]:
    file_type = detect_file_type(path)
    data = path.read_bytes()

    if file_type == "html":
        return read_html_tables(decode_html(data))

    if file_type == "xlsx":
        sheets = pd.read_excel(
            io.BytesIO(data),
            sheet_name=None,
            header=None,
            engine="openpyxl",
        )
        return [TableEntry(name=str(name), df=df) for name, df in sheets.items()]

    if file_type == "xls":
        sheets = pd.read_excel(
            io.BytesIO(data),
            sheet_name=None,
            header=None,
            engine="xlrd",
        )
        return [TableEntry(name=str(name), df=df) for name, df in sheets.items()]

    raise RuntimeError("Unsupported CODAL Excel file type")


# ============================================================
# TABLE NORMALIZATION / IDENTIFICATION
# ============================================================


def normalize_table(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()

    header_rows: list[list[str]] = []
    if isinstance(df.columns, pd.MultiIndex):
        for level in range(df.columns.nlevels):
            row: list[str] = []
            for column in df.columns:
                value = normalize_text(column[level])
                if value.startswith("Unnamed:") or value.lower() == "nan":
                    value = ""
                row.append(value)
            header_rows.append(row)

    body_rows = [
        [normalize_text(value) for value in row]
        for row in df.itertuples(index=False, name=None)
    ]

    if not header_rows and not body_rows:
        return pd.DataFrame()

    result = pd.DataFrame(header_rows + body_rows)

    if result.empty:
        return result

    result = result[
        ~result.apply(
            lambda row: all(normalize_text(v) == "" for v in row),
            axis=1,
        )
    ]
    result = result.loc[
        :,
        ~result.apply(
            lambda col: all(normalize_text(v) == "" for v in col),
            axis=0,
        ),
    ]
    result = result.reset_index(drop=True)
    result.columns = range(result.shape[1])
    return result


def table_text(df: pd.DataFrame) -> str:
    if df.empty:
        return ""
    return " ".join(normalize_text(value) for value in df.fillna("").values.flatten())


INCOME_KEYWORDS = [
    ("درآمدهای عملیاتی", 8),
    ("بهای تمام شده", 7),
    ("سود (زیان) ناخالص", 6),
    ("سود(زیان) ناخالص", 6),
    ("هزینه های فروش", 4),
    ("سود (زیان) عملیاتی", 6),
    ("سود(زیان) عملیاتی", 6),
    ("سود (زیان) خالص", 7),
    ("سود(زیان) خالص", 7),
    ("سود پایه هر سهم", 3),
]

POSITION_KEYWORDS = [
    ("دارایی های جاری", 7),
    ("دارایی های غیرجاری", 7),
    ("جمع دارایی", 7),
    ("بدهی های جاری", 7),
    ("بدهی های غیرجاری", 7),
    ("حقوق مالکانه", 8),
    ("جمع بدهی", 6),
    ("وجه نقد", 3),
    ("سرمایه", 3),
]


def statement_score(
    entry: TableEntry,
    kind: str,
    report_scope: str,
) -> tuple[int, pd.DataFrame]:
    df = normalize_table(entry.df)
    if df.empty:
        return -999, df

    text = table_text(df)
    name = normalize_text(entry.name)
    haystack = f"{name} {text}"

    score = 0

    if kind == "income_statement":
        for keyword, points in INCOME_KEYWORDS:
            if keyword in haystack:
                score += points

        if "صورت سود و زیان" in haystack:
            score += 15
        if "سود و زیان جامع" in haystack:
            score -= 18
        if "صورت وضعیت مالی" in haystack:
            score -= 15

    elif kind == "financial_position":
        for keyword, points in POSITION_KEYWORDS:
            if keyword in haystack:
                score += points

        if "صورت وضعیت مالی" in haystack or "ترازنامه" in haystack:
            score += 15
        if "صورت سود و زیان" in haystack:
            score -= 15

    has_consolidated = "تلفیقی" in haystack
    if report_scope == "consolidated":
        score += 18 if has_consolidated else -4
    else:
        score += -18 if has_consolidated else 4

    return score, df


def find_best_statement_table(
    tables: list[TableEntry],
    kind: str,
    report_scope: str,
) -> tuple[TableEntry, pd.DataFrame, int]:
    best_entry: TableEntry | None = None
    best_df: pd.DataFrame | None = None
    best_score = -999999

    for entry in tables:
        score, df = statement_score(entry, kind, report_scope)
        if score > best_score:
            best_score = score
            best_entry = entry
            best_df = df

    if best_entry is None or best_df is None:
        raise RuntimeError(f"No table found for {kind}")

    threshold = 16 if kind == "income_statement" else 16
    if best_score < threshold:
        raise RuntimeError(
            f"Could not confidently identify {kind} table; best score={best_score}"
        )

    return best_entry, best_df, best_score


# ============================================================
# FINANCIAL STATEMENT PARSING
# ============================================================


def detect_unit(df: pd.DataFrame) -> tuple[str, float]:
    text = table_text(df)[:100000]

    if "میلیون ریال" in text:
        return "میلیون ریال", 1.0
    if "هزار ریال" in text:
        return "هزار ریال", 1.0 / 1000.0

    # Avoid interpreting "میلیون ریال" as plain rial (already handled above).
    if " ریال" in text or text.endswith("ریال"):
        return "ریال", 1.0 / 1_000_000.0

    return "unknown", 1.0


def detect_label_column(df: pd.DataFrame, kind: str) -> int:
    keywords = (
        [item[0] for item in INCOME_KEYWORDS]
        if kind == "income_statement"
        else [item[0] for item in POSITION_KEYWORDS]
    )

    best_col = 0
    best_score = -1

    for col in df.columns:
        score = 0
        nonempty = 0
        for value in df[col].head(260):
            text = normalize_text(value)
            if not text:
                continue
            nonempty += 1
            if any(keyword in text for keyword in keywords):
                score += 6
            if len(text) > 3 and to_number(text) is None:
                score += 0.15

        # A label column is text-heavy but should not be just a sparse title column.
        if nonempty >= 4:
            score += min(nonempty, 40) * 0.02

        if score > best_score:
            best_score = score
            best_col = int(col)

    return best_col


def column_header_text(df: pd.DataFrame, col: int, max_rows: int = 12) -> str:
    values = [normalize_text(v) for v in df[col].head(max_rows).tolist()]
    return " ".join(v for v in values if v)


def detect_period_columns(df: pd.DataFrame) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []

    for col in df.columns:
        header = column_header_text(df, int(col))
        if not header:
            continue

        dates = DATE_RE.findall(header)
        if not dates:
            continue

        # The most specific date is usually the first date in the merged header.
        period = normalize_date(dates[0])
        restated = (
            "تجدید ارائه" in header
            or "اصلاح شده" in header
            or "اصلاح2شده" in header
        )

        result.append(
            {
                "col": int(col),
                "header": header,
                "period": period,
                "restated": restated,
            }
        )

    return result


def looks_like_line_item(label: str) -> bool:
    if not label:
        return False

    lowered = label.lower()
    blocked_exact = {
        "شرح",
        "یادداشت",
        "درصد تغییر",
        "حسابرسی شده",
        "حسابرسی نشده",
        "دارایی ها",
        "بدهی ها",
        "حقوق مالکانه و بدهی ها",
        "عملیات در حال تداوم",
    }
    if label in blocked_exact:
        return False

    if any(token in lowered for token in ["دوره منتهی", "تجدید ارائه", "درصد تغییر"]):
        return False
    if DATE_RE.search(label):
        return False
    if to_number(label) is not None:
        return False
    if len(label) < 2:
        return False

    return True


def extract_column_values(
    df: pd.DataFrame,
    kind: str,
    label_col: int,
    value_col: int,
    unit_factor: float,
) -> tuple[dict[str, float | None], dict[str, set[str]], dict[str, float]]:
    values: dict[str, float | None] = {}
    raw_labels: dict[str, set[str]] = defaultdict(set)
    position_samples: dict[str, list[float]] = defaultdict(list)

    rows = df.fillna("").values.tolist()
    total_rows = max(len(rows) - 1, 1)

    for row_index, row in enumerate(rows):
        raw_label = normalize_line_item(safe_get(row, label_col))
        if not looks_like_line_item(raw_label):
            continue

        number = to_number(safe_get(row, value_col))
        if number is None:
            # Skip headings / decorative rows that never carry a number.
            continue

        label = canonical_line_item(kind, raw_label)
        if not label:
            continue

        number *= unit_factor

        # Avoid silently overwriting an exact duplicate label in the same table.
        # If CODAL truly duplicates a row label, keep the first and expose the
        # duplicate through the audit raw-label list for manual review.
        if label not in values:
            values[label] = number
            position_samples[label].append(row_index / total_rows)
        elif values[label] == number:
            position_samples[label].append(row_index / total_rows)
        else:
            duplicate_label = f"{label} [ردیف تکراری]"
            suffix = 2
            while duplicate_label in values:
                duplicate_label = f"{label} [ردیف تکراری {suffix}]"
                suffix += 1
            values[duplicate_label] = number
            raw_labels[duplicate_label].add(raw_label)
            position_samples[duplicate_label].append(row_index / total_rows)
            continue

        raw_labels[label].add(raw_label)

    positions = {
        label: median(samples)
        for label, samples in position_samples.items()
        if samples
    }
    return values, dict(raw_labels), positions


def parse_statement(
    tables: list[TableEntry],
    kind: str,
    report_period: str,
    report_scope: str,
) -> StatementExtraction:
    entry, df, _score = find_best_statement_table(tables, kind, report_scope)

    label_col = detect_label_column(df, kind)
    period_columns = detect_period_columns(df)
    if not period_columns:
        raise RuntimeError(f"No period columns detected in {kind}")

    # Prefer an exact current-period match that is not marked restated.
    current_candidates = [
        item
        for item in period_columns
        if item["period"] == report_period and not item["restated"]
    ]
    if not current_candidates:
        current_candidates = [
            item for item in period_columns if item["period"] == report_period
        ]
    if not current_candidates:
        raise RuntimeError(
            f"Current period {report_period} not found in {kind} table headers"
        )

    current_col = current_candidates[0]["col"]
    unit_text, unit_factor = detect_unit(df)

    values, raw_labels, positions = extract_column_values(
        df,
        kind,
        label_col,
        current_col,
        unit_factor,
    )

    restated_columns: list[tuple[str, dict[str, float | None]]] = []
    for item in period_columns:
        if not item["restated"]:
            continue
        target_period = item["period"]
        if target_period == report_period:
            continue

        restated_values, restated_raw, _ = extract_column_values(
            df,
            kind,
            label_col,
            item["col"],
            unit_factor,
        )

        if restated_values:
            restated_columns.append((target_period, restated_values))
            for label, variants in restated_raw.items():
                raw_labels.setdefault(label, set()).update(variants)

    return StatementExtraction(
        kind=kind,
        current_period=report_period,
        values=values,
        raw_labels=raw_labels,
        positions=positions,
        restated_columns=restated_columns,
        unit_text=unit_text,
        unit_factor_to_million=unit_factor,
        source_table_name=entry.name,
    )


# ============================================================
# MASTER LINE-ITEM ORDER / RESTATEMENTS
# ============================================================


def build_master_order(
    rows: dict[tuple[str, str], dict[str, float | None]],
    position_history: dict[str, list[float]],
) -> list[str]:
    all_items: set[str] = set()
    first_seen: dict[str, int] = {}
    counter = 0

    for values in rows.values():
        for item in values:
            all_items.add(item)
            if item not in first_seen:
                first_seen[item] = counter
                counter += 1

    def sort_key(item: str) -> tuple[float, int, str]:
        positions = position_history.get(item, [])
        position = median(positions) if positions else 999.0
        return (position, first_seen.get(item, 999999), item)

    return sorted(all_items, key=sort_key)


def apply_restatement(
    target_rows: dict[tuple[str, str], dict[str, float | None]],
    company: str,
    symbol: str,
    target_period: str,
    restated_values: dict[str, float | None],
    kind: str,
    source_period: str,
    source_title: str,
    restatement_log: list[dict[str, Any]],
) -> None:
    key = (symbol, target_period)
    row = target_rows.setdefault(key, {})

    changed = 0
    created = 0

    for item, value in restated_values.items():
        if value is None:
            continue
        if item in row:
            if row[item] != value:
                changed += 1
        else:
            created += 1
        row[item] = value

    restatement_log.append(
        {
            "company": company,
            "symbol": symbol,
            "statement": kind,
            "target_period": target_period,
            "source_period": source_period,
            "changed_existing_items": changed,
            "created_missing_items": created,
            "source_title": source_title,
        }
    )


# ============================================================
# EXCEL OUTPUT
# ============================================================


def period_type_text(months: int | None) -> str:
    if months == 6:
        return "6 ماهه"
    if months == 12:
        return "12 ماهه"
    return ""


def ordered_keys(
    companies: list[tuple[str, str]],
    rows: dict[tuple[str, str], dict[str, float | None]],
) -> list[tuple[str, str]]:
    """Keep companies in companies.xlsx order and periods ascending inside each company."""
    company_order = {symbol: index for index, (symbol, _) in enumerate(companies)}
    return sorted(
        rows,
        key=lambda key: (company_order.get(key[0], 999999), key[1]),
    )


def apply_sheet_style(
    ws,
    numeric_from_col: int | None = None,
    wrap_cols: set[int] | None = None,
) -> None:
    wrap_cols = wrap_cols or set()
    ws.sheet_view.rightToLeft = True
    ws.sheet_view.showGridLines = False

    header_fill = PatternFill("solid", fgColor="1F4E78")
    header_font = Font(color="FFFFFF", bold=True)
    thin = Side(style="thin", color="D9E2F3")
    header_border = Border(bottom=Side(style="thin", color="9EADBA"))

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )
        cell.border = header_border

    ws.row_dimensions[1].height = 42
    ws.auto_filter.ref = ws.dimensions

    for row in ws.iter_rows(min_row=2):
        for idx, cell in enumerate(row, start=1):
            cell.alignment = Alignment(
                horizontal="right" if idx in wrap_cols else "center",
                vertical="center",
                wrap_text=(idx in wrap_cols),
            )
            cell.border = Border(bottom=thin)
            if (
                numeric_from_col is not None
                and idx >= numeric_from_col
                and isinstance(cell.value, (int, float))
            ):
                cell.number_format = NUMBER_FORMAT


def write_comparison_sheet(
    wb: Workbook,
    sheet_name: str,
    companies: list[tuple[str, str]],
    rows: dict[tuple[str, str], dict[str, float | None]],
    master_items: list[str],
    period_lengths: dict[tuple[str, str], int | None],
) -> None:
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = "E2"

    companies_map = {symbol: company for symbol, company in companies}
    headers = ["نام شرکت", "نماد", "دوره", "نوع دوره"] + master_items
    ws.append(headers)

    for symbol, period in ordered_keys(companies, rows):
        values = rows[(symbol, period)]
        row = [
            companies_map.get(symbol, symbol),
            symbol,
            period,
            period_type_text(period_lengths.get((symbol, period))),
        ]
        row.extend(values.get(item) for item in master_items)
        ws.append(row)

    apply_sheet_style(ws, numeric_from_col=5, wrap_cols={1})

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 12

    for col_index in range(5, 5 + len(master_items)):
        letter = get_column_letter(col_index)
        item = master_items[col_index - 5]
        ws.column_dimensions[letter].width = min(max(len(item) * 0.85 + 4, 16), 34)


def write_detail_sheet(
    wb: Workbook,
    sheet_name: str,
    companies: list[tuple[str, str]],
    rows: dict[tuple[str, str], dict[str, float | None]],
    master_items: list[str],
    period_lengths: dict[tuple[str, str], int | None],
) -> None:
    """Long format: company -> period -> line items, then next company."""
    ws = wb.create_sheet(sheet_name)
    ws.freeze_panes = "A2"

    companies_map = {symbol: company for symbol, company in companies}
    ws.append(["نام شرکت", "نماد", "دوره", "نوع دوره", "شرح", "مبلغ (میلیون ریال)"])

    for symbol, period in ordered_keys(companies, rows):
        values = rows[(symbol, period)]
        for item in master_items:
            # The master item still appears even when a company does not have it.
            # This keeps all companies directly comparable in long format too.
            ws.append(
                [
                    companies_map.get(symbol, symbol),
                    symbol,
                    period,
                    period_type_text(period_lengths.get((symbol, period))),
                    item,
                    values.get(item),
                ]
            )

    apply_sheet_style(ws, numeric_from_col=6, wrap_cols={1, 5})

    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 14
    ws.column_dimensions["C"].width = 15
    ws.column_dimensions["D"].width = 12
    ws.column_dimensions["E"].width = 48
    ws.column_dimensions["F"].width = 22


def create_output_workbook(
    companies: list[tuple[str, str]],
    income_rows: dict[tuple[str, str], dict[str, float | None]],
    position_rows: dict[tuple[str, str], dict[str, float | None]],
    income_order: list[str],
    position_order: list[str],
    period_lengths: dict[tuple[str, str], int | None],
) -> Path:
    wb = Workbook()
    wb.remove(wb.active)

    write_comparison_sheet(
        wb,
        SHEET_INCOME_COMPARISON,
        companies,
        income_rows,
        income_order,
        period_lengths,
    )
    write_comparison_sheet(
        wb,
        SHEET_POSITION_COMPARISON,
        companies,
        position_rows,
        position_order,
        period_lengths,
    )
    write_detail_sheet(
        wb,
        SHEET_INCOME_DETAIL,
        companies,
        income_rows,
        income_order,
        period_lengths,
    )
    write_detail_sheet(
        wb,
        SHEET_POSITION_DETAIL,
        companies,
        position_rows,
        position_order,
        period_lengths,
    )

    wb.save(OUTPUT_FILE)
    return OUTPUT_FILE


# ============================================================
# AUDIT OUTPUTS
# ============================================================


def write_csv(rows: Iterable[dict[str, Any]], path: Path) -> None:
    rows = list(rows)
    if not rows:
        path.write_text("", encoding="utf-8-sig")
        return
    pd.DataFrame(rows).to_csv(path, index=False, encoding="utf-8-sig")


def write_line_item_audit(
    raw_label_history: dict[str, set[str]],
    position_history: dict[str, list[float]],
    path: Path,
) -> None:
    rows = []
    for canonical, raw_variants in raw_label_history.items():
        rows.append(
            {
                "canonical_label": canonical,
                "raw_variants": " | ".join(sorted(raw_variants)),
                "variant_count": len(raw_variants),
                "median_relative_position": median(position_history.get(canonical, [999.0])),
            }
        )

    rows.sort(key=lambda r: (r["median_relative_position"], r["canonical_label"]))
    write_csv(rows, path)


# ============================================================
# MAIN PROCESSING
# ============================================================


def main() -> None:
    print("=" * 72)
    print("CODAL AUDITED 6/12-MONTH FINANCIAL STATEMENTS")
    print("=" * 72)

    companies = load_companies()
    print(f"Loaded {len(companies)} companies from companies.xlsx")

    income_rows: dict[tuple[str, str], dict[str, float | None]] = {}
    position_rows: dict[tuple[str, str], dict[str, float | None]] = {}
    period_lengths: dict[tuple[str, str], int | None] = {}

    income_positions: dict[str, list[float]] = defaultdict(list)
    position_positions: dict[str, list[float]] = defaultdict(list)

    income_raw_labels: dict[str, set[str]] = defaultdict(set)
    position_raw_labels: dict[str, set[str]] = defaultdict(set)

    selection_log: list[dict[str, Any]] = []
    scope_conflicts: list[dict[str, Any]] = []
    processing_errors: list[dict[str, Any]] = []
    restatement_log: list[dict[str, Any]] = []

    # We store restatement events and apply them after all base rows exist.
    pending_restatements: list[
        tuple[str, str, str, str, str, str, int | None, dict[str, float | None]]
    ] = []

    for company_index, (symbol, company) in enumerate(companies, start=1):
        print()
        print(f"[{company_index}/{len(companies)}] {symbol} - {company}")

        try:
            letters = search_financial_reports(symbol)
            selected_reports = select_reports_for_company(
                symbol,
                company,
                letters,
                scope_conflicts,
                selection_log,
            )
            print(f"    Selected {len(selected_reports)} audited 6/12-month periods")
        except Exception as error:
            processing_errors.append(
                {
                    "company": company,
                    "symbol": symbol,
                    "period": "",
                    "stage": "search_or_selection",
                    "error_type": type(error).__name__,
                    "error": str(error),
                }
            )
            print(f"    ERROR: {type(error).__name__}: {error}")
            continue

        for report_index, candidate in enumerate(selected_reports, start=1):
            print(
                f"    ({report_index}/{len(selected_reports)}) "
                f"{candidate.report_period} | {candidate.scope}"
                + (" | correction" if candidate.is_correction else "")
            )
            period_lengths[(symbol, candidate.report_period)] = candidate.length_months

            try:
                path = download_report(candidate)
                tables = read_codal_file(path)
            except Exception as error:
                processing_errors.append(
                    {
                        "company": company,
                        "symbol": symbol,
                        "period": candidate.report_period,
                        "stage": "download_or_read",
                        "title": candidate.title,
                        "error_type": type(error).__name__,
                        "error": str(error),
                    }
                )
                print(f"        ERROR: {type(error).__name__}: {error}")
                continue

            # ---------------- Income statement ----------------
            try:
                extraction = parse_statement(
                    tables,
                    "income_statement",
                    candidate.report_period,
                    candidate.scope,
                )
                income_rows[(symbol, candidate.report_period)] = dict(extraction.values)

                for label, pos in extraction.positions.items():
                    income_positions[label].append(pos)
                for label, variants in extraction.raw_labels.items():
                    income_raw_labels[label].update(variants)

                for target_period, restated_values in extraction.restated_columns:
                    pending_restatements.append(
                        (
                            "income_statement",
                            company,
                            symbol,
                            target_period,
                            candidate.report_period,
                            candidate.title,
                            candidate.length_months,
                            restated_values,
                        )
                    )

                print(
                    f"        Income: {len(extraction.values)} rows | "
                    f"table={extraction.source_table_name} | "
                    f"unit={extraction.unit_text}"
                )
            except Exception as error:
                processing_errors.append(
                    {
                        "company": company,
                        "symbol": symbol,
                        "period": candidate.report_period,
                        "stage": "parse_income_statement",
                        "title": candidate.title,
                        "error_type": type(error).__name__,
                        "error": str(error),
                    }
                )
                print(f"        Income ERROR: {type(error).__name__}: {error}")

            # ---------------- Financial position ----------------
            try:
                extraction = parse_statement(
                    tables,
                    "financial_position",
                    candidate.report_period,
                    candidate.scope,
                )
                position_rows[(symbol, candidate.report_period)] = dict(extraction.values)

                for label, pos in extraction.positions.items():
                    position_positions[label].append(pos)
                for label, variants in extraction.raw_labels.items():
                    position_raw_labels[label].update(variants)

                for target_period, restated_values in extraction.restated_columns:
                    pending_restatements.append(
                        (
                            "financial_position",
                            company,
                            symbol,
                            target_period,
                            candidate.report_period,
                            candidate.title,
                            candidate.length_months,
                            restated_values,
                        )
                    )

                print(
                    f"        Position: {len(extraction.values)} rows | "
                    f"table={extraction.source_table_name} | "
                    f"unit={extraction.unit_text}"
                )
            except Exception as error:
                processing_errors.append(
                    {
                        "company": company,
                        "symbol": symbol,
                        "period": candidate.report_period,
                        "stage": "parse_financial_position",
                        "title": candidate.title,
                        "error_type": type(error).__name__,
                        "error": str(error),
                    }
                )
                print(f"        Position ERROR: {type(error).__name__}: {error}")

        if company_index < len(companies):
            time.sleep(COMPANY_DELAY)

    # Apply explicit restated / corrected comparative columns after base extraction.
    print()
    print("Applying restated historical columns...")

    for event in pending_restatements:
        (
            kind,
            company,
            symbol,
            target_period,
            source_period,
            source_title,
            source_length_months,
            restated_values,
        ) = event

        # A comparative restated period has the same 6/12-month duration as
        # the source statement unless a direct filing already set it.
        period_lengths.setdefault((symbol, target_period), source_length_months)

        # Respect optional year filters for the historical target too.
        if not in_year_range(target_period):
            continue

        if kind == "income_statement":
            apply_restatement(
                income_rows,
                company,
                symbol,
                target_period,
                restated_values,
                kind,
                source_period,
                source_title,
                restatement_log,
            )
        else:
            apply_restatement(
                position_rows,
                company,
                symbol,
                target_period,
                restated_values,
                kind,
                source_period,
                source_title,
                restatement_log,
            )

    income_order = build_master_order(income_rows, income_positions)
    position_order = build_master_order(position_rows, position_positions)

    # Audit files
    write_csv(selection_log, AUDIT_DIR / "report_selection.csv")
    write_csv(scope_conflicts, AUDIT_DIR / "report_scope_conflicts.csv")
    write_csv(processing_errors, AUDIT_DIR / "processing_errors.csv")
    write_csv(restatement_log, AUDIT_DIR / "restatements_applied.csv")
    write_line_item_audit(
        income_raw_labels,
        income_positions,
        AUDIT_DIR / "line_items_income.csv",
    )
    write_line_item_audit(
        position_raw_labels,
        position_positions,
        AUDIT_DIR / "line_items_financial_position.csv",
    )

    output_path = create_output_workbook(
        companies,
        income_rows,
        position_rows,
        income_order,
        position_order,
        period_lengths,
    )

    print()
    print("=" * 72)
    print("DONE")
    print(f"Output: {output_path}")
    print(f"Income periods: {len(income_rows)}")
    print(f"Financial-position periods: {len(position_rows)}")
    print(f"Income master line items: {len(income_order)}")
    print(f"Position master line items: {len(position_order)}")
    print(f"Scope conflicts: {len(scope_conflicts)}")
    print(f"Processing errors: {len(processing_errors)}")
    print(f"Restatements applied: {len(restatement_log)}")
    print("Workbook sheets: 4 (2 comparison + 2 detailed)")
    print("Check the audit folder before treating the workbook as final.")
    print("=" * 72)


if __name__ == "__main__":
    main()
