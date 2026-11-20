import io
import re
import json
import time
import base64
import html
import requests
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import streamlit as st
from streamlit_autorefresh import st_autorefresh
from openpyxl import load_workbook
from supabase import create_client, ClientOptions

# ============================================================
# CONFIG
# ============================================================

SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
SERVICE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]

BUCKET = "excel-files"
MASTER_FILE = "master.xlsx"
EXCEL_MIME = (
    "application/vnd.openxmlformats-officedocument."
    "spreadsheetml.sheet"
)

# Authentication clients must be isolated per Streamlit browser session.
# A module-global Supabase Auth client can share/overwrite tokens when users
# log in concurrently on the same Streamlit worker.

# SERVICE client is used for admin/profile/storage operations.
admin = create_client(
    SUPABASE_URL,
    SERVICE_KEY
)

# ============================================================
# SESSION
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "profile" not in st.session_state:
    st.session_state.profile = None

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False

if "recovery" not in st.session_state:
    st.session_state.recovery = False

if "session_login_started_at" not in st.session_state:
    st.session_state.session_login_started_at = None

# Each browser session owns its own Supabase Auth client and token state.
# The service-role client below remains server-side and is used only for
# controlled admin/profile/storage operations.
if "_supabase_auth_client" not in st.session_state:
    st.session_state["_supabase_auth_client"] = create_client(
        SUPABASE_URL,
        SUPABASE_KEY,
        options=ClientOptions(flow_type="pkce")
    )

sb = st.session_state["_supabase_auth_client"]


# ============================================================
# FUNCTIONS
# ============================================================

def get_profile(uid):
    r = admin.table("profiles").select("*").eq(
        "id", uid
    ).execute()

    return r.data[0] if r.data else None


CONSTITUENT_MANIFEST_PATH = "index-constituents/manifest.json"
CONSTITUENT_ARCHIVE_PATH = "index-constituents/constituents.json"

@st.cache_data(ttl=86400, max_entries=4, show_spinner=False)
def load_constituent_archive():
    """Cache the shared constituent archive to avoid repeated Supabase reads."""
    try:
        raw = admin.storage.from_(BUCKET).download(CONSTITUENT_ARCHIVE_PATH)
        payload = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
        if isinstance(payload, dict):
            payload.setdefault("indices", {})
            return payload
    except Exception:
        pass
    return {"indices": {}}

def save_constituent_archive(archive):
    admin.storage.from_(BUCKET).upload(
        CONSTITUENT_ARCHIVE_PATH,
        json.dumps(archive, ensure_ascii=False).encode("utf-8"),
        {"content-type": "application/json", "upsert": "true"}
    )
    load_constituent_archive.clear()

@st.cache_data(ttl=86400, max_entries=4, show_spinner=False)
def load_constituent_manifest():
    """Cache the eligible-index manifest instead of downloading it on every rerun."""
    try:
        raw = admin.storage.from_(BUCKET).download(CONSTITUENT_MANIFEST_PATH)
        payload = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
        return payload if isinstance(payload, dict) else {}
    except Exception:
        return {}

def save_index_constituents(index_name, members, source):
    try:
        archive = load_constituent_archive()
        key = " ".join(re.sub(r"[^A-Z0-9]+", " ", str(index_name).upper()).split())
        previous = archive["indices"].get(key, {})
        old_symbols = [row.get("NSE Symbol") for row in previous.get("members", [])]
        new_symbols = [row.get("NSE Symbol") for row in members]
        now = datetime.now(timezone.utc).isoformat()
        record = dict(previous)
        record.update({
            "index_name": index_name, "members": members, "status": "available",
            "source": source, "last_verified_at": now
        })
        if old_symbols != new_symbols or not previous.get("last_updated_at"):
            record["last_updated_at"] = now
            if old_symbols:
                record["previous_symbols"] = old_symbols
        archive["indices"][key] = record
        save_constituent_archive(archive)
    except Exception as error:
        st.warning(f"Could not save {index_name} constituents to Supabase: {error}")
    return members

def save_index_manifest(index_names):
    try:
        names = sorted({str(name).strip().upper() for name in index_names if str(name).strip()})
        old_manifest = load_constituent_manifest()
        if old_manifest.get("index_names") == names:
            return
        manifest = {
            "index_names": names,
            "updated_at": datetime.now(timezone.utc).isoformat(),
            "filter": "CURRENT > 10000"
        }
        admin.storage.from_(BUCKET).upload(
            CONSTITUENT_MANIFEST_PATH,
            json.dumps(manifest, ensure_ascii=False).encode("utf-8"),
            {"content-type": "application/json", "upsert": "true"}
        )
        load_constituent_manifest.clear()
    except Exception as error:
        st.warning(f"Could not save the eligible-index list to Supabase: {error}")


@st.cache_data(ttl=900, max_entries=2, show_spinner=False)
def get_master():
    """Cache Master Excel bytes for 15 minutes to reduce Supabase Storage reads."""
    try:
        return admin.storage.from_(BUCKET).download(MASTER_FILE)
    except Exception:
        return None


def save_master(data):
    admin.storage.from_(BUCKET).upload(
        MASTER_FILE,
        data,
        {
            "content-type": EXCEL_MIME,
            "upsert": "true"
        }
    )
    get_master.clear()


def delete_master():
    admin.storage.from_(BUCKET).remove([MASTER_FILE])
    get_master.clear()


INDEX_CONSTITUENTS_CSV_URL = "https://raw.githubusercontent.com/Servejit/Vardaan/main/index_constituents.csv"

def normalize_index_name(value):
    """Match full NSE index labels to the short group names in the CSV."""
    name = " ".join(re.sub(r"[^A-Z0-9]+", " ", str(value or "").upper()).split())
    if name.startswith("NIFTY "):
        name = name[6:]
    aliases = {
        "50": "50",
        "MIDCAP 50": "MID 50",
        "FINANCIAL SERVICES": "FIN",
        "FINANCIAL SERVICE": "FIN",
        "OIL GAS": "OIL",
        "OIL AND GAS": "OIL",
        "PSU BANK": "PSU",
        "CAPITAL MARKETS": "CAPITAL",
        "CONSUMER DURABLES": "CONSUMER DURABLE",
        "INDIA CEMENT": "CEMENT",
    }
    return aliases.get(name, name)


@st.cache_data(ttl=3600, max_entries=1, show_spinner=False)
def load_github_index_groups():
    """Load all index/group columns from the shared GitHub constituent CSV."""
    import csv
    import io
    import requests

    response = requests.get(INDEX_CONSTITUENTS_CSV_URL, timeout=20)
    response.raise_for_status()
    text_data = response.content.decode("utf-8-sig")
    groups = {}

    # Re-read the small CSV for each header so every column is processed fully.
    headers = next(csv.reader(io.StringIO(text_data)), [])
    for column_index, column_name in enumerate(headers):
        group_name = normalize_index_name(column_name)
        if not group_name:
            continue
        rows = csv.reader(io.StringIO(text_data))
        next(rows, None)
        members = []
        seen_symbols = set()
        for row in rows:
            raw_symbol = row[column_index] if column_index < len(row) else ""
            symbol = normalize_stock_symbol(raw_symbol)
            if symbol and symbol not in seen_symbols:
                seen_symbols.add(symbol)
                members.append({
                    "No.": len(members) + 1,
                    "NSE Symbol": symbol,
                    "Company": "",
                    "Industry": ""
                })
        groups[group_name] = members

    if not groups:
        raise ValueError("The GitHub index constituent CSV has no index columns.")
    return groups


def parse_uploaded_index_fallback(file_bytes, file_name):
    """Read fallback index rows from an uploaded CSV/XLSX without using Supabase."""
    import csv

    required = ["INDEX", "CURRENT", "%CHANGE", "OPEN", "HIGH", "LOW"]

    def header_key(value):
        text = " ".join(re.sub(r"[^A-Z0-9%]+", " ", str(value or "").upper()).split())
        aliases = {
            "CHANGE": "%CHANGE",
            "CHANGE %": "%CHANGE",
            "% CHANGE": "%CHANGE",
            "PERCENT CHANGE": "%CHANGE",
            "INDEX NAME": "INDEX",
            "INDEX VALUE": "CURRENT",
            "LAST": "CURRENT",
        }
        return aliases.get(text, text)

    def extract_rows(raw_rows):
        if not raw_rows:
            return []
        header_row = None
        header_position = None
        for position, row in enumerate(raw_rows[:30]):
            mapped = {header_key(value): index for index, value in enumerate(row) if value is not None}
            if all(column in mapped for column in required):
                header_row = mapped
                header_position = position
                break
        if header_row is None:
            return []

        output_rows = []
        for row in raw_rows[header_position + 1:]:
            values = {}
            for column in required:
                index = header_row[column]
                values[column] = row[index] if index < len(row) else None
            if values.get("INDEX") is None or values.get("CURRENT") is None:
                continue
            output_rows.append(values)
        return output_rows

    extension = str(file_name or "").lower().rsplit(".", 1)[-1]
    if extension == "csv":
        decoded = None
        for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                decoded = file_bytes.decode(encoding)
                break
            except UnicodeDecodeError:
                continue
        if decoded is None:
            raise ValueError("Could not decode the uploaded Index CSV.")
        sample = decoded[:8192]
        try:
            dialect = csv.Sniffer().sniff(sample, delimiters=",;\\t|")
        except csv.Error:
            dialect = csv.excel
        rows = list(csv.reader(io.StringIO(decoded), dialect))
        result = extract_rows(rows)
    elif extension == "xlsx":
        workbook = load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
        result = []
        try:
            for sheet in workbook.worksheets:
                rows = list(sheet.iter_rows(values_only=True))
                result = extract_rows(rows)
                if result:
                    break
        finally:
            workbook.close()
    else:
        raise ValueError("Upload an Index file in .xlsx or .csv format.")

    if not result:
        raise ValueError(
            "No usable index rows found. The file must contain INDEX, CURRENT, %CHANGE, OPEN, HIGH and LOW columns."
        )
    return result


def merge_uploaded_index_fallback(yahoo_csv_bytes, fallback_bytes, fallback_name):
    """Keep Yahoo rows first and add only index rows missing from Yahoo."""
    import csv

    columns = ["INDEX", "CURRENT", "%CHANGE", "OPEN", "HIGH", "LOW"]
    merged = {}
    if yahoo_csv_bytes:
        reader = csv.DictReader(io.StringIO(yahoo_csv_bytes.decode("utf-8-sig")))
        for row in reader:
            key = normalize_index_name(row.get("INDEX"))
            if key:
                merged[key] = {column: row.get(column, "") for column in columns}

    fallback_added = 0
    fallback_error = None
    if fallback_bytes and fallback_name:
        try:
            fallback_rows = parse_uploaded_index_fallback(fallback_bytes, fallback_name)
            allowed_groups = set(load_github_index_groups().keys())
            for row in fallback_rows:
                key = normalize_index_name(row.get("INDEX"))
                if not key or key not in allowed_groups or key in merged:
                    continue
                merged[key] = {
                    column: row.get(column, "")
                    for column in columns
                }
                fallback_added += 1
        except Exception as error:
            fallback_error = str(error)

    output = io.StringIO()
    writer = csv.DictWriter(output, fieldnames=columns)
    writer.writeheader()
    for row in merged.values():
        writer.writerow(row)
    return output.getvalue().encode("utf-8"), len(merged), fallback_added, fallback_error


@st.cache_data(ttl=60, max_entries=1, show_spinner=False)
def load_live_yahoo_index_csv():
    """
    Fetch Yahoo Finance intraday index values for configured groups.
    Yahoo does not provide every custom group in the GitHub constituent CSV.
    Unsupported or unavailable tickers are skipped, never estimated.
    """
    import csv
    import yfinance as yf
    from concurrent.futures import ThreadPoolExecutor, as_completed
    from zoneinfo import ZoneInfo

    # Yahoo Finance symbols for supported NSE indices.
    # Custom constituent groups without a verified Yahoo index ticker are
    # intentionally omitted instead of substituting a different index.
    yahoo_index_tickers = {
        "50": "^NSEI",
        "MID 50": "^NSEMDCP50",
        "AUTO": "^CNXAUTO",
        "FIN": "^CNXFIN",
        "FMCG": "^CNXFMCG",
        "IT": "^CNXIT",
        "MEDIA": "^CNXMEDIA",
        "METAL": "^CNXMETAL",
        "PHARMA": "^CNXPHARMA",
        "PSU": "^CNXPSUBANK",
        "BANK": "^NSEBANK",
        "REALTY": "^CNXREALTY",
        "ENERGY": "^CNXENERGY",
    }

    configured_groups = load_github_index_groups()
    jobs = {
        group: ticker
        for group, ticker in yahoo_index_tickers.items()
        if group in configured_groups
    }
    if not jobs:
        raise ValueError("No configured groups have a Yahoo Finance ticker mapping.")

    def fetch_one(group, ticker):
        history = yf.Ticker(ticker).history(
            period="5d",
            interval="1m",
            auto_adjust=False,
            actions=False,
            prepost=False,
            timeout=15,
        )
        if history is None or history.empty or "Close" not in history:
            return group, ticker, None

        history = history.dropna(subset=["Close"])
        if history.empty:
            return group, ticker, None

        # Use only the latest trading session for current OHLC values.
        latest_date = history.index[-1].date()
        current_session = history[history.index.date == latest_date]
        previous_sessions = history[history.index.date < latest_date]
        if current_session.empty:
            return group, ticker, None

        current = float(current_session["Close"].iloc[-1])
        previous_close = (
            float(previous_sessions["Close"].iloc[-1])
            if not previous_sessions.empty else None
        )
        change_pct = (
            ((current - previous_close) / previous_close) * 100
            if previous_close not in (None, 0) else None
        )

        first_open = (
            float(current_session["Open"].dropna().iloc[0])
            if "Open" in current_session and not current_session["Open"].dropna().empty
            else None
        )
        session_high = (
            float(current_session["High"].max())
            if "High" in current_session and not current_session["High"].dropna().empty
            else None
        )
        session_low = (
            float(current_session["Low"].min())
            if "Low" in current_session and not current_session["Low"].dropna().empty
            else None
        )

        return group, ticker, {
            "INDEX": {
                "50": "NIFTY 50",
                "MID 50": "NIFTY MIDCAP 50",
                "FIN": "NIFTY FINANCIAL SERVICES",
                "PSU": "NIFTY PSU BANK",
            }.get(group, "NIFTY " + group),
            "CURRENT": current,
            "%CHANGE": change_pct,
            "OPEN": first_open,
            "HIGH": session_high,
            "LOW": session_low,
        }

    fetched = {}
    failures = []
    with ThreadPoolExecutor(max_workers=min(8, len(jobs))) as executor:
        futures = {
            executor.submit(fetch_one, group, ticker): (group, ticker)
            for group, ticker in jobs.items()
        }
        for future in as_completed(futures):
            group, ticker = futures[future]
            try:
                result_group, result_ticker, row = future.result()
                if row is not None:
                    fetched[result_group] = row
                else:
                    failures.append(f"{group} ({ticker}): no intraday data")
            except Exception as exc:
                failures.append(f"{group} ({ticker}): {str(exc)[:160]}")

    if not fetched:
        details = "; ".join(failures[:4])
        raise ValueError(
            "Yahoo Finance returned no usable intraday index values. "
            + details
        )

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["INDEX", "CURRENT", "%CHANGE", "OPEN", "HIGH", "LOW"])
    for group in configured_groups:
        if group in fetched:
            row = fetched[group]
            writer.writerow([
                row["INDEX"], row["CURRENT"], row["%CHANGE"],
                row["OPEN"], row["HIGH"], row["LOW"],
            ])

    fetched_at = datetime.now(timezone.utc).astimezone(
        ZoneInfo("Asia/Kolkata")
    ).strftime("%d %b %Y, %I:%M:%S %p")
    return output.getvalue().encode("utf-8"), fetched_at, len(fetched), len(jobs)


def normalize_stock_symbol(value):
    """Normalize common NSE symbol formats before matching Master column A."""
    if value is None:
        return ""
    symbol = str(value).strip().upper()
    symbol = re.sub(r"^NSE\s*[:_-]\s*", "", symbol)
    symbol = re.sub(r"\.NS$", "", symbol)
    symbol = re.sub(r"-(EQ|BE|SM)$", "", symbol)
    symbol = re.sub(r"[^A-Z0-9&-]", "", symbol)
    return symbol




# ============================================================
# AUTOGREEN MASTER WORKBOOK SIGNALS
# The workbook is maintained in Servejit/Signal and is read-only here.
# Only the AutoGreen worksheet is used for these Buy cards.
# ============================================================

SIGNAL_GITHUB_API = "https://api.github.com/repos/Servejit/Signal"
SIGNAL_GITHUB_BRANCH = "main"
SIGNAL_MASTER_PATH = "master.xlsx"
SIGNAL_GITHUB_TOKEN = str(
    st.secrets.get("SIGNAL_GITHUB_TOKEN", st.secrets.get("GITHUB_TOKEN", "")) or ""
).strip()


def get_autogreen_master_bytes():
    """Fetch private Signal/master.xlsx through GitHub's authenticated API."""
    if not SIGNAL_GITHUB_TOKEN:
        raise RuntimeError(
            "SIGNAL_GITHUB_TOKEN is missing from Streamlit Secrets. "
            "Add a GitHub token with Contents: Read access to Servejit/Signal."
        )

    headers = {
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "Authorization": f"Bearer {SIGNAL_GITHUB_TOKEN}",
    }
    response = requests.get(
        f"{SIGNAL_GITHUB_API}/contents/{SIGNAL_MASTER_PATH}",
        params={"ref": SIGNAL_GITHUB_BRANCH},
        headers=headers,
        timeout=30,
    )
    if response.status_code == 404:
        raise RuntimeError(
            "GitHub could not access Servejit/Signal/master.xlsx. Check that "
            "SIGNAL_GITHUB_TOKEN has Contents: Read access to the private repository."
        )
    response.raise_for_status()
    payload = response.json()
    encoded = payload.get("content", "")
    if encoded:
        workbook_bytes = base64.b64decode(encoded)
    else:
        blob_sha = payload.get("sha")
        if not blob_sha:
            raise ValueError("GitHub did not return master.xlsx content or blob SHA.")
        blob_response = requests.get(
            f"{SIGNAL_GITHUB_API}/git/blobs/{blob_sha}",
            headers={
                **headers,
                "Accept": "application/vnd.github.raw",
            },
            timeout=60,
        )
        blob_response.raise_for_status()
        workbook_bytes = blob_response.content

    if not workbook_bytes:
        raise ValueError("GitHub master.xlsx is empty.")
    return workbook_bytes


def _autogreen_number(value):
    """Convert Excel numeric cells and formatted numeric strings safely."""
    if value is None:
        return None
    if isinstance(value, bool):
        return None
    try:
        cleaned = str(value).replace(",", "").replace("%", "").strip()
        if not cleaned or cleaned in {"-", "—", "N/A", "NA"}:
            return None
        return float(cleaned)
    except (TypeError, ValueError):
        return None


def _autogreen_db_parenthetical_number(value):
    """Read the final numeric value in parentheses from the formatted D-B cell."""
    if value is None or isinstance(value, bool):
        return None
    numeric_value = _autogreen_number(value)
    if numeric_value is not None:
        return numeric_value
    cleaned = str(value).replace(",", "")
    matches = re.findall(r"\(\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s*\)", cleaned)
    if not matches:
        return None
    try:
        return float(matches[-1])
    except (TypeError, ValueError):
        return None


def load_autogreen_buy_signals():
    """
    Read only AutoGreen from the GitHub master workbook.
    Base: both volume expansions > 100000 and Vol.Expand 1 > Vol.Expand 2.
    Colour priority: Green (4 negative), Blue (first 3 negative),
    Purple (first 2 negative).
    """
    workbook_bytes = get_autogreen_master_bytes()
    workbook = load_workbook(
        io.BytesIO(workbook_bytes),
        data_only=True,
        read_only=True
    )
    try:
        sheet = next(
            (ws for ws in workbook.worksheets
             if ws.title.strip().casefold() == "autogreen"),
            None
        )
        if sheet is None:
            raise ValueError(
                "The GitHub workbook does not contain an AutoGreen worksheet."
            )

        def header_key(value):
            return " ".join(str(value or "").strip().casefold().split())

        required = [
            "Vol.Expand 1", "Vol.Expand 2",
            "%Chg.1", "%Chg.2", "%Chg.3", "%Chg.4"
        ]
        wanted = {header_key(name): name for name in required}
        header_row_number = None
        column_map = {}

        for row_number, row in enumerate(
            sheet.iter_rows(
                min_row=1,
                max_row=min(sheet.max_row or 1, 30),
                values_only=True
            ),
            start=1
        ):
            found = {
                header_key(value): index
                for index, value in enumerate(row)
                if value is not None
            }
            if all(key in found for key in wanted):
                header_row_number = row_number
                column_map = {
                    original: found[normalized]
                    for normalized, original in wanted.items()
                }
                symbol_index = next(
                    (
                        found[header_key(label)]
                        for label in ("Symbol", "NSE Symbol", "Ticker")
                        if header_key(label) in found
                    ),
                    0
                )
                break

        if header_row_number is None:
            raise ValueError(
                "Could not find all six required headings in AutoGreen."
            )

        base_qualified = []
        signals = []
        for row in sheet.iter_rows(
            min_row=header_row_number + 1,
            values_only=True
        ):
            if not row or symbol_index >= len(row):
                continue
            symbol = normalize_stock_symbol(row[symbol_index])
            if not symbol:
                continue

            values = {}
            for name in required:
                index = column_map[name]
                values[name] = (
                    _autogreen_number(row[index])
                    if index < len(row) else None
                )

            vol1 = values["Vol.Expand 1"]
            vol2 = values["Vol.Expand 2"]
            if (
                vol1 is None or vol2 is None
                or vol1 <= 100000 or vol2 <= 100000
                or vol1 <= vol2
            ):
                continue

            record = {
                "Symbol": symbol,
                "Vol.Expand 1": vol1,
                "Vol.Expand 2": vol2,
                "%Chg.1": values["%Chg.1"],
                "%Chg.2": values["%Chg.2"],
                "%Chg.3": values["%Chg.3"],
                "%Chg.4": values["%Chg.4"],
            }
            base_qualified.append(record)

            changes = [
                values["%Chg.1"], values["%Chg.2"],
                values["%Chg.3"], values["%Chg.4"]
            ]
            if all(value is not None and value < 0 for value in changes):
                record["Signal"] = "GREEN BUY"
                record["Card Colour"] = "Green"
                signals.append(record)
            elif all(
                value is not None and value < 0 for value in changes[:3]
            ):
                record["Signal"] = "BLUE BUY"
                record["Card Colour"] = "Blue"
                signals.append(record)
            elif all(
                value is not None and value < 0 for value in changes[:2]
            ):
                record["Signal"] = "PURPLE BUY"
                record["Card Colour"] = "Purple"
                signals.append(record)

        priority = {"Green": 0, "Blue": 1, "Purple": 2}
        signals.sort(key=lambda item: (
            priority[item["Card Colour"]],
            item["Symbol"]
        ))
        return base_qualified, signals
    finally:
        workbook.close()


@st.cache_data(ttl=300, max_entries=1, show_spinner=False)
def load_stock_of_the_day_rows(db_threshold=-2.0):
    """Use AutoGreen volume-qualified stocks only; never apply the High - 0.20% layer."""
    # Take the base volume-expansion list, not colour-qualified or High - 0.20% candidates.
    base_rows, _ = load_autogreen_buy_signals()
    if not base_rows:
        return []

    workbook_bytes = get_autogreen_master_bytes()
    workbook = load_workbook(
        io.BytesIO(workbook_bytes),
        data_only=True,
        read_only=True
    )
    try:
        def header_key(value):
            return " ".join(str(value or "").strip().casefold().split())

        # The two O2H readings are the first two columns immediately after
        # Vol.Expand 2 on the AutoGreen worksheet.
        auto_sheet = next(
            (ws for ws in workbook.worksheets
             if ws.title.strip().casefold() == "autogreen"),
            None
        )
        if auto_sheet is None:
            raise ValueError("The Master workbook does not contain an AutoGreen worksheet.")

        auto_header_row = None
        auto_symbol_col = 0
        o2h_columns = []
        for row_number, row in enumerate(
            auto_sheet.iter_rows(
                min_row=1,
                max_row=min(auto_sheet.max_row or 1, 30),
                values_only=True
            ),
            start=1
        ):
            found = {
                header_key(value): index
                for index, value in enumerate(row)
                if value is not None
            }
            vol2_key = header_key("Vol.Expand 2")
            if vol2_key in found:
                auto_header_row = row_number
                vol2_col = found[vol2_key]
                auto_symbol_col = next(
                    (found[header_key(label)]
                     for label in ("Symbol", "NSE Symbol", "Ticker")
                     if header_key(label) in found),
                    0
                )
                o2h_columns = [vol2_col + 1, vol2_col + 2]
                break

        if auto_header_row is None:
            raise ValueError("Could not locate Vol.Expand 2 on the AutoGreen worksheet.")

        o2h_by_symbol = {}
        for row in auto_sheet.iter_rows(
            min_row=auto_header_row + 1,
            values_only=True
        ):
            if not row or auto_symbol_col >= len(row):
                continue
            symbol = normalize_stock_symbol(row[auto_symbol_col])
            if not symbol:
                continue
            readings = [
                _autogreen_number(row[index]) if index < len(row) else None
                for index in o2h_columns
            ]
            o2h_by_symbol[symbol] = readings

        # Read D-B and every dated O2L column from the AutoGreen worksheet.
        # Stock of the Day uses the same qualifying-stock source, not a separate Summary tab.
        # The first two O2L readings are mandatory; later readings are tie-break advantages.
        summary_sheet = auto_sheet

        summary_header_row = None
        summary_symbol_col = 0
        db_col = None
        o2l_columns = []
        o2l_headings = []
        db_key = header_key("16< D-B / Avg.4")
        for row_number, row in enumerate(
            summary_sheet.iter_rows(
                min_row=1,
                max_row=min(summary_sheet.max_row or 1, 30),
                values_only=True
            ),
            start=1
        ):
            found = {
                header_key(value): index
                for index, value in enumerate(row)
                if value is not None
            }
            dated_o2l = [
                (index, str(value).strip())
                for index, value in enumerate(row)
                if value is not None and re.search(r"\bO2L\b", str(value), re.IGNORECASE)
            ]
            if db_key in found and len(dated_o2l) >= 2:
                summary_header_row = row_number
                db_col = found[db_key]
                summary_symbol_col = next(
                    (found[header_key(label)]
                     for label in ("Symbol", "NSE Symbol", "Ticker")
                     if header_key(label) in found),
                    0
                )
                o2l_columns = [index for index, _ in dated_o2l]
                o2l_headings = [heading for _, heading in dated_o2l]
                break

        if summary_header_row is None:
            raise ValueError(
                "Could not locate 16< D-B / Avg.4 and at least two dated O2L columns on AutoGreen."
            )

        summary_by_symbol = {}
        for row in summary_sheet.iter_rows(
            min_row=summary_header_row + 1,
            values_only=True
        ):
            if not row or summary_symbol_col >= len(row):
                continue
            symbol = normalize_stock_symbol(row[summary_symbol_col])
            if not symbol:
                continue
            db_value = _autogreen_db_parenthetical_number(row[db_col]) if db_col < len(row) else None
            o2l_values = [
                _autogreen_number(row[index]) if index < len(row) else None
                for index in o2l_columns
            ]
            summary_by_symbol[symbol] = (db_value, o2l_values)

        matches = []
        for item in base_rows:
            symbol = item["Symbol"]
            o2h_values = o2h_by_symbol.get(symbol, [None, None])
            db_value, o2l_values = summary_by_symbol.get(
                symbol, (None, [None] * len(o2l_columns))
            )
            change1 = item.get("%Chg.1")
            change2 = item.get("%Chg.2")
            change3 = item.get("%Chg.3")
            change4 = item.get("%Chg.4")

            # Mandatory Stock of the Day conditions.
            if (
                len(o2h_values) < 2
                or any(value is None or value >= 0.50 for value in o2h_values[:2])
                or db_value is None or db_value >= db_threshold
                or len(o2l_values) < 2
                or any(value is None or value >= -1.50 for value in o2l_values[:2])
                or change1 is None or change1 >= 0
                or change2 is None or change2 >= 0
            ):
                continue

            # Tie-break score: each additional dated O2L below -1.50,
            # plus %Chg.3 and %Chg.4 below zero, adds one advantage.
            extra_o2l_advantages = sum(
                1 for value in o2l_values[2:]
                if value is not None and value < -1.50
            )
            extra_change_advantages = sum(
                1 for value in (change3, change4)
                if value is not None and value < 0
            )
            advantage_score = extra_o2l_advantages + extra_change_advantages

            record = {
                "Stock": symbol,
                "O2H 1": round(o2h_values[0], 2),
                "O2H 2": round(o2h_values[1], 2),
                "16< D-B / Avg.4": round(db_value, 2),
            }
            dated_o2l_position = 0
            for position, value in enumerate(o2l_values, start=1):
                heading = o2l_headings[position - 1] if position <= len(o2l_headings) else ""
                if re.search(r"sum\s*o2l[.\s]*10", heading, re.IGNORECASE):
                    column_name = "Sum O2L.10"
                else:
                    dated_o2l_position += 1
                    column_name = f"O2L.{dated_o2l_position}" + (f" · {heading}" if heading else "")
                record[column_name] = round(value, 2) if value is not None else "—"
            record["%Chg.1"] = round(change1, 2)
            record["%Chg.2"] = round(change2, 2)
            record["%Chg.3"] = round(change3, 2) if change3 is not None else "—"
            record["%Chg.4"] = round(change4, 2) if change4 is not None else "—"
            record["Extra Advantages"] = advantage_score
            matches.append((advantage_score, record))

        # Every stock that passes ALL mandatory conditions is a Stock of the Day.
        # Extra advantages only determine display order; they never disqualify a match.
        selected_matches = [
            record for _, record in sorted(
                matches,
                key=lambda item: (-item[0], item[1]["Stock"])
            )
        ]
        return selected_matches

    finally:
        workbook.close()


def render_stock_of_the_day():
    """Show only Stock of the Day BUY cards in the main signal-card presentation."""
    st.title("6thSense-Vardaan")

    try:
        master_bytes = get_autogreen_master_bytes()
        ranks = get_autogreen_sum_o2l_ranks(master_bytes)
    except Exception:
        st.error("Master File Not Available")
        return

    try:
        db_threshold = float(st.session_state.get("stock_day_db_threshold_value", -2.0))
        rows = load_stock_of_the_day_rows(db_threshold)
        if not rows:
            st.markdown(
                """
                <style>
                @keyframes sotdNoSignalShimmer {
                    0% { background-position:200% 50%; }
                    100% { background-position:-100% 50%; }
                }
                @keyframes sotdNoSignalFloat {
                    0%,100% { transform:translateY(0); }
                    50% { transform:translateY(-3px); }
                }
                @keyframes sotdNoSignalGlow {
                    0% { transform:translate3d(-35%,-5%,0) rotate(18deg); opacity:0; }
                    18% { opacity:.62; }
                    52% { transform:translate3d(115%,3%,0) rotate(18deg); opacity:.42; }
                    82% { opacity:.18; }
                    100% { transform:translate3d(190%,-4%,0) rotate(18deg); opacity:0; }
                }
                @media (prefers-reduced-motion: reduce) {
                    .stock-day-no-signal-card { animation:none !important; }
                }
                </style>
                <div class="stock-day-no-signal-card" style="
                    position:relative;overflow:hidden;isolation:isolate;
                    border:2px solid transparent;border-radius:21px;
                    padding:30px 23px 28px;margin:12px 0 18px;color:#202833;text-align:center;
                    background:
                        linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,
                        linear-gradient(110deg,#202833 0%,#687383 24%,#FFFFFF 38%,
                                        #111827 54%,#202833 76%,#687383 100%) border-box;
                    background-size:100% 100%,300% 300%;
                    animation:sotdNoSignalShimmer 5.8s linear infinite,
                              sotdNoSignalFloat 3.1s ease-in-out infinite;
                    box-shadow:0 0 0 1px rgba(255,255,255,.92) inset,
                               0 14px 30px rgba(17,24,39,.18),
                               0 0 22px rgba(17,24,39,.30),
                               0 0 38px rgba(104,115,131,.22);
                ">
                  <div aria-hidden="true" style="position:absolute;top:-55%;left:-72%;width:72%;height:210%;
                              pointer-events:none;z-index:0;filter:blur(7px);
                              background:linear-gradient(105deg,transparent 0%,rgba(17,24,39,.13) 25%,
                              rgba(17,24,39,.72) 48%,rgba(255,255,255,.9) 58%,transparent 82%);
                              animation:sotdNoSignalGlow 6s ease-in-out infinite;"></div>
                  <div style="position:absolute;right:-35px;top:-75px;width:205px;height:205px;
                              border-radius:50%;background:radial-gradient(circle,rgba(17,24,39,.18),transparent 70%);
                              pointer-events:none;"></div>
                  <div style="position:relative;z-index:1;text-align:center;font-size:11px;
                              font-weight:900;letter-spacing:3px;color:#202833;margin:0 0 15px;">
                    6THSENSE · VARDAAN
                  </div>
                  <div style="position:relative;z-index:1;font-size:clamp(30px,5vw,42px);
                              line-height:1.1;font-weight:950;letter-spacing:2.2px;
                              margin:13px 0 10px;color:#111827;text-shadow:0 2px 12px rgba(17,24,39,.12);">
                    NO SIGNAL
                  </div>
                  <div style="position:relative;z-index:1;width:74px;height:1px;margin:0 auto 15px;
                              background:linear-gradient(90deg,transparent,#202833,transparent);"></div>
                  <div style="position:relative;z-index:1;font-size:12px;font-weight:900;
                              letter-spacing:1.65px;color:#374151;">
                    WAIT FOR A QUALIFIED SETUP
                  </div>
                  <div style="position:relative;z-index:1;font-size:13px;line-height:1.65;
                              font-weight:600;color:#596273;margin-top:10px;">
                    No Stock of the Day BUY signal is available.
                  </div>
                </div>
                """,
                unsafe_allow_html=True
            )



        symbols = list(dict.fromkeys(
            normalize_stock_symbol(row.get("Stock"))
            for row in rows
            if normalize_stock_symbol(row.get("Stock"))
        ))
        with st.spinner("Preparing Stock of the Day BUY card..."):
            quotes = get_autogreen_live_prices(tuple(symbols))
    except Exception as error:
        st.error(f"Stock of the Day could not be loaded: {error}")
        return

    # This page contains only the Stock of the Day signal, so its metallic
    # black/silver card is the topmost card and controls the page theme.
    st.markdown(
        """
        <style>
        .stApp {
            background:radial-gradient(ellipse at 10% 0%,#DDE2E8 0%,transparent 44%),
                       linear-gradient(135deg,#E7EAEE 0%,#D7DCE2 52%,#E9ECEF 100%) !important;
        }
        [data-testid="stHeader"] { background:rgba(225,229,235,.94) !important; }
        [data-testid="stSidebar"] {
            background:linear-gradient(180deg,#D8DDE4 0%,#EEF0F3 100%) !important;
            border-right:1px solid #929EAC !important;
        }
        h1,h2,h3 { color:#202833 !important; }
        @keyframes stockDayCardShimmer {
            0% { background-position:200% 50%; }
            100% { background-position:-100% 50%; }
        }
        @keyframes stockDayCardFloat {
            0%,100% { transform:translateY(0); }
            50% { transform:translateY(-3px); }
        }
        @keyframes stockDayCardGlow {
            0% { transform:translate3d(-35%,-5%,0) rotate(18deg); opacity:0; }
            18% { opacity:.62; }
            52% { transform:translate3d(115%,3%,0) rotate(18deg); opacity:.42; }
            82% { opacity:.18; }
            100% { transform:translate3d(190%,-4%,0) rotate(18deg); opacity:0; }
        }
        @media (prefers-reduced-motion: reduce) {
            .stock-day-buy-card { animation:none !important; }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    # Alphabet-based Stock of the Day border and shimmer colours.
    stock_day_palettes = {
        "A": ("#A8D5A2", "#63C96B"), "B": ("#A9DDF2", "#4FC4F2"),
        "C": ("#CBB6EF", "#9869E2"), "D": ("#E8D18A", "#D5A32D"),
        "E": ("#E9A0B4", "#E34F78"), "F": ("#F4B8D0", "#EF68A5"),
        "G": ("#A9CFF5", "#4D9FF0"), "H": ("#C9D0D8", "#8794A5"),
        "I": ("#A8DED4", "#36BDA8"), "J": ("#A9E8E8", "#35C5C5"),
        "K": ("#B8E2C1", "#55BD79"), "L": ("#D5E99B", "#A6CE39"),
        "M": ("#F4C7A1", "#ED9858"), "N": ("#F3B0A8", "#E87868"),
        "O": ("#F2D39A", "#E5AD3D"), "P": ("#D8C4F0", "#A987DF"),
        "Q": ("#BCC8F2", "#8298E8"), "R": ("#E7A0A8", "#D94C60"),
        "S": ("#C6D9B7", "#8EBA68"), "T": ("#A5DDD9", "#42B8B0"),
        "U": ("#D0B9EF", "#9A72D8"), "V": ("#D7B5D9", "#AE70B5"),
        "W": ("#B7DDF0", "#68B9E5"), "X": ("#D5D9DE", "#9AA4B0"),
        "Y": ("#F1E3A0", "#D9C13F"), "Z": ("#B9C5D4", "#71839B"),
    }

    # Match the page heading to the same alphabet-based text colour as the BUY card.
    if rows:
        heading_symbol = normalize_stock_symbol(rows[0].get("Stock", "A"))
        heading_colour = stock_day_palettes.get(heading_symbol[:1].upper(), stock_day_palettes["A"])[1]
    else:
        # With no qualifying stock, match the black text used by the NO SIGNAL card.
        heading_colour = "#202833"
    st.markdown(
        f"""
        <style>
        h1 {{ color:{heading_colour} !important; }}
        </style>
        """,
        unsafe_allow_html=True
    )

    rendered = 0
    for row in rows:
        symbol = normalize_stock_symbol(row.get("Stock"))
        border_colour, shimmer_colour = stock_day_palettes.get(symbol[:1].upper(), stock_day_palettes["A"])
        quote = quotes.get(symbol, {})
        day_open = quote.get("open")
        day_high = quote.get("high")
        live_price = quote.get("price")
        if day_open is None or day_open <= 0 or day_high is None or day_high <= 0:
            continue

        # Stock of the Day uses the actual session Open as entry.
        entry = day_open
        target1 = entry * 1.005
        target2 = entry * 1.01
        target1_achieved = day_high >= target1
        # Intraday High confirms targets already reached; live price takes
        # precedence when it is currently trading above Target 2.
        if live_price is not None and live_price > target2:
            status_text = "TODAY TARGET HIT · PRICE RUNNING ABOVE TARGET PRICE"
        elif day_high > target2:
            status_text = "TARGET 1 & 2 ACHIEVED FOR THE DAY"
        elif day_high >= target2:
            status_text = "TARGET 2 ACHIEVED"
        elif day_high >= target1:
            status_text = "TARGET 1 ACHIEVED"
        else:
            status_text = "TARGETS PENDING"

        rank_value = ranks.get(symbol)
        rank_text = f"{rank_value:g}" if isinstance(rank_value, (int, float)) else "—"

        card_html = f"""
        <div class="stock-day-buy-card" style="
            position:relative;overflow:hidden;isolation:isolate;
            border:2px solid transparent;border-radius:21px;
            padding:23px 23px 22px;margin:12px 0 18px;color:#253047;
            background:
                linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,
                linear-gradient(110deg,{border_colour} 0%,{shimmer_colour} 24%,#FFFFFF 38%,
                                {shimmer_colour} 54%,{border_colour} 76%,{shimmer_colour} 100%) border-box;
            background-size:100% 100%,300% 300%;
            animation:stockDayCardShimmer 5.8s linear infinite,
                      stockDayCardFloat 3.1s ease-in-out infinite;
            box-shadow:0 0 0 1px rgba(255,255,255,.92) inset,
                       0 14px 30px {shimmer_colour}30,
                       0 0 22px {shimmer_colour}88,
                       0 0 38px {shimmer_colour}55;
        ">
          <div aria-hidden="true" style="position:absolute;top:-55%;left:-72%;width:72%;height:210%;
                      pointer-events:none;z-index:0;filter:blur(7px);
                      background:linear-gradient(105deg,transparent 0%,{shimmer_colour}22 25%,
                      {shimmer_colour}CC 48%,rgba(255,255,255,.72) 58%,transparent 82%);
                      animation:stockDayCardGlow 6s ease-in-out infinite;"></div>
          <div style="position:absolute;right:-35px;top:-75px;width:205px;height:205px;
                      border-radius:50%;background:radial-gradient(circle,{shimmer_colour}66,transparent 70%);
                      pointer-events:none;"></div>
          <div style="position:relative;z-index:1;text-align:center;font-size:11px;
                      font-weight:900;letter-spacing:3px;
                      color:{shimmer_colour};margin:0 0 15px;">
            6THSENSE · VARDAAN
          </div>
          <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;
                      position:relative;z-index:1;">
            <div style="font-size:1.78rem;line-height:1.15;font-weight:950;letter-spacing:.7px;
                        color:{shimmer_colour};overflow-wrap:anywhere;">{html.escape(symbol)}</div>
            <div style="font-size:.72rem;font-weight:950;letter-spacing:1.15px;padding:8px 12px;
                        border:1px solid {border_colour};border-radius:999px;color:{shimmer_colour};
                        background:linear-gradient(135deg,#F1F8EA,#FFF9E8);white-space:nowrap;">
              BUY SIGNAL
            </div>
          </div>
          <div style="font-size:.72rem;font-weight:900;letter-spacing:1.25px;color:#C5CFDC;
                      margin-top:9px;position:relative;z-index:1;color:{shimmer_colour};">RANK {rank_text}</div>
          <div style="height:1px;background:linear-gradient(90deg,transparent,#C69B43,
                      rgba(255,236,184,.95),#C69B43,transparent);margin:19px 0 17px;
                      position:relative;z-index:1;"></div>
          <div style="display:grid;grid-template-columns:repeat(3,minmax(0,1fr));
                      align-items:center;gap:10px;position:relative;z-index:1;">
            <div style="min-width:0;">
              <div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">ENTRY PRICE</div>
              <div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};
                          margin-top:7px;white-space:nowrap;">₹{entry:,.2f}</div>
            </div>
            <div style="min-width:0;">
              <div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">TARGET 1</div>
              <div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};
                          margin-top:7px;white-space:nowrap;">₹{target1:,.2f}</div>
            </div>
            <div style="min-width:0;">
              <div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">TARGET 2</div>
              <div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};
                          margin-top:7px;white-space:nowrap;">₹{target2:,.2f}</div>
            </div>
          </div>
          <div style="position:relative;z-index:1;margin-top:18px;padding:10px 12px;
                      border:1px solid {border_colour};border-radius:11px;
                      background:#FFFFFF;text-align:center;
                      background-clip:padding-box;
                      color:{shimmer_colour};font-size:.72rem;font-weight:950;letter-spacing:.8px;">
            {html.escape(status_text)}
          </div>
        </div>
        """
        st.markdown(card_html, unsafe_allow_html=True)
        rendered += 1

    # Do not show a separate message when live price data is unavailable.

    # Match MORE DETAILS text and its border to the Stock of the Day card palette.
    if rendered > 0:
        details_symbol = normalize_stock_symbol(rows[0].get("Stock", "A"))
        details_border, details_shimmer = stock_day_palettes.get(details_symbol[:1].upper(), stock_day_palettes["A"])
        st.markdown(
            f"""
            <style>
            @keyframes sotdDetailsBorderOnlyFlow {{
                0% {{ background-position:0% 50%; }}
                100% {{ background-position:300% 50%; }}
            }}
            div[data-testid="stExpander"] {{
                background:linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,
                           linear-gradient(110deg,{details_border} 0%,{details_shimmer} 24%,#FFFFFF 38%,
                                           {details_shimmer} 54%,{details_border} 76%,{details_shimmer} 100%) border-box !important;
                border:2px solid transparent !important;
                border-radius:14px !important;
                box-shadow:0 8px 22px {details_shimmer}30,0 0 18px {details_shimmer}35,
                           inset 0 1px 0 rgba(255,255,255,.95);
                overflow:hidden;
                background-size:100% 100%,300% 300% !important;
                animation:sotdDetailsBorderOnlyFlow 7s linear infinite !important;
            }}
            div[data-testid="stExpander"] summary {{
                background:none !important;
                color:{details_shimmer} !important;
                font-weight:950 !important;
                letter-spacing:1.2px;
                padding:12px 16px !important;
            }}
            div[data-testid="stExpander"] summary:hover {{
                background:{details_shimmer}12 !important;
            }}
            div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
                background:#FFFFFF;
                border-top:1px solid {details_shimmer}88;
                color:{details_shimmer};
                padding:12px 16px 16px;
            }}
            div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] p,
            div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] li {{
                color:{details_shimmer} !important;
            }}
            div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] strong {{
                color:{details_shimmer} !important;
            }}
            </style>
            """,
            unsafe_allow_html=True
        )
        with st.expander("MORE DETAILS", expanded=False):
            # Present selected stock metrics in a compact black/silver data panel.
            if rows:
                table_headers = list(rows[0].keys())
            core_columns = [column for column in ("O2H 1", "O2H 2", "16< D-B / Avg.4") if column in table_headers]
            o2l_columns = [column for column in table_headers if column == "Sum O2L.10" or re.match(r"^O2L\.[1-4](?:\s|$)", column)]
            change_columns = [column for column in table_headers if column.startswith("%Chg.")]

            def metric_tiles(row, columns):
                return "".join(
                    f'<div class="sotd-metric-tile"><div class="sotd-metric-label">{html.escape(str(column))}</div>'
                    f'<div class="sotd-metric-value">{html.escape("" if row.get(column) is None else str(row.get(column)))}</div></div>'
                    for column in columns
                )

            stock_panels = []
            for row in rows:
                raw_symbol = normalize_stock_symbol(row.get("Stock", "STOCK"))
                symbol = html.escape(raw_symbol)
                panel_border, panel_shimmer = stock_day_palettes.get(raw_symbol[:1].upper(), stock_day_palettes["A"])
                groups = []
                for heading, columns, css_class in (
                    ("CORE READINGS", core_columns, "core"),
                    ("O2L HISTORY", o2l_columns, "history"),
                    ("% CHANGE HISTORY", change_columns, "change"),
                ):
                    if columns:
                        groups.append(
                            f'<section class="sotd-metric-group {css_class}">'
                            f'<div class="sotd-group-heading">{heading}</div>'
                            f'<div class="sotd-metric-grid">{metric_tiles(row, columns)}</div></section>'
                        )
                stock_panels.append(
                    f'<article class="sotd-stock-panel" style="--sotd-border:{panel_border};--sotd-shimmer:{panel_shimmer};--sotd-text:{panel_shimmer};background:linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,linear-gradient(110deg,{panel_border} 0%,{panel_shimmer} 24%,#FFFFFF 38%,{panel_shimmer} 54%,{panel_border} 76%,{panel_shimmer} 100%) border-box;">'
                    f'<div aria-hidden="true" class="sotd-static-glow" style="background:radial-gradient(circle,{panel_shimmer}45,transparent 72%);"></div><div class="sotd-stock-panel-content"><div class="sotd-stock-panel-head">'
                    f'<span class="sotd-stock-symbol" style="color:{panel_shimmer};">{symbol}</span>'
                    f'<span class="sotd-stock-tag" style="color:{panel_shimmer};">STOCK OF THE DAY</span></div>'
                    f'{"".join(groups)}</div></article>'
                )

            table_symbol = normalize_stock_symbol(rows[0].get("Stock", "STOCK")) if rows else "STOCK"
            table_border, table_shimmer = stock_day_palettes.get(table_symbol[:1].upper(), stock_day_palettes["A"])
            st.markdown(
                f"""
                <div class="sotd-table-heading" style="--sotd-table-border:{table_border};--sotd-table-shimmer:{table_shimmer};background:linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,linear-gradient(110deg,{table_border} 0%,{table_shimmer} 24%,#FFFFFF 38%,{table_shimmer} 54%,{table_border} 76%,{table_shimmer} 100%) border-box;">
                  <div>
                    <div class="sotd-table-title">STOCK DATA</div>
                    <div class="sotd-table-subtitle">All qualifying readings · grouped for quick review</div>
                  </div>
                  <div class="sotd-table-count">{len(rows)} STOCK{"S" if len(rows) != 1 else ""}</div>
                </div>
                <style>
                @keyframes sotdBorderOnlyFlow {{
                    0% {{ background-position:0% 50%; }}
                    100% {{ background-position:300% 50%; }}
                }}
                .sotd-table-heading,
                .sotd-stock-panel {{
                    background-size:100% 100%,300% 300% !important;
                    animation:sotdBorderOnlyFlow 7s linear infinite !important;
                }}
                .sotd-table-heading {{
                    position:relative;overflow:hidden;isolation:isolate;
                    display:flex;align-items:center;justify-content:space-between;gap:12px;
                    margin:18px 0 12px;padding:14px 16px;border:2px solid transparent;
                    border-radius:18px;background-clip:padding-box;
                    box-shadow:0 0 0 1px rgba(255,255,255,.92) inset,
                               0 14px 30px var(--sotd-table-shimmer)30,
                               0 0 22px var(--sotd-table-shimmer)55,
                               0 0 34px var(--sotd-table-shimmer)30;
                }}
                .sotd-table-title {{position:relative;z-index:1;color:var(--sotd-table-shimmer);font-size:.86rem;font-weight:950;letter-spacing:1.65px;}}
                .sotd-table-subtitle {{position:relative;z-index:1;color:var(--sotd-table-shimmer);font-size:.72rem;font-weight:600;margin-top:4px;}}
                .sotd-table-count {{position:relative;z-index:1;flex:0 0 auto;padding:6px 10px;border-radius:999px;
                    border:1px solid var(--sotd-table-border);background:#FFFFFF;
                    color:var(--sotd-table-shimmer);font-size:.65rem;font-weight:900;letter-spacing:.8px;}}
                .sotd-stock-panel {{position:relative;isolation:isolate;overflow:hidden;margin:0 0 16px;padding:15px;
                    border:2px solid transparent;border-radius:19px;background-clip:padding-box;
                    box-shadow:0 0 0 1px rgba(255,255,255,.92) inset,
                               0 14px 30px var(--sotd-shimmer)30,
                               0 0 22px var(--sotd-shimmer)66,
                               0 0 36px var(--sotd-shimmer)40;}}
                .sotd-static-glow {{position:absolute;right:-38px;top:-76px;width:190px;height:190px;
                    border-radius:50%;pointer-events:none;z-index:0;}}
                .sotd-stock-panel-content {{position:relative;z-index:1;}}
                .sotd-stock-panel-head {{display:flex;align-items:center;justify-content:space-between;gap:10px;
                    padding:0 2px 10px;border-bottom:1px solid var(--sotd-border);margin-bottom:12px;}}
                .sotd-stock-symbol {{color:var(--sotd-text,#76539A);font-size:1rem;font-weight:950;letter-spacing:.8px;}}
                .sotd-stock-tag {{color:var(--sotd-text,#76539A);font-size:.58rem;font-weight:900;letter-spacing:1px;
                    padding:5px 8px;border:1px solid var(--sotd-border,#D5B66D);border-radius:20px;background:#FFFFFF;}}
                .sotd-metric-group {{margin-top:9px;}}
                .sotd-group-heading {{color:var(--sotd-text,#76539A);font-size:.68rem;font-weight:950;letter-spacing:1.25px;
                    margin:0 0 7px 2px;}}
                .sotd-metric-grid {{display:grid;grid-template-columns:repeat(auto-fit,minmax(92px,1fr));gap:7px;}}
                .sotd-metric-tile {{min-width:0;padding:9px 8px;border:1px solid var(--sotd-border);
                    border-radius:10px;background:linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 72%,var(--sotd-border)18);}}
                .sotd-metric-label {{color:var(--sotd-text,#806C48);font-size:.66rem;font-weight:850;line-height:1.3;
                    overflow-wrap:anywhere;}}
                .sotd-metric-value {{color:var(--sotd-text,#24466F);font-size:.95rem;font-weight:900;font-variant-numeric:tabular-nums;
                    margin-top:5px;overflow-wrap:anywhere;}}
                .sotd-metric-group.history .sotd-metric-tile {{background:#FFFFFF;}}
                .sotd-metric-group.change .sotd-metric-value {{color:var(--sotd-text,#21835B);}}
                @media(max-width:640px) {{
                    .sotd-table-heading {{padding:11px 12px;}}
                    .sotd-table-subtitle {{font-size:.64rem;}}
                    .sotd-stock-panel {{padding:11px;}}
                    .sotd-metric-grid {{grid-template-columns:repeat(auto-fit,minmax(82px,1fr));gap:6px;}}
                    .sotd-metric-tile {{padding:7px 5px;}}
                    .sotd-metric-label {{font-size:.61rem;}}
                    .sotd-metric-value {{font-size:.84rem;}}
                }}
                </style>
                {''.join(stock_panels)}
                """,
                unsafe_allow_html=True
            )

    if not rows:
        st.markdown(
            """
            <style>
            @keyframes sotdThresholdBorderFlow {
                0% { background-position:0% 50%; }
                100% { background-position:300% 50%; }
            }
            div[data-testid="stExpander"] {
                background:linear-gradient(135deg,#FFFFFF 0%,#FFFFFF 48%,#FFFFFF 100%) padding-box,
                           linear-gradient(110deg,#202833 0%,#687383 24%,#FFFFFF 38%,
                                           #111827 54%,#202833 76%,#687383 100%) border-box !important;
                border:2px solid transparent !important;
                border-radius:14px !important;
                box-shadow:0 8px 22px rgba(17,24,39,.18),0 0 18px rgba(32,40,51,.25),
                           inset 0 1px 0 rgba(255,255,255,.95);
                overflow:hidden;
                background-size:100% 100%,300% 300% !important;
                animation:sotdThresholdBorderFlow 7s linear infinite !important;
            }
            div[data-testid="stExpander"] summary {
                background:none !important;
                color:#202833 !important;
                font-weight:950 !important;
                letter-spacing:1.2px;
                padding:12px 16px !important;
            }
            div[data-testid="stExpander"] summary:hover {
                background:rgba(32,40,51,.08) !important;
            }
            div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
                background:#FFFFFF;
                border-top:1px solid #687383;
                color:#202833;
                padding:12px 16px 16px;
            }
            div[data-testid="stExpander"] label,
            div[data-testid="stExpander"] label p,
            div[data-testid="stExpander"] [data-testid="stMarkdownContainer"],
            div[data-testid="stExpander"] [data-testid="stMarkdownContainer"] p,
            div[data-testid="stExpander"] [data-testid="stTickBar"] div,
            div[data-testid="stExpander"] [data-testid="stThumbValue"],
            div[data-testid="stExpander"] [role="slider"] {
                color:#000000 !important;
            }
            div[data-testid="stExpander"] [data-baseweb="slider"] div[role="slider"] {
                background-color:#000000 !important;
                border-color:#000000 !important;
            }
            div[data-testid="stExpander"] [data-baseweb="slider"] div[data-testid="track"],
            div[data-testid="stExpander"] [data-baseweb="slider"] div[data-testid="track"] > div {
                background-color:#000000 !important;
            }
            div[data-testid="stExpander"] [data-baseweb="slider"] div[aria-hidden="true"] {
                border-color:#000000 !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
        with st.expander("THRESHOLD FREQUENCY", expanded=False):
            selected_db_threshold = st.slider(
                "D-B threshold",
                min_value=-5.0,
                max_value=0.0,
                value=db_threshold,
                step=0.5,
                format="%.1f",
                key="stock_day_db_threshold_slider",
                help="Adjust the D-B qualification threshold when the NO SIGNAL card is displayed."
            )
            st.session_state["stock_day_db_threshold_value"] = selected_db_threshold
    
def get_autogreen_sum_o2l_ranks(workbook_bytes):
    """Rank all Master stocks by Sum O2L.10; lowest value receives rank 1."""
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=True, read_only=True)
    try:
        def key(value):
            return " ".join(str(value or "").strip().casefold().split())

        selected_sheet = None
        header_row = None
        symbol_col = 0
        rank_col = None
        for ws in sorted(workbook.worksheets,
                         key=lambda item: (item.title.casefold() != "summary", item.title.casefold())):
            for row_number, row in enumerate(
                ws.iter_rows(min_row=1, max_row=min(ws.max_row or 1, 30), values_only=True),
                start=1
            ):
                found = {key(value): index for index, value in enumerate(row) if value is not None}
                if key("Sum O2L.10") in found:
                    selected_sheet = ws
                    header_row = row_number
                    rank_col = found[key("Sum O2L.10")]
                    symbol_col = next(
                        (found[key(label)] for label in ("Symbol", "NSE Symbol", "Ticker")
                         if key(label) in found), 0
                    )
                    break
            if selected_sheet is not None:
                break

        if selected_sheet is None:
            raise ValueError("Could not find the Sum O2L.10 heading in the Master workbook.")

        values = []
        for row in selected_sheet.iter_rows(min_row=header_row + 1, values_only=True):
            if not row or symbol_col >= len(row):
                continue
            symbol = normalize_stock_symbol(row[symbol_col])
            raw = row[rank_col] if rank_col < len(row) else None
            number = _autogreen_number(raw)
            if symbol and number is not None:
                values.append((number, symbol))

        values.sort(key=lambda item: item[0])
        ranks = {}
        previous = None
        current_rank = 0
        for position, (value, symbol) in enumerate(values, start=1):
            if previous is None or value != previous:
                current_rank = position
            ranks[symbol] = current_rank
            previous = value
        return ranks
    finally:
        workbook.close()


@st.cache_data(ttl=10, max_entries=1, show_spinner=False)
def get_autogreen_live_prices(symbols):
    """Fetch live NSE quotes, with Yahoo fallbacks for delayed/missing intraday bars."""
    import yfinance as yf
    from concurrent.futures import ThreadPoolExecutor, as_completed

    def fetch_one(symbol):
        empty = {
            "price": None, "open": None, "high": None, "low": None,
            "previous_close": None, "change_pct": None
        }
        try:
            ticker = yf.Ticker(f"{symbol}.NS")
            history = ticker.history(
                period="1d", interval="1m", auto_adjust=False,
                actions=False, prepost=False, timeout=15
            )

            # Yahoo sometimes returns no 1-minute bars; retry with 5-minute bars.
            if history is None or history.empty or "Close" not in history:
                history = ticker.history(
                    period="5d", interval="5m", auto_adjust=False,
                    actions=False, prepost=False, timeout=15
                )
            if history is None or history.empty or "Close" not in history:
                return symbol, empty

            history = history.dropna(subset=["Close"])
            if history.empty:
                return symbol, empty

            # The 5-day fallback must use only its latest session for today's OHLC.
            latest_session_date = history.index[-1].date()
            current_session = history[
                history.index.date == latest_session_date
            ]
            if current_session.empty:
                return symbol, empty

            closes = current_session["Close"].dropna()
            highs = current_session["High"].dropna() if "High" in current_session else []
            opens = current_session["Open"].dropna() if "Open" in current_session else []
            lows = current_session["Low"].dropna() if "Low" in current_session else []

            price = _autogreen_number(closes.iloc[-1]) if not closes.empty else None
            high = _autogreen_number(highs.max()) if len(highs) else None
            open_price = _autogreen_number(opens.iloc[0]) if len(opens) else None
            low = _autogreen_number(lows.min()) if len(lows) else None

            daily = ticker.history(
                period="5d", interval="1d", auto_adjust=False,
                actions=False, prepost=False, timeout=15
            )
            # Keep today's OHLC internally consistent: prefer all four values
            # from the same current-session daily candle, not a mix of daily and
            # 1-minute bars. Intraday OHLC above remains the fallback.
            if daily is not None and not daily.empty:
                daily_session_rows = daily[
                    daily.index.date == latest_session_date
                ]
                if not daily_session_rows.empty:
                    daily_session = daily_session_rows.iloc[-1]
                    daily_open = _autogreen_number(daily_session.get("Open"))
                    daily_high = _autogreen_number(daily_session.get("High"))
                    daily_low = _autogreen_number(daily_session.get("Low"))
                    if daily_open is not None and daily_open > 0:
                        open_price = daily_open
                    if daily_high is not None and daily_high > 0:
                        high = daily_high
                    if daily_low is not None and daily_low > 0:
                        low = daily_low

            daily_closes = (
                daily["Close"].dropna()
                if daily is not None and not daily.empty and "Close" in daily
                else []
            )
            previous_close = (
                _autogreen_number(daily_closes.iloc[-2])
                if len(daily_closes) >= 2 else None
            )

            # If Yahoo supplies only one daily candle, use its explicit prior close.
            if previous_close is None:
                try:
                    fast_info = ticker.fast_info
                    previous_close = _autogreen_number(
                        fast_info.get("previous_close")
                    )
                except Exception:
                    previous_close = None

            change_pct = (
                ((price - previous_close) / previous_close) * 100
                if price is not None and previous_close is not None and previous_close > 0
                else None
            )
            return symbol, {
                "price": price, "open": open_price, "high": high, "low": low,
                "previous_close": previous_close, "change_pct": change_pct
            }
        except Exception:
            return symbol, empty

    results = {}
    unique_symbols = list(dict.fromkeys(symbols))
    if not unique_symbols:
        return results
    with ThreadPoolExecutor(max_workers=min(8, len(unique_symbols))) as pool:
        futures = [pool.submit(fetch_one, symbol) for symbol in unique_symbols]
        for future in as_completed(futures):
            symbol, quote = future.result()
            results[symbol] = quote
    return results

def render_accent_table(rows, accent, ink, tint, min_width=720):
    """Render a light, readable table with an accent inherited from the top signal card."""
    if not rows:
        return

    columns = list(rows[0].keys())
    header_html = "".join(
        f'<th style="padding:12px 14px;text-align:left;white-space:nowrap;'
        f'font-size:13px;font-weight:900;letter-spacing:.7px;color:{ink};'
        f'background:linear-gradient(180deg,{tint} 0%,#FFFFFF 100%);'
        f'border-bottom:2px solid {accent}66;">'
        f'{html.escape(str(column))}</th>'
        for column in columns
    )
    body_rows = []
    for row_number, record in enumerate(rows):
        cells = "".join(
            f'<td style="padding:11px 14px;white-space:nowrap;'
            f'font-size:13px;font-weight:650;color:#263342;'
            f'border-bottom:1px solid #E8ECF1;">'
            f'{html.escape(str(record.get(column, "—")))}</td>'
            for column in columns
        )
        row_background = "#FFFFFF" if row_number % 2 == 0 else tint
        body_rows.append(f'<tr style="background:{row_background};">{cells}</tr>')

    table_html = f"""
    <div class="signal-shimmer-surface" style="position:relative;padding:2px;border-radius:13px;overflow:hidden;
                background:linear-gradient(115deg,{accent}55 0%,#FFFFFF 22%,
                {accent}88 52%,#FFFFFF 78%,{accent}66 100%);
                box-shadow:0 5px 18px rgba(31,45,61,.09),0 0 0 1px {accent}22;">
      <div style="overflow-x:auto;border-radius:11px;background:#FFFFFF;">
        <table style="width:100%;min-width:{min_width}px;border-collapse:separate;
                      border-spacing:0;background:#FFFFFF;">
          <thead><tr>{header_html}</tr></thead>
          <tbody>{''.join(body_rows)}</tbody>
        </table>
      </div>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)


def render_qualified_stock_table(rows, accent, ink, tint):
    """Render qualified stocks in the same light adaptive palette as the top signal card."""
    render_accent_table(rows, accent, ink, tint, min_width=720)


def render_autogreen_dashboard():
    """Select BUY and CONFIRMATION independently from all four base-qualified target stocks."""
    st.title("6thSense-Vardaan")


    try:
        with st.spinner("Reading AutoGreen and Master ranking data from GitHub..."):
            base_rows, signal_rows = load_autogreen_buy_signals()
            master_bytes = get_autogreen_master_bytes()
            sum_o2l_ranks = get_autogreen_sum_o2l_ranks(master_bytes)
    except Exception:
        st.error("Master File Not Available")
        return

    ranked_candidates = []
    for item in signal_rows:
        rank = sum_o2l_ranks.get(item["Symbol"])
        ranked_candidates.append({**item, "Sum O2L.10 Rank": rank})

    # Fetch live price and today's High for EVERY AutoGreen-qualified stock.
    # The High - 0.10% metric uses the full qualified universe, not only colour-qualified stocks.
    quote_map = {}
    if base_rows:
        with st.spinner("Fetching live NSE prices and today's highs from Yahoo Finance..."):
            quote_map = get_autogreen_live_prices(tuple(item["Symbol"] for item in base_rows))

    high_010_qualified_count = 0
    for item in base_rows:
        quote = quote_map.get(item["Symbol"], {})
        live_price = quote.get("price")
        day_high = quote.get("high")
        if live_price is None or day_high is None or day_high <= 0:
            continue
        high_gap_pct = ((day_high - live_price) / day_high) * 100
        if high_gap_pct >= 0.10:
            high_010_qualified_count += 1

    # Build the four target stocks from the four lowest qualified D-B values.
    qualified_detail_rows, qualified_detail_error = get_qualified_stock_detail_rows(
        base_rows, master_bytes
    )
    lowest_db_targets = []
    if not qualified_detail_error and qualified_detail_rows:
        # Select strictly the four lowest D-B values; ignore %Chg history.
        def parse_target_db(detail):
            raw_db = str(detail.get("16< D-B / Avg.4", "")).strip().replace("−", "-")
            outer_match = re.search(r"\(([^()]*)\)\s*\)?\s*$", raw_db)
            numeric_match = (
                re.search(r"([+\-]?\d+(?:,\d{3})*(?:\.\d+)?)\s*$", outer_match.group(1).strip())
                if outer_match else None
            )
            try:
                return float(numeric_match.group(1).replace(",", "")) if numeric_match else None
            except (TypeError, ValueError):
                return None

        target_candidates = []
        for detail in qualified_detail_rows:
            db_value = parse_target_db(detail)
            if db_value is not None:
                target_candidates.append((db_value, detail["Stock"]))

        target_candidates.sort(key=lambda item: (item[0], item[1]))
        lowest_db_targets = target_candidates[:4]

    # Keep the target list limited to the first four preferred stocks.
    target_quotes = {}
    target_table_rows = []
    if lowest_db_targets:
        first_symbols = [symbol for _, symbol in lowest_db_targets[:4]]
        with st.spinner("Fetching live prices for the four target stocks..."):
            target_quotes = get_autogreen_live_prices(tuple(first_symbols))

    # Target selection is based on all four stocks from the base-qualified AutoGreen list.
    # Colour-qualified signals are preferred for their existing card palette, but
    # must not exclude a target stock from the BUY competition.
    base_by_symbol = {item["Symbol"]: item for item in base_rows}
    signal_by_symbol = {item["Symbol"]: item for item in ranked_candidates}
    target_records = []
    for position, (_, symbol) in enumerate(lowest_db_targets, start=1):
        quote = target_quotes.get(symbol, {})
        rank = sum_o2l_ranks.get(symbol)
        open_price = quote.get("open")
        day_high = quote.get("high")
        day_low = quote.get("low")
        live_price = quote.get("price")
        o2h_pct = ((day_high - open_price) / open_price) * 100 if open_price and day_high is not None else None
        o2l_pct = ((day_low - open_price) / open_price) * 100 if open_price and day_low is not None else None
        target_status = "Target Completed" if o2h_pct is not None and o2h_pct >= 1.0 else "Signal Awaited"
        target_table_rows.append({
            "Role": f"TARGET #{position}",
            "Target Status": target_status,
            "Symbol": symbol,
            "Open": round(open_price, 2) if open_price is not None else "Unavailable",
            "High": round(day_high, 2) if day_high is not None else "Unavailable",
            "Low": round(day_low, 2) if day_low is not None else "Unavailable",
            "Live Price": round(live_price, 2) if live_price is not None else "Unavailable",
            "O2H": f"{o2h_pct:+.2f}%" if o2h_pct is not None else "Unavailable",
            "O2L": f"{o2l_pct:+.2f}%" if o2l_pct is not None else "Unavailable",
            "% Chg": f'{quote["change_pct"]:+.2f}%' if quote.get("change_pct") is not None else "Unavailable",
            "Rank": rank,
        })

        # Every one of the four D-B target stocks must reach BUY evaluation.
        # Do not drop a target merely because it has no colour-qualified signal row.
        signal = (
            signal_by_symbol.get(symbol)
            or base_by_symbol.get(symbol)
            or {"Symbol": symbol, "Card Colour": "Green"}
        )
        price = quote.get("price")
        high = quote.get("high")
        change_pct = quote.get("change_pct")
        if price is None or high is None or high <= 0 or change_pct is None:
            continue
        gap_pct = ((high - price) / high) * 100
        target_records.append({
            **signal,
            "Live Price": price,
            "Live % Chg": change_pct,
            "Day Open": quote.get("open"),
            "Day High": high,
            "Day Low": quote.get("low"),
            "High Gap %": gap_pct,
            "Sum O2L.10 Rank": rank,
            "_Target Position": position,
        })

    # BUY CONFIRMATION is evaluated independently for every one of the four target stocks.
    # O2L must be below -1%, O2H must remain below +1%, and live price must be above Open.
    # A stock whose O2H reaches +1% is marked Target Completed in the target table.
    confirmation_candidates = [
        item for item in target_records
        if item.get("Day Open") is not None
        and item.get("Day High") is not None
        and item.get("Day Low") is not None
        and item["Day Open"] > 0
        and ((item["Day Low"] - item["Day Open"]) / item["Day Open"]) * 100 < -1.0
        and ((item["Day High"] - item["Day Open"]) / item["Day Open"]) * 100 < 1.0
        and item["Live Price"] > item["Day Open"]
    ]
    confirmation_candidates.sort(key=lambda item: (
        -item["Live % Chg"],
        item["Sum O2L.10 Rank"] if item["Sum O2L.10 Rank"] is not None else float("inf"),
        item["Symbol"]
    ))
    for confirmation_row in confirmation_candidates:
        confirmation_row["_Role"] = "CONFIRMATION"

    # Display every target stock that satisfies the new BUY CONFIRMATION conditions.
    selected_rows = confirmation_candidates

    # Keep the existing High - 0.20% metric populated from the four target stocks.
    high_gap_candidates = [
        item for item in target_records
        if item["High Gap %"] >= 0.20
    ]


    # The metrics panel always inherits the first (topmost) signal card palette.
    # With no active signal, it follows the Rose-Cherry NO SIGNAL card.
    metric_palettes = {
        "Green": ("#062d20", "#087447", "#063d2b", "#E6C76A", "#CFFFE4", "#8FFFF0"),
        "Blue": ("#071b3d", "#075b91", "#092d57", "#78D9FF", "#D9F4FF", "#BDEBFF"),
        "Purple": ("#29113f", "#713a91", "#35134f", "#E0B5FF", "#F2DFFF", "#EAD7FF"),
    }
    if selected_rows:
        metric_colour_name = selected_rows[0].get("Card Colour", "Green")
        metric_c1, metric_c2, metric_c3, metric_border, metric_accent, metric_shimmer = metric_palettes.get(
            metric_colour_name, metric_palettes["Green"]
        )
        metric_outer = metric_border
        metric_glow = metric_c2 + "66"
    else:
        metric_c1, metric_c2, metric_c3 = "#3B1025", "#7E2342", "#35101F"
        metric_border, metric_accent, metric_shimmer = "#A83E59", "#F0B5C7", "#EAD7DD"
        metric_outer, metric_glow = "#A83E59", "rgba(220,92,126,.30)"

    # Light pearl dashboard; the topmost signal card supplies the accent colour.
    metric_colour_name = selected_rows[0].get("Card Colour", "Green") if selected_rows else "Rose-Cherry"
    light_accents = {
        "Green": ("#287A59", "#1F674B", "#E5F1E9", "#DCEAE1"),
        "Blue": ("#2877A8", "#205F88", "#E4EEF6", "#D9E7F1"),
        "Purple": ("#8050A2", "#684083", "#EEE6F4", "#E4D9EE"),
        "Rose-Cherry": ("#B04B68", "#913B55", "#F3E5EA", "#EBD8DF"),
    }
    metric_ink, metric_accent_text, metric_tint, metric_canvas = light_accents.get(
        metric_colour_name, light_accents["Rose-Cherry"]
    )
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:radial-gradient(ellipse at 8% 0%,{metric_canvas} 0%,transparent 42%),linear-gradient(135deg,#E6E9ED 0%,#DDE2E7 52%,#E9ECEF 100%) !important;
            color:#263342 !important;
        }}
        [data-testid="stHeader"] {{ background:rgba(231,234,238,.96) !important; }}
        /* 6thSense-Vardaan heading follows the active AutoGreen theme colour */
        .stApp h1 {{ color:{metric_ink} !important; }}
        /* MORE DASHBOARD follows the active signal-card colour */
        [data-testid="stExpander"] {{
            background:linear-gradient(145deg,#FFFFFF 0%,{metric_tint} 58%,{metric_canvas} 100%) !important;
            border:1.5px solid {metric_ink}88 !important;
            border-left:5px solid {metric_ink} !important;
            border-radius:14px !important;
            box-shadow:0 7px 22px rgba(35,51,68,.12),0 0 0 1px {metric_ink}18 !important;
            overflow:hidden !important;
        }}
        [data-testid="stExpander"] details {{
            background:transparent !important;
        }}
        [data-testid="stExpander"] summary {{
            color:{metric_ink} !important;
            background:linear-gradient(90deg,{metric_tint} 0%,{metric_canvas} 72%,#FFFFFF 100%) !important;
            font-weight:900 !important;
            letter-spacing:.45px !important;
        }}
        [data-testid="stExpander"] summary:hover {{
            color:{metric_accent_text} !important;
            background:linear-gradient(90deg,{metric_canvas} 0%,{metric_tint} 100%) !important;
        }}
        [data-testid="stExpander"] summary svg {{
            color:{metric_ink} !important;
            fill:{metric_ink} !important;
        }}
        [data-testid="stExpander"] [data-testid="stExpanderDetails"] {{
            border-top:1px solid {metric_ink}55 !important;
            background:linear-gradient(180deg,{metric_canvas} 0%,#FFFFFF 100%) !important;
        }}
        [data-testid="stSidebar"] {{
            background:linear-gradient(180deg,{metric_canvas} 0%,#E1E5E9 100%) !important;
            border-right:1px solid {metric_ink}66 !important;
            box-shadow:inset -8px 0 24px {metric_ink}0A !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ color:#344456 !important; }}
        h1,h2,h3,[data-testid="stSubheader"] {{ color:#263342 !important; }}
        h2,h3 {{
            border-bottom:1px solid {metric_ink}55 !important;
            border-left:4px solid {metric_ink} !important;
            padding:5px 0 8px 12px !important;
            background:linear-gradient(90deg,{metric_tint} 0%,transparent 78%) !important;
            border-radius:0 0 0 5px !important;
        }}
        @keyframes dashboardDiagonalShimmer {{
            0% {{ transform:translate3d(-140%,0,0) skewX(-18deg); opacity:0; }}
            12% {{ opacity:.38; }}
            55% {{ opacity:.22; }}
            100% {{ transform:translate3d(240%,0,0) skewX(-18deg); opacity:0; }}
        }}
        @keyframes dashboardFloatGlow {{
            0%,100% {{ background-position:0% 50%; }}
            50% {{ background-position:100% 50%; }}
        }}
        .signal-shimmer-surface::after,
        [data-testid="stMetric"]::after,
        [data-testid="stDataFrame"]::after,
        [data-testid="stTable"]::after {{
            content:"";
            position:absolute;
            inset:-35% auto -35% -45%;
            width:34%;
            pointer-events:none;
            z-index:5;
            background:linear-gradient(105deg,transparent 0%,rgba(255,255,255,0) 25%,
                rgba(255,255,255,.72) 48%,rgba(255,255,255,.18) 58%,transparent 82%);
            filter:blur(1.5px);
            animation:dashboardDiagonalShimmer 6.5s ease-in-out infinite;
        }}
        .signal-shimmer-surface,
        [data-testid="stMetric"],
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {{
            position:relative !important;
            overflow:hidden !important;
            isolation:isolate;
        }}
        [data-testid="stMetric"]::after {{ animation-delay:-2s; }}
        [data-testid="stDataFrame"]::after {{ animation-delay:-3.5s; }}
        [data-testid="stTable"]::after {{ animation-delay:-1s; }}
                [data-testid="stMetric"] {{
            background:linear-gradient(145deg,#FFFFFF 0%,{metric_tint} 68%,{metric_canvas} 100%) !important;
            border:1px solid {metric_ink}66 !important;
            border-top:4px solid {metric_ink} !important;
            border-radius:14px !important;
            padding:12px 14px !important;
            color:#263342 !important;
            box-shadow:0 7px 20px rgba(35,51,68,.10),0 0 0 1px {metric_ink}12 !important;
        }}
        [data-testid="stMetric"] label,[data-testid="stMetric"] [data-testid="stMetricLabel"] {{ color:#657487 !important; }}
        [data-testid="stMetric"] [data-testid="stMetricValue"] {{ color:#263342 !important; }}
        [data-testid="stDataFrame"],[data-testid="stTable"] {{
            border:1px solid {metric_ink}55 !important;
            border-radius:11px !important;
            overflow:hidden !important;
            box-shadow:0 7px 22px rgba(35,51,68,.10),0 0 0 1px {metric_ink}12 !important;
            background:#FFFFFF !important;
        }}
        .stButton > button,.stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {{
            border:1px solid {metric_ink}88 !important;
            color:{metric_ink} !important;
            background:linear-gradient(145deg,#FFFFFF 0%,{metric_tint} 75%,{metric_canvas} 100%) !important;
        }}
        .stButton > button:hover,.stDownloadButton > button:hover,
        [data-testid="stFormSubmitButton"] > button:hover {{
            border-color:{metric_ink} !important;
            box-shadow:0 0 12px {metric_ink}33 !important;
        }}
        [data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,
        [data-testid="stSelectbox"] > div,[data-testid="stFileUploader"] section,
        [data-testid="stExpander"] {{
            border-color:#D5DDE6 !important;
            background-color:#FFFFFF !important;
            color:#263342 !important;
        }}
        [data-testid="stExpander"] {{ border-radius:11px !important; }}
        hr {{ border-color:#DCE2E9 !important; }}
        [data-testid="stAlert"] {{ border-color:{metric_ink}66 !important; }}
        </style>
        """,
        unsafe_allow_html=True
    )



    st.markdown(
        """
        <style>
        @keyframes dependentSignalShimmer {
            0% { background-position:0% 50%; filter:brightness(1); }
            50% { background-position:100% 50%; filter:brightness(1.12); }
            100% { background-position:0% 50%; filter:brightness(1); }
        }
        @keyframes noSignalGreyShimmer {
            0% { background-position:200% 50%; }
            100% { background-position:-100% 50%; }
        }
        @keyframes noSignalFloatingGlow {
            0% { transform:translate3d(-35%,-5%,0) rotate(18deg); opacity:0; }
            18% { opacity:.62; }
            52% { transform:translate3d(115%,3%,0) rotate(18deg); opacity:.42; }
            82% { opacity:.18; }
            100% { transform:translate3d(190%,-4%,0) rotate(18deg); opacity:0; }
        }
        @keyframes dependentSignalFloat {
            0%,100% { transform:translateY(0); }
            50% { transform:translateY(-2px); }
        }
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
    f"""
    <div style="position:relative;overflow:hidden;width:100%;box-sizing:border-box;margin:12px 0 20px;padding:17px 12px;
                border:1px solid {metric_ink}88;border-top:3px solid {metric_ink};border-radius:15px;
                background:linear-gradient(120deg,#FFFFFF 0%,{metric_tint} 48%,#FFFFFF 100%);
                box-shadow:inset 0 1px 0 #FFFFFF,0 8px 22px rgba(35,51,68,.09),0 0 13px {metric_ink}18;">
      <div style="position:absolute;left:-72%;top:-55%;width:72%;height:210%;pointer-events:none;
                  background:linear-gradient(110deg,transparent 28%,{metric_shimmer}22 42%,#FFFFFF88 50%,{metric_shimmer}55 56%,transparent 70%);
                  filter:blur(7px);animation:noSignalFloatingGlow 6s ease-in-out infinite;"></div>
      <div style="position:relative;z-index:1;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));align-items:center;">
        <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
          <div style="font-size:clamp(12px,1.9vw,14px);font-weight:850;letter-spacing:.35px;color:{metric_ink};line-height:1.35;">AutoGreen qualified</div>
          <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#263342;margin-top:5px;line-height:1.1;">{len(base_rows)}</div>
        </div>
        <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
          <div style="font-size:clamp(12px,1.9vw,14px);font-weight:850;letter-spacing:.35px;color:{metric_ink};line-height:1.35;">Colour-qualified stocks</div>
          <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#263342;margin-top:5px;line-height:1.1;">{len(ranked_candidates)}</div>
        </div>
        <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
          <div style="font-size:clamp(12px,1.9vw,14px);font-weight:850;letter-spacing:.35px;color:{metric_ink};line-height:1.35;">High − 0.20% qualified</div>
          <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#263342;margin-top:5px;line-height:1.1;">{len(high_gap_candidates)}</div>
        </div>
        <div style="text-align:center;padding:4px 7px;">
          <div style="font-size:clamp(12px,1.9vw,14px);font-weight:850;letter-spacing:.35px;color:{metric_ink};line-height:1.35;">High − 0.10% qualified</div>
          <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#263342;margin-top:5px;line-height:1.1;">{high_010_qualified_count:02d}/{len(base_rows):02d}</div>
        </div>
      </div>
    </div>
    """,
    unsafe_allow_html=True
)

    if not selected_rows:
        no_signal_card = """
        <div style="max-width:620px;margin:18px auto 24px;padding:1.5px;border-radius:20px;
                    background:#A83E59;
                    box-shadow:0 0 0 1px rgba(255,255,255,.07),0 14px 34px rgba(0,0,0,.42),
                    0 0 22px rgba(220,92,126,.30);">
          <div style="position:relative;overflow:hidden;text-align:center;padding:30px 24px 28px;
                      border-radius:18px;background:radial-gradient(ellipse at 50% 0%,
                      rgba(255,225,235,.34),transparent 54%),linear-gradient(145deg,#3B1025 0%,#7E2342 48%,#35101F 100%);
                      color:#FFF4F6;animation:dependentSignalFloat 5s ease-in-out infinite;">
            <div aria-hidden="true" style="position:absolute;top:-55%;left:-72%;width:72%;height:210%;
                        pointer-events:none;z-index:0;filter:blur(7px);
                        background:linear-gradient(105deg,transparent 0%,rgba(170,176,184,.12) 25%,
                        rgba(232,237,244,.42) 48%,rgba(255,255,255,.25) 58%,transparent 82%);
                        animation:noSignalFloatingGlow 6s ease-in-out infinite;"></div>
            <div style="position:relative;z-index:1;font-size:11px;font-weight:900;letter-spacing:3px;color:#F0B8C8;
                        text-shadow:0 0 12px rgba(255,210,225,.34);">6THSENSE · VARDAAN</div>
            <div style="position:relative;z-index:1;font-size:clamp(30px,5vw,42px);line-height:1.1;font-weight:950;
                        letter-spacing:2.2px;margin:13px 0 10px;color:#FFF0F4;
                        text-shadow:0 2px 18px rgba(255,190,210,.34);">NO SIGNAL</div>
            <div style="position:relative;z-index:1;width:74px;height:1px;margin:0 auto 15px;
                        background:linear-gradient(90deg,transparent,#F2B5C7,transparent);"></div>
            <div style="position:relative;z-index:1;font-size:12px;font-weight:900;letter-spacing:1.65px;color:#F0B5C7;">
              WAIT FOR A QUALIFIED SETUP
            </div>
            <div style="position:relative;z-index:1;font-size:13px;line-height:1.65;font-weight:550;color:#E5C0CB;
                        margin-top:10px;">
              No stock currently meets the BUY CONFIRMATION conditions.
            </div>
          </div>
        </div>
        """
        st.markdown(no_signal_card, unsafe_allow_html=True)
    else:
        for index, item in enumerate(selected_rows):
            is_confirmation = item.get("_Role") == "CONFIRMATION"
            symbol = item["Symbol"]
            entry = item["Live Price"]
            target = entry * 1.0035
            rank_value = item.get("Sum O2L.10 Rank")
            rank_text = f"{rank_value:g}" if isinstance(rank_value, (int, float)) else "—"
            role_text = "BUY CONFIRMATION" if is_confirmation else "BUY SIGNAL"
            palette = {
                "Green": ("#062d20", "#087447", "#063d2b", "#E6C76A", "#CFFFE4", "#8FFFF0"),
                "Blue": ("#071b3d", "#075b91", "#092d57", "#78D9FF", "#D9F4FF", "#BDEBFF"),
                "Purple": ("#29113f", "#713a91", "#35134f", "#E0B5FF", "#F2DFFF", "#EAD7FF"),
            }
            c1, c2, c3, border_colour, accent_colour, shimmer_colour = palette.get(
                item.get("Card Colour"), palette["Green"]
            )
            card_html = f"""
            <div style="
                position:relative;overflow:hidden;
                background:
                    linear-gradient(125deg,{c1} 0%,{c2} 46%,{c3} 100%) padding-box,
                    linear-gradient(110deg,{border_colour} 0%,{shimmer_colour} 24%,#FFFFFF 34%,
                                    {shimmer_colour} 44%,{border_colour} 62%,{shimmer_colour} 82%,
                                    {border_colour} 100%) border-box;
                background-size:200% 200%,300% 300%;
                background-origin:padding-box,border-box;
                background-clip:padding-box,border-box;
                animation:dependentSignalShimmer 4.2s linear infinite,
                          dependentSignalFloat 2.8s ease-in-out infinite;
                border:2px solid transparent;
                border-radius:21px;
                padding:23px 23px 22px;
                margin:7px 0 17px;
                color:#fff;
                box-shadow:0 0 0 1px rgba(255,255,255,.13) inset,
                           0 10px 28px rgba(0,0,0,.24),
                           0 8px 25px {c2}88,
                           0 0 23px {c2}70;
            ">
              <div style="position:absolute;right:-30px;top:-62px;width:190px;height:190px;border-radius:50%;
                          background:radial-gradient(circle,rgba(255,255,255,.22),transparent 69%);pointer-events:none;"></div>
              <div style="position:relative;z-index:1;text-align:center;font-size:11px;font-weight:900;letter-spacing:3px;color:{accent_colour};margin:0 0 15px;text-shadow:0 0 12px {c2}88;">6THSENSE · VARDAAN</div>
            <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;position:relative;z-index:1;">
                <div style="font-size:1.78rem;line-height:1.15;font-weight:950;letter-spacing:.7px;color:#fff;
                            text-shadow:0 2px 12px rgba(0,0,0,.24);">
                  {symbol}
                </div>
                <div style="font-size:.72rem;font-weight:950;letter-spacing:1.15px;padding:8px 12px;
                            border:1px solid {border_colour};border-radius:999px;color:{accent_colour};
                            background:rgba(0,0,0,.20);white-space:nowrap;
                            box-shadow:0 0 13px {c2}66;">
                  {role_text}
                </div>
              </div>
              <div style="font-size:.72rem;font-weight:900;letter-spacing:1.25px;color:{accent_colour};
                          margin-top:9px;position:relative;z-index:1;text-shadow:0 1px 8px rgba(0,0,0,.25);">
                RANK {rank_text}
              </div>
              <div style="height:1px;background:linear-gradient(90deg,transparent,{border_colour},rgba(255,255,255,.42),{border_colour},transparent);
                          margin:19px 0 17px;position:relative;z-index:1;"></div>
              <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));
                          align-items:center;gap:18px;position:relative;z-index:1;">
                <div style="min-width:0;">
                  <div style="font-size:.70rem;font-weight:850;letter-spacing:1.15px;opacity:.84;">ENTRY PRICE</div>
                  <div style="font-size:1.48rem;line-height:1.25;font-weight:950;color:#fff;margin-top:7px;
                              white-space:nowrap;text-shadow:0 2px 10px rgba(0,0,0,.22);">
                    ₹{entry:,.2f}
                  </div>
                </div>
                <div style="min-width:0;text-align:left;">
                  <div style="font-size:.70rem;font-weight:850;letter-spacing:1.15px;opacity:.84;">TARGET PRICE</div>
                  <div style="font-size:1.48rem;line-height:1.25;font-weight:950;color:{accent_colour};margin-top:7px;
                              white-space:nowrap;text-shadow:0 0 12px {c2}88;">
                    ₹{target:,.2f}
                  </div>
                </div>
              </div>
              <div style="position:absolute;left:-38px;bottom:-92px;width:190px;height:190px;border-radius:50%;
                          background:radial-gradient(circle,rgba(255,255,255,.12),transparent 70%);pointer-events:none;"></div>
            </div>
            """
            st.markdown(card_html, unsafe_allow_html=True)

    st.subheader("Target Stocks")
    if target_table_rows:
        render_accent_table(target_table_rows, metric_ink, metric_accent_text, metric_tint, min_width=680)
    else:
        st.info("No four qualified target stocks with a readable numeric value are available yet.")


    with st.expander("MORE DASHBOARD", expanded=False):

        st.subheader("Qualified AutoGreen Stocks")
        if base_rows:
            if qualified_detail_error:
                st.warning(f"Qualified-stock details could not be loaded: {qualified_detail_error}")
            elif qualified_detail_rows:
                render_qualified_stock_table(qualified_detail_rows, metric_ink, metric_accent_text, metric_tint)
            else:
                st.caption("No qualified-stock details are available.")
        else:
            st.caption("No stocks passed both volume-expansion conditions.")


def get_qualified_stock_detail_rows(base_rows, master_bytes):
    """Build the lean qualified-stock view from AutoGreen and Master workbook columns."""
    if not base_rows:
        return [], None

    def clean_header(value):
        return " ".join(str(value or "").strip().casefold().split())

    def display_value(value, parentheses=False, decimals=None):
        if value is None or str(value).strip() == "":
            return "—"
        try:
            number = float(str(value).replace(",", "").replace("%", "").strip())
            rendered = f"{number:.{decimals}f}" if decimals is not None else f"{number:g}"
        except (TypeError, ValueError):
            rendered = str(value).strip()
        return f"({rendered})" if parentheses else rendered

    symbols = {item["Symbol"]: item for item in base_rows}

    def signed_change_history(item):
        history_values = []
        for key in ("%Chg.1", "%Chg.2", "%Chg.3", "%Chg.4"):
            value = item.get(key)
            if value is None:
                history_values.append("—")
                continue
            try:
                history_values.append(f"{float(value):+.2f}")
            except (TypeError, ValueError):
                history_values.append("—")
        return ", ".join(history_values)

    details = {
        symbol: {
            "Stock": symbol,
            "% Chg History": signed_change_history(item),
        }
        for symbol, item in symbols.items()
    }

    # Read the two dated O2H columns immediately following Vol.Expand 2
    # from the AutoGreen sheet, retaining the workbook's actual date headings.
    auto_wb = None
    master_wb = None
    try:
        auto_wb = load_workbook(
            io.BytesIO(master_bytes),
            data_only=True,
            read_only=True
        )
        auto_sheet = next(
            (ws for ws in auto_wb.worksheets if ws.title.strip().casefold() == "autogreen"),
            None
        )
        if auto_sheet is not None:
            header_row_number = None
            symbol_col = 0
            vol2_col = None
            o2h_columns = []
            for row_number, row in enumerate(
                auto_sheet.iter_rows(min_row=1, max_row=min(auto_sheet.max_row or 1, 30), values_only=True),
                start=1
            ):
                found = {clean_header(value): index for index, value in enumerate(row) if value is not None}
                if clean_header("Vol.Expand 2") in found:
                    header_row_number = row_number
                    vol2_col = found[clean_header("Vol.Expand 2")]
                    symbol_col = next(
                        (found[clean_header(label)] for label in ("Symbol", "NSE Symbol", "Ticker")
                         if clean_header(label) in found),
                        0
                    )
                    o2h_columns = [
                        (index, str(row[index]).strip() if row[index] is not None else f"O2H {index - vol2_col}")
                        for index in (vol2_col + 1, vol2_col + 2)
                        if index < len(row)
                    ]
                    break

            if header_row_number is not None:
                for row in auto_sheet.iter_rows(min_row=header_row_number + 1, values_only=True):
                    if not row or symbol_col >= len(row):
                        continue
                    symbol = normalize_stock_symbol(row[symbol_col])
                    if symbol not in details:
                        continue
                    for position, (column_index, heading) in enumerate(o2h_columns, start=1):
                        value = row[column_index] if column_index < len(row) else None
                        details[symbol][f"O2H {position} · {heading}"] = display_value(value, decimals=2)

        # Read the two requested D-B / C-B values from the Master workbook.
        master_wb = load_workbook(
            io.BytesIO(master_bytes),
            data_only=True,
            read_only=True
        )
        selected_sheet = None
        header_row_number = None
        header_columns = {}
        required_headers = ("16< D-B / Avg.4", "16> C-B / Avg.4")
        for sheet in sorted(
            master_wb.worksheets,
            key=lambda ws: (ws.title.casefold() != "summary", ws.title.casefold())
        ):
            for row_number, row in enumerate(
                sheet.iter_rows(min_row=1, max_row=min(sheet.max_row or 1, 30), values_only=True),
                start=1
            ):
                found = {clean_header(value): index for index, value in enumerate(row) if value is not None}
                if all(clean_header(name) in found for name in required_headers):
                    selected_sheet = sheet
                    header_row_number = row_number
                    header_columns = {name: found[clean_header(name)] for name in required_headers}
                    break
            if selected_sheet is not None:
                break

        if selected_sheet is not None:
            for row in selected_sheet.iter_rows(min_row=header_row_number + 1, values_only=True):
                if not row:
                    continue
                symbol = normalize_stock_symbol(row[0] if len(row) else None)
                if symbol not in details:
                    continue
                for heading in required_headers:
                    index = header_columns[heading]
                    value = row[index] if index < len(row) else None
                    details[symbol][heading] = display_value(value, parentheses=True)

        output = []
        for symbol in symbols:
            record = details[symbol]
            output.append({
                "Stock": symbol,
                "% Chg History": record.get("% Chg History", "—"),
                "16< D-B / Avg.4": record.get("16< D-B / Avg.4", "—"),
                "16> C-B / Avg.4": record.get("16> C-B / Avg.4", "—"),
                **{
                    key: value for key, value in record.items()
                    if key.startswith("O2H ")
                }
            })
        # Sort qualified stocks by the numeric D-B value inside parentheses:
        # lowest value first; missing/non-numeric values remain at the bottom.
        def db_sort_value(record):
            # Values may look like "(xxxxxxxx(y))". Sort by the final
            # numeric value in the innermost trailing parentheses only.
            raw = str(record.get("16< D-B / Avg.4", "")).strip().replace("−", "-")
            match = re.search(r"\(([^()]*)\)\s*\)?\s*$", raw)
            if not match:
                return (1, 0.0)
            numeric_match = re.search(
                r"([+\-]?\d+(?:,\d{3})*(?:\.\d+)?)\s*$",
                match.group(1).strip()
            )
            if not numeric_match:
                return (1, 0.0)
            try:
                return (0, float(numeric_match.group(1).replace(",", "")))
            except (TypeError, ValueError):
                return (1, 0.0)

        output.sort(key=db_sort_value)
        return output, None
    except Exception as error:
        return [], str(error)
    finally:
        if auto_wb is not None:
            auto_wb.close()
        if master_wb is not None:
            master_wb.close()

def get_top_stocks_master_values(top_stocks, master_bytes, rank_source_column):
    """
    Find the requested columns in the Master workbook (prefer the summary
    sheet, but search every sheet and the first 30 rows for the header row).
    Stock symbols are matched against column A.
    """
    # Keep the requested panel-specific display order.
    # H2L: O2L, O2H, O2L Rank, D-B, C-B.
    # L2H: O2H, O2L, O2H Rank, C-B, D-B (existing order).
    if rank_source_column == "Sum O2L.10":
        output_columns = [
            "Sum O2L.10",
            "Sum O2H.10",
            "16< D-B / Avg.4",
            "16> C-B / Avg.4",
        ]
    else:
        output_columns = [
            "Sum O2H.10",
            "Sum O2L.10",
            "16> C-B / Avg.4",
            "16< D-B / Avg.4",
        ]
    if not master_bytes:
        return [], "Master Excel is not available in Supabase."

    workbook = None
    try:
        workbook = load_workbook(
            io.BytesIO(master_bytes),
            data_only=True,
            read_only=True
        )

        def clean_header(value):
            return " ".join(str(value or "").strip().casefold().split())

        wanted_headers = {clean_header(name): name for name in output_columns}
        sheet_candidates = sorted(
            workbook.worksheets,
            key=lambda sheet: (sheet.title.casefold() != "summary", sheet.title.casefold())
        )

        selected_sheet = None
        header_row_number = None
        header_columns = None

        # Locate the actual header row instead of assuming row 1 or a fixed sheet.
        for sheet in sheet_candidates:
            for row_number, row in enumerate(
                sheet.iter_rows(min_row=1, max_row=min(sheet.max_row, 30), values_only=True),
                start=1
            ):
                found = {
                    clean_header(value): column_number
                    for column_number, value in enumerate(row, start=1)
                    if value is not None
                }
                if all(name in found for name in wanted_headers):
                    selected_sheet = sheet
                    header_row_number = row_number
                    header_columns = {
                        original: found[normalized]
                        for normalized, original in wanted_headers.items()
                    }
                    break
            if selected_sheet is not None:
                break

        if selected_sheet is None:
            available = ", ".join(sheet.title for sheet in workbook.worksheets)
            return [], (
                "Could not find all four requested headers in the first 30 rows "
                f"of any Master sheet. Sheets checked: {available}"
            )

        master_rows = {}
        for row in selected_sheet.iter_rows(
            min_row=header_row_number + 1,
            values_only=True
        ):
            if not row:
                continue
            symbol = normalize_stock_symbol(row[0] if len(row) else None)
            if symbol:
                master_rows[symbol] = row

        # Rank every stock using the requested panel-specific direction:
        # H2L uses Sum O2L.10 with the LOWEST value ranked 1.
        # L2H uses Sum O2H.10 with the HIGHEST value ranked 1.
        # Ties share the same competition rank.
        rank_source_index = header_columns[rank_source_column] - 1

        def numeric_value(value):
            if value is None:
                return None
            try:
                cleaned = str(value).replace(",", "").replace("%", "").strip()
                return float(cleaned)
            except (TypeError, ValueError):
                return None

        ranked_values = []
        for master_symbol, master_row in master_rows.items():
            raw_rank_value = (
                master_row[rank_source_index]
                if rank_source_index < len(master_row) else None
            )
            numeric_rank_value = numeric_value(raw_rank_value)
            if numeric_rank_value is not None:
                ranked_values.append((numeric_rank_value, master_symbol))

        ranked_values.sort(
            key=lambda item: item[0],
            reverse=(rank_source_column != "Sum O2L.10")
        )
        master_ranks = {}
        previous_value = None
        current_rank = 0
        for position, (rank_value, master_symbol) in enumerate(ranked_values, start=1):
            if previous_value is None or rank_value != previous_value:
                current_rank = position
                previous_value = rank_value
            master_ranks[master_symbol] = current_rank

        results = []
        matched_count = 0
        for _, stock in top_stocks.iterrows():
            display_symbol = stock.get("NSE Symbol")
            symbol = normalize_stock_symbol(display_symbol)
            master_row = master_rows.get(symbol)
            result = {
                "NSE Symbol": display_symbol,
                "% Chg": stock.get("% Chg")
            }

            if master_row is None:
                for column_name in output_columns[:2]:
                    result[column_name] = "Not found"
                result[f"{rank_source_column} Rank"] = ""
                for column_name in output_columns[2:]:
                    result[column_name] = "Not found"
            else:
                matched_count += 1
                for column_name in output_columns[:2]:
                    column_index = header_columns[column_name] - 1
                    raw_value = (
                        master_row[column_index]
                        if column_index < len(master_row) else None
                    )
                    # Display Sum O2L.10 and Sum O2H.10 consistently
                    # with exactly two decimal places when the value is numeric.
                    try:
                        numeric_display_value = float(
                            str(raw_value).replace(",", "").replace("%", "").strip()
                        ) if raw_value is not None else None
                        result[column_name] = (
                            f"{numeric_display_value:.2f}"
                            if numeric_display_value is not None
                            else raw_value
                        )
                    except (TypeError, ValueError):
                        result[column_name] = raw_value

                # Display the rank calculated from this panel's source column.
                result[f"{rank_source_column} Rank"] = master_ranks.get(symbol, "")

                # Show only text/numbers enclosed in parentheses in the last two columns.
                for column_name in output_columns[2:]:
                    column_index = header_columns[column_name] - 1
                    raw_value = (
                        master_row[column_index]
                        if column_index < len(master_row) else None
                    )
                    matches = re.findall(r"\(([^()]*)\)", str(raw_value or ""))
                    result[column_name] = ", ".join(
                        value.strip() for value in matches if value.strip()
                    ) if matches else ""

            results.append(result)

        if matched_count == 0:
            return results, (
                "None of the four selected stock symbols matched column A in "
                f"Master sheet '{selected_sheet.title}'. Check that Master column A "
                "contains NSE symbols (for example, SBIN)."
            )

        return results, None
    except Exception as error:
        return [], f"Could not read Master Excel: {error}"
    finally:
        if workbook is not None:
            workbook.close()


def clear_index_session_data():
    """
    Remove uploaded Index data and signal candidates from this browser session.
    Streamlit session_state is isolated per connected user session.
    """
    for key in (
        "index_file_bytes",
        "index_file_name",
        "index_file_owner_id",
        "trade_candidates_H2L",
        "trade_candidates_L2H",
        "trade_candidates_index2_H2L",
        "trade_candidates_index2_L2H",
    ):
        st.session_state.pop(key, None)


def logout():
    try:
        sb.auth.sign_out()
    except Exception:
        pass

    clear_index_session_data()
    st.session_state.user = None
    st.session_state.profile = None
    st.session_state.is_admin = False
    st.session_state.recovery = False
    st.session_state.session_login_started_at = None
    st.session_state.pop("session_expired_notice", None)
    st.rerun()


def expire_login_session():
    """Sign out the current user after the one-hour login limit."""
    try:
        sb.auth.sign_out()
    except Exception:
        pass

    clear_index_session_data()
    st.session_state.user = None
    st.session_state.profile = None
    st.session_state.is_admin = False
    st.session_state.recovery = False
    st.session_state.session_login_started_at = None
    st.session_state.session_expired_notice = True


# ============================================================
# PASSWORD RECOVERY
# ============================================================

if "code" in st.query_params:

    try:
        sb.auth.exchange_code_for_session(
            {
                "auth_code":
                st.query_params["code"]
            }
        )

        st.session_state.recovery = True
        st.query_params.clear()
        st.rerun()

    except Exception as e:
        st.error(f"Recovery error: {e}")
        st.stop()


if st.session_state.recovery:

    st.title("Reset Password")

    p1 = st.text_input(
        "New Password",
        type="password"
    )

    p2 = st.text_input(
        "Confirm Password",
        type="password"
    )

    if st.button(
        "Update Password",
        use_container_width=True
    ):

        if len(p1) < 6:
            st.error(
                "Password must be at least 6 characters."
            )

        elif p1 != p2:
            st.error(
                "Passwords do not match."
            )

        else:

            try:
                sb.auth.update_user(
                    {"password": p1}
                )

                st.success(
                    "Password updated successfully."
                )

                st.session_state.user = None
                st.session_state.profile = None
                st.session_state.recovery = False

                st.rerun()

            except Exception as e:
                st.error(str(e))

    st.stop()


# ============================================================
# LOGIN
# ============================================================

# Enforce an absolute one-hour login period for every account, including admins.
login_started_at = st.session_state.get("session_login_started_at")
if (
    st.session_state.user is not None
    and login_started_at is not None
    and time.time() - login_started_at >= 60 * 60
):
    expire_login_session()

if st.session_state.user is None:

    st.title(
        "6thSense Excel Manager"
    )

    st.subheader("Login")

    if st.session_state.pop("session_expired_notice", False):
        st.info("Your 1-hour login session has expired. Please log in again.")

    with st.form("login_form"):

        email = st.text_input(
            "Email / User ID"
        )

        password = st.text_input(
            "Password",
            type="password"
        )

        login_button = st.form_submit_button(
            "Login",
            use_container_width=True
        )

    if login_button:

        try:

            result = sb.auth.sign_in_with_password(
                {
                    "email": email.strip(),
                    "password": password
                }
            )

            user = result.user
            prof = get_profile(user.id)

            if not prof:
                st.error(
                    "Profile not found."
                )

            elif not prof["is_active"]:
                st.error(
                    "Your account is inactive."
                )

            else:

                st.session_state.user = user
                st.session_state.profile = prof
                st.session_state.is_admin = (
                    prof["role"] == "admin"
                )
                st.session_state.session_login_started_at = time.time()
                st.session_state.pop("session_expired_notice", None)

                st.rerun()

        except Exception:
            st.error(
                "Invalid email or password."
            )

    # --------------------------------------------------------
    # FORGOT PASSWORD
    # --------------------------------------------------------

    st.divider()

    st.subheader(
        "Forgot Password"
    )

    reset_email = st.text_input(
        "Email for password reset"
    )

    if st.button(
        "Send Reset Link",
        use_container_width=True
    ):

        try:

            host = st.context.headers.get(
                "host",
                ""
            )

            redirect_url = (
                "https://" + host
            )

            sb.auth.reset_password_for_email(
                reset_email.strip(),
                {
                    "redirect_to":
                    redirect_url
                }
            )

            st.success(
                "Password reset link sent."
            )

        except Exception as e:
            st.error(str(e))

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user
profile = st.session_state.profile
is_admin = st.session_state.is_admin

# Start the one-hour clock for sessions that were already active during deployment.
if st.session_state.session_login_started_at is None:
    st.session_state.session_login_started_at = time.time()

# Check the one-hour limit automatically every 30 seconds while the app is open.
@st.fragment(run_every="30s")
def session_timeout_monitor():
    started_at = st.session_state.get("session_login_started_at")
    if (
        st.session_state.get("user") is not None
        and started_at is not None
        and time.time() - started_at >= 60 * 60
    ):
        expire_login_session()
        st.rerun()


session_timeout_monitor()

# ============================================================
# BLUE & SILVER DASHBOARD THEME
# Green H2L and orange L2H priority panels keep their own colors.
# ============================================================

st.markdown("""
<style>
/* Cool silver canvas and restrained blue accents for the general UI. */
.stApp {
    background: linear-gradient(135deg, #f7f9fc 0%, #edf2f8 52%, #f8fafc 100%);
}
[data-testid="stHeader"] {
    background: rgba(247, 249, 252, 0.88);
}
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #eaf0f7 0%, #f5f7fa 100%);
    border-right: 1px solid #d5deea;
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
[data-testid="stSidebar"] label {
    color: #263b55;
}
h1, h2, h3 {
    color: #183b63;
    letter-spacing: -0.025em;
}
h2, h3 {
    border-bottom: 1px solid #d7e0eb;
    padding-bottom: 0.35rem;
}
[data-testid="stMetric"] {
    background: linear-gradient(145deg, #ffffff 0%, #eef3f8 100%);
    border: 1px solid #d7e0eb;
    border-radius: 12px;
    padding: 0.75rem 0.9rem;
    box-shadow: 0 3px 12px rgba(34, 65, 100, 0.045);
}
[data-testid="stDataFrame"],
[data-testid="stTable"] {
    border: 1px solid #d6dfeb;
    border-radius: 10px;
    overflow: hidden;
}
.stButton > button,
.stDownloadButton > button,
[data-testid="stFormSubmitButton"] > button {
    border: 1px solid #9db5d1;
    border-radius: 9px;
    color: #173e68;
    background: linear-gradient(180deg, #ffffff 0%, #e8eef6 100%);
    font-weight: 600;
}
.stButton > button:hover,
.stDownloadButton > button:hover,
[data-testid="stFormSubmitButton"] > button:hover {
    border-color: #4779ad;
    color: #12365c;
    background: #dce8f5;
}
[data-testid="stTextInput"] input,
[data-testid="stNumberInput"] input,
[data-testid="stSelectbox"] > div,
[data-testid="stFileUploader"] section {
    border-color: #cbd7e5;
}
hr {
    border-color: #d4deea;
}
[data-testid="stExpander"] {
    border: 1px solid #d5dfeb;
    border-radius: 10px;
    background: rgba(255,255,255,0.58);
}
[data-testid="stAlert"] {
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)



st.caption(
    f"Logged in as: {user.email}"
)

# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.write(
        f"**User:** {user.email}"
    )

    st.write(
        f"**Role:** "
        f"{'Admin' if is_admin else 'User'}"
    )

    if st.button(
        "Log out",
        use_container_width=True,
        type="primary",
        key="sidebar_logout_button"
    ):
        logout()

    st.divider()
    page = st.radio(
        "Dashboard",
        ["AutoGreen", "Stock of the Day", "Master Excel"],
        index=0,
        key="main_dashboard_default_index"
    )

if page == "AutoGreen":
    # Refresh the AutoGreen market data automatically during NSE trading hours.
    market_now = datetime.now(ZoneInfo("Asia/Kolkata"))
    market_open = (
        market_now.weekday() < 5
        and market_now.hour * 60 + market_now.minute >= 9 * 60 + 15
        and market_now.hour * 60 + market_now.minute <= 15 * 60 + 30
    )
    if market_open:
        st_autorefresh(
            interval=15_000,
            limit=None,
            key="autogreen_live_market_refresh",
        )
    render_autogreen_dashboard()

elif page == "Stock of the Day":
    # Refresh live Stock of the Day prices and target status during NSE market hours.
    market_now = datetime.now(ZoneInfo("Asia/Kolkata"))
    market_open = (
        market_now.weekday() < 5
        and market_now.hour * 60 + market_now.minute >= 9 * 60 + 15
        and market_now.hour * 60 + market_now.minute <= 15 * 60 + 30
    )
    if market_open:
        st_autorefresh(
            interval=15_000,
            limit=None,
            key="stock_of_the_day_live_market_refresh",
        )
    render_stock_of_the_day()


else:

# MASTER FILE
    # ============================================================

    master = get_master()

    st.header(
        "Master Excel"
    )

    if master:

        st.success(
            "Master file is available."
        )

        st.download_button(
            "Download Master Excel",
            data=master,
            file_name=MASTER_FILE,
            mime=EXCEL_MIME,
            use_container_width=True
        )

    else:

        st.warning(
            "Master file is not uploaded."
        )


    # ============================================================
    # ADMIN PANEL
    # ============================================================

    if is_admin:

        st.divider()

        st.header(
            "Admin Panel"
        )

        # ========================================================
        # UPLOAD / REPLACE
        # ========================================================

        st.subheader(
            "Upload / Replace Master File"
        )

        uploaded = st.file_uploader(
            "Select .xlsx file",
            type=["xlsx"]
        )

        if uploaded:

            if st.button(
                "Save Master File",
                use_container_width=True
            ):

                try:

                    # Save ORIGINAL bytes directly.
                    # No pandas rewriting.
                    save_master(
                        uploaded.getvalue()
                    )

                    st.success(
                        "Master file uploaded successfully."
                    )

                    st.rerun()

                except Exception as e:
                    st.error(
                        f"Upload error: {e}"
                    )

        # ========================================================
        # MODIFY
        # ========================================================

        if master:

            st.divider()

            st.subheader(
                "Modify Master Excel"
            )

            try:

                wb = load_workbook(
                    io.BytesIO(master),
                    data_only=False
                )

                sheet = st.selectbox(
                    "Select Sheet",
                    wb.sheetnames
                )

                ws = wb[sheet]

                cell = st.text_input(
                    "Cell Address",
                    value="A1"
                ).strip().upper()

                try:

                    old_value = ws[cell].value

                    st.write(
                        f"Current value: "
                        f"`{old_value}`"
                    )

                    new_value = st.text_input(
                        "New Value",
                        value=(
                            ""
                            if old_value is None
                            else str(old_value)
                        )
                    )

                    if st.button(
                        "Save Cell Change",
                        use_container_width=True
                    ):

                        # ONLY change cell value.
                        # Existing Excel formatting remains.
                        ws[cell].value = new_value

                        output = io.BytesIO()

                        wb.save(output)

                        save_master(
                            output.getvalue()
                        )

                        st.success(
                            "Cell updated successfully."
                        )

                        st.rerun()

                except Exception:
                    st.error(
                        "Invalid cell address."
                    )

            except Exception as e:
                st.error(
                    f"Excel error: {e}"
                )

        # ========================================================
        # DELETE MASTER
        # ========================================================

        st.divider()

        st.subheader(
            "Delete Master File"
        )

        if master:

            if st.button(
                "Delete Master Excel",
                use_container_width=True
            ):

                try:

                    delete_master()

                    st.success(
                        "Master file deleted."
                    )

                    st.rerun()

                except Exception as e:
                    st.error(
                        f"Delete error: {e}"
                    )

        else:

            st.info(
                "There is no Master file to delete."
            )

        # ========================================================
        # ADD USER
        # ========================================================

        st.divider()

        st.subheader(
            "Add User"
        )

        new_email = st.text_input(
            "New User Email"
        )

        new_password = st.text_input(
            "New User Password",
            type="password"
        )

        new_role = st.selectbox(
            "New User Role",
            ["user", "admin"]
        )

        if st.button(
            "Create User",
            use_container_width=True
        ):

            try:

                if not new_email.strip():
                    st.error(
                        "Enter an email."
                    )

                elif len(new_password) < 6:
                    st.error(
                        "Password must be at least 6 characters."
                    )

                else:

                    # Create Auth user
                    result = admin.auth.admin.create_user(
                        {
                            "email":
                            new_email.strip(),

                            "password":
                            new_password,

                            "email_confirm":
                            True
                        }
                    )

                    uid = result.user.id

                    # Create profile using SERVICE client.
                    # This avoids the RLS 403 for admin creation.
                    admin.table(
                        "profiles"
                    ).insert(
                        {
                            "id": uid,
                            "email":
                            new_email.strip(),
                            "role":
                            new_role,
                            "is_active":
                            True
                        }
                    ).execute()

                    st.success(
                        "User created successfully."
                    )

                    st.rerun()

            except Exception as e:
                st.error(
                    f"Create user error: {e}"
                )

        # ========================================================
        # USER MANAGEMENT
        # ========================================================

        st.divider()

        st.subheader(
            "User Management"
        )

        try:

            users = admin.table(
                "profiles"
            ).select(
                "*"
            ).order(
                "email"
            ).execute().data

            for u in users:

                uid = u["id"]

                c1, c2, c3, c4 = st.columns(
                    [4, 1, 1, 1]
                )

                c1.write(
                    f"**{u['email']}**  \n"
                    f"Role: {u['role']} | "
                    f"{'Active' if u['is_active'] else 'Banned'}"
                )

                # ------------------------------------------------
                # ROLE
                # ------------------------------------------------

                if c2.button(
                    "Role",
                    key=f"role_{uid}"
                ):

                    new_role = (
                        "user"
                        if u["role"] == "admin"
                        else "admin"
                    )

                    admin.table(
                        "profiles"
                    ).update(
                        {
                            "role":
                            new_role
                        }
                    ).eq(
                        "id",
                        uid
                    ).execute()

                    st.rerun()

                # ------------------------------------------------
                # BAN / ACTIVATE
                # ------------------------------------------------

                if c3.button(
                    (
                        "Ban"
                        if u["is_active"]
                        else "Activate"
                    ),
                    key=f"active_{uid}"
                ):

                    admin.table(
                        "profiles"
                    ).update(
                        {
                            "is_active":
                            not u["is_active"]
                        }
                    ).eq(
                        "id",
                        uid
                    ).execute()

                    st.rerun()

                # ------------------------------------------------
                # DELETE
                # ------------------------------------------------

                if c4.button(
                    "Delete",
                    key=f"delete_{uid}"
                ):

                    if uid == user.id:

                        st.error(
                            "You cannot delete yourself."
                        )

                    else:

                        # Delete profile
                        admin.table(
                            "profiles"
                        ).delete().eq(
                            "id",
                            uid
                        ).execute()

                        # Delete Auth account
                        try:
                            admin.auth.admin.delete_user(
                                uid
                            )
                        except Exception:
                            pass

                        st.success(
                            "User deleted."
                        )

                        st.rerun()


        except Exception as e:

            st.error(
                f"User management error: {e}"
            )


    # ============================================================
    # NORMAL USER
    # ============================================================

    else:

        st.info(
            "USER ACCESS â€” Read-only"
        )
