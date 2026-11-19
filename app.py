import io
import re
import json
import time
import base64
import html
import requests
from datetime import datetime, timezone
import streamlit as st
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


@st.cache_data(ttl=300, max_entries=1, show_spinner=False)
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


@st.cache_data(ttl=300, max_entries=1, show_spinner=False)
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


@st.cache_data(ttl=60, max_entries=1, show_spinner=False)
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

def render_qualified_stock_table(rows):
    """Render the qualified-stock list in the dashboard's premium black metallic style."""
    if not rows:
        return

    columns = list(rows[0].keys())
    header_html = "".join(
        f'<th style="padding:13px 15px;text-align:left;white-space:nowrap;'
        f'font-size:11px;font-weight:900;letter-spacing:1px;color:#E8D7A5;'
        f'background:linear-gradient(180deg,#242424 0%,#111 100%);'
        f'border-bottom:1px solid rgba(222,190,120,.48);">'
        f'{html.escape(str(column))}</th>'
        for column in columns
    )
    body_rows = []
    for row_number, record in enumerate(rows):
        cells = "".join(
            f'<td style="padding:12px 15px;white-space:nowrap;'
            f'font-size:13px;font-weight:650;color:#F1F1F1;'
            f'border-bottom:1px solid rgba(255,255,255,.075);">'
            f'{html.escape(str(record.get(column, "—")))}</td>'
            for column in columns
        )
        row_background = "rgba(255,255,255,.035)" if row_number % 2 == 0 else "rgba(0,0,0,.18)"
        body_rows.append(
            f'<tr style="background:{row_background};">{cells}</tr>'
        )

    table_html = f"""
    <div style="padding:1px;border-radius:15px;overflow:hidden;
                background:linear-gradient(115deg,#6F6248 0%,#E7D6A5 18%,
                #5E523D 38%,#B9A77A 58%,#39332A 78%,#D8C38D 100%);
                box-shadow:0 0 0 1px rgba(255,255,255,.08),
                0 10px 28px rgba(0,0,0,.34),0 0 18px rgba(190,160,95,.12);">
      <div style="overflow-x:auto;border-radius:14px;background:#090A0C;">
        <table style="width:100%;min-width:720px;border-collapse:separate;
                      border-spacing:0;background:#090A0C;">
          <thead><tr>{header_html}</tr></thead>
          <tbody>{''.join(body_rows)}</tbody>
        </table>
      </div>
    </div>
    """
    st.markdown(table_html, unsafe_allow_html=True)


def render_autogreen_dashboard():
    """Select BUY and CONFIRMATION independently from all four base-qualified target stocks."""
    st.title("6thSense-Vardaan")


    try:
        with st.spinner("Reading AutoGreen and Master ranking data from GitHub..."):
            base_rows, signal_rows = load_autogreen_buy_signals()
            master_bytes = get_autogreen_master_bytes()
            sum_o2l_ranks = get_autogreen_sum_o2l_ranks(master_bytes)
    except Exception as error:
        st.error(f"AutoGreen master data is unavailable: {error}")
        st.info("Signals are paused; no old or estimated workbook values are used.")
        return

    ranked_candidates = []
    for item in signal_rows:
        rank = sum_o2l_ranks.get(item["Symbol"])
        ranked_candidates.append({**item, "Sum O2L.10 Rank": rank})

    quote_map = {}
    if ranked_candidates:
        with st.spinner("Fetching live NSE prices and today's highs from Yahoo Finance..."):
            quote_map = get_autogreen_live_prices(tuple(item["Symbol"] for item in ranked_candidates))

    # Build the four target stocks from the four lowest qualified D-B values.
    qualified_detail_rows, qualified_detail_error = get_qualified_stock_detail_rows(
        base_rows, master_bytes
    )
    lowest_db_targets = []
    if not qualified_detail_error and qualified_detail_rows:
        for detail in qualified_detail_rows:
            raw_db = str(detail.get("16< D-B / Avg.4", "")).strip().replace("−", "-")
            outer_match = re.search(r"\(([^()]*)\)\s*\)?\s*$", raw_db)
            numeric_match = (
                re.search(r"([+\-]?\d+(?:,\d{3})*(?:\.\d+)?)\s*$", outer_match.group(1).strip())
                if outer_match else None
            )
            if numeric_match:
                try:
                    lowest_db_targets.append((
                        float(numeric_match.group(1).replace(",", "")),
                        detail["Stock"]
                    ))
                except (TypeError, ValueError):
                    pass
        lowest_db_targets.sort(key=lambda item: (item[0], item[1]))
        lowest_db_targets = lowest_db_targets[:4]

    target_symbols = [symbol for _, symbol in lowest_db_targets]
    target_quotes = {}
    target_table_rows = []
    if target_symbols:
        with st.spinner("Fetching live prices for the four target stocks..."):
            target_quotes = get_autogreen_live_prices(tuple(target_symbols))

    # Target selection is based on all four stocks from the base-qualified AutoGreen list.
    # Colour-qualified signals are preferred for their existing card palette, but
    # must not exclude a target stock from the BUY competition.
    base_by_symbol = {item["Symbol"]: item for item in base_rows}
    signal_by_symbol = {item["Symbol"]: item for item in ranked_candidates}
    target_records = []
    for position, (_, symbol) in enumerate(lowest_db_targets, start=1):
        quote = target_quotes.get(symbol, {})
        rank = sum_o2l_ranks.get(symbol)
        target_table_rows.append({
            "Role": f"TARGET #{position}",
            "Symbol": symbol,
            "Open": round(quote["open"], 2) if quote.get("open") is not None else "Unavailable",
            "High": round(quote["high"], 2) if quote.get("high") is not None else "Unavailable",
            "Low": round(quote["low"], 2) if quote.get("low") is not None else "Unavailable",
            "Live Price": round(quote["price"], 2) if quote.get("price") is not None else "Unavailable",
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

    # All four target stocks (#1-#4) compete equally for BUY.
    # Require price at least 0.20% below today's high; choose highest live
    # % change, then lowest rank as the tie-breaker.
    buy_candidates = [
        item for item in target_records
        if item["High Gap %"] >= 0.20
    ]
    buy_candidates.sort(key=lambda item: (
        -item["Live % Chg"],
        item["Sum O2L.10 Rank"] if item["Sum O2L.10 Rank"] is not None else float("inf"),
        item["Symbol"]
    ))
    best_buy = buy_candidates[0] if buy_candidates else None

    # All four target stocks compete independently for CONFIRMATION.
    # Apply the same High - 0.20% eligibility and rank by highest live % change,
    # using the lowest Sum O2L.10 rank as the tie-breaker.
    confirmation_candidates = [
        item for item in target_records
        if item["High Gap %"] >= 0.20
    ]
    confirmation_candidates.sort(key=lambda item: (
        -item["Live % Chg"],
        item["Sum O2L.10 Rank"] if item["Sum O2L.10 Rank"] is not None else float("inf"),
        item["Symbol"]
    ))
    # Keep an independent copy because the same stock may qualify as both BUY
    # and CONFIRMATION under these independent selection rules.
    confirmation_row = (
        dict(confirmation_candidates[0]) if confirmation_candidates else None
    )
    if confirmation_row:
        confirmation_row["_Role"] = "CONFIRMATION"
    if best_buy:
        best_buy["_Role"] = "BUY"
    selected_rows = ([best_buy] if best_buy else []) + (
        [confirmation_row] if confirmation_row else []
    )
    high_gap_candidates = buy_candidates
    buy_rows = [best_buy] if best_buy else []

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

    # Apply the same topmost-card palette to the complete dashboard interface.
    # The original signal cards keep their own card styling; all surrounding UI
    # Premium graphite dashboard: the active card colour is used as a restrained accent,
    # while the full interface stays clean, dark and consistent.
    st.markdown(
        f"""
        <style>
        .stApp {{
            background:radial-gradient(ellipse at 50% 0%,{metric_c1}24 0%,transparent 42%),
                       linear-gradient(145deg,#090B0F 0%,#101318 52%,#090B0F 100%) !important;
            color:#F3F5F8 !important;
        }}
        [data-testid="stHeader"] {{ background:#0B0D11 !important; }}
        [data-testid="stSidebar"] {{
            background:linear-gradient(180deg,#11151B 0%,#090B0F 100%) !important;
            border-right:1px solid #343A43 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ color:#E7EAF0 !important; }}
        h1,h2,h3,[data-testid="stSubheader"] {{ color:#F3F5F8 !important; }}
        h2,h3 {{ border-bottom:1px solid {metric_border}88 !important; }}
        [data-testid="stMetric"] {{
            background:linear-gradient(145deg,#191E25 0%,#101419 100%) !important;
            border:1px solid {metric_border}AA !important;
            border-top:2px solid {metric_accent} !important;
            border-radius:12px !important;
            color:#FFFFFF !important;
            box-shadow:0 5px 18px rgba(0,0,0,.28),0 0 10px {metric_glow}55 !important;
        }}
        [data-testid="stDataFrame"],[data-testid="stTable"] {{
            border:1px solid #343A43 !important;
            border-radius:10px !important;
            overflow:hidden !important;
            box-shadow:0 4px 18px rgba(0,0,0,.24) !important;
        }}
        .stButton > button,.stDownloadButton > button,
        [data-testid="stFormSubmitButton"] > button {{
            border:1px solid {metric_border} !important;
            color:#F7F8FA !important;
            background:linear-gradient(145deg,#202630 0%,#11161D 100%) !important;
        }}
        .stButton > button:hover,.stDownloadButton > button:hover,
        [data-testid="stFormSubmitButton"] > button:hover {{
            border-color:{metric_accent} !important;
            box-shadow:0 0 12px {metric_glow} !important;
        }}
        [data-testid="stTextInput"] input,[data-testid="stNumberInput"] input,
        [data-testid="stSelectbox"] > div,[data-testid="stFileUploader"] section,
        [data-testid="stExpander"] {{
            border-color:#414954 !important;
            background-color:#151A21 !important;
            color:#FFFFFF !important;
        }}
        [data-testid="stExpander"] {{ border-radius:10px !important; }}
        hr {{ border-color:#343A43 !important; }}
        [data-testid="stAlert"] {{ border-color:{metric_border}AA !important; }}
        </style>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div style="position:relative;overflow:hidden;width:100%;box-sizing:border-box;margin:12px 0 20px;padding:17px 12px;
                    border:1px solid {metric_outer};border-radius:17px;
                    background:linear-gradient(145deg,#1B2028 0%,#101419 58%,#1A1F27 100%);
                    box-shadow:inset 0 1px 0 rgba(255,255,255,.12),0 12px 30px rgba(0,0,0,.32),0 0 0 1px {metric_border}66,0 0 14px {metric_glow}55;">
          <div style="position:absolute;left:-72%;top:-55%;width:72%;height:210%;pointer-events:none;
                      background:linear-gradient(110deg,transparent 28%,{metric_shimmer}22 42%,#FFFFFF88 50%,{metric_shimmer}55 56%,transparent 70%);
                      filter:blur(7px);animation:noSignalFloatingGlow 6s ease-in-out infinite;"></div>
          <div style="position:relative;z-index:1;display:grid;grid-template-columns:repeat(4,minmax(0,1fr));align-items:center;">
            <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
              <div style="font-size:clamp(10px,1.8vw,12px);font-weight:800;letter-spacing:.35px;color:{metric_accent};line-height:1.35;">AutoGreen qualified</div>
              <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#FFFFFF;margin-top:5px;line-height:1.1;">{len(base_rows)}</div>
            </div>
            <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
              <div style="font-size:clamp(10px,1.8vw,12px);font-weight:800;letter-spacing:.35px;color:{metric_accent};line-height:1.35;">Colour-qualified stocks</div>
              <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#FFFFFF;margin-top:5px;line-height:1.1;">{len(ranked_candidates)}</div>
            </div>
            <div style="text-align:center;padding:4px 7px;border-right:1px solid {metric_border}66;">
              <div style="font-size:clamp(10px,1.8vw,12px);font-weight:800;letter-spacing:.35px;color:{metric_accent};line-height:1.35;">High − 0.20% qualified</div>
              <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#FFFFFF;margin-top:5px;line-height:1.1;">{len(high_gap_candidates)}</div>
            </div>
            <div style="text-align:center;padding:4px 7px;">
              <div style="font-size:clamp(10px,1.8vw,12px);font-weight:800;letter-spacing:.35px;color:{metric_accent};line-height:1.35;">BUY / Confirmation</div>
              <div style="font-size:clamp(23px,3.2vw,30px);font-weight:950;color:#FFFFFF;margin-top:5px;line-height:1.1;">{len(buy_rows)} / {1 if confirmation_row else 0}</div>
            </div>
          </div>
        </div>
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
                        text-shadow:0 0 12px rgba(255,210,225,.34);">6THSENSE · AUTOGREEN</div>
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
              No stock currently meets all BUY and confirmation conditions.
            </div>
          </div>
        </div>
        """
        st.markdown(no_signal_card, unsafe_allow_html=True)
    else:
        for index, item in enumerate(selected_rows):
            is_confirmation = item.get("_Role") == "CONFIRMATION"
            symbol = item["Symbol"]
            entry = item["Day High"] * (1 - 0.002)
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

    st.subheader("Selected BUY stocks and confirmation")
    if selected_rows:
        st.dataframe(
            [{
                "Role": "CONFIRMATION ONLY" if item.get("_Role") == "CONFIRMATION" else "BUY",
                "Symbol": item["Symbol"],
                "Open": round(item["Day Open"], 2) if item.get("Day Open") is not None else "Unavailable",
                "High": round(item["Day High"], 2),
                "Low": round(item["Day Low"], 2) if item.get("Day Low") is not None else "Unavailable",
                "Live Price": round(item["Live Price"], 2),
                "% Chg": f'{item["Live % Chg"]:+.2f}%',
                "Rank": item["Sum O2L.10 Rank"],
            } for i, item in enumerate(selected_rows)],
            use_container_width=True,
            hide_index=True
        )
    st.subheader("Target Stocks")
    if target_table_rows:
        st.dataframe(target_table_rows, use_container_width=True, hide_index=True)
    else:
        st.info("No four qualified target stocks with a readable numeric value are available yet.")

    with st.expander("View qualifying AutoGreen stock data", expanded=False):
        if base_rows:
            if qualified_detail_error:
                st.warning(f"Qualified-stock details could not be loaded: {qualified_detail_error}")
            elif qualified_detail_rows:
                render_qualified_stock_table(qualified_detail_rows)
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
        ["AutoGreen", "Master Excel"],
        index=0,
        key="main_dashboard_default_index"
    )

if page == "AutoGreen":
    render_autogreen_dashboard()


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
