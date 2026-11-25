import io
import re
import json
import time
import base64
import html
import requests
from datetime import datetime, timezone
from zoneinfo import ZoneInfo
import pandas as pd
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

# Reset once per Streamlit rerun so the message is shown at most once per dashboard load.
st.session_state["_master_file_unavailable_shown"] = False

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


def render_master_missing_no_signal_card():
    """Apply the teal dashboard theme and show the matching NO SIGNAL card."""
    # Keep the bottom dashboard shortcuts in the same teal palette as this
    # unavailable-data screen instead of retaining the previous dashboard palette.
    st.session_state["_vardaan_active_theme"] = {
        "deep": "#075B59",
        "primary": "#0B6663",
        "accent": "#27B7A8",
        "ink": "#123F40",
        "accent_text": "#137F77",
        "tint": "#D2F3EE",
        "canvas": "#E9FFFB",
        "border": "#58BDB1",
        "soft_text": "#164D4B",
    }
    no_master_card = """
    <div style="max-width:620px;margin:18px auto 24px;padding:1.5px;border-radius:20px;
                background:linear-gradient(110deg,#075B59 0%,#27B7A8 28%,#B8F3E8 48%,#138F86 72%,#075B59 100%);
                background-size:300% 300%;animation:masterTealBorder 6s linear infinite;
                box-shadow:0 0 0 1px rgba(255,255,255,.07),0 14px 34px rgba(0,0,0,.42),
                0 0 22px rgba(32,190,174,.30);">
      <div style="position:relative;overflow:hidden;text-align:center;padding:30px 24px 28px;
                  border-radius:18px;
                  background:radial-gradient(ellipse at 50% 0%,rgba(255,220,130,.20),transparent 54%),
                  linear-gradient(145deg,#082D31 0%,#0B6663 48%,#061F27 100%);
                  color:#F2FFFC;animation:masterTealFloat 5s ease-in-out infinite;">
        <div aria-hidden="true" style="position:absolute;top:-55%;left:-72%;width:72%;height:210%;
                    pointer-events:none;z-index:0;filter:blur(7px);
                    background:linear-gradient(105deg,transparent 0%,rgba(65,210,190,.12) 25%,
                    rgba(205,255,246,.48) 48%,rgba(255,255,255,.25) 58%,transparent 82%);
                    animation:masterTealGlow 6s ease-in-out infinite;"></div>
        <div style="position:relative;z-index:1;font-size:11px;font-weight:900;letter-spacing:3px;
                    color:#A7F3E8;text-shadow:0 0 12px rgba(90,230,210,.34);">6THSENSE · VARDAAN</div>
        <div style="position:relative;z-index:1;font-size:clamp(30px,5vw,42px);line-height:1.1;
                    font-weight:950;letter-spacing:2.2px;margin:13px 0 10px;color:#D8FFF7;
                    text-shadow:0 2px 18px rgba(70,220,195,.30);">NO SIGNAL</div>
        <div style="position:relative;z-index:1;width:74px;height:1px;margin:0 auto 15px;
                    background:linear-gradient(90deg,transparent,#42CDBB,transparent);"></div>
        <div style="position:relative;z-index:1;font-size:12px;font-weight:900;letter-spacing:1.65px;
                    color:#8FE5D8;">MASTER FILE NOT AVAILABLE</div>
        <div style="position:relative;z-index:1;font-size:13px;line-height:1.65;font-weight:550;
                    color:#B7E9E2;margin-top:10px;">
          Signal cards are hidden until the Master file is available.
        </div>
      </div>
    </div>
    <style>
    .stApp { --brand-flower-color:#27B7A8 !important; }
    [data-testid="stSidebar"] { --brand-flower-color:#27B7A8 !important; }
    
    /* Full dashboard theme: teal replaces the Rose-Cherry palette. */
    .stApp {
        background:radial-gradient(ellipse at 8% 0%,#C9F7F0 0%,transparent 42%),
                   linear-gradient(135deg,#E9FFFB 0%,#C9F0EA 52%,#F2FFFC 100%) !important;
        color:#123F40 !important;
    }
    [data-testid="stHeader"] {
        background:rgba(229,251,247,.96) !important;
        border-bottom:1px solid #58BDB1 !important;
    }
    [data-testid="stSidebar"] {
        background:linear-gradient(180deg,#A9E7DE 0%,#D2F3EE 48%,#EEFCF9 100%) !important;
        border-right:1.5px solid #238E85 !important;
        box-shadow:inset -8px 0 24px rgba(20,125,115,.12) !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
        color:#164D4B !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label,
    [data-testid="stSidebar"] [data-baseweb="select"] > div,
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        background:linear-gradient(135deg,#F3FFFC 0%,#D2F3EE 100%) !important;
        border-color:#64BFB4 !important;
        color:#123F40 !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {
        border-color:#137F77 !important;
        box-shadow:0 3px 12px rgba(19,127,119,.18) !important;
    }
    h1,h2,h3,h4,p,[data-testid="stMarkdownContainer"] {
        color:#123F40;
    }
    div[data-testid="stExpander"] {
        background:linear-gradient(135deg,#F7FFFD 0%,#DDF6F1 100%) padding-box,
                   linear-gradient(110deg,#137F77 0%,#64D6C8 48%,#137F77 100%) border-box !important;
        border:1.5px solid transparent !important;
        border-radius:14px !important;
        box-shadow:0 8px 22px rgba(18,100,92,.12) !important;
    }
    div[data-testid="stExpander"] summary,
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        color:#164D4B !important;
    }
    [data-testid="stTextInput"] input,
    [data-testid="stNumberInput"] input,
    [data-testid="stTextArea"] textarea,
    [data-testid="stSelectbox"] [data-baseweb="select"] > div {
        background:#F7FFFD !important;
        border-color:#64BFB4 !important;
        color:#123F40 !important;
    }
    [data-testid="stButton"] button,
    [data-testid="stDownloadButton"] button {
        border-color:#238E85 !important;
        color:#164D4B !important;
    }
    @keyframes masterTealBorder { 0% {background-position:200% 50%;} 100% {background-position:-100% 50%;} }
    @keyframes masterTealFloat { 0%,100% {transform:translateY(0);} 50% {transform:translateY(-3px);} }
    @keyframes masterTealGlow {
      0% {transform:translate3d(-35%,-5%,0) rotate(18deg);opacity:0;}
      18% {opacity:.62;} 52% {transform:translate3d(115%,3%,0) rotate(18deg);opacity:.42;}
      82% {opacity:.18;} 100% {transform:translate3d(190%,-4%,0) rotate(18deg);opacity:0;}
    }
    @media (prefers-reduced-motion:reduce) {
      @keyframes masterTealBorder { from {background-position:0 50%;} to {background-position:0 50%;} }
      @keyframes masterTealFloat { from {transform:none;} to {transform:none;} }
      @keyframes masterTealGlow { from {opacity:0;} to {opacity:0;} }
    }
    </style>
    """
    st.markdown(no_master_card, unsafe_allow_html=True)


# ============================================================

def show_master_file_unavailable_once():
    """Display the Master-file warning no more than once per dashboard run."""
    if not st.session_state.get("_master_file_unavailable_shown", False):
        st.error("Master file not available")
        st.session_state["_master_file_unavailable_shown"] = True


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
    # Clear workbook-derived results so the newly uploaded file is used immediately.
    for cache_name in ("load_autogreen_buy_signals", "load_stock_of_the_day_rows"):
        cached_function = globals().get(cache_name)
        if cached_function is not None:
            cached_function.clear()


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
    """Read the shared Master Excel directly from Vardaan Supabase Storage."""
    workbook_bytes = get_master()
    if not workbook_bytes:
        raise RuntimeError()
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


@st.cache_data(ttl=60, max_entries=1, show_spinner=False)
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


@st.cache_data(ttl=300, max_entries=10, show_spinner=False)
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
        sum_o2l_10_col = None
        sum_i_col = None
        o2l_columns = []
        o2l_headings = []
        db_key = header_key("16< D-B / Avg.4")
        sum_o2l_10_key = header_key("Sum O2L.10")
        sum_i_key = header_key("Sum I")
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
            if (
                db_key in found
                and sum_o2l_10_key in found
                and sum_i_key in found
                and len(dated_o2l) >= 2
            ):
                summary_header_row = row_number
                db_col = found[db_key]
                sum_o2l_10_col = found[sum_o2l_10_key]
                sum_i_col = found[sum_i_key]
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
                "Could not locate 16< D-B / Avg.4, Sum O2L.10, Sum I, and at least two O2L columns on AutoGreen."
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
            sum_o2l_10_value = (
                _autogreen_number(row[sum_o2l_10_col])
                if sum_o2l_10_col < len(row) else None
            )
            sum_i_value = (
                _autogreen_number(row[sum_i_col])
                if sum_i_col < len(row) else None
            )
            o2l_values = [
                _autogreen_number(row[index]) if index < len(row) else None
                for index in o2l_columns
            ]
            summary_by_symbol[symbol] = (
                db_value, o2l_values, sum_o2l_10_value, sum_i_value
            )

        matches = []
        for item in base_rows:
            symbol = item["Symbol"]
            o2h_values = o2h_by_symbol.get(symbol, [None, None])
            db_value, o2l_values, sum_o2l_10_value, sum_i_value = summary_by_symbol.get(
                symbol, (None, [None] * len(o2l_columns), None, None)
            )
            change1 = item.get("%Chg.1")
            change2 = item.get("%Chg.2")
            change3 = item.get("%Chg.3")
            change4 = item.get("%Chg.4")

            # Mandatory Stock of the Day conditions.
            if (
                len(o2h_values) < 2
                or any(value is None or value >= 0.50 for value in o2h_values[:2])
                or db_value is None or db_value > db_threshold
                or sum_o2l_10_value is None or sum_o2l_10_value >= -20.00
                or sum_i_value is None or sum_i_value >= 0
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
                "Sum O2L.10": round(sum_o2l_10_value, 2),
                "Sum I": round(sum_i_value, 2),
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
    st.markdown(
        """
        <style>
        [data-testid="stMainBlockContainer"] { padding-top: 3.5rem !important; }
        .block-container { padding-top: 3.5rem !important; }
        h1 { position: relative; }
        h1::before {
            content: ""; position: absolute; left: 50%; top: 0;
            width: 76px; height: 52px;
            transform: translate(-50%, -50%);
            z-index: 2;
            background: linear-gradient(110deg, var(--brand-flower-color, #C58A3A) 0%, var(--brand-flower-color, #C58A3A) 30%, rgba(255,255,255,.96) 48%, var(--brand-flower-color, #C58A3A) 66%, var(--brand-flower-color, #C58A3A) 100%);
            background-size: 300% 300%;
            animation: dependentSignalShimmer 4.2s linear infinite;
            -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
            mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
        }

        
        @keyframes dependentSignalShimmer {
            0% { background-position:0% 50%; filter:brightness(1); }
            50% { background-position:100% 50%; filter:brightness(1.12); }
            100% { background-position:0% 50%; filter:brightness(1); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("6thSense-Vardaan")

    try:
        master_bytes = get_autogreen_master_bytes()
        ranks = get_autogreen_sum_o2l_ranks(master_bytes)
    except Exception:
        render_master_missing_no_signal_card()
        return

    try:
        is_next_session_mode = is_stock_of_day_next_session_window()
        db_threshold = float(st.session_state.get("stock_day_db_threshold_value", -2.0))
        rows = load_stock_of_the_day_rows(db_threshold)

        # From 5 PM to 9 AM, Stock of the Day uses only stocks
        # with explicitly white font in AutoGreen Column A.
        if is_next_session_mode:
            white_font_symbols = get_white_font_master_column_a_symbols(master_bytes)
            existing_rows_by_symbol = {
                normalize_stock_symbol(row.get("Stock")): row
                for row in rows
                if normalize_stock_symbol(row.get("Stock"))
            }
            rows = [
                dict(existing_rows_by_symbol.get(symbol, {"Stock": symbol}), Stock=symbol)
                for symbol in white_font_symbols
            ]

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
    # Use the first Stock of the Day card's alphabet palette for the full sidebar.
    if rows:
        sidebar_border, sidebar_accent = stock_day_palettes.get(
            heading_symbol[:1].upper(), stock_day_palettes["A"]
        )
        sidebar_soft = {
            "A": "#E7F2E5", "B": "#E5F3F9", "C": "#F0EAF8", "D": "#F8F1DF",
            "E": "#F8E9ED", "F": "#FBEAF1", "G": "#E9F1FB", "H": "#EDF0F3",
            "I": "#E6F3F0", "J": "#E5F5F5", "K": "#E9F4EC", "L": "#F1F6E4",
            "M": "#FBF0E7", "N": "#F9EBE9", "O": "#F8F1E4", "P": "#F1ECF8",
            "Q": "#EBEEF9", "R": "#F8E9EB", "S": "#EEF3E9", "T": "#E7F3F2",
            "U": "#F0EBF8", "V": "#F3EAF3", "W": "#E8F3F8", "X": "#EEF0F2",
            "Y": "#F8F5E5", "Z": "#EAEFF4",
        }.get(heading_symbol[:1].upper(), "#E7F2E5")
        sidebar_text = "#344456"
    else:
        sidebar_border, sidebar_accent, sidebar_soft, sidebar_text = (
            "#8794A5", "#202833", "#E7EAEE", "#344456"
        )
    st.session_state["_vardaan_active_theme"] = {
        "deep": sidebar_text, "primary": heading_colour, "accent": sidebar_border,
        "ink": sidebar_text, "accent_text": sidebar_text,
        "tint": sidebar_soft, "canvas": "#FFFFFF",
        "border": sidebar_border, "soft_text": sidebar_text,
    }
    st.markdown(
        f"""
        <style>
        .stApp {{ --brand-flower-color:{heading_colour}; }}
        h1 {{ color:{heading_colour} !important; }}
        /* Match the complete workspace shell to the first visible Stock of the Day card. */
        .vardaan-welcome {{
            border-color:{sidebar_border}88 !important;
            background:radial-gradient(circle at 88% 12%,rgba(255,255,255,.30),transparent 27%),
                       linear-gradient(115deg,{sidebar_text} 0%,{heading_colour} 52%,{sidebar_border} 100%) !important;
            box-shadow:0 14px 34px {sidebar_border}28,inset 0 1px 0 rgba(255,255,255,.22) !important;
        }}
        .vardaan-kicker {{ color:#FFFFFF !important; }}
        .vardaan-heading {{ color:#FFFFFF !important; }}
        .vardaan-subheading {{ color:rgba(255,255,255,.88) !important; }}
        .vardaan-nav-label {{ color:{sidebar_text} !important; }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] > div[role="radiogroup"] {{
            border-color:{sidebar_border}88 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
            box-shadow:0 5px 18px {sidebar_border}18 !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label p {{
            color:{sidebar_text} !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
            border-color:{sidebar_border} !important;
            background:linear-gradient(135deg,{sidebar_text} 0%,{heading_colour} 100%) !important;
            box-shadow:0 5px 13px {sidebar_border}35 !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) p {{
            color:#FFFFFF !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stMetric"] {{
            border-color:{sidebar_border}88 !important;
            background:linear-gradient(145deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
            box-shadow:0 6px 18px {sidebar_border}18 !important;
        }}
        [data-testid="stMainBlockContainer"] .stButton > button,
        [data-testid="stMainBlockContainer"] .stDownloadButton > button {{
            border-color:{sidebar_border} !important;
            color:{sidebar_text} !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
        }}
        /* Sidebar follows the current topmost Stock of the Day card palette. */
        [data-testid="stSidebar"] {{
            background:linear-gradient(180deg,{sidebar_soft} 0%,#FFFFFF 100%) !important;
            border-right:1.5px solid {sidebar_border} !important;
            box-shadow:inset -8px 0 24px {sidebar_border}18 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{
            color:{sidebar_text} !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="border-bottom"] {{
            border-color:{sidebar_border}88 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="border:1px solid"] {{
            border-color:{sidebar_border}99 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="font-size:1.35rem"] {{
            color:{sidebar_accent} !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {{
            border:1px solid {sidebar_border}88 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
            box-shadow:0 2px 8px {sidebar_border}18 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {{
            border-color:{sidebar_accent} !important;
            background:linear-gradient(110deg,{sidebar_soft} 0%,#FFFFFF 100%) !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
            border:1.5px solid {sidebar_accent} !important;
            background:linear-gradient(110deg,{sidebar_soft} 0%,#FFFFFF 100%) !important;
            box-shadow:inset 4px 0 0 {sidebar_accent},0 3px 10px {sidebar_border}28 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label p {{
            color:{sidebar_text} !important;
        }}
        [data-testid="stSidebar"] .stButton > button {{
            border-color:{sidebar_border} !important;
            color:{sidebar_text} !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{sidebar_soft} 100%) !important;
        }}
        [data-testid="stSidebar"] .stButton > button:hover {{
            border-color:{sidebar_accent} !important;
            background:linear-gradient(110deg,{sidebar_soft} 0%,#FFFFFF 100%) !important;
        }}
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
        if not is_next_session_mode and (
            day_open is None or day_open <= 0 or day_high is None or day_high <= 0
        ):
            continue

        if is_next_session_mode:
            # No live quote is required after hours; the card announces the
            # next-session plan using the next session's opening price.
            entry = None
            target1 = None
            target2 = None
            status_text = "NEXT SESSION · BUY CONFIRMATION"
            rank_label = "Stock of the Day- Next Session"
            entry_display = "Open Price"
            target1_label = "TARGET PRICE"
            target1_display = "Open Price +0.50%"
            target2_html = ""
            grid_columns = 2
        else:
            # Preserve the existing daytime Stock of the Day calculations.
            entry = day_open
            target1 = entry * 1.005
            target2 = entry * 1.01
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
            rank_label = f"RANK {rank_text}"
            entry_display = f"₹{entry:,.2f}"
            target1_label = "TARGET 1"
            target1_display = f"₹{target1:,.2f}"
            target2_html = (
                f'<div style="min-width:0;">'
                f'<div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">TARGET 2</div>'
                f'<div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};margin-top:7px;white-space:nowrap;">₹{target2:,.2f}</div>'
                f'</div>'
            )
            grid_columns = 3

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
              {"BUY CONFIRMATION" if is_next_session_mode else "BUY SIGNAL"}
            </div>
          </div>
          <div style="font-size:.72rem;font-weight:900;letter-spacing:1.25px;color:#C5CFDC;
                      margin-top:9px;position:relative;z-index:1;color:{shimmer_colour};">{rank_label}</div>
          <div style="height:1px;background:linear-gradient(90deg,transparent,#C69B43,
                      rgba(255,236,184,.95),#C69B43,transparent);margin:19px 0 17px;
                      position:relative;z-index:1;"></div>
          <div style="display:grid;grid-template-columns:repeat({grid_columns},minmax(0,1fr));
                      align-items:center;gap:10px;position:relative;z-index:1;">
            <div style="min-width:0;">
              <div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">ENTRY PRICE</div>
              <div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};
                          margin-top:7px;white-space:nowrap;">{entry_display}</div>
            </div>
            <div style="min-width:0;">
              <div style="font-size:.66rem;font-weight:850;letter-spacing:.8px;color:{shimmer_colour};">{target1_label}</div>
              <div style="font-size:1.25rem;line-height:1.25;font-weight:950;color:{shimmer_colour};
                          margin-top:7px;white-space:nowrap;">{target1_display}</div>
            </div>
{target2_html}
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
            st.markdown(
                '<div style="font-weight:850;color:var(--text-color);">Find Next Stock at Your Own Risk</div>',
                unsafe_allow_html=True
            )
            # Restore the standard slider, keeping the permitted range -4.00 to -1.00.
            threshold_options = [-4.0, -3.5, -3.0, -2.5, -2.0, -1.5, -1.0]
            db_threshold = max(-4.0, min(-1.0, float(db_threshold)))
            saved_threshold = st.session_state.get(
                "stock_day_db_threshold_radio",
                st.session_state.get(
                    "stock_day_db_threshold_slider_v2",
                    st.session_state.get("stock_day_db_threshold_slider", -2.0)
                )
            )
            try:
                saved_threshold = max(-4.0, min(-1.0, float(saved_threshold)))
            except (TypeError, ValueError):
                saved_threshold = -2.0
            saved_threshold = round(round((saved_threshold + 4.0) / 0.5) * 0.5 - 4.0, 2)
            # Count stocks at every level from the complete -1.00 eligible set.
            # Each lower threshold is a stricter subset; all other mandatory conditions stay applied.
            threshold_base_rows = load_stock_of_the_day_rows(-1.0)
            threshold_counts = {
                option: sum(
                    1 for stock in threshold_base_rows
                    if float(stock.get("16< D-B / Avg.4", 9999)) <= option
                )
                for option in threshold_options
            }
            selected_db_threshold = st.slider(
                "D-B threshold",
                min_value=-4.0,
                max_value=-1.0,
                value=saved_threshold,
                step=0.5,
                format="%.2f",
                key="stock_day_db_threshold_slider_v2",
                help="Move the slider from -4.00 to -1.00. Stock counts update for the selected threshold."
            )
            selected_db_threshold = max(-4.0, min(-1.0, float(selected_db_threshold)))
            st.session_state["stock_day_db_threshold_value"] = selected_db_threshold
    

def is_stock_of_day_next_session_window():
    """Return True from 5:00 PM until 9:00 AM the next day, India time."""
    current_time = datetime.now(ZoneInfo("Asia/Kolkata")).time()
    return current_time >= datetime.strptime("17:00", "%H:%M").time() or current_time < datetime.strptime("09:00", "%H:%M").time()


def is_white_font_color(color):
    """Recognise explicit white font colours in common Excel encodings."""
    if color is None:
        return False
    if color.type == "rgb":
        return str(color.rgb or "").upper()[-6:] == "FFFFFF"
    if color.type == "indexed":
        # Excel's standard indexed palette uses 1 and 9 for white.
        return color.indexed in (1, 9)
    if color.type == "theme":
        # Theme 0 is Light 1 (white) in the standard Office theme.
        return color.theme == 0 and float(color.tint or 0) == 0
    return False


def get_white_font_master_column_a_symbols(workbook_bytes):
    """Collect every stock with an explicitly white font in AutoGreen Column A."""
    workbook = load_workbook(io.BytesIO(workbook_bytes), data_only=True, read_only=False)
    try:
        if "AutoGreen" not in workbook.sheetnames:
            return []

        ws = workbook["AutoGreen"]
        selected_symbols = []
        ignored_labels = {
            "SYMBOL", "STOCK", "STOCKNAME", "STOCK_NAME", "NAME",
            "SCRIP", "NSE", "TICKER", "COMPANY",
        }

        for row_number in range(1, (ws.max_row or 0) + 1):
            cell = ws.cell(row=row_number, column=1)
            if not isinstance(cell.value, str):
                continue

            symbol = normalize_stock_symbol(cell.value)
            if not symbol or symbol in ignored_labels or len(symbol) > 30:
                continue

            if is_white_font_color(cell.font.color):
                selected_symbols.append(symbol)

        return list(dict.fromkeys(selected_symbols))
    finally:
        workbook.close()


@st.cache_data(ttl=300, max_entries=2, show_spinner=False)
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


def get_index_toppers_master_values(symbols):
    """Read the two Index Toppers metrics from the same Master sheet and symbol column used by the Master dashboard."""
    wanted_headers = ("16< D-B / Avg.4", "16> C-B / Avg.4")
    wanted_symbols = {
        normalize_stock_symbol(symbol)
        for symbol in symbols
        if normalize_stock_symbol(symbol)
    }
    details = {}
    workbook = None

    def clean_header(value):
        return " ".join(str(value or "").strip().casefold().split())

    def display_master_value(value):
        if value is None or str(value).strip() == "":
            return "—"
        raw_value = str(value).strip().replace(",", "").replace("−", "-")
        # Master cells can contain a value with a parenthetical annotation;
        # use the numeric value inside the parentheses when present.
        matches = re.findall(
            r"\(\s*([-+]?(?:\d+(?:\.\d*)?|\.\d+))\s*\)",
            raw_value,
        )
        if matches:
            try:
                return "{:+.2f}".format(float(matches[-1]))
            except (TypeError, ValueError):
                pass
        try:
            return "{:+.2f}".format(float(raw_value.replace("%", "").strip()))
        except (TypeError, ValueError):
            return "—"

    try:
        master_bytes = get_autogreen_master_bytes()
        if not master_bytes:
            return {}
        workbook = load_workbook(
            io.BytesIO(master_bytes),
            data_only=True,
            read_only=True,
        )

        # Match the Master dashboard's sheet selection: prefer Summary,
        # then use the first sheet containing both required column headings.
        sheets = sorted(
            workbook.worksheets,
            key=lambda ws: (ws.title.casefold() != "summary", ws.title.casefold()),
        )
        selected_sheet = None
        header_row_number = None
        header_columns = {}

        for sheet in sheets:
            for row_number, row in enumerate(
                sheet.iter_rows(
                    min_row=1,
                    max_row=min(sheet.max_row or 1, 30),
                    values_only=True,
                ),
                start=1,
            ):
                found = {
                    clean_header(value): index
                    for index, value in enumerate(row)
                    if value is not None
                }
                if all(clean_header(name) in found for name in wanted_headers):
                    selected_sheet = sheet
                    header_row_number = row_number
                    header_columns = {
                        name: found[clean_header(name)] for name in wanted_headers
                    }
                    break
            if selected_sheet is not None:
                break

        if selected_sheet is None:
            return {}

        # The Master dashboard identifies each stock from column A.
        for row in selected_sheet.iter_rows(
            min_row=header_row_number + 1,
            values_only=True,
        ):
            if not row:
                continue
            matched_symbol = normalize_stock_symbol(row[0] if len(row) else None)
            if matched_symbol not in wanted_symbols:
                continue

            values = details.setdefault(matched_symbol, {})
            for heading in wanted_headers:
                column_index = header_columns[heading]
                raw_value = row[column_index] if column_index < len(row) else None
                rendered = display_master_value(raw_value)
                if rendered != "—" or heading not in values:
                    values[heading] = rendered

        return details
    except Exception:
        return {}
    finally:
        if workbook is not None:
            workbook.close()

def render_index_toppers_dashboard():
    """Rank the three requested NSE indices and build the live Top Choice card."""
    from zoneinfo import ZoneInfo

    # Refresh every 15 seconds during NSE hours; outside hours refresh once a minute.
    market_now = datetime.now(ZoneInfo("Asia/Kolkata"))
    market_open = (
        market_now.weekday() < 5
        and market_now.hour * 60 + market_now.minute >= 9 * 60 + 15
        and market_now.hour * 60 + market_now.minute <= 15 * 60 + 30
    )
    st_autorefresh(
        interval=15_000 if market_open else 60_000,
        limit=None,
        key="index_toppers_live_refresh",
    )

    st.markdown("""
    <style>
    /* Complete Chocolate Brown theme for the Index Toppers dashboard. */
    .stApp {
        background:radial-gradient(ellipse at 8% 0%,#E8D5C7 0%,transparent 44%),
                   linear-gradient(135deg,#F7F0EA 0%,#EAD9CC 52%,#FCF8F4 100%) !important;
        color:#3B2118 !important;
    }
    [data-testid="stHeader"] {
        background:rgba(247,240,234,.97) !important;
        border-bottom:1px solid #A97858 !important;
    }
    /* Light AutoGreen-style sidebar: bright, clean and easy to read. */
    [data-testid="stSidebar"] {
        background:linear-gradient(180deg,#FFFDFB 0%,#F7F0EA 58%,#F1DFD0 100%) !important;
        border-right:1.5px solid #D6B49A !important;
        box-shadow:inset -6px 0 18px rgba(138,86,56,.07) !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] [data-testid="stWidgetLabel"] {
        color:#4A2B20 !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {
        border:1px solid #D6B49A !important;
        border-radius:10px !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#F7F0EA 100%) !important;
        box-shadow:0 2px 7px rgba(74,43,32,.07) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {
        border-color:#A97858 !important;
        background:linear-gradient(110deg,#F7F0EA 0%,#EAD9CC 100%) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {
        border:1.5px solid #A97858 !important;
        background:linear-gradient(110deg,#EAD9CC 0%,#F7F0EA 100%) !important;
        box-shadow:inset 4px 0 0 #8A5638,0 3px 10px rgba(74,43,32,.10) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label p {
        color:#3B2118 !important;
        font-weight:750 !important;
    }
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button {
        border:1px solid #C99D7B !important;
        color:#4A2B20 !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#EAD9CC 100%) !important;
        box-shadow:0 3px 9px rgba(74,43,32,.08) !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover {
        border-color:#8A5638 !important;
        background:linear-gradient(110deg,#EAD9CC 0%,#D9B89C 100%) !important;
        color:#3B2118 !important;
    }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea,
    [data-testid="stSidebar"] [data-baseweb="select"] > div {
        border-color:#D6B49A !important;
        background:#FFFFFF !important;
        color:#3B2118 !important;
    }
    h1,h2,h3,p,[data-testid="stMarkdownContainer"] {color:#3B2118;}
    [data-testid="stDataFrame"],[data-testid="stTable"] {
        border:1px solid #A97858 !important;
        border-radius:12px !important;
    }
    /* Match Signal's light expander and white table surface in Index Toppers. */
    div[data-testid="stExpander"] {
        background:#FFFFFF !important;
        border:1px solid #D6B49A !important;
        border-radius:14px !important;
        overflow:hidden !important;
    }
    div[data-testid="stExpander"] summary {
        background:#F7F0EA !important;
        color:#4A2B20 !important;
        font-weight:700 !important;
    }
    div[data-testid="stExpander"] summary:hover {
        background:#EAD9CC !important;
    }
    div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
        background:#FFFFFF !important;
        color:#3B2118 !important;
        border-top:1px solid #E1CDBD !important;
    }
    [data-testid="stMainBlockContainer"] .stButton > button,
    [data-testid="stMainBlockContainer"] [data-testid="stDownloadButton"] button {
        border:1px solid #8A5638 !important;
        color:#FFF4EB !important;
        background:linear-gradient(110deg,#4A2B20 0%,#754A34 55%,#8A5638 100%) !important;
        box-shadow:0 5px 14px rgba(74,43,32,.18) !important;
        border-radius:10px !important;
    }
    [data-testid="stMainBlockContainer"] .stButton > button:hover,
    [data-testid="stMainBlockContainer"] [data-testid="stDownloadButton"] button:hover {
        border-color:#C99D7B !important;
        filter:brightness(1.08) !important;
    }
    .index-topper-hero {
        padding:23px 25px;margin:5px 0 18px;border-radius:19px;
        color:#FFF8F2;border:1px solid #A97858;
        background:radial-gradient(circle at 88% 12%,rgba(255,255,255,.24),transparent 27%),
                   linear-gradient(115deg,#3B2118 0%,#754A34 48%,#4A2B20 100%);
        box-shadow:0 14px 34px rgba(59,33,24,.22),inset 0 1px 0 rgba(255,255,255,.18);
    }
    .index-topper-kicker {font-size:10px;font-weight:900;letter-spacing:2.5px;color:#E7CBB7;}
    .index-topper-title {font-size:clamp(27px,4vw,36px);font-weight:950;line-height:1.12;margin:7px 0;color:#FFFFFF;}
    .index-topper-sub {font-size:13px;line-height:1.55;color:#F1DFD0;}
    .index-topper-card {
        position:relative;overflow:hidden;margin:12px 0 20px;padding:23px 24px;border-radius:20px;
        color:#FFF8F2;border:2px solid transparent;
        background:linear-gradient(130deg,#3B2118 0%,#754A34 48%,#4A2B20 100%) padding-box,
                   linear-gradient(115deg,#8A5638 0%,#E7CBB7 35%,#FFFFFF 52%,#C99D7B 75%,#4A2B20 100%) border-box;
        box-shadow:0 12px 32px rgba(59,33,24,.22),0 0 25px rgba(138,86,56,.16);
    }
    .index-topper-card-label {font-size:11px;font-weight:950;letter-spacing:2.4px;color:#E7CBB7;text-align:center;}
    .index-topper-symbol {font-size:clamp(26px,4vw,35px);font-weight:950;color:#FFFFFF;letter-spacing:.7px;}
    .index-topper-pill {font-size:11px;font-weight:950;letter-spacing:1.1px;color:#3B2118;
        background:linear-gradient(110deg,#F1DFD0,#D6B49A);padding:8px 13px;border-radius:999px;}
    .index-topper-metric-label {font-size:10px;font-weight:850;letter-spacing:1.3px;color:#E7CBB7;}
    .index-topper-metric-value {font-size:clamp(21px,3vw,28px);font-weight:950;color:#FFFFFF;margin-top:5px;}
    .index-topper-target {color:#E7CBB7 !important;}
    .index-topper-table-wrap {overflow-x:auto;border:1px solid #A97858;border-radius:12px;background:#F7F0EA;}
    .index-topper-table {width:100%;min-width:900px;border-collapse:collapse;color:#3B2118;}
    .index-topper-table th {padding:11px 12px;text-align:left;white-space:nowrap;background:linear-gradient(110deg,#D9B89C,#EAD9CC);
        color:#4A2B20;font-size:12px;font-weight:900;border-bottom:2px solid #8A5638;}
    .index-topper-table td {padding:10px 12px;white-space:nowrap;font-size:12px;font-weight:650;border-bottom:1px solid #E1CDBD;}
    .index-topper-table tr:nth-child(even) {background:#FAF3ED;}
    .index-topper-table tr.top-three {background:#EAD9CC;font-weight:850;}
    .index-topper-table tr.top-choice {background:#C99D7B;font-weight:950;}
    .index-topper-section {color:#754A34;font-size:14px;font-weight:900;letter-spacing:1px;margin:18px 0 9px;}
    [data-testid="stMainBlockContainer"] {
        --signal-nav-border:#A97858;
        --signal-nav-accent:#754A34;
        --signal-nav-soft:#EAD9CC;
        --signal-nav-canvas:#F1DFD0;
        --signal-nav-text:#3B2118;
    }
    

    [data-testid="stMainBlockContainer"] { padding-top: 3.5rem !important; }
    .block-container { padding-top: 3.5rem !important; }
    h1 { position: relative; }
    h1::before {
        content: ""; position: absolute; left: 50%; top: 0;
        width: 76px; height: 52px;
        transform: translate(-50%, -50%);
            z-index: 2;
        background: linear-gradient(110deg, var(--brand-flower-color, #8A5638) 0%, var(--brand-flower-color, #8A5638) 30%, rgba(255,255,255,.96) 48%, var(--brand-flower-color, #8A5638) 66%, var(--brand-flower-color, #8A5638) 100%);
            background-size: 300% 300%;
            animation: dependentSignalShimmer 4.2s linear infinite;
        -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
        mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
    }


    
        @keyframes dependentSignalShimmer {
            0% { background-position:0% 50%; filter:brightness(1); }
            50% { background-position:100% 50%; filter:brightness(1.12); }
            100% { background-position:0% 50%; filter:brightness(1); }
        }
        </style>
    """, unsafe_allow_html=True)

    st.title("6thSense-Vardaan")

    st.markdown("""
    <div class="index-topper-hero">
      <div class="index-topper-kicker">6THSENSE · LIVE MARKET SELECTION</div>
      <div class="index-topper-title">Index Toppers</div>
    </div>
    """, unsafe_allow_html=True)

    index_specs = [
        ("NIFTY 50", "50"),
        ("NIFTY NEXT 50", "NEXT 50"),
        ("NIFTY MIDCAP 50", "MID 50"),
    ]

    try:
        index_groups = load_github_index_groups()
        missing_groups = [label for label, key in index_specs if key not in index_groups]
        if missing_groups:
            raise ValueError("Constituent list is missing: " + ", ".join(missing_groups))

        candidates = []
        symbols = []
        for index_label, group_key in index_specs:
            for member in index_groups[group_key]:
                symbol = normalize_stock_symbol(member.get("NSE Symbol"))
                if symbol and symbol not in symbols:
                    symbols.append(symbol)

        with st.spinner("Fetching live quotes for the three indices..."):
            quotes = get_autogreen_live_prices(tuple(symbols))

        for index_label, group_key in index_specs:
            index_rows = []
            for member in index_groups[group_key]:
                symbol = normalize_stock_symbol(member.get("NSE Symbol"))
                quote = quotes.get(symbol, {})
                change = quote.get("change_pct")
                if not symbol or change is None:
                    continue
                index_rows.append({
                    "Index": index_label,
                    "Stock Name": symbol,
                    "% Chg": float(change),
                    "Price": quote.get("price"),
                    "Open": quote.get("open"),
                    "High": quote.get("high"),
                    "Low": quote.get("low"),
                })
            index_rows.sort(key=lambda row: row["% Chg"], reverse=True)
            for row in index_rows[:5]:
                open_price = row.get("Open")
                low_price = row.get("Low")
                high_price = row.get("High")
                if open_price is not None and open_price > 0 and low_price is not None:
                    row["O2L%"] = ((low_price - open_price) / open_price) * 100
                else:
                    row["O2L%"] = None
                if open_price is not None and open_price > 0 and high_price is not None:
                    row["O2H%"] = ((high_price - open_price) / open_price) * 100
                else:
                    row["O2H%"] = None
                candidates.append(row)

        # Load Master values for all candidates before selection so Avg.O2L.4
        # can control each stock's individual Open-based price cap.
        index_toppers_master_values = get_index_toppers_master_values(
            tuple(dict.fromkeys(row["Stock Name"] for row in candidates))
        )
        master_values = index_toppers_master_values

        def get_numeric_avg_o2l4(stock_row):
            master_row = master_values.get(
                normalize_stock_symbol(stock_row["Stock Name"]), {}
            )
            raw_value = master_row.get("16< D-B / Avg.4", "—")
            match = re.search(
                r"[-+]?(?:\d+(?:\.\d*)?|\.\d+)",
                str(raw_value).replace("−", "-").replace(",", ""),
            )
            try:
                return float(match.group(0)) if match else None
            except (TypeError, ValueError):
                return None

        for candidate in candidates:
            candidate["_avg_o2l4_numeric"] = get_numeric_avg_o2l4(candidate)

        # Rank all candidates by lowest (most negative) O2L%.
        o2l_ranked = sorted(
            [row for row in candidates if row.get("O2L%") is not None],
            key=lambda row: row["O2L%"]
        )
        lowest_three = o2l_ranked[:3]
        next_two = o2l_ranked[3:5]

        def get_eligible_stocks(stock_rows):
            eligible_rows = []
            for row in stock_rows:
                open_price = row.get("Open")
                live_price = row.get("Price")
                if (
                    open_price is None or open_price <= 0
                    or live_price is None
                    or row.get("O2H%") is None
                ):
                    continue

                # Avg.O2L.4 below -3.00 gives this stock its own +1.50% cap.
                cap_multiplier = (
                    1.015 if row.get("_avg_o2l4_numeric") is not None
                    and row["_avg_o2l4_numeric"] < -3.00 else 1.01
                )
                row["_open_cap_multiplier"] = cap_multiplier

                if open_price < live_price < open_price * cap_multiplier:
                    eligible_rows.append(row)

            eligible_rows.sort(
                key=lambda row: (round(row["O2H%"], 8), row["% Chg"])
            )
            return eligible_rows

        # First attempt: the three lowest O2L% stocks.
        primary_eligible = get_eligible_stocks(lowest_three)
        fallback_mode = False
        fallback_eligible = []

        # If none qualifies, check the next two lowest O2L% stocks (#4 and #5).
        if not primary_eligible:
            fallback_eligible = get_eligible_stocks(next_two)
            if fallback_eligible:
                fallback_mode = True

        eligible = primary_eligible if primary_eligible else fallback_eligible
        top_choice = eligible[0] if eligible else None
        highlighted_candidates = lowest_three + (next_two if fallback_mode else [])

        st.session_state['_index_toppers_pistachio_mode'] = bool(fallback_mode)

        if fallback_mode:
            st.markdown("""
            <style>
            :root { --brand-flower-color:#8FAF72 !important; }
            .stApp {
                background:radial-gradient(ellipse at 8% 0%,#E7F0D9 0%,transparent 44%),
                           linear-gradient(135deg,#F4F8EF 0%,#DDEBCF 52%,#FBFDF7 100%) !important;
                color:#365B3B !important;
            }
            [data-testid="stHeader"] {
                background:rgba(255,247,237,.97) !important;
                border-bottom:1px solid #8FAF72 !important;
            }
            [data-testid="stMainBlockContainer"] {
                --signal-nav-border:#B7D7A8;
                --signal-nav-accent:#658B4E;
                --signal-nav-soft:#DDEBCF;
                --signal-nav-canvas:#F4F8EF;
                --signal-nav-text:#365B3B;
            }
            .index-topper-hero {
                border-color:#8FAF72 !important;
                background:radial-gradient(circle at 88% 12%,rgba(255,255,255,.28),transparent 27%),
                           linear-gradient(115deg,#365B3B 0%,#658B4E 48%,#4D7440 100%) !important;
                box-shadow:0 14px 34px rgba(194,65,12,.22),inset 0 1px 0 rgba(255,255,255,.18) !important;
            }
            .index-topper-kicker,.index-topper-metric-label { color:#C8DDB5 !important; }
            .index-topper-sub { color:#DDEBCF !important; }
            .index-topper-card {
                background:linear-gradient(130deg,#365B3B 0%,#658B4E 48%,#4D7440 100%) padding-box,
                           linear-gradient(115deg,#8FAF72 0%,#B7D7A8 35%,#FFFFFF 52%,#A9C78E 75%,#365B3B 100%) border-box !important;
                box-shadow:0 12px 32px rgba(194,65,12,.24),0 0 25px rgba(249,115,22,.2) !important;
            }
            .index-topper-card-label { color:#C8DDB5 !important; }
            .index-topper-pill {
                color:#365B3B !important;
                background:linear-gradient(110deg,#DDEBCF,#B7D7A8) !important;
            }
            .index-topper-target { color:#C8DDB5 !important; }
            .index-topper-table-wrap { border-color:#8FAF72 !important; background:#F4F8EF !important; }
            .index-topper-table { color:#365B3B !important; }
            .index-topper-table th {
                background:linear-gradient(110deg,#B7D7A8,#DDEBCF) !important;
                color:#365B3B !important;border-bottom-color:#658B4E !important;
            }
            .index-topper-table td { border-bottom-color:#C8DDB5 !important; }
            .index-topper-table tr:nth-child(even) { background:#F4F8EF !important; }
            .index-topper-table tr.top-three { background:#DDEBCF !important; }
            .index-topper-table tr.top-choice { background:#B7D7A8 !important; }
            .index-topper-section { color:#4D7440 !important; }
            div[data-testid="stExpander"] { border-color:#B7D7A8 !important; }
            div[data-testid="stExpander"] summary {
                background:#F4F8EF !important;color:#365B3B !important;
            }
            div[data-testid="stExpander"] summary:hover { background:#DDEBCF !important; }
            div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
                background:#FFFFFF !important;color:#365B3B !important;
                border-top-color:#C8DDB5 !important;
            }
            [data-testid="stMainBlockContainer"] .stButton > button,
            [data-testid="stMainBlockContainer"] [data-testid="stDownloadButton"] button {
                border-color:#658B4E !important;color:#F4F8EF !important;
                background:linear-gradient(110deg,#365B3B 0%,#658B4E 55%,#8FAF72 100%) !important;
                box-shadow:0 5px 14px rgba(194,65,12,.2) !important;
            }
            [data-testid="stMainBlockContainer"] .stButton > button:hover,
            [data-testid="stMainBlockContainer"] [data-testid="stDownloadButton"] button:hover {
                border-color:#B7D7A8 !important;filter:brightness(1.08) !important;
            }

            /* Pistachio applies to the #4/#5 fallback stock dashboard surfaces. */
            [data-testid="stSidebar"] {background:linear-gradient(180deg,#FBFDF7 0%,#F4F8EF 58%,#E7F0D9 100%) !important;border-right:1.5px solid #C8DDB5 !important;}
            [data-testid="stSidebar"] [data-testid="stMarkdownContainer"],[data-testid="stSidebar"] label,[data-testid="stSidebar"] [data-testid="stWidgetLabel"] {color:#365B3B !important;}
            [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {border-color:#C8DDB5 !important;background:linear-gradient(135deg,#FFFFFF,#F4F8EF) !important;}
            [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {border-color:#8FAF72 !important;background:#DDEBCF !important;box-shadow:inset 4px 0 0 #658B4E,0 3px 10px rgba(54,91,59,.12) !important;}
            [data-testid="stSidebar"] .stButton > button,[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button {border-color:#8FAF72 !important;color:#365B3B !important;background:linear-gradient(135deg,#FFFFFF,#DDEBCF) !important;}
            [data-testid="stSidebar"] .stButton > button:hover,[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover {background:linear-gradient(110deg,#DDEBCF,#B7D7A8) !important;}
            [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label {border-color:#C8DDB5 !important;background:linear-gradient(135deg,#FFFFFF,#F4F8EF) !important;}
            [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {border-color:#8FAF72 !important;background:#DDEBCF !important;box-shadow:inset 4px 0 0 #658B4E,0 3px 10px rgba(54,91,59,.12) !important;}
            [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label,[data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label * {color:#365B3B !important;-webkit-text-fill-color:#365B3B !important;}
            </style>
            """, unsafe_allow_html=True)


        # Build the display table rows first. The Top Choice card reads its
        # two Master metrics directly from this same table row.
        master_values = index_toppers_master_values
        table_rows = []
        for row in candidates:
            table_rows.append({
                "Stock Name": row["Stock Name"],
                "Index": row["Index"],
                "% Chg": row["% Chg"],
                "Price": row.get("Price"),
                "O2L%": row.get("O2L%"),
                "O2H%": row.get("O2H%"),
                "Open": row.get("Open"),
                "High": row.get("High"),
                "Low": row.get("Low"),
                "Avg.O2L.4": f'({master_values.get(normalize_stock_symbol(row["Stock Name"]), {}).get("16< D-B / Avg.4", "—")})',
                "Avg.O2H.4": f'({master_values.get(normalize_stock_symbol(row["Stock Name"]), {}).get("16> C-B / Avg.4", "—")})',
            })

        if top_choice:
            top_choice_table_row = next(
                (
                    row for row in table_rows
                    if normalize_stock_symbol(row["Stock Name"])
                    == normalize_stock_symbol(top_choice["Stock Name"])
                ),
                {},
            )
            top_choice_avg_o2l = top_choice_table_row.get("Avg.O2L.4", "(—)")
            top_choice_avg_o2h = top_choice_table_row.get("Avg.O2H.4", "(—)")
            top_choice_master_parts = []
            if top_choice_avg_o2l != "—":
                top_choice_master_parts.append(
                    f"Avg.O2L.4 {html.escape(top_choice_avg_o2l)}"
                )
            if top_choice_avg_o2h != "—":
                top_choice_master_parts.append(
                    f"Avg.O2H.4 {html.escape(top_choice_avg_o2h)}"
                )
            top_choice_master_line = (
                " · " + " · ".join(top_choice_master_parts)
                if top_choice_master_parts else ""
            )
            entry = top_choice["Price"]
            entry_based_target = entry * 1.0035
            open_based_cap = (
                top_choice["Open"] * top_choice.get("_open_cap_multiplier", 1.01)
                if top_choice.get("Open") is not None and top_choice["Open"] > 0
                else entry_based_target
            )
            target = min(entry_based_target, open_based_cap)
            st.markdown(f"""
            <div class="index-topper-card">
              <div class="index-topper-card-label">6THSENSE · INDEX TOPPERS SIGNAL</div>
              <div style="display:flex;align-items:center;justify-content:space-between;gap:12px;margin-top:16px;flex-wrap:wrap;">
                <div class="index-topper-symbol">{html.escape(top_choice["Stock Name"])}</div>
                <div class="index-topper-pill">TOP CHOICE</div>
              </div>
              <div style="font-size:11px;color:#D4D5D9;margin-top:8px;letter-spacing:.8px;">
                {html.escape(top_choice["Index"])} · Live % Chg {top_choice["% Chg"]:+.2f}% · O2L {top_choice["O2L%"]:+.2f}% · O2H {top_choice["O2H%"]:+.2f}%{top_choice_master_line}
              </div>
              <div style="height:1px;background:linear-gradient(90deg,transparent,#D8B65A,#D9DCE1,transparent);margin:20px 0 17px;"></div>
              <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:16px;">
                <div><div class="index-topper-metric-label">ENTRY PRICE</div>
                  <div class="index-topper-metric-value">₹{entry:,.2f}</div></div>
                <div><div class="index-topper-metric-label">TARGET PRICE</div>
                  <div class="index-topper-metric-value index-topper-target">₹{target:,.2f}</div></div>
              </div>
            </div>
            """, unsafe_allow_html=True)
        else:
            st.markdown("""
            <div class="index-topper-card" style="text-align:center;">
              <div class="index-topper-card-label">6THSENSE-VARDAAN</div>
              <div style="font-size:30px;font-weight:950;margin:14px 0;color:#F4E8C5;">TOP CHOICE STOCK AWAITED</div>
              <div style="font-size:13px;color:#D4D5D9;">No Stock for Top Choice is available.</div>
            </div>
            """, unsafe_allow_html=True)

        with st.expander("TOP 5 LIVE % GAINERS · EACH INDEX", expanded=False):
            if candidates:
                columns = ["Stock Name", "Index", "% Chg", "Price", "O2L%", "O2H%", "Open", "High", "Low", "Avg.O2L.4", "Avg.O2H.4"]
                # Native Streamlit dataframe enables sorting by clicking every column heading.
                # Preserve the existing highlights for the lowest-three candidates and Top Choice.
                table_df = pd.DataFrame(table_rows, columns=columns)

                # Pistachio is used only for the #4/#5 fallback dashboard.
                # The normal dashboard keeps its existing chocolate table palette.
                pistachio_table = bool(fallback_mode)
                top_choice_row_style = (
                    "background-color: #B7D7A8; color: #365B3B; font-weight: 900;"
                    if pistachio_table else
                    "background-color: #C99D7B; color: #3B2118; font-weight: 900;"
                )
                highlighted_row_style = (
                    "background-color: #DDEBCF; color: #365B3B; font-weight: 800;"
                    if pistachio_table else
                    "background-color: #EAD9CC; color: #4A2B20; font-weight: 800;"
                )
                base_row_style = (
                    "background-color: #F4F8EF; color: #365B3B;"
                    if pistachio_table else
                    "background-color: #FFFFFF; color: #3B2118;"
                )

                def highlight_index_topper_row(row):
                    stock_name = row.get("Stock Name")
                    if top_choice and stock_name == top_choice["Stock Name"]:
                        style = top_choice_row_style
                    elif any(stock_name == item["Stock Name"] for item in highlighted_candidates):
                        style = highlighted_row_style
                    else:
                        style = base_row_style
                    return [style] * len(row)

                table_header_background = "#B7D7A8" if pistachio_table else "#D9B89C"
                table_header_text = "#365B3B" if pistachio_table else "#4A2B20"
                table_header_border = "#658B4E" if pistachio_table else "#8A5638"
                table_cell_border = "#C8DDB5" if pistachio_table else "#E1CDBD"
                table_background = "#F4F8EF" if pistachio_table else "#F7F0EA"
                table_border = "#8FAF72" if pistachio_table else "#A97858"
                table_cell_text = "#365B3B" if pistachio_table else "#3B2118"

                styled_table = (
                    table_df.style
                    .set_table_styles([
                        {
                            "selector": "th",
                            "props": [
                                ("background-color", table_header_background),
                                ("color", table_header_text),
                                ("font-weight", "900"),
                                ("border-bottom", f"2px solid {table_header_border}"),
                            ],
                        },
                        {
                            "selector": "tbody tr:nth-child(odd) td",
                            "props": [
                                ("background-color", "#F4F8EF" if pistachio_table else "#FFFFFF"),
                                ("color", table_cell_text),
                                ("border-bottom", f"1px solid {table_cell_border}"),
                            ],
                        },
                        {
                            "selector": "tbody tr:nth-child(even) td",
                            "props": [
                                ("background-color", "#FFFFFF"),
                                ("color", table_cell_text),
                                ("border-bottom", f"1px solid {table_cell_border}"),
                            ],
                        },
                        {
                            "selector": "table",
                            "props": [
                                ("background-color", table_background),
                                ("border", f"1px solid {table_border}"),
                                ("border-collapse", "collapse"),
                            ],
                        },
                    ])
                    .set_table_styles([
                        {
                            "selector": "th:nth-child(10), td:nth-child(10), th:nth-child(11), td:nth-child(11)",
                            "props": [
                                ("width", "112px"),
                                ("min-width", "112px"),
                                ("max-width", "112px"),
                                ("white-space", "nowrap"),
                                ("text-align", "center"),
                            ],
                        },
                    ], overwrite=False)
                    .apply(highlight_index_topper_row, axis=1)
                    .format({
                        "% Chg": "{:+.2f}%",
                        "Price": "₹{:,.2f}",
                        "O2L%": "{:+.2f}%",
                        "O2H%": "{:+.2f}%",
                        "Open": "₹{:,.2f}",
                        "High": "₹{:,.2f}",
                        "Low": "₹{:,.2f}",
                    }, na_rep="—")
                )
                st.dataframe(
                    styled_table,
                    hide_index=True,
                    use_container_width=True,
                )
            else:
                st.warning("Live stock quotes are not available for the selected indices right now.")

    except Exception as error:
        st.error(f"Index Toppers could not be loaded: {error}")
        st.info("The dashboard will retry automatically. Check the index constituent CSV and live quote availability if this continues.")



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
    st.markdown(
        """
        <style>
        [data-testid="stMainBlockContainer"] { padding-top: 3.5rem !important; }
        .block-container { padding-top: 3.5rem !important; }
        h1 { position: relative; }
        h1::before {
            content: ""; position: absolute; left: 50%; top: 0;
            width: 76px; height: 52px;
            transform: translate(-50%, -50%);
            z-index: 2;
            background: linear-gradient(110deg, var(--brand-flower-color, #C58A3A) 0%, var(--brand-flower-color, #C58A3A) 30%, rgba(255,255,255,.96) 48%, var(--brand-flower-color, #C58A3A) 66%, var(--brand-flower-color, #C58A3A) 100%);
            background-size: 300% 300%;
            animation: dependentSignalShimmer 4.2s linear infinite;
            -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
            mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
        }

        
        @keyframes dependentSignalShimmer {
            0% { background-position:0% 50%; filter:brightness(1); }
            50% { background-position:100% 50%; filter:brightness(1.12); }
            100% { background-position:0% 50%; filter:brightness(1); }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    st.title("6thSense-Vardaan")




    try:
        with st.spinner("Reading AutoGreen and Master ranking data..."):
            base_rows, signal_rows = load_autogreen_buy_signals()
            master_bytes = get_autogreen_master_bytes()
            sum_o2l_ranks = get_autogreen_sum_o2l_ranks(master_bytes)
    except Exception:
        render_master_missing_no_signal_card()
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
        first_history_change = base_by_symbol.get(symbol, {}).get("%Chg.1")
        try:
            first_history_change = float(
                str(first_history_change).replace(",", "").replace("%", "").strip()
            )
        except (TypeError, ValueError):
            first_history_change = None
        # Positive first %Chg history gets a tighter +0.75% O2H target;
        # all other stocks retain the standard +1.00% target.
        target_o2h_pct = 0.75 if first_history_change is not None and first_history_change > 0 else 1.0
        target_status = (
            "Target Completed"
            if o2h_pct is not None and o2h_pct >= target_o2h_pct
            else "Signal Awaited"
        )
        # Use the same BUY-card target rule: +0.35% from live entry,
        # capped at this stock's Open-based O2H target (+0.75% or +1.00%).
        entry_based_target = live_price * 1.0035 if live_price is not None else None
        open_based_max_target = (
            open_price * (1 + target_o2h_pct / 100)
            if open_price is not None and open_price > 0
            else entry_based_target
        )
        target_price = (
            min(entry_based_target, open_based_max_target)
            if entry_based_target is not None
            else None
        )
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
            "Target Price": f"₹{target_price:,.2f}" if target_price is not None else "Unavailable",
        })

        # Every one of the four D-B target stocks must reach BUY evaluation.
        # Do not drop a target merely because it has no colour-qualified signal row.
        signal = dict(
            signal_by_symbol.get(symbol)
            or base_by_symbol.get(symbol)
            or {"Symbol": symbol}
        )
        signal.setdefault("Card Colour", "Copper-Orange")
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
            "O2H Target %": target_o2h_pct,
            "Sum O2L.10 Rank": rank,
            "_Target Position": position,
        })

    # BUY CONFIRMATION is evaluated independently for every one of the four target stocks.
    # O2L must be below -1%, O2H must remain below that stock's assigned target,
    # and live price must be above Open. Positive first %Chg history uses +0.75%;
    # all other stocks use +1.00%.
    confirmation_candidates = [
        item for item in target_records
        if item.get("Day Open") is not None
        and item.get("Day High") is not None
        and item.get("Day Low") is not None
        and item["Day Open"] > 0
        and ((item["Day Low"] - item["Day Open"]) / item["Day Open"]) * 100 < -1.0
        and ((item["Day High"] - item["Day Open"]) / item["Day Open"]) * 100
            < item.get("O2H Target %", 1.0)
        and item["Live Price"] > item["Day Open"]
    ]
    confirmation_candidates.sort(key=lambda item: (
        -item["Live % Chg"],
        item["Sum O2L.10 Rank"] if item["Sum O2L.10 Rank"] is not None else float("inf"),
        item["Symbol"]
    ))
    for confirmation_row in confirmation_candidates:
        confirmation_row["_Role"] = "CONFIRMATION"

    # If a target has a visible BUY CONFIRMATION card, reflect that in its target status.
    confirmation_symbols = {item["Symbol"] for item in confirmation_candidates}
    for target_row in target_table_rows:
        if (
            target_row["Symbol"] in confirmation_symbols
            and target_row["Target Status"] != "Target Completed"
        ):
            target_row["Target Status"] = "Signal Generated"

    # Once the assigned O2H target is reached, replace the numeric target price
    # with the completion label in the four-stock Target Stocks table.
    for target_row in target_table_rows:
        if target_row["Target Status"] == "Target Completed":
            target_row["Target Price"] = "Target Completed"

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
        "Copper-Orange": ("#3D1D0B", "#A64B16", "#54250C", "#F2A15A", "#FFF0E2", "#FFD0A6"),
    }
    if selected_rows:
        metric_colour_name = selected_rows[0].get("Card Colour", "Copper-Orange")
        metric_c1, metric_c2, metric_c3, metric_border, metric_accent, metric_shimmer = metric_palettes.get(
            metric_colour_name, metric_palettes["Copper-Orange"]
        )
        metric_outer = metric_border
        metric_glow = metric_c2 + "66"
    else:
        metric_c1, metric_c2, metric_c3 = "#3B1025", "#7E2342", "#35101F"
        metric_border, metric_accent, metric_shimmer = "#A83E59", "#F0B5C7", "#EAD7DD"
        metric_outer, metric_glow = "#A83E59", "rgba(220,92,126,.30)"

    # Light pearl dashboard; the topmost signal card supplies the accent colour.
    metric_colour_name = selected_rows[0].get("Card Colour", "Copper-Orange") if selected_rows else "Rose-Cherry"
    light_accents = {
        "Green": ("#287A59", "#1F674B", "#E5F1E9", "#DCEAE1"),
        "Blue": ("#2877A8", "#205F88", "#E4EEF6", "#D9E7F1"),
        "Purple": ("#8050A2", "#684083", "#EEE6F4", "#E4D9EE"),
        "Copper-Orange": ("#A6531D", "#874014", "#F7E9DF", "#F1D9C8"),
        "Rose-Cherry": ("#B04B68", "#913B55", "#F3E5EA", "#EBD8DF"),
    }
    metric_ink, metric_accent_text, metric_tint, metric_canvas = light_accents.get(
        metric_colour_name, light_accents["Rose-Cherry"]
    )
    st.session_state["_vardaan_active_theme"] = {
        "deep": metric_c1, "primary": metric_c2, "accent": metric_c3,
        "ink": metric_ink, "accent_text": metric_accent_text,
        "tint": metric_tint, "canvas": metric_canvas,
        "border": metric_ink, "soft_text": metric_accent,
    }
    st.markdown(
        f"""
        <style>
        .stApp {{
            --brand-flower-color:{metric_ink};
            background:radial-gradient(ellipse at 8% 0%,{metric_canvas} 0%,transparent 42%),linear-gradient(135deg,#E6E9ED 0%,#DDE2E7 52%,#E9ECEF 100%) !important;
            color:#263342 !important;
        }}
        [data-testid="stHeader"] {{ background:rgba(231,234,238,.96) !important; }}
        /* Apply the topmost AutoGreen signal-card palette across the entire workspace. */
        .stApp h1 {{ color:{metric_ink} !important; }}
        .vardaan-welcome {{
            border-color:{metric_ink}66 !important;
            background:radial-gradient(circle at 88% 12%,rgba(255,255,255,.30),transparent 27%),
                       linear-gradient(115deg,{metric_c1} 0%,{metric_c2} 48%,{metric_c3} 100%) !important;
            box-shadow:0 14px 34px {metric_ink}22,inset 0 1px 0 rgba(255,255,255,.22) !important;
        }}
        .vardaan-kicker {{ color:{metric_accent} !important; }}
        .vardaan-heading {{ color:#FFFFFF !important; }}
        .vardaan-subheading {{ color:{metric_shimmer} !important; }}
        .vardaan-nav-label {{ color:{metric_ink} !important; }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] > div[role="radiogroup"] {{
            border-color:{metric_ink}55 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{metric_tint} 100%) !important;
            box-shadow:0 5px 18px {metric_ink}14 !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label p {{
            color:{metric_accent_text} !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
            border-color:{metric_ink} !important;
            background:linear-gradient(135deg,{metric_c1} 0%,{metric_c2} 100%) !important;
            box-shadow:0 5px 13px {metric_ink}30 !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) p {{
            color:#FFFFFF !important;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stMetric"] {{
            border-color:{metric_ink}55 !important;
            background:linear-gradient(145deg,#FFFFFF 0%,{metric_tint} 100%) !important;
            box-shadow:0 6px 18px {metric_ink}14 !important;
        }}
        [data-testid="stMainBlockContainer"] .stButton > button,
        [data-testid="stMainBlockContainer"] .stDownloadButton > button {{
            border-color:{metric_ink}88 !important;
            color:{metric_accent_text} !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{metric_tint} 100%) !important;
        }}
        [data-testid="stMainBlockContainer"] .stButton > button:hover,
        [data-testid="stMainBlockContainer"] .stDownloadButton > button:hover {{
            border-color:{metric_ink} !important;
            background:linear-gradient(110deg,{metric_tint} 0%,{metric_canvas} 100%) !important;
        }}
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
        /* Sidebar follows the topmost AutoGreen signal-card palette. */
        [data-testid="stSidebar"] {{
            background:linear-gradient(180deg,{metric_tint} 0%,{metric_canvas} 48%,#FFFFFF 100%) !important;
            border-right:1.5px solid {metric_ink}88 !important;
            box-shadow:inset -8px 0 24px {metric_ink}12 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
        [data-testid="stSidebar"] label,
        [data-testid="stSidebar"] [data-testid="stCaptionContainer"] {{ color:{metric_accent_text} !important; }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="border-bottom"] {{
            border-color:{metric_ink}55 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="border:1px solid"] {{
            border-color:{metric_ink}66 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{metric_tint} 100%) !important;
        }}
        [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div[style*="font-size:1.35rem"] {{
            color:{metric_ink} !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {{
            border:1px solid {metric_ink}55 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{metric_canvas} 100%) !important;
            box-shadow:0 2px 8px {metric_ink}12 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {{
            border-color:{metric_ink} !important;
            background:linear-gradient(110deg,{metric_tint} 0%,#FFFFFF 100%) !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
            border:1.5px solid {metric_ink} !important;
            background:linear-gradient(110deg,{metric_tint} 0%,{metric_canvas} 100%) !important;
            box-shadow:inset 4px 0 0 {metric_ink},0 3px 10px {metric_ink}20 !important;
        }}
        [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label p {{
            color:{metric_accent_text} !important;
        }}
        [data-testid="stSidebar"] .stButton > button {{
            border-color:{metric_ink}77 !important;
            color:{metric_accent_text} !important;
            background:linear-gradient(135deg,#FFFFFF 0%,{metric_tint} 100%) !important;
        }}
        [data-testid="stSidebar"] .stButton > button:hover {{
            border-color:{metric_ink} !important;
            background:linear-gradient(110deg,{metric_tint} 0%,{metric_canvas} 100%) !important;
        }}
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
        @keyframes signalCardSweep {
            0% { transform:translate3d(-145%,-90%,0) skewX(-18deg); opacity:0; }
            16% { opacity:.48; }
            54% { opacity:.24; }
            100% { transform:translate3d(330%,90%,0) skewX(-18deg); opacity:0; }
        }
        .index-signal-card {
            position:relative;
            overflow:hidden;
            isolation:isolate;
        }
        .index-signal-card::after {
            content:"";
            position:absolute;
            top:-35%;
            bottom:-35%;
            left:-48%;
            width:34%;
            z-index:0;
            pointer-events:none;
            background:linear-gradient(105deg,transparent 0%,rgba(255,255,255,0) 25%,
                rgba(255,255,255,.42) 48%,rgba(255,255,255,.16) 58%,transparent 82%);
            filter:blur(2px);
            animation:signalCardSweep 6.2s ease-in-out infinite;
        }
        @media (prefers-reduced-motion:reduce) {
            .index-signal-card::after { animation:none !important; display:none !important; }
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
            is_next_session_stock = bool(item.get("_Next Session Stock"))
            entry = item.get("Live Price")
            # Card target is +0.35% from Entry, capped at the stock-specific
            # maximum calculated from the session Open (+0.75% or +1.00%).
            entry_based_target = entry * 1.0035 if entry is not None else None
            day_open = item.get("Day Open")
            o2h_cap_pct = float(item.get("O2H Target %", 1.0))
            open_based_max_target = (
                day_open * (1 + o2h_cap_pct / 100)
                if day_open is not None and day_open > 0
                else entry_based_target
            )
            target = min(entry_based_target, open_based_max_target) if entry is not None else None
            rank_value = item.get("Sum O2L.10 Rank")
            rank_text = f"{rank_value:g}" if isinstance(rank_value, (int, float)) else "—"
            rank_label = "Stock of the Day- Next Session" if is_next_session_stock else f"RANK {rank_text}"
            entry_display = "Open Price" if is_next_session_stock else (f"₹{entry:,.2f}" if entry is not None else "Unavailable")
            target_display = "Open Price +0.50%" if is_next_session_stock else (f"₹{target:,.2f}" if target is not None else "Unavailable")
            role_text = "BUY CONFIRMATION" if is_confirmation else "BUY SIGNAL"
            palette = {
                "Green": ("#062d20", "#087447", "#063d2b", "#E6C76A", "#CFFFE4", "#8FFFF0"),
                "Blue": ("#071b3d", "#075b91", "#092d57", "#78D9FF", "#D9F4FF", "#BDEBFF"),
                "Purple": ("#29113f", "#713a91", "#35134f", "#E0B5FF", "#F2DFFF", "#EAD7FF"),
                "Copper-Orange": ("#3D1D0B", "#A64B16", "#54250C", "#F2A15A", "#FFF0E2", "#FFD0A6"),
            }
            c1, c2, c3, border_colour, accent_colour, shimmer_colour = palette.get(
                item.get("Card Colour"), palette["Copper-Orange"]
            )
            card_html = f"""
            <div class="index-signal-card" style="
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
                {rank_label}
              </div>
              <div style="height:1px;background:linear-gradient(90deg,transparent,{border_colour},rgba(255,255,255,.42),{border_colour},transparent);
                          margin:19px 0 17px;position:relative;z-index:1;"></div>
              <div style="display:grid;grid-template-columns:repeat(2,minmax(0,1fr));
                          align-items:center;gap:18px;position:relative;z-index:1;">
                <div style="min-width:0;">
                  <div style="font-size:.70rem;font-weight:850;letter-spacing:1.15px;opacity:.84;">ENTRY PRICE</div>
                  <div style="font-size:1.48rem;line-height:1.25;font-weight:950;color:#fff;margin-top:7px;
                              white-space:nowrap;text-shadow:0 2px 10px rgba(0,0,0,.22);">
                    {entry_display}
                  </div>
                </div>
                <div style="min-width:0;text-align:left;">
                  <div style="font-size:.70rem;font-weight:850;letter-spacing:1.15px;opacity:.84;">TARGET PRICE</div>
                  <div style="font-size:1.48rem;line-height:1.25;font-weight:950;color:{accent_colour};margin-top:7px;
                              white-space:nowrap;text-shadow:0 0 12px {c2}88;">
                    {target_display}
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




def clear_inactive_plan_request_session():
    """Keep a plan-request confirmation private to the inactive account flow."""
    for key in (
        "inactive_account_request",
        "inactive_plan_request_sent",
        "inactive_plan_choice",
    ):
        st.session_state.pop(key, None)


def logout():
    try:
        sb.auth.sign_out()
    except Exception:
        pass

    clear_index_session_data()
    clear_inactive_plan_request_session()
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
    clear_inactive_plan_request_session()
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
# INACTIVE ACCOUNT PLAN REQUEST
# ============================================================

PLAN_REQUEST_WEBHOOK = "https://script.google.com/macros/s/AKfycbyQQuarJ89YwLFKu00yJRQIujMZEcPYujIeWLazGN9ej1rG1HWjreddIJ-NXUoxCDU8bg/exec"


@st.dialog("Your Trial Session Has Ended", width="medium")
def show_inactive_plan_dialog():
    pending = st.session_state.get("inactive_account_request") or {}
    if not pending:
        return

    st.markdown(
        """
        <style>
        [data-testid="stDialog"] [role="dialog"] {
            border: 1px solid #D8E8E2;
            border-radius: 24px;
            background: linear-gradient(145deg, #FFFFFF 0%, #F7FCF9 100%);
            box-shadow: 0 24px 70px rgba(15, 60, 45, .18);
        }
        [data-testid="stDialog"] [data-testid="stDialogHeader"] {
            border-bottom: 1px solid #E6F0EB;
            padding-bottom: .75rem;
        }
        [data-testid="stDialog"] [data-testid="stRadio"] > label {
            color: #23483C;
            font-weight: 800;
            font-size: .9rem;
            margin-bottom: .65rem;
        }
        [data-testid="stDialog"] [data-testid="stRadio"] [role="radiogroup"] {
            gap: 10px;
        }
        [data-testid="stDialog"] [data-testid="stRadio"] [role="radiogroup"] > label {
            align-items: center;
            border: 1px solid #DCEAE3;
            border-radius: 15px;
            background: #FFFFFF;
            padding: 13px 15px;
            min-height: 56px;
            box-shadow: 0 3px 10px rgba(20, 70, 50, .045);
            transition: border-color .18s ease, background .18s ease, box-shadow .18s ease;
        }
        [data-testid="stDialog"] [data-testid="stRadio"] [role="radiogroup"] > label:hover {
            border-color: #55A987;
            background: #F4FBF7;
        }
        [data-testid="stDialog"] [data-testid="stRadio"] [role="radiogroup"] > label:has(input:checked) {
            border: 1.6px solid #16845B;
            background: linear-gradient(105deg, #EAF8F0 0%, #F8FDF9 100%);
            box-shadow: inset 4px 0 0 #16845B, 0 5px 15px rgba(22, 132, 91, .10);
        }
        [data-testid="stDialog"] [data-testid="stRadio"] [role="radiogroup"] > label p {
            color: #173F31 !important;
            font-weight: 750;
            line-height: 1.35;
            margin: 0;
        }
        [data-testid="stDialog"] [data-testid="stButton"] button {
            min-height: 48px;
            border: 0;
            border-radius: 13px;
            color: #FFFFFF;
            font-weight: 850;
            letter-spacing: .01em;
            background: linear-gradient(100deg, #075B42 0%, #13865C 55%, #28A879 100%);
            box-shadow: 0 8px 18px rgba(19, 134, 92, .22);
        }
        [data-testid="stDialog"] [data-testid="stButton"] button:hover {
            border: 0;
            background: linear-gradient(100deg, #064B38 0%, #0F754F 55%, #219568 100%);
            box-shadow: 0 10px 22px rgba(19, 134, 92, .28);
        }
        </style>
        <div style="text-align:center;padding:8px 8px 18px;">
          <div style="display:inline-flex;align-items:center;justify-content:center;width:58px;height:58px;
                      border-radius:18px;background:linear-gradient(145deg,#E2F7EB,#F3FBF6);
                      border:1px solid #CBE9D8;font-size:28px;margin-bottom:12px;">⏳</div>
          <div style="font-size:11px;font-weight:900;letter-spacing:2.2px;color:#16845B;margin-bottom:7px;">
            6THSENSE · VARDAAN
          </div>
          <div style="font-size:24px;line-height:1.2;font-weight:900;letter-spacing:-.6px;color:#173F31;margin-bottom:9px;">
            Your Trial Session Has Ended
          </div>
          <div style="max-width:420px;margin:0 auto;font-size:13px;line-height:1.65;color:#61766B;">
            Choose an access plan below to request an extension. We’ll email your request to the administrator for review.
          </div>
        </div>
        <div style="height:1px;background:linear-gradient(90deg,transparent,#DCEAE3,transparent);margin:0 0 15px;"></div>
        """,
        unsafe_allow_html=True,
    )

    plans = {
        "₹250  ·  15 days  ·  Trial extension": "15_days",
        "₹500  ·  1 month": "1_month",
        "₹1,200  ·  3 months": "3_months",
    }

    if st.session_state.get("inactive_plan_request_sent"):
        st.markdown(
            """
            <div style="text-align:center;padding:18px 8px 12px;">
              <div style="font-size:34px;margin-bottom:8px;">✉️</div>
              <div style="font-size:18px;font-weight:850;color:#176B4C;">Request sent successfully</div>
              <div style="font-size:13px;line-height:1.6;color:#64766D;margin-top:7px;">
                Your plan request has been emailed to the administrator. Your account will remain inactive until approval.
              </div>
            </div>
            """,
            unsafe_allow_html=True,
        )
        if st.button("Close", use_container_width=True):
            st.session_state.pop("inactive_account_request", None)
            st.session_state.pop("inactive_plan_request_sent", None)
            st.rerun()
        return

    selected_label = st.radio(
        "SELECT YOUR ACCESS PLAN",
        list(plans.keys()),
        key="inactive_plan_choice",
    )

    st.markdown(
        """
        <div style="display:flex;align-items:center;gap:8px;padding:10px 12px;margin:8px 0 14px;
                    border-radius:12px;background:#F1F8F4;border:1px solid #DFEEE5;
                    color:#526B5E;font-size:12px;line-height:1.5;">
          <span style="font-size:16px;">🔒</span>
          <span>Your account stays inactive until the administrator reviews and approves your request.</span>
        </div>
        """,
        unsafe_allow_html=True,
    )

    if st.button("Submit Extension Request  →", type="primary", use_container_width=True):
        request_token = str(st.secrets.get("PLAN_REQUEST_TOKEN", "") or "").strip()
        if not request_token:
            st.error(
                "Plan request email is not configured yet. "
                "Please contact the administrator."
            )
            return

        payload = {
            "action": "plan_request",
            "token": request_token,
            "plan": plans[selected_label],
            "account_email": str(pending.get("email", "")).strip(),
            "account_id": str(pending.get("user_id", "")).strip(),
        }

        try:
            with st.spinner("Sending your request..."):
                response = requests.post(
                    PLAN_REQUEST_WEBHOOK,
                    json=payload,
                    timeout=25,
                )
                response.raise_for_status()
                result = response.json()

            if result.get("success"):
                st.session_state["inactive_plan_request_sent"] = True
                st.rerun()
            else:
                st.error(
                    result.get("message", "Could not send the request. Please try again.")
                )
        except Exception:
            st.error(
                "We could not send your request right now. Please try again later "
                "or contact the administrator."
            )


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

    # Login palette changes automatically according to Indian Standard Time.
    st_autorefresh(interval=30_000, key="vardaan_login_theme_refresh")
    login_clock = datetime.now(ZoneInfo("Asia/Kolkata"))
    login_minute = login_clock.hour * 60 + login_clock.minute

    if 9 * 60 <= login_minute <= 11 * 60:
        login_theme = {
            "name": "Emerald Sunrise",
            "page_a": "#ECFDF5", "page_b": "#F7FEE7",
            "deep": "#064E3B", "primary": "#10B981",
            "accent": "#A3E635", "highlight": "#FACC15",
            "heading": "#064E3B", "heading_end": "#FACC15",
            "card": "#F4FFF8", "border": "#86D6AE",
            "focus": "#10B981", "shadow": "16, 110, 75",
        }
    elif 11 * 60 + 1 <= login_minute <= 13 * 60:
        login_theme = {
            "name": "Ocean Spectrum",
            "page_a": "#E0F2FE", "page_b": "#ECFEFF",
            "deep": "#082F49", "primary": "#0284C7",
            "accent": "#22D3EE", "highlight": "#818CF8",
            "heading": "#082F49", "heading_end": "#818CF8",
            "card": "#F2FBFF", "border": "#7DD3FC",
            "focus": "#0284C7", "shadow": "2, 132, 199",
        }
    elif 13 * 60 + 1 <= login_minute <= 15 * 60 + 30:
        login_theme = {
            "name": "Royal Fusion",
            "page_a": "#F3E8FF", "page_b": "#FCE7F3",
            "deep": "#3B0764", "primary": "#9333EA",
            "accent": "#EC4899", "highlight": "#F59E0B",
            "heading": "#581C87", "heading_end": "#EC4899",
            "card": "#FFF7FD", "border": "#D8B4FE",
            "focus": "#9333EA", "shadow": "147, 51, 234",
        }
    else:
        login_theme = {
            "name": "Sunset Luxe",
            "page_a": "#FFF7ED", "page_b": "#FCE7F3",
            "deep": "#7C2D12", "primary": "#EA580C",
            "accent": "#DC4568", "highlight": "#D6B36A",
            "heading": "#7C2D12", "heading_end": "#DC4568",
            "card": "#FFFAF5", "border": "#FDBA74",
            "focus": "#EA580C", "shadow": "234, 88, 12",
        }

    st.markdown(
        f"""
        <style>
        [data-testid="stAppViewContainer"] {{
            background:
                radial-gradient(ellipse at 8% 8%, {login_theme["accent"]}30, transparent 36%),
                radial-gradient(ellipse at 92% 18%, {login_theme["highlight"]}35, transparent 34%),
                linear-gradient(145deg, {login_theme["page_a"]} 0%, #FFFFFF 48%, {login_theme["page_b"]} 100%);
        }}
        .stApp {{
            --brand-flower-color: {login_theme["primary"]};
        }}
        [data-testid="stMainBlockContainer"] {{
            max-width: 680px;
            padding-top: 4.5rem;
        }}
        [data-testid="stMainBlockContainer"] h1 {{
            position: relative;
            text-align: center;
            font-weight: 900;
            letter-spacing: -0.045em;
            background: linear-gradient(100deg, {login_theme["deep"]} 0%, {login_theme["primary"]} 38%, {login_theme["accent"]} 70%, {login_theme["highlight"]} 100%);
            -webkit-background-clip: text;
            background-clip: text;
            color: transparent;
        }}
        [data-testid="stMainBlockContainer"] h1::before {{
            content: "";
            position: absolute;
            left: 50%;
            top: 0;
            width: 76px;
            height: 52px;
            transform: translate(-50%, calc(-100% - 4px));
            background: linear-gradient(110deg, var(--brand-flower-color) 0%, var(--brand-flower-color) 30%, rgba(255,255,255,.96) 48%, var(--brand-flower-color) 66%, var(--brand-flower-color) 100%);
            background-size: 300% 300%;
            animation: dependentSignalShimmer 4.2s linear infinite;
            -webkit-mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
            mask: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 120 72' fill='none' stroke='white' stroke-width='3' stroke-linecap='round' stroke-linejoin='round'%3E%3Cpath d='M60 58 C43 43 43 24 60 8 C77 24 77 43 60 58Z'/%3E%3Cpath d='M55 58 C34 51 22 35 27 20 C43 26 54 40 55 58Z'/%3E%3Cpath d='M65 58 C86 51 98 35 93 20 C77 26 66 40 65 58Z'/%3E%3Cpath d='M49 59 C30 60 16 52 10 40 C28 38 43 45 49 59Z'/%3E%3Cpath d='M71 59 C90 60 104 52 110 40 C92 38 77 45 71 59Z'/%3E%3Cpath d='M34 64 Q60 69 86 64'/%3E%3C/svg%3E") center / contain no-repeat;
        }}
        [data-testid="stMainBlockContainer"] h2 {{
            color: {login_theme["heading"]};
            text-align: center;
            font-weight: 800;
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stForm"] {{
            background: linear-gradient(145deg, rgba(255,255,255,.99), {login_theme["card"]} 58%, {login_theme["page_b"]} 100%);
            border: 1px solid {login_theme["border"]};
            border-radius: 20px;
            padding: 1.35rem 1.35rem 1.1rem;
            box-shadow: 0 14px 38px rgba({login_theme["shadow"]},.12), 0 2px 8px {login_theme["highlight"]}20;
        }}
        [data-testid="stMainBlockContainer"] input {{
            border-radius: 11px;
            border-color: {login_theme["border"]};
            background: #ffffff;
        }}
        [data-testid="stMainBlockContainer"] input:focus {{
            border-color: {login_theme["focus"]};
            box-shadow: 0 0 0 1px {login_theme["focus"]};
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button,
        [data-testid="stMainBlockContainer"] [data-testid="stButton"] button {{
            color: #ffffff;
            font-weight: 800;
            border: 0;
            border-radius: 12px;
            background: linear-gradient(105deg, {login_theme["deep"]} 0%, {login_theme["primary"]} 42%, {login_theme["accent"]} 76%, {login_theme["highlight"]} 100%);
            box-shadow: 0 5px 14px rgba({login_theme["shadow"]},.22);
        }}
        [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button:hover,
        [data-testid="stMainBlockContainer"] [data-testid="stButton"] button:hover {{
            filter: brightness(1.06);
            box-shadow: 0 7px 18px rgba({login_theme["shadow"]},.29);
        }}
        
        @keyframes dependentSignalShimmer {{
            0% {{ background-position:0% 50%; filter:brightness(1); }}
            50% {{ background-position:100% 50%; filter:brightness(1.12); }}
            100% {{ background-position:0% 50%; filter:brightness(1); }}
        }}
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.title(
        "6thSense-Vardaan"
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
                st.session_state["inactive_account_request"] = {
                    "email": str(getattr(user, "email", "") or email).strip(),
                    "user_id": str(user.id),
                }
                st.session_state.pop("inactive_plan_request_sent", None)
                try:
                    sb.auth.sign_out()
                except Exception:
                    pass
                st.rerun()

            else:

                # Never carry a previous account's plan-request confirmation
                # into a newly authenticated admin or user session.
                clear_inactive_plan_request_session()
                st.session_state.user = user
                st.session_state.profile = prof
                st.session_state.is_admin = (
                    prof["role"] == "admin"
                )
                # Start each fresh login on AutoGreen, with both navigation
                # controls visibly selecting the same dashboard.
                st.session_state["main_dashboard_default_index_v2"] = "AutoGreen"
                st.session_state["sidebar_dashboard_nav_v1"] = "AutoGreen"
                st.session_state["bottom_dashboard_nav_v1"] = "AutoGreen"
                st.session_state.session_login_started_at = time.time()
                st.session_state.pop("session_expired_notice", None)

                st.rerun()

        except Exception:
            st.error(
                "Invalid email or password."
            )

    if st.session_state.get("inactive_account_request"):
        show_inactive_plan_dialog()

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

/* Professional sidebar navigation */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #e8eef6 0%, #f7f9fc 100%);
}
/* Compact sidebar width while preserving its existing theme and navigation. */
@media (min-width: 768px) {
    [data-testid="stSidebar"] {
        width: 260px !important;
        min-width: 260px !important;
        max-width: 260px !important;
        flex: 0 0 260px !important;
    }
}
@media (max-width: 767px) {
    [data-testid="stSidebar"] {
        width: min(82vw, 260px) !important;
        min-width: min(82vw, 260px) !important;
        max-width: min(82vw, 260px) !important;
    }
}
[data-testid="stSidebar"] [data-testid="stMarkdownContainer"] h2 {
    color: #173e68;
    border: 0;
    padding: 0;
    margin-bottom: 0.15rem;
}
[data-testid="stSidebar"] [data-testid="stRadio"] > label {
    color: #60758e;
    font-size: 0.72rem;
    font-weight: 800;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    margin: 0.75rem 0 0.55rem;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] {
    gap: 0.45rem;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {
    display: flex;
    align-items: center;
    min-height: 3rem;
    padding: 0.55rem 0.8rem;
    margin: 0;
    border: 1px solid #d7e1ed;
    border-radius: 12px;
    background: rgba(255,255,255,0.78);
    box-shadow: 0 2px 7px rgba(30, 61, 95, 0.035);
    transition: all 150ms ease;
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {
    border-color: #8eabc9;
    background: #ffffff;
    transform: translateX(2px);
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {
    border: 1px solid #376b9e;
    background: linear-gradient(110deg, #dce9f7 0%, #f5f8fc 100%);
    box-shadow: inset 3px 0 0 #285d91, 0 3px 10px rgba(30, 61, 95, 0.07);
}
[data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label p {
    color: #294663;
    font-weight: 650;
    font-size: 0.94rem;
}
[data-testid="stSidebar"] .stButton > button {
    min-height: 2.75rem;
    border-radius: 11px;
    border: 1px solid #c6d3e2;
    font-weight: 700;
    letter-spacing: 0.01em;
}
[data-testid="stSidebar"] .stButton > button:hover {
    border-color: #4779ad;
    box-shadow: 0 4px 12px rgba(35, 77, 120, 0.10);
}
</style>
""", unsafe_allow_html=True)




# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown(
        """
        <div style="
            padding: 0.55rem 0 0.9rem;
            border-bottom: 1px solid #d5dfeb;
            margin-bottom: 0.9rem;
        ">
            <div style="color:#71849a;font-size:0.68rem;font-weight:800;
                        letter-spacing:0.18em;">6THSENSE</div>
            <div style="color:#173e68;font-size:1.35rem;font-weight:850;
                        letter-spacing:-0.04em;line-height:1.25;">VARDAAN</div>
            <div style="color:#71849a;font-size:0.76rem;margin-top:0.2rem;">
                Your market workspace
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        f"""
        <div style="
            padding:0.75rem 0.8rem;
            border:1px solid #d7e1ed;
            border-radius:12px;
            background:rgba(255,255,255,0.78);
            margin-bottom:0.65rem;
        ">
            <div style="font-size:0.68rem;color:#71849a;font-weight:800;
                        letter-spacing:0.09em;text-transform:uppercase;">
                Signed in as
            </div>
            <div style="font-size:0.86rem;color:#294663;font-weight:650;
                        overflow-wrap:anywhere;margin-top:0.18rem;">
                {html.escape(str(user.email or ""))}
            </div>
            <div style="font-size:0.75rem;color:#60758e;margin-top:0.35rem;">
                {'Administrator' if is_admin else 'User'}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    # Default every signed-in account (including admins) to AutoGreen once
    # per browser session. Versioned marker also corrects sessions created by
    # the earlier default logic, and both visible selectors start in sync.
    dashboard_default_initialized_key = f"_dashboard_default_initialized_v3_{user.id}"
    if not st.session_state.get(dashboard_default_initialized_key):
        st.session_state["main_dashboard_default_index_v2"] = "AutoGreen"
        st.session_state["sidebar_dashboard_nav_v1"] = "AutoGreen"
        st.session_state["bottom_dashboard_nav_v1"] = "AutoGreen"
        st.session_state[dashboard_default_initialized_key] = True

    # Recover safely if a selector key is missing without changing an existing choice.
    if "sidebar_dashboard_nav_v1" not in st.session_state:
        st.session_state["sidebar_dashboard_nav_v1"] = st.session_state.get(
            "main_dashboard_default_index_v2", "AutoGreen"
        )

    if "bottom_dashboard_nav_v1" not in st.session_state:
        st.session_state["bottom_dashboard_nav_v1"] = st.session_state.get(
            "main_dashboard_default_index_v2", "AutoGreen"
        )

    def sync_sidebar_dashboard_to_main():
        selected_dashboard = st.session_state["sidebar_dashboard_nav_v1"]
        st.session_state["main_dashboard_default_index_v2"] = selected_dashboard
        st.session_state["bottom_dashboard_nav_v1"] = selected_dashboard

    def sync_bottom_dashboard_to_main():
        selected_dashboard = st.session_state["bottom_dashboard_nav_v1"]
        st.session_state["main_dashboard_default_index_v2"] = selected_dashboard
        st.session_state["sidebar_dashboard_nav_v1"] = selected_dashboard

    st.radio(
        "Sidebar dashboard navigation",
        ["AutoGreen", "Stock of the Day", "Index Toppers", "Master Excel"],
        format_func=lambda item: {
            "AutoGreen": "🌿  AutoGreen",
            "Stock of the Day": "⭐  Stock of the Day",
            "Index Toppers": "🏆  Index Toppers",
            "Master Excel": "📊  Master Excel",
        }[item],
        key="sidebar_dashboard_nav_v1",
        on_change=sync_sidebar_dashboard_to_main,
        label_visibility="collapsed",
    )

    st.divider()
    if st.button(
        "↪  Log out",
        use_container_width=True,
        type="secondary",
        key="sidebar_logout_button"
    ):
        logout()

page = st.session_state.get("main_dashboard_default_index_v2", "AutoGreen")

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
    # Refresh often during market hours and every minute outside market
    # hours so the 5 PM / 9 AM Stock of the Day mode switch happens automatically.
    st_autorefresh(
        interval=15_000 if market_open else 60_000,
        limit=None,
        key="stock_of_the_day_live_market_refresh",
    )
    render_stock_of_the_day()


elif page == "Index Toppers":
    render_index_toppers_dashboard()

elif page == "Master Excel":

# MASTER FILE
    # ============================================================

    st.markdown("""
    <style>
    .masterx-hero {
        position:relative; overflow:hidden; padding:25px 26px; margin:4px 0 18px;
        border-radius:20px; color:#173B50;
        background:linear-gradient(105deg,#EAF7FF 0%,#DDF2FC 52%,#EFF9FF 100%);
        box-shadow:0 10px 26px rgba(56,130,165,.10);
        border:1px solid #B9DDED;
    }
    .masterx-hero:after {
        content:""; position:absolute; width:190px; height:190px; right:-45px; top:-95px;
        border:24px solid rgba(22,143,192,.10); border-radius:50%;
    }
    .masterx-kicker {font-size:10px;letter-spacing:2.1px;font-weight:850;color:#168FC0;text-transform:uppercase;}
    .masterx-title {font-size:clamp(25px,3vw,34px);font-weight:900;line-height:1.12;margin:7px 0 6px;color:#123B5D;}
    .masterx-subtitle {font-size:13px;line-height:1.5;color:#557084;max-width:680px;}
    .masterx-status {
        display:flex;align-items:center;gap:12px;padding:14px 16px;margin:0 0 18px;
        border-radius:15px;background:linear-gradient(105deg,#EAF7FF,#DDF2FC 52%,#EFF9FF);
        border:1px solid #B9DDED;color:#173B50;
    }
    .masterx-status-dot {width:11px;height:11px;flex:0 0 11px;border-radius:50%;background:#38BDF8;box-shadow:0 0 0 5px rgba(56,189,248,.16);}
    .masterx-status-dot.offline {background:#38BDF8;box-shadow:0 0 0 5px rgba(56,189,248,.16);}
    .masterx-status-title {font-size:14px;font-weight:850;line-height:1.3;}
    .masterx-status-note {font-size:11px;color:#557084;margin-top:3px;}
    .masterx-section {
        display:flex;align-items:center;gap:11px;padding:12px 15px;margin:18px 0 12px;
        border-radius:13px;background:linear-gradient(100deg,#E7F6FF,#DCEFFD 48%,#F0FAFF);
        border:1px solid #C6E2F1;color:#173B50;font-weight:850;font-size:14px;
    }
    .masterx-section-icon {
        display:flex;align-items:center;justify-content:center;width:30px;height:30px;
        border-radius:10px;background:linear-gradient(135deg,#FFFFFF,#DDF2FC,#B9E8F5);
        color:#155E91;font-size:15px;
    }
    div[data-testid="stDownloadButton"] button {
        border:1px solid #B9DDED!important;color:#155E91!important;font-weight:800!important;
        background:linear-gradient(100deg,#FFFFFF,#EAF7FF 55%,#DDF2FC)!important;
        box-shadow:0 4px 12px rgba(56,130,165,.10)!important;border-radius:12px!important;
    }
    div[data-testid="stDownloadButton"] button:hover {background:#DDF2FC!important;border-color:#168FC0!important;}

    /* Master Excel page: keep the complete sidebar in the fixed ice-blue theme. */
    .stApp { --signal-nav-border:#B9DDED; --signal-nav-accent:#168FC0; --signal-nav-soft:#DDF2FC; --signal-nav-canvas:#EAF7FF; --signal-nav-text:#173B50; }
    [data-testid="stSidebar"] {
        background:linear-gradient(180deg,#DDF2FC 0%,#EAF7FF 42%,#F7FCFF 100%) !important;
        border-right:1.5px solid #9CCFE5 !important;
        box-shadow:inset -8px 0 24px rgba(23,101,139,.08) !important;
    }
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] p,
    [data-testid="stSidebar"] [data-testid="stMarkdownContainer"] div,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] [data-testid="stCaptionContainer"],
    [data-testid="stSidebar"] [data-testid="stText"],
    [data-testid="stSidebar"] [data-testid="stAlert"] {
        color:#173B50 !important;
    }
    [data-testid="stSidebar"] div[style*="font-size:1.35rem"] {
        color:#155E91 !important;
    }
    [data-testid="stSidebar"] div[style*="border-bottom"] {
        border-color:#A9D7EA !important;
    }
    [data-testid="stSidebar"] div[style*="border:1px solid"] {
        border-color:#B9DDED !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#EAF7FF 100%) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label {
        border:1px solid #B9DDED !important;
        border-radius:11px !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#EAF7FF 100%) !important;
        box-shadow:0 2px 8px rgba(23,101,139,.08) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:hover {
        border-color:#38BDF8 !important;
        background:linear-gradient(110deg,#DDF2FC 0%,#FFFFFF 100%) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {
        border:1.5px solid #168FC0 !important;
        background:linear-gradient(110deg,#DDF2FC 0%,#FFFFFF 100%) !important;
        box-shadow:inset 4px 0 0 #168FC0,0 3px 10px rgba(23,101,139,.12) !important;
    }
    [data-testid="stSidebar"] [data-testid="stRadio"] [role="radiogroup"] label p {
        color:#173B50 !important;
        font-weight:750 !important;
    }
    [data-testid="stSidebar"] .stButton > button,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button {
        border:1px solid #A9D7EA !important;
        border-radius:10px !important;
        color:#155E91 !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#DDF2FC 100%) !important;
        box-shadow:0 3px 10px rgba(23,101,139,.08) !important;
    }
    [data-testid="stSidebar"] .stButton > button:hover,
    [data-testid="stSidebar"] [data-testid="stFormSubmitButton"] button:hover {
        border-color:#168FC0 !important;
        background:linear-gradient(110deg,#DDF2FC 0%,#FFFFFF 100%) !important;
        color:#123B5D !important;
    }
    [data-testid="stSidebar"] input,
    [data-testid="stSidebar"] textarea {
        border-color:#A9D7EA !important;
        background:#FFFFFF !important;
        color:#173B50 !important;
    }
    [data-testid="stSidebar"] input:focus,
    [data-testid="stSidebar"] textarea:focus {
        border-color:#168FC0 !important;
        box-shadow:0 0 0 2px rgba(56,189,248,.18) !important;
    }
    
    /* Keep every Master Excel control within the same ice-blue and white palette. */
    [data-testid="stMainBlockContainer"] .stButton > button,
    [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button {
        border:1px solid #B9DDED !important;
        border-radius:11px !important;
        color:#155E91 !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#EAF7FF 58%,#DDF2FC 100%) !important;
        box-shadow:0 3px 10px rgba(56,130,165,.09) !important;
        font-weight:800 !important;
    }
    [data-testid="stMainBlockContainer"] .stButton > button:hover,
    [data-testid="stMainBlockContainer"] [data-testid="stFormSubmitButton"] button:hover {
        border-color:#168FC0 !important;
        background:linear-gradient(110deg,#DDF2FC 0%,#FFFFFF 100%) !important;
        color:#123B5D !important;
    }
    [data-testid="stFileUploader"] section {
        background:#FFFFFF !important;
        border:1px dashed #A9D7EA !important;
        border-radius:12px !important;
        color:#173B50 !important;
    }
    [data-testid="stFileUploader"] section *,
    [data-testid="stFileUploader"] button {
        color:#155E91 !important;
    }
    [data-testid="stFileUploader"] button {
        border:1px solid #A9D7EA !important;
        background:linear-gradient(135deg,#FFFFFF 0%,#DDF2FC 100%) !important;
    }
    [data-testid="stMainBlockContainer"] input,
    [data-testid="stMainBlockContainer"] textarea,
    [data-testid="stMainBlockContainer"] [data-baseweb="select"] > div {
        border-color:#B9DDED !important;
        background-color:#FFFFFF !important;
        color:#173B50 !important;
    }
    </style>
    """, unsafe_allow_html=True)


    master = get_master()

    st.markdown("""
    <div class="masterx-hero">
      <div class="masterx-kicker">6THSENSE • DATA CONTROL CENTRE</div>
      <div class="masterx-title">Master Excel</div>
      <div class="masterx-subtitle">Your central workbook, organised in one place. View, download and manage the master data securely.</div>
    </div>
    """, unsafe_allow_html=True)

    if master:

        st.markdown("""
        <div class="masterx-status">
          <span class="masterx-status-dot"></span>
          <div><div class="masterx-status-title">Master workbook is available</div>
          <div class="masterx-status-note">The latest stored Excel file is ready to download.</div></div>
        </div>
        """, unsafe_allow_html=True)

        st.download_button(
            "Download Master Excel",
            data=master,
            file_name=MASTER_FILE,
            mime=EXCEL_MIME,
            use_container_width=True
        )

    else:

        st.markdown("""
        <div class="masterx-status">
          <span class="masterx-status-dot offline"></span>
          <div><div class="masterx-status-title">Master workbook not uploaded</div>
          <div class="masterx-status-note">An administrator can upload the .xlsx workbook below.</div></div>
        </div>
        """, unsafe_allow_html=True)


    # ============================================================
    # ADMIN PANEL
    # ============================================================

    if is_admin:

        st.divider()

        st.markdown("""
        <div class="masterx-section"><span class="masterx-section-icon">⚙</span> ADMIN CONTROL PANEL</div>
        """, unsafe_allow_html=True)

        # ========================================================
        # UPLOAD / REPLACE
        # ========================================================

        st.markdown("""
        <div class="masterx-section"><span class="masterx-section-icon">⇧</span> UPLOAD / REPLACE MASTER FILE</div>
        """, unsafe_allow_html=True)

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

            st.markdown("""
            <div class="masterx-section"><span class="masterx-section-icon">✎</span> MODIFY MASTER EXCEL</div>
            """, unsafe_allow_html=True)

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

        st.markdown("""
        <div class="masterx-section"><span class="masterx-section-icon">⌫</span> DELETE MASTER FILE</div>
        """, unsafe_allow_html=True)

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
        # ========================================================
        # USER MANAGEMENT — ICE BLUE ADMIN CENTRE
        # ========================================================

        st.markdown("""
        <style>
        .userx-hero {
            padding:18px 20px; margin:8px 0 16px; border-radius:17px;
            color:#173B50; border:1px solid #B9DDED;
            background:linear-gradient(112deg,#EAF7FF 0%,#DDF2FC 48%,#F7FCFF 100%);
            box-shadow:0 7px 20px rgba(23,101,139,.08);
        }
        .userx-kicker {font-size:10px;letter-spacing:1.8px;font-weight:850;color:#168FC0;text-transform:uppercase;}
        .userx-title {font-size:24px;font-weight:900;line-height:1.2;margin:5px 0;color:#123B5D;}
        .userx-note {font-size:12px;color:#557084;line-height:1.5;}
        .userx-stat {
            padding:13px 14px;border-radius:14px;border:1px solid #C6E2F1;
            background:linear-gradient(135deg,#FFFFFF 0%,#EAF7FF 100%);
        }
        .userx-stat-label {font-size:10px;letter-spacing:1px;text-transform:uppercase;font-weight:800;color:#557084;}
        .userx-stat-value {font-size:25px;font-weight:900;color:#155E91;line-height:1.2;margin-top:4px;}
        .userx-section {
            display:flex;align-items:center;gap:11px;padding:12px 15px;margin:10px 0 12px;
            border-radius:13px;background:linear-gradient(100deg,#E7F6FF,#DCEFFD 48%,#F0FAFF);
            border:1px solid #C6E2F1;color:#173B50;font-weight:850;font-size:14px;
        }
        .userx-stat {
            border:1px solid #C6E2F1 !important;border-radius:14px !important;
            background:linear-gradient(135deg,#FFFFFF 0%,#EAF7FF 100%) !important;
            box-shadow:0 5px 15px rgba(23,101,139,.07) !important;
        }
        div[data-testid="stTabs"] [role="tablist"] {
            gap:8px !important;border-bottom:1px solid #B9DDED !important;
        }
        div[data-testid="stTabs"] button[role="tab"] {
            color:#155E91 !important;font-weight:800 !important;
            border-radius:10px 10px 0 0 !important;
        }
        div[data-testid="stTabs"] button[role="tab"][aria-selected="true"] {
            background:linear-gradient(110deg,#DDF2FC 0%,#FFFFFF 100%) !important;
            border:1px solid #B9DDED !important;border-bottom:2px solid #168FC0 !important;
        }
        div[data-testid="stExpander"] {
            border:1px solid #C6E2F1 !important;border-radius:13px !important;
            background:linear-gradient(135deg,#FFFFFF 0%,#F4FBFF 100%) !important;
            box-shadow:0 4px 13px rgba(23,101,139,.06) !important;
        }
        div[data-testid="stExpander"] summary {
            color:#173B50 !important;font-weight:800 !important;
        }
        div[data-testid="stExpander"] [data-testid="stExpanderDetails"] {
            background:#F7FCFF !important;border-top:1px solid #C6E2F1 !important;
        }
        div[data-testid="stForm"] {
            border:1px solid #C6E2F1 !important;border-radius:14px !important;
            padding:14px !important;background:#F7FCFF !important;
        }
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] > div,
        div[data-testid="stCheckbox"] label,
        div[data-testid="stFileUploader"] section {
            border-color:#A9D7EA !important;
            color:#173B50 !important;
        }
        div[data-testid="stTextInput"] input,
        div[data-testid="stSelectbox"] > div {
            background:#FFFFFF !important;
        }
        div[data-testid="stButton"] button,
        div[data-testid="stFormSubmitButton"] button {
            border:1px solid #A9D7EA !important;border-radius:10px !important;
            color:#155E91 !important;font-weight:800 !important;
            background:linear-gradient(135deg,#FFFFFF 0%,#DDF2FC 100%) !important;
            box-shadow:0 3px 10px rgba(23,101,139,.08) !important;
        }
        div[data-testid="stButton"] button:hover,
        div[data-testid="stFormSubmitButton"] button:hover {
            border-color:#168FC0 !important;background:linear-gradient(110deg,#DDF2FC,#FFFFFF) !important;
            color:#123B5D !important;
        }
        </style>
        """, unsafe_allow_html=True)

        st.markdown("""
        <div class="userx-hero">
          <div class="userx-kicker">6THSENSE • ADMINISTRATION</div>
          <div class="userx-title">User Management</div>
          <div class="userx-note">Manage accounts, access roles and account status from one place. These controls are available only to administrators.</div>
        </div>
        """, unsafe_allow_html=True)

        try:
            users = admin.table("profiles").select("*").order("email").execute().data or []
            total_users = len(users)
            active_users = sum(1 for item in users if bool(item.get("is_active", True)))
            admin_users = sum(1 for item in users if str(item.get("role", "user")).lower() == "admin")

            metric_cols = st.columns(3)
            metric_values = [
                ("TOTAL USERS", total_users),
                ("ACTIVE USERS", active_users),
                ("ADMIN ACCOUNTS", admin_users),
            ]
            for metric_col, (metric_label, metric_value) in zip(metric_cols, metric_values):
                with metric_col:
                    st.markdown(
                        '<div class="userx-stat"><div class="userx-stat-label">'
                        + str(metric_label)
                        + '</div><div class="userx-stat-value">'
                        + str(metric_value)
                        + '</div></div>',
                        unsafe_allow_html=True,
                    )

            st.write("")
            manage_tab, add_tab = st.tabs(["Manage Users", "＋ Add User"])

            with manage_tab:
                st.markdown('<div class="userx-section">ACCOUNT DIRECTORY</div>', unsafe_allow_html=True)
                filter_cols = st.columns([2, 1])
                search_email = filter_cols[0].text_input(
                    "Search by email",
                    placeholder="Type an email address…",
                    key="vardaan_user_search",
                ).strip().lower()
                status_filter = filter_cols[1].selectbox(
                    "Show",
                    ["All users", "Active", "Disabled", "Admins", "Users"],
                    key="vardaan_user_status_filter",
                )

                visible_users = []
                for item in users:
                    item_email = str(item.get("email", "") or "")
                    item_role = str(item.get("role", "user") or "user").lower()
                    item_active = bool(item.get("is_active", True))
                    if search_email and search_email not in item_email.lower():
                        continue
                    if status_filter == "Active" and not item_active:
                        continue
                    if status_filter == "Disabled" and item_active:
                        continue
                    if status_filter == "Admins" and item_role != "admin":
                        continue
                    if status_filter == "Users" and item_role != "user":
                        continue
                    visible_users.append(item)

                st.caption(f"Showing {len(visible_users)} of {total_users} accounts")

                if not visible_users:
                    st.info("No accounts match this search or filter.")
                else:
                    for item in visible_users:
                        uid = item["id"]
                        email = str(item.get("email", "") or "Email unavailable")
                        current_role = str(item.get("role", "user") or "user").lower()
                        current_active = bool(item.get("is_active", True))
                        status_label = "Active" if current_active else "Disabled"
                        role_label = "Admin" if current_role == "admin" else "User"
                        is_current_admin = str(uid) == str(user.id)

                        with st.expander(
                            f"{'🩵' if current_active else '⚪'}  {email}  ·  {role_label}  ·  {status_label}",
                            expanded=False,
                        ):
                            detail_cols = st.columns([1.25, 1])
                            detail_cols[0].markdown("**Email address**")
                            detail_cols[0].write(email)
                            detail_cols[1].markdown("**Current access**")
                            detail_cols[1].write(f"{role_label} · {status_label}")

                            if is_current_admin:
                                st.info("This is your signed-in administrator account. Your own role and access cannot be changed here.")

                            current_login_email = str(email or "").strip().lower()
                            credential_success_key = f"vardaan_credentials_success_{uid}"
                            if st.session_state.get(credential_success_key):
                                st.success(st.session_state.pop(credential_success_key))
                            access_success_key = f"vardaan_access_success_{uid}"
                            if st.session_state.get(access_success_key):
                                st.success(st.session_state.pop(access_success_key))

                            with st.form(f"vardaan_credentials_form_{uid}", clear_on_submit=True):
                                st.markdown("**Change login username and password**")
                                st.info(f"Selected account — current login username: {current_login_email or 'Unavailable'}")
                                st.caption(
                                    "To change only the password, leave the username field empty. "
                                    "To change the username too, enter the new email / User ID."
                                )
                                replacement_email = st.text_input(
                                    "New username (leave empty to keep current username)",
                                    value="",
                                    key=f"vardaan_new_login_email_{uid}",
                                ).strip().lower()
                                replacement_password = st.text_input(
                                    "New password (leave blank to keep current password)",
                                    type="password",
                                    key=f"vardaan_new_login_password_{uid}",
                                )
                                credentials_submitted = st.form_submit_button(
                                    "Update username / password",
                                    use_container_width=True,
                                )

                            if credentials_submitted:
                                if replacement_email and not re.fullmatch(
                                    r"[^\s@]+@[^\s@]+\.[^\s@]+",
                                    replacement_email,
                                ):
                                    st.error("Enter a valid new username email address, or leave it empty to keep the current username.")
                                elif replacement_password and len(replacement_password) < 6:
                                    st.error("New password must be at least 6 characters.")
                                elif not replacement_email and not replacement_password:
                                    st.info("Enter a new password or a new username.")
                                else:
                                    try:
                                        auth_updates = {}
                                        if replacement_email:
                                            auth_updates["email"] = replacement_email
                                            auth_updates["email_confirm"] = True
                                        if replacement_password:
                                            auth_updates["password"] = replacement_password

                                        admin.auth.admin.update_user_by_id(uid, auth_updates)

                                        if replacement_email:
                                            admin.table("profiles").update({
                                                "email": replacement_email,
                                            }).eq("id", uid).execute()

                                        st.session_state[credential_success_key] = (
                                            ("Login username updated" if replacement_email else "")
                                            + (" and " if replacement_email and replacement_password else "")
                                            + ("password changed successfully." if replacement_password else "updated successfully.")
                                        )
                                        st.rerun()
                                    except Exception as credential_error:
                                        st.error(f"Could not update login details: {credential_error}")

                            edit_cols = st.columns(2)
                            selected_role = edit_cols[0].selectbox(
                                "User role",
                                ["user", "admin"],
                                index=0 if current_role == "user" else 1,
                                key=f"vardaan_edit_role_{uid}",
                                disabled=is_current_admin,
                                format_func=lambda value: "User" if value == "user" else "Admin",
                            )
                            selected_active = edit_cols[1].checkbox(
                                "Account active",
                                value=current_active,
                                key=f"vardaan_edit_active_{uid}",
                                disabled=is_current_admin,
                            )

                            action_cols = st.columns([1, 1])
                            if action_cols[0].button(
                                "Save changes",
                                key=f"vardaan_save_user_{uid}",
                                use_container_width=True,
                                disabled=is_current_admin or (
                                    selected_role == current_role and selected_active == current_active
                                ),
                            ):
                                admin.table("profiles").update({
                                    "role": selected_role,
                                    "is_active": selected_active,
                                }).eq("id", uid).execute()
                                st.session_state[access_success_key] = f"Changes saved successfully for {email}."
                                st.rerun()

                            if is_current_admin:
                                st.caption("Your own account is protected from deletion.")
                            else:
                                confirm_delete = st.checkbox(
                                    "I confirm that I want to permanently delete this account",
                                    key=f"vardaan_confirm_delete_{uid}",
                                )
                                if action_cols[1].button(
                                    "Delete user",
                                    key=f"vardaan_delete_user_{uid}",
                                    use_container_width=True,
                                    disabled=not confirm_delete,
                                ):
                                    admin.table("profiles").delete().eq("id", uid).execute()
                                    try:
                                        admin.auth.admin.delete_user(uid)
                                    except Exception as delete_auth_error:
                                        st.warning(
                                            "The profile was deleted, but the authentication account may need "
                                            f"administrator cleanup: {delete_auth_error}"
                                        )
                                    else:
                                        st.success(f"Deleted {email}.")
                                    st.rerun()

            with add_tab:
                st.markdown('<div class="userx-section">CREATE A NEW ACCOUNT</div>', unsafe_allow_html=True)
                st.caption("Create the login and assign its access level. New accounts are active immediately.")
                with st.form("vardaan_add_user_form", clear_on_submit=True):
                    new_email = st.text_input("User email", placeholder="name@example.com")
                    new_password = st.text_input("Temporary password", type="password")
                    new_role = st.selectbox(
                        "Access role",
                        ["user", "admin"],
                        format_func=lambda value: "User" if value == "user" else "Admin",
                    )
                    create_user_submitted = st.form_submit_button(
                        "Create user account",
                        use_container_width=True,
                    )

                if create_user_submitted:
                    clean_email = new_email.strip().lower()
                    if not clean_email or "@" not in clean_email:
                        st.error("Enter a valid email address.")
                    elif len(new_password) < 6:
                        st.error("Password must be at least 6 characters.")
                    else:
                        try:
                            created = admin.auth.admin.create_user({
                                "email": clean_email,
                                "password": new_password,
                                "email_confirm": True,
                            })
                            created_uid = created.user.id
                            admin.table("profiles").insert({
                                "id": created_uid,
                                "email": clean_email,
                                "role": new_role,
                                "is_active": True,
                            }).execute()
                            st.success(f"Account created successfully for {clean_email}.")
                            st.rerun()
                        except Exception as create_error:
                            st.error(f"Could not create the account: {create_error}")

        except Exception as e:
            st.error(f"User management error: {e}")


    # ============================================================
    # NORMAL USER
    # ============================================================

    else:

        st.info(
            "USER ACCESS â€” Read-only"
        )



else:
    # Never show Master Excel for an unknown/stale navigation value.
    st.session_state["main_dashboard_default_index_v2"] = "AutoGreen"
    render_autogreen_dashboard()


# Matching dashboard shortcuts: same dimensions and states as sidebar navigation.
_bottom_nav_page = st.session_state.get("main_dashboard_default_index_v2", "AutoGreen")
_master_excel_ice_blue_theme = {
    "deep": "#123B5D",
    "primary": "#17658B",
    "accent": "#168FC0",
    "ink": "#173B50",
    "accent_text": "#155E91",
    "tint": "#DDF2FC",
    "canvas": "#EAF7FF",
    "border": "#B9DDED",
    "soft_text": "#557084",
}
_index_toppers_sidebar_theme = {
    # Exact Chocolate Brown palette used by Index Toppers sidebar navigation.
    "deep": "#3B2118",
    "primary": "#8A5638",
    "accent": "#A97858",
    "ink": "#3B2118",
    "accent_text": "#3B2118",
    "tint": "#EAD9CC",
    "canvas": "#F7F0EA",
    "border": "#D6B49A",
    "soft_text": "#4A2B20",
    "left_accent": "#8A5638",
}
_bottom_nav_theme = (
    _master_excel_ice_blue_theme
    if _bottom_nav_page == "Master Excel"
    else ({**_index_toppers_sidebar_theme, "deep":"#365B3B", "primary":"#658B4E", "accent":"#8FAF72", "ink":"#365B3B", "accent_text":"#365B3B", "tint":"#DDEBCF", "canvas":"#F4F8EF", "border":"#C8DDB5", "soft_text":"#4F7045", "left_accent":"#658B4E"} if st.session_state.get("_index_toppers_pistachio_mode") else _index_toppers_sidebar_theme)
    if _bottom_nav_page == "Index Toppers"
    else st.session_state.get("_vardaan_active_theme") or _master_excel_ice_blue_theme
)
_bottom_nav_left_accent = _bottom_nav_theme.get("left_accent", _bottom_nav_theme["accent"])
st.markdown(
    f"""
    <style>
    .vardaan-bottom-nav-label {{
        color:{_bottom_nav_theme["soft_text"]} !important;
        font-size:.72rem;font-weight:800;letter-spacing:.12em;
        text-transform:uppercase;margin:.75rem 0 .55rem 2px;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav {{ width:100%; }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] > div[role="radiogroup"] {{
        display:flex !important;flex-direction:row !important;
        align-items:stretch !important;gap:.45rem !important;width:100% !important;
        padding:0 !important;border:0 !important;border-radius:0 !important;
        background:transparent !important;box-shadow:none !important;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label {{
        flex:1 1 0 !important;min-width:0 !important;
        display:flex !important;align-items:center !important;justify-content:center !important;
        min-height:3rem !important;padding:.55rem .8rem !important;margin:0 !important;
        border:1px solid {_bottom_nav_theme["border"]} !important;
        border-radius:11px !important;
        background:linear-gradient(135deg,#FFFFFF 0%,{_bottom_nav_theme["canvas"]} 100%) !important;
        box-shadow:0 2px 8px rgba(23,101,139,.08) !important;
        transition:all 150ms ease !important;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label:hover {{
        border-color:{_bottom_nav_theme["accent"]} !important;
        background:linear-gradient(110deg,{_bottom_nav_theme["tint"]} 0%,#FFFFFF 100%) !important;
        transform:translateX(2px) !important;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) {{
        border:1.5px solid {_bottom_nav_theme["accent"]} !important;
        background:linear-gradient(110deg,{_bottom_nav_theme["tint"]} 0%,#FFFFFF 100%) !important;
        box-shadow:inset 4px 0 0 {_bottom_nav_left_accent},0 3px 10px rgba(23,101,139,.12) !important;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label p {{
        color:{_bottom_nav_theme["accent_text"]} !important;
        font-weight:750 !important;font-size:.94rem !important;
        text-align:center !important;margin:0 !important;
    }}
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label,
    [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label * {{
        color:{_bottom_nav_theme["accent_text"]} !important;
    }}

    /* Keep shortcut text on the dashboard theme colour, including Streamlit's selected state. */
    .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label,
    .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label *,
    .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked),
    .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label:has(input:checked) * {{
        color:{_bottom_nav_theme["accent_text"]} !important;
        -webkit-text-fill-color:{_bottom_nav_theme["accent_text"]} !important;
        opacity:1 !important;
    }}

    @media(max-width:640px) {{
        [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] {{ gap:.35rem !important; }}
        [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label {{
            min-height:3rem !important;padding:.5rem .3rem !important;
        }}
        [data-testid="stMainBlockContainer"] .st-key-vardaan-bottom-nav [data-testid="stRadio"] [role="radiogroup"] label p {{ font-size:.78rem !important; }}
    }}
    </style>
    """,
    unsafe_allow_html=True,
)
with st.container(key="vardaan-bottom-nav"):
    st.markdown(
        '<div class="vardaan-bottom-nav-label">DASHBOARDS</div>',
        unsafe_allow_html=True,
    )
    st.radio(
        "Dashboard shortcuts",
        ["AutoGreen", "Stock of the Day", "Index Toppers", "Master Excel"],
        format_func=lambda item: {
            "AutoGreen": "🌿  AutoGreen",
            "Stock of the Day": "⭐  Stock of the Day",
            "Index Toppers": "🏆  Index Toppers",
            "Master Excel": "📊  Master Excel",
        }[item],
        key="bottom_dashboard_nav_v1",
        on_change=sync_bottom_dashboard_to_main,
        horizontal=True,
        label_visibility="collapsed",
    )
