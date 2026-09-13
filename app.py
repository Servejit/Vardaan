import io
import os
import re
import hashlib
import secrets
from datetime import datetime, date

import pandas as pd
import numpy as np
import requests
import yfinance as yf
import streamlit as st

from supabase import create_client, Client

from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="6thSense (6S-FO200) Vardaan",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)


# ============================================================
# CONSTANTS
# ============================================================

APP_TITLE = "6thSense (6S-FO200) Vardaan"

BUCKET_NAME = "vardaan-master"
MASTER_PATH = "Master.xlsx"

OUTPUT_FILENAME = "6thsense(6S-FO200)Vardaan.xlsx"

DEFAULT_BLUE = "ADD8E6"
DEFAULT_GREEN = "90EE90"


# ============================================================
# SUPABASE CONNECTION
# ============================================================

@st.cache_resource
def get_supabase() -> Client:

    url = st.secrets["SUPABASE_URL"]
    key = st.secrets["SUPABASE_SECRET_KEY"]

    return create_client(
        url,
        key
    )


supabase = get_supabase()


# ============================================================
# CSS
# ============================================================

st.markdown(
    """
    <style>

    .main-title {
        text-align:center;
        font-size:32px;
        font-weight:700;
        margin-top:5px;
        margin-bottom:20px;
    }

    .status-box {
        padding:12px;
        border-radius:10px;
        border:1px solid #dddddd;
        background:#f7f7f7;
        margin-bottom:12px;
    }

    .pulse {
        animation: pulse-green 1.2s infinite;
        border-radius:10px;
        padding:12px;
        font-weight:700;
        text-align:center;
    }

    @keyframes pulse-green {
        0%   { opacity:1; }
        50%  { opacity:0.35; }
        100% { opacity:1; }
    }

    div.stButton > button {
        width:100%;
        font-weight:600;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# PASSWORD FUNCTIONS
# ============================================================

def hash_password(password, salt=None):

    if salt is None:
        salt = secrets.token_hex(16)

    dk = hashlib.pbkdf2_hmac(
        "sha256",
        password.encode("utf-8"),
        salt.encode("utf-8"),
        120000
    )

    return f"{salt}${dk.hex()}"


def verify_password(password, stored):

    try:

        salt, stored_hash = stored.split("$", 1)

        dk = hashlib.pbkdf2_hmac(
            "sha256",
            password.encode("utf-8"),
            salt.encode("utf-8"),
            120000
        )

        return secrets.compare_digest(
            dk.hex(),
            stored_hash
        )

    except Exception:

        return False


# ============================================================
# ADMIN CREDENTIALS
# ============================================================

def get_admin_credentials():

    username = st.secrets.get(
        "ADMIN_USERNAME",
        "VardaanAdmin"
    )

    password = st.secrets.get(
        "ADMIN_PASSWORD",
        ""
    )

    return str(username), str(password)


# ============================================================
# DATABASE SETTINGS
# ============================================================

def get_setting(key, default=""):

    try:

        response = (
            supabase
            .table("vardaan_settings")
            .select("setting_value")
            .eq("setting_key", key)
            .limit(1)
            .execute()
        )

        if response.data:

            value = response.data[0].get(
                "setting_value"
            )

            if value is not None:
                return value

    except Exception:
        pass

    return default


def set_setting(key, value):

    supabase.table(
        "vardaan_settings"
    ).upsert(
        {
            "setting_key": key,
            "setting_value": str(value),
            "updated_at": datetime.utcnow().isoformat()
        }
    ).execute()


# ============================================================
# USER MANAGEMENT
# ============================================================

def get_users():

    try:

        response = (
            supabase
            .table("vardaan_users")
            .select(
                "username,enabled,created_at,updated_at"
            )
            .order(
                "username"
            )
            .execute()
        )

        return response.data or []

    except Exception:

        return []


def add_user(username, password):

    username = username.strip()

    if not username:
        return False, "User ID is required."

    if not password:
        return False, "Password is required."

    admin_username, _ = get_admin_credentials()

    if username.lower() == admin_username.lower():

        return (
            False,
            "This User ID is reserved for Admin."
        )

    existing = (
        supabase
        .table("vardaan_users")
        .select("username")
        .eq("username", username)
        .limit(1)
        .execute()
    )

    if existing.data:

        return (
            False,
            "User already exists."
        )

    supabase.table(
        "vardaan_users"
    ).insert(
        {
            "username": username,
            "password_hash": hash_password(password),
            "enabled": True
        }
    ).execute()

    return True, "User created successfully."


def change_user_password(username, password):

    if not password:

        return (
            False,
            "New password is required."
        )

    response = (
        supabase
        .table("vardaan_users")
        .update(
            {
                "password_hash": hash_password(password),
                "updated_at": datetime.utcnow().isoformat()
            }
        )
        .eq("username", username)
        .execute()
    )

    if response.data:

        return (
            True,
            "Password changed successfully."
        )

    return (
        False,
        "User not found."
    )


def set_user_enabled(username, enabled):

    response = (
        supabase
        .table("vardaan_users")
        .update(
            {
                "enabled": bool(enabled),
                "updated_at": datetime.utcnow().isoformat()
            }
        )
        .eq("username", username)
        .execute()
    )

    if response.data:

        return (
            True,
            "User status updated."
        )

    return (
        False,
        "User not found."
    )


def delete_user(username):

    response = (
        supabase
        .table("vardaan_users")
        .delete()
        .eq("username", username)
        .execute()
    )

    if response.data:

        return (
            True,
            "User deleted successfully."
        )

    return (
        False,
        "User not found."
    )


# ============================================================
# AUTHENTICATION
# ============================================================

def authenticate(username, password):

    username = username.strip()

    admin_username, admin_password = (
        get_admin_credentials()
    )

    if (
        username == admin_username
        and password == admin_password
    ):

        return True, "admin"

    try:

        response = (
            supabase
            .table("vardaan_users")
            .select(
                "username,password_hash,enabled"
            )
            .eq("username", username)
            .limit(1)
            .execute()
        )

        if not response.data:
            return False, None

        user = response.data[0]

        if not user.get("enabled", False):

            return False, "disabled"

        if verify_password(
            password,
            user["password_hash"]
        ):

            return True, "user"

    except Exception:

        return False, None

    return False, None


# ============================================================
# SUPABASE MASTER STORAGE
# ============================================================

def master_exists():

    try:

        result = (
            supabase
            .storage
            .from_(BUCKET_NAME)
            .list(
                path=""
            )
        )

        for item in result:

            if item.get("name") == MASTER_PATH:
                return True

    except Exception:

        pass

    return False


def upload_master(uploaded_file):

    data = uploaded_file.getvalue()

    if not data:

        raise ValueError(
            "Uploaded file is empty."
        )

    # Validate first.
    wb = load_workbook(
        io.BytesIO(data),
        read_only=True,
        data_only=True
    )

    sheet_names = [
        s.lower()
        for s in wb.sheetnames
    ]

    if "summary" not in sheet_names:

        wb.close()

        raise ValueError(
            "Master Excel must contain a Summary sheet."
        )

    wb.close()

    # Replace existing Master.
    try:

        supabase.storage.from_(
            BUCKET_NAME
        ).remove(
            [MASTER_PATH]
        )

    except Exception:
        pass

    supabase.storage.from_(
        BUCKET_NAME
    ).upload(
        MASTER_PATH,
        data,
        {
            "content-type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            "upsert": "true"
        }
    )


def download_master():

    return (
        supabase
        .storage
        .from_(BUCKET_NAME)
        .download(MASTER_PATH)
    )


def clear_master():

    try:

        supabase.storage.from_(
            BUCKET_NAME
        ).remove(
            [MASTER_PATH]
        )

    except Exception:
        pass


# ============================================================
# FIND SUMMARY SHEET
# ============================================================

def get_summary_sheet(wb):

    for sheet in wb.sheetnames:

        if str(sheet).strip().lower() == "summary":

            return wb[sheet]

    return None


# ============================================================
# HEADER HELPERS
# ============================================================

def normalize_header(value):

    if value is None:
        return ""

    return re.sub(
        r"[^a-z0-9]+",
        "",
        str(value).strip().lower()
    )


def find_heading(ws, names):

    wanted = {
        normalize_header(x)
        for x in names
    }

    for col in range(
        1,
        ws.max_column + 1
    ):

        value = ws.cell(
            1,
            col
        ).value

        if normalize_header(value) in wanted:
            return col

    return None


def find_heading_contains(ws, text):

    target = normalize_header(text)

    for col in range(
        1,
        ws.max_column + 1
    ):

        value = ws.cell(
            1,
            col
        ).value

        if target in normalize_header(value):
            return col

    return None


# ============================================================
# NUMBER HELPERS
# ============================================================

def numeric_value(value):

    if value is None:
        return None

    if isinstance(
        value,
        (int, float, np.integer, np.floating)
    ):

        if pd.isna(value):
            return None

        return float(value)

    text = str(value).replace(
        ",",
        ""
    )

    match = re.search(
        r"[-+]?\d+(?:\.\d+)?",
        text
    )

    if not match:
        return None

    try:
        return float(
            match.group(0)
        )
    except Exception:
        return None


def first_number(value):

    if value is None:
        return None

    match = re.search(
        r"[-+]?\d+(?:\.\d+)?",
        str(value)
    )

    if not match:
        return None

    return float(
        match.group(0)
    )


def number_inside_parentheses(value):

    if value is None:
        return None

    match = re.search(
        r"\(\s*([-+]?\d+(?:\.\d+)?)\s*\)",
        str(value)
    )

    if not match:
        return None

    return float(
        match.group(1)
    )


# ============================================================
# DATE / RECENT O2L-P2L COLUMN
# ============================================================

MONTHS = {
    "jan": 1,
    "feb": 2,
    "mar": 3,
    "apr": 4,
    "may": 5,
    "jun": 6,
    "jul": 7,
    "aug": 8,
    "sep": 9,
    "oct": 10,
    "nov": 11,
    "dec": 12
}


def extract_date(value):

    if value is None:
        return None

    text = str(value)

    match = re.search(
        r"\b(\d{1,2})\s+"
        r"(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"\b",
        text,
        re.IGNORECASE
    )

    if match:

        day = int(
            match.group(1)
        )

        month = MONTHS[
            match.group(2).lower()[:3]
        ]

        year = datetime.now().year

        try:

            return date(
                year,
                month,
                day
            )

        except Exception:
            return None

    return None


def find_recent_o2l_p2l(ws):

    candidates = []

    for col in range(
        1,
        ws.max_column + 1
    ):

        heading = ws.cell(
            1,
            col
        ).value

        if heading is None:
            continue

        text = str(
            heading
        ).upper()

        # Your example uses O2L.
        # Also accepts P2L if present.
        if (
            "O2L" not in text
            and "P2L" not in text
        ):
            continue

        dt = extract_date(
            heading
        )

        if dt:

            candidates.append(
                (
                    dt,
                    col
                )
            )

    if not candidates:
        return None

    candidates.sort(
        key=lambda x: x[0],
        reverse=True
    )

    return candidates[0][1]


# ============================================================
# EXCEL COLOR
# ============================================================

def clean_hex(value, fallback):

    if value is None:
        return fallback

    value = str(value).replace(
        "#",
        ""
    ).upper()

    if re.fullmatch(
        r"[0-9A-F]{6}",
        value
    ):

        return value

    return fallback


def make_fill(hex_color):

    return PatternFill(
        fill_type="solid",
        fgColor=hex_color
    )


# ============================================================
# MAIN EXCEL GENERATOR
# ============================================================

def generate_vardaan(master_bytes):

    # --------------------------------------------------------
    # Read Master with calculated values only
    # --------------------------------------------------------

    source_wb = load_workbook(
        io.BytesIO(master_bytes),
        data_only=True
    )

    source_ws = get_summary_sheet(
        source_wb
    )

    if source_ws is None:

        raise ValueError(
            "Summary sheet not found."
        )

    # --------------------------------------------------------
    # New workbook
    # --------------------------------------------------------

    output_wb = Workbook()

    output_ws = output_wb.active

    output_ws.title = "summary"

    # --------------------------------------------------------
    # COPY VALUES ONLY
    # --------------------------------------------------------

    for row in source_ws.iter_rows():

        for cell in row:

            output_ws.cell(
                row=cell.row,
                column=cell.column,
                value=cell.value
            )

    # --------------------------------------------------------
    # COLORS
    # --------------------------------------------------------

    blue = clean_hex(
        get_setting(
            "blue_color",
            DEFAULT_BLUE
        ),
        DEFAULT_BLUE
    )

    green = clean_hex(
        get_setting(
            "green_color",
            DEFAULT_GREEN
        ),
        DEFAULT_GREEN
    )

    blue_fill = make_fill(
        blue
    )

    green_fill = make_fill(
        green
    )

    # --------------------------------------------------------
    # HEADINGS
    # --------------------------------------------------------

    sum_i_col = find_heading(
        output_ws,
        ["Sum I"]
    )

    gt_col = find_heading(
        output_ws,
        ["16> C-B / Avg.4"]
    )

    lt_col = find_heading(
        output_ws,
        ["16< D-B / Avg.4"]
    )

    sum_o2h_col = find_heading_contains(
        output_ws,
        "Sum O2H.10"
    )

    sum_o2l_col = find_heading_contains(
        output_ws,
        "Sum O2L.10"
    )

    recent_o2l_col = find_recent_o2l_p2l(
        output_ws
    )

    # --------------------------------------------------------
    # PRE-CALCULATE O2H/O2L AVERAGES
    # --------------------------------------------------------

    o2h_average = None
    o2l_average = None

    if sum_o2h_col:

        values = []

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            v = numeric_value(
                output_ws.cell(
                    row,
                    sum_o2h_col
                ).value
            )

            if v is not None:
                values.append(v)

        if values:
            o2h_average = float(
                np.mean(values)
            )

    if sum_o2l_col:

        values = []

        for row in range(
            2,
            output_ws.max_row + 1
        ):

            v = numeric_value(
                output_ws.cell(
                    row,
                    sum_o2l_col
                ).value
            )

            if v is not None:
                values.append(v)

        if values:
            o2l_average = float(
                np.mean(values)
            )

    # --------------------------------------------------------
    # CONDITION COUNTS
    # --------------------------------------------------------

    condition_counts = {}

    for row in range(
        2,
        output_ws.max_row + 1
    ):

        count = 0

        # ====================================================
        # CONDITION 1
        # Sum I < -4.00
        # ====================================================

        if sum_i_col:

            value = numeric_value(
                output_ws.cell(
                    row,
                    sum_i_col
                ).value
            )

            if (
                value is not None
                and value < -4.00
            ):

                output_ws.cell(
                    row,
                    sum_i_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        # ====================================================
        # CONDITION 2
        #
        # 16>0.25, Avg.4 (0.44)
        #
        # Parentheses value < 0.50
        # ====================================================

        if gt_col:

            cell_value = output_ws.cell(
                row,
                gt_col
            ).value

            avg_value = (
                number_inside_parentheses(
                    cell_value
                )
            )

            if (
                avg_value is not None
                and avg_value < 0.50
            ):

                output_ws.cell(
                    row,
                    gt_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        # ====================================================
        # CONDITION 3
        #
        # 16<-1.09, Avg.4 (-1.13)
        #
        # -1.13 < -1.09
        # AND
        # -1.13 < -1.00
        # ====================================================

        if lt_col:

            cell_value = output_ws.cell(
                row,
                lt_col
            ).value

            first_val = first_number(
                cell_value
            )

            avg_val = (
                number_inside_parentheses(
                    cell_value
                )
            )

            if (
                first_val is not None
                and avg_val is not None
                and avg_val < first_val
                and avg_val < -1.00
            ):

                output_ws.cell(
                    row,
                    lt_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        # ====================================================
        # CONDITION 4
        # Sum O2H.10 below average
        # ====================================================

        if (
            sum_o2h_col
            and o2h_average is not None
        ):

            value = numeric_value(
                output_ws.cell(
                    row,
                    sum_o2h_col
                ).value
            )

            if (
                value is not None
                and value < o2h_average
            ):

                output_ws.cell(
                    row,
                    sum_o2h_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        # ====================================================
        # CONDITION 5
        # Sum O2L.10 below average
        # ====================================================

        if (
            sum_o2l_col
            and o2l_average is not None
        ):

            value = numeric_value(
                output_ws.cell(
                    row,
                    sum_o2l_col
                ).value
            )

            if (
                value is not None
                and value < o2l_average
            ):

                output_ws.cell(
                    row,
                    sum_o2l_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        # ====================================================
        # CONDITION 6
        #
        # Most recent O2L/P2L < -1.00
        # ====================================================

        if recent_o2l_col:

            value = numeric_value(
                output_ws.cell(
                    row,
                    recent_o2l_col
                ).value
            )

            if (
                value is not None
                and value < -1.00
            ):

                output_ws.cell(
                    row,
                    recent_o2l_col
                ).fill = blue_fill

                output_ws.cell(
                    row,
                    1
                ).fill = blue_fill

                count += 1

        condition_counts[row] = count

    # ========================================================
    # MOST CONDITIONS = GREEN
    # ========================================================

    max_conditions = 0

    if condition_counts:

        max_conditions = max(
            condition_counts.values()
        )

    if max_conditions > 0:

        for row, count in condition_counts.items():

            if count == max_conditions:

                for col in range(
                    1,
                    output_ws.max_column + 1
                ):

                    output_ws.cell(
                        row,
                        col
                    ).fill = green_fill

    # ========================================================
    # HEADER FORMAT
    # ========================================================

    for cell in output_ws[1]:

        cell.font = Font(
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    # ========================================================
    # AUTO FILTER / SORT BUTTONS
    # ========================================================

    output_ws.freeze_panes = "A2"

    output_ws.auto_filter.ref = (
        f"A1:"
        f"{get_column_letter(output_ws.max_column)}"
        f"{output_ws.max_row}"
    )

    # ========================================================
    # COLUMN WIDTHS
    # ========================================================

    for col in range(
        1,
        output_ws.max_column + 1
    ):

        maximum = 0

        for row in range(
            1,
            min(
                output_ws.max_row,
                500
            ) + 1
        ):

            value = output_ws.cell(
                row,
                col
            ).value

            if value is None:
                continue

            maximum = max(
                maximum,
                len(str(value))
            )

        output_ws.column_dimensions[
            get_column_letter(col)
        ].width = min(
            max(
                maximum + 2,
                10
            ),
            35
        )

    # ========================================================
    # SAVE
    # ========================================================

    output = io.BytesIO()

    output_wb.save(
        output
    )

    output.seek(0)

    source_wb.close()
    output_wb.close()

    return output.getvalue(), max_conditions


# ============================================================
# TELEGRAM
# ============================================================

def telegram_send(message):

    token = get_setting(
        "telegram_bot_token",
        ""
    ).strip()

    chat_id = get_setting(
        "telegram_chat_id",
        ""
    ).strip()

    if not token or not chat_id:

        return (
            False,
            "Telegram Bot Token or Chat ID is missing."
        )

    url = (
        "https://api.telegram.org/bot"
        + token
        + "/sendMessage"
    )

    try:

        response = requests.post(
            url,
            data={
                "chat_id": chat_id,
                "text": message
            },
            timeout=20
        )

        if response.ok:

            return (
                True,
                "Telegram message sent."
            )

        return (
            False,
            response.text
        )

    except Exception as e:

        return (
            False,
            str(e)
        )


def calculate_live_o2l(symbol):

    symbol = str(
        symbol
    ).strip().upper()

    if not symbol:
        return None

    if not symbol.endswith(".NS"):
        symbol += ".NS"

    try:

        df = yf.download(
            symbol,
            period="1d",
            interval="5m",
            progress=False,
            auto_adjust=False,
            threads=False
        )

        if df is None or df.empty:
            return None

        if isinstance(
            df.columns,
            pd.MultiIndex
        ):

            df.columns = (
                df.columns
                .get_level_values(0)
            )

        if (
            "Open" not in df.columns
            or "Low" not in df.columns
        ):
            return None

        df = df.dropna(
            subset=[
                "Open",
                "Low"
            ]
        )

        if df.empty:
            return None

        day_open = float(
            df["Open"].iloc[0]
        )

        day_low = float(
            df["Low"].min()
        )

        if day_open == 0:
            return None

        return (
            (day_low - day_open)
            / day_open
            * 100
        )

    except Exception:

        return None


def telegram_o2l_scan():

    if (
        get_setting(
            "telegram_enabled",
            "0"
        ) != "1"
    ):

        return (
            False,
            "Telegram alerts are disabled by Admin."
        )

    try:

        master = download_master()

    except Exception:

        return (
            False,
            "Master Excel is not available."
        )

    wb = load_workbook(
        io.BytesIO(master),
        data_only=True
    )

    ws = get_summary_sheet(
        wb
    )

    if ws is None:

        return (
            False,
            "Summary sheet not found."
        )

    symbol_col = find_heading(
        ws,
        [
            "Symbol",
            "Stock",
            "Scrip",
            "Ticker",
            "Name"
        ]
    )

    if not symbol_col:

        return (
            False,
            "Symbol column not found."
        )

    operator = get_setting(
        "alert_operator",
        "<"
    )

    try:

        threshold = float(
            get_setting(
                "alert_value",
                "-1.00"
            )
        )

    except Exception:

        threshold = -1.00

    matches = []

    for row in range(
        2,
        ws.max_row + 1
    ):

        symbol = ws.cell(
            row,
            symbol_col
        ).value

        if not symbol:
            continue

        value = calculate_live_o2l(
            symbol
        )

        if value is None:
            continue

        condition = False

        if operator == "<":
            condition = value < threshold

        elif operator == "<=":
            condition = value <= threshold

        elif operator == ">":
            condition = value > threshold

        elif operator == ">=":
            condition = value >= threshold

        if condition:

            matches.append(
                (
                    str(symbol),
                    value
                )
            )

    wb.close()

    if not matches:

        return (
            True,
            "No stocks met the Telegram condition."
        )

    lines = [
        "📊 6thSense (6S-FO200) Vardaan",
        "",
        "O2L ALERT",
        ""
    ]

    for symbol, value in matches:

        lines.append(
            f"{symbol}: {value:.2f}%"
        )

    ok, message = telegram_send(
        "\n".join(lines)
    )

    if ok:

        return (
            True,
            f"{len(matches)} alert(s) sent."
        )

    return False, message


# ============================================================
# SESSION
# ============================================================

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "role" not in st.session_state:
    st.session_state.role = None

if "username" not in st.session_state:
    st.session_state.username = None

if "output_bytes" not in st.session_state:
    st.session_state.output_bytes = None

if "max_conditions" not in st.session_state:
    st.session_state.max_conditions = 0


# ============================================================
# LOGOUT
# ============================================================

def logout():

    st.session_state.logged_in = False
    st.session_state.role = None
    st.session_state.username = None
    st.session_state.output_bytes = None
    st.session_state.max_conditions = 0

    st.rerun()


# ============================================================
# LOGIN
# ============================================================

def login_page():

    st.markdown(
        f"""
        <div class="main-title">
        {APP_TITLE}
        </div>
        """,
        unsafe_allow_html=True
    )

    st.subheader(
        "Login"
    )

    username = st.text_input(
        "User ID",
        key="login_user"
    )

    password = st.text_input(
        "Password",
        type="password",
        key="login_password"
    )

    if st.button(
        "Login",
        type="primary"
    ):

        if not username or not password:

            st.error(
                "Enter User ID and Password."
            )

            return

        success, role = authenticate(
            username,
            password
        )

        if success:

            st.session_state.logged_in = True
            st.session_state.role = role
            st.session_state.username = (
                username.strip()
            )

            st.rerun()

        elif role == "disabled":

            st.error(
                "Your account has been disabled by Admin."
            )

        else:

            st.error(
                "Invalid User ID or Password."
            )


# ============================================================
# ADMIN PAGE
# ============================================================

def admin_page():

    st.markdown(
        f"""
        <div class="main-title">
        {APP_TITLE}
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(
        [5, 1]
    )

    with c1:

        st.success(
            f"Admin: {st.session_state.username}"
        )

    with c2:

        if st.button(
            "Logout",
            key="admin_logout"
        ):

            logout()

    # ========================================================
    # MASTER
    # ========================================================

    st.header(
        "Master Excel"
    )

    if master_exists():

        st.success(
            "Master Excel is available."
        )

    else:

        st.warning(
            "Master Excel is not available."
        )

    uploaded = st.file_uploader(
        "Upload / Replace Master",
        type=[
            "xlsx",
            "xlsm"
        ],
        key="admin_upload"
    )

    if uploaded is not None:

        if st.button(
            "Save Master",
            key="save_master"
        ):

            try:

                upload_master(
                    uploaded
                )

                st.success(
                    "Master Excel saved successfully."
                )

                st.rerun()

            except Exception as e:

                st.error(
                    f"Master upload failed: {e}"
                )

    if master_exists():

        if st.button(
            "Clear Master",
            key="admin_clear_master"
        ):

            clear_master()

            st.session_state.output_bytes = None

            st.success(
                "Master Excel cleared."
            )

            st.rerun()

    # ========================================================
    # GENERATE
    # ========================================================

    st.divider()

    st.header(
        "Generate"
    )

    if master_exists():

        if st.button(
            "Generate Vardaan",
            type="primary",
            key="admin_generate"
        ):

            try:

                master = download_master()

                output, count = generate_vardaan(
                    master
                )

                st.session_state.output_bytes = (
                    output
                )

                st.session_state.max_conditions = (
                    count
                )

                st.success(
                    "Vardaan Excel generated."
                )

            except Exception as e:

                st.error(
                    f"Generation error: {e}"
                )

        if st.session_state.output_bytes:

            st.download_button(
                "Download Vardaan Excel",
                data=st.session_state.output_bytes,
                file_name=OUTPUT_FILENAME,
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                ),
                key="admin_download"
            )

            if st.session_state.max_conditions > 0:

                st.markdown(
                    f"""
                    <div class="pulse">
                    Maximum matching conditions:
                    {st.session_state.max_conditions}
                    </div>
                    """,
                    unsafe_allow_html=True
                )

    else:

        st.info(
            "Upload Master Excel first."
        )

    # ========================================================
    # USERS
    # ========================================================

    st.divider()

    st.header(
        "Users"
    )

    tabs = st.tabs(
        [
            "Add",
            "Password",
            "Status",
            "Delete"
        ]
    )

    # --------------------------------------------------------
    # ADD
    # --------------------------------------------------------

    with tabs[0]:

        new_user = st.text_input(
            "User ID",
            key="new_user"
        )

        new_pass = st.text_input(
            "Password",
            type="password",
            key="new_pass"
        )

        if st.button(
            "Add User",
            key="add_user"
        ):

            ok, message = add_user(
                new_user,
                new_pass
            )

            if ok:
                st.success(message)
            else:
                st.error(message)

    # --------------------------------------------------------
    # PASSWORD
    # --------------------------------------------------------

    with tabs[1]:

        users = get_users()

        if users:

            names = [
                x["username"]
                for x in users
            ]

            selected = st.selectbox(
                "User",
                names,
                key="password_user_select"
            )

            password = st.text_input(
                "New Password",
                type="password",
                key="password_new"
            )

            if st.button(
                "Change Password",
                key="change_user_password"
            ):

                ok, message = (
                    change_user_password(
                        selected,
                        password
                    )
                )

                if ok:
                    st.success(message)
                else:
                    st.error(message)

        else:

            st.info(
                "No users created."
            )

    # --------------------------------------------------------
    # STATUS
    # --------------------------------------------------------

    with tabs[2]:

        users = get_users()

        if users:

            names = [
                x["username"]
                for x in users
            ]

            selected = st.selectbox(
                "User",
                names,
                key="status_user_select"
            )

            selected_data = next(
                (
                    x
                    for x in users
                    if x["username"] == selected
                ),
                None
            )

            current = bool(
                selected_data["enabled"]
            )

            status = st.radio(
                "Status",
                [
                    "Enabled",
                    "Disabled"
                ],
                index=0 if current else 1,
                key="status_radio"
            )

            if st.button(
                "Save Status",
                key="save_status"
            ):

                ok, message = (
                    set_user_enabled(
                        selected,
                        status == "Enabled"
                    )
                )

                if ok:
                    st.success(message)
                else:
                    st.error(message)

        else:

            st.info(
                "No users created."
            )

    # --------------------------------------------------------
    # DELETE
    # --------------------------------------------------------

    with tabs[3]:

        users = get_users()

        if users:

            names = [
                x["username"]
                for x in users
            ]

            selected = st.selectbox(
                "User",
                names,
                key="delete_user_select"
            )

            confirm = st.checkbox(
                "Confirm permanent deletion",
                key="delete_confirm"
            )

            if st.button(
                "Delete User",
                key="delete_user"
            ):

                if not confirm:

                    st.warning(
                        "Confirm deletion first."
                    )

                else:

                    ok, message = delete_user(
                        selected
                    )

                    if ok:
                        st.success(message)
                    else:
                        st.error(message)

        else:

            st.info(
                "No users created."
            )

    # ========================================================
    # USER LIST
    # ========================================================

    st.subheader(
        "Current Users"
    )

    users = get_users()

    if users:

        display = []

        for user in users:

            display.append(
                {
                    "User ID": user["username"],
                    "Status": (
                        "Enabled"
                        if user["enabled"]
                        else "Disabled"
                    ),
                    "Created": user["created_at"]
                }
            )

        st.dataframe(
            pd.DataFrame(display),
            use_container_width=True,
            hide_index=True
        )

    # ========================================================
    # COLORS
    # ========================================================

    st.divider()

    st.header(
        "Excel Colors"
    )

    current_blue = get_setting(
        "blue_color",
        DEFAULT_BLUE
    )

    current_green = get_setting(
        "green_color",
        DEFAULT_GREEN
    )

    blue = st.text_input(
        "Light Blue HEX",
        value=current_blue,
        key="admin_blue"
    )

    green = st.text_input(
        "Light Green HEX",
        value=current_green,
        key="admin_green"
    )

    pulse = st.checkbox(
        "Enable green pulse on screen",
        value=True,
        key="green_pulse"
    )

    if st.button(
        "Save Colors",
        key="save_colors"
    ):

        blue = clean_hex(
            blue,
            DEFAULT_BLUE
        )

        green = clean_hex(
            green,
            DEFAULT_GREEN
        )

        set_setting(
            "blue_color",
            blue
        )

        set_setting(
            "green_color",
            green
        )

        set_setting(
            "screen_pulse",
            "1" if pulse else "0"
        )

        st.success(
            "Colors saved."
        )

    # ========================================================
    # TELEGRAM
    # ========================================================

    st.divider()

    st.header(
        "Telegram"
    )

    enabled = (
        get_setting(
            "telegram_enabled",
            "0"
        ) == "1"
    )

    telegram_enabled_value = st.checkbox(
        "Enable Telegram",
        value=enabled,
        key="telegram_enabled"
    )

    token = st.text_input(
        "Bot Token",
        value=get_setting(
            "telegram_bot_token",
            ""
        ),
        type="password",
        key="telegram_token"
    )

    chat = st.text_input(
        "Chat ID",
        value=get_setting(
            "telegram_chat_id",
            ""
        ),
        key="telegram_chat"
    )

    operators = [
        "<",
        "<=",
        ">",
        ">="
    ]

    current_operator = get_setting(
        "alert_operator",
        "<"
    )

    if current_operator not in operators:
        current_operator = "<"

    operator = st.selectbox(
        "O2L Operator",
        operators,
        index=operators.index(
            current_operator
        ),
        key="telegram_operator"
    )

    try:

        threshold = float(
            get_setting(
                "alert_value",
                "-1.00"
            )
        )

    except Exception:

        threshold = -1.00

    alert_value = st.number_input(
        "O2L Alert Value",
        value=threshold,
        step=0.10,
        format="%.2f",
        key="telegram_alert_value"
    )

    if st.button(
        "Save Telegram",
        key="save_telegram"
    ):

        set_setting(
            "telegram_enabled",
            "1"
            if telegram_enabled_value
            else "0"
        )

        set_setting(
            "telegram_bot_token",
            token.strip()
        )

        set_setting(
            "telegram_chat_id",
            chat.strip()
        )

        set_setting(
            "alert_operator",
            operator
        )

        set_setting(
            "alert_value",
            f"{alert_value:.2f}"
        )

        st.success(
            "Telegram settings saved."
        )

    if st.button(
        "Telegram Test",
        key="telegram_test"
    ):

        ok, message = telegram_send(
            "📊 6thSense (6S-FO200) Vardaan\n"
            "Telegram test successful."
        )

        if ok:
            st.success(message)
        else:
            st.error(message)


# ============================================================
# NORMAL USER PAGE
# ============================================================

def user_page():

    st.markdown(
        f"""
        <div class="main-title">
        {APP_TITLE}
        </div>
        """,
        unsafe_allow_html=True
    )

    c1, c2 = st.columns(
        [5, 1]
    )

    with c1:

        st.success(
            f"User: {st.session_state.username}"
        )

    with c2:

        if st.button(
            "Logout",
            key="user_logout"
        ):

            logout()

    st.divider()

    # ========================================================
    # NO UPLOAD
    # NO CLEAR
    # NO ADMIN
    # ========================================================

    if not master_exists():

        st.warning(
            "Master Excel is currently unavailable. "
            "Please contact Admin."
        )

        return

    st.success(
        "Master Excel is ready."
    )

    # ========================================================
    # ONLY GENERATE
    # ========================================================

    if st.button(
        "Generate Vardaan",
        type="primary",
        key="user_generate"
    ):

        try:

            master = download_master()

            output, count = generate_vardaan(
                master
            )

            st.session_state.output_bytes = (
                output
            )

            st.session_state.max_conditions = (
                count
            )

            st.success(
                "Vardaan Excel generated."
            )

        except Exception as e:

            st.error(
                f"Generation error: {e}"
            )

    if st.session_state.output_bytes:

        st.download_button(
            "Download Vardaan Excel",
            data=st.session_state.output_bytes,
            file_name=OUTPUT_FILENAME,
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            key="user_download"
        )

        if (
            st.session_state.max_conditions > 0
            and get_setting(
                "screen_pulse",
                "1"
            ) == "1"
        ):

            st.markdown(
                f"""
                <div class="pulse">
                Maximum matching conditions:
                {st.session_state.max_conditions}
                </div>
                """,
                unsafe_allow_html=True
            )

    # ========================================================
    # TELEGRAM SCAN
    # ========================================================

    if (
        get_setting(
            "telegram_enabled",
            "0"
        ) == "1"
    ):

        st.divider()

        if st.button(
            "Run O2L Alert",
            key="user_telegram"
        ):

            ok, message = telegram_o2l_scan()

            if ok:
                st.success(message)
            else:
                st.error(message)


# ============================================================
# ROUTING
# ============================================================

if not st.session_state.logged_in:

    login_page()

elif st.session_state.role == "admin":

    admin_page()

elif st.session_state.role == "user":

    user_page()

else:

    logout()
