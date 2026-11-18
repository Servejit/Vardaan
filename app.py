import io
import re
import json
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

sb = create_client(
    SUPABASE_URL,
    SUPABASE_KEY,
    options=ClientOptions(flow_type="pkce")
)

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

def load_constituent_archive():
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
        try:
            raw = admin.storage.from_(BUCKET).download(CONSTITUENT_MANIFEST_PATH)
            old_manifest = json.loads(raw.decode("utf-8") if isinstance(raw, bytes) else raw)
        except Exception:
            old_manifest = {}
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
    except Exception as error:
        st.warning(f"Could not save the eligible-index list to Supabase: {error}")


def get_master():
    try:
        return admin.storage.from_(BUCKET).download(
            MASTER_FILE
        )
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


def delete_master():
    admin.storage.from_(BUCKET).remove(
        [MASTER_FILE]
    )


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
    st.rerun()


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

if st.session_state.user is None:

    st.title(
        "6thSense Excel Manager"
    )

    st.subheader("Login")

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

st.title(
    "6thSense Excel Manager"
)

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
        "Logout",
        use_container_width=True
    ):
        logout()

    st.divider()
    page = st.radio(
        "Dashboard",
        ["Master Excel", "Index", "Index 2"],
        index=1,
        key="main_dashboard_default_index"
    )



# ============================================================
# INDEX DASHBOARD
# ============================================================

if page == "Index":

    import csv

    st.header("Index Dashboard")
    st.caption(
        "Upload an Excel (.xlsx) or CSV (.csv) file. "
        "The uploaded file is processed in Streamlit memory and is not saved to Supabase."
    )

    # Each connected Streamlit session has its own session_state. The owner
    # marker also prevents a different account using the same browser session
    # from inheriting the previous account's uploaded Index file or signals.
    session_user_id = str(user.id)
    if st.session_state.get("index_file_owner_id") != session_user_id:
        clear_index_session_data()
        st.session_state["index_file_owner_id"] = session_user_id

    index_upload = st.file_uploader(
        "Upload Index File to replace the saved file",
        type=["xlsx", "csv"],
        key=f"index_excel_upload_{session_user_id}"
    )

    # The trade signal is rendered here so it stays directly below the uploader,
    # rather than inside either H2L or L2H analysis panel.
    trade_signal_placeholder = st.empty()

    try:
        # Keep the latest Index upload in this user's Streamlit session only.
        # It is never written to Supabase or shared with other users.
        if index_upload is not None:
            new_file_bytes = index_upload.getvalue()
            new_file_name = index_upload.name
            if (
                new_file_bytes != st.session_state.get("index_file_bytes")
                or new_file_name != st.session_state.get("index_file_name")
            ):
                # A new upload invalidates old H2L/L2H candidates so signals
                # cannot briefly use stocks calculated from the previous file.
                st.session_state.pop("trade_candidates_H2L", None)
                st.session_state.pop("trade_candidates_L2H", None)
            st.session_state["index_file_bytes"] = new_file_bytes
            st.session_state["index_file_name"] = new_file_name
            st.session_state["index_file_owner_id"] = session_user_id

        file_bytes = st.session_state.get("index_file_bytes")
        file_name = st.session_state.get("index_file_name")

        if file_bytes is None or file_name is None:
            st.info("No Index file in this session yet. Upload an Excel or CSV file above.")
            st.stop()
        extension = file_name.lower().rsplit(".", 1)[-1]

        if extension == "csv":
            decoded = None
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    decoded = file_bytes.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if decoded is None:
                raise ValueError("Could not decode this CSV file.")

            sample = decoded[:8192]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\\t|")
            except csv.Error:
                dialect = csv.excel

            csv_reader = csv.reader(io.StringIO(decoded), dialect)
            csv_rows = list(csv_reader)

            # Remove completely blank lines.
            csv_rows = [
                row for row in csv_rows
                if any(str(value).strip() for value in row)
            ]

            if not csv_rows:
                raise ValueError("The CSV file is empty.")

            raw_headers = csv_rows[0]
            data_values = csv_rows[1:]

            index_wb = None
            sheet_names = ["CSV Data"]
            selected_sheet = "CSV Data"

        else:
            index_wb = load_workbook(
                io.BytesIO(file_bytes),
                read_only=True,
                data_only=True
            )
            sheet_names = index_wb.sheetnames
            selected_sheet = st.selectbox(
                "Select worksheet",
                sheet_names,
                key="index_sheet_selector"
            )
            index_ws = index_wb[selected_sheet]
            all_rows = list(
                index_ws.iter_rows(values_only=True)
            )
            all_rows = [
                list(row) for row in all_rows
                if any(value is not None and str(value).strip() for value in row)
            ]

            if not all_rows:
                raise ValueError("The selected worksheet is empty.")

            raw_headers = all_rows[0]
            data_values = all_rows[1:]

        headers = []
        used_headers = set()

        for position, value in enumerate(raw_headers, start=1):
            header = str(value).strip().replace("\\n", " ") if value is not None else ""
            header = " ".join(header.split())
            if not header:
                header = f"Column {position}"

            base_header = header
            suffix = 2
            while header in used_headers:
                header = f"{base_header}_{suffix}"
                suffix += 1

            used_headers.add(header)
            headers.append(header)

        records = []
        for row in data_values:
            values = list(row)
            if len(values) < len(headers):
                values.extend([None] * (len(headers) - len(values)))
            records.append({
                headers[i]: values[i]
                for i in range(len(headers))
            })

        st.success(f"Loaded: {file_name}")

        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("Worksheets", len(sheet_names))
        metric2.metric("Data Rows", len(records))
        metric3.metric("Columns in Output", 6)

        st.divider()
        st.subheader("Index Data")

        # Display only the six requested columns, regardless of
        # extra spaces or line breaks in the uploaded file headers.
        required_columns = [
            "INDEX",
            "CURRENT",
            "%CHANGE",
            "OPEN",
            "HIGH",
            "LOW"
        ]

        normalized_header_map = {
            " ".join(str(header).upper().split()): header
            for header in headers
        }

        missing_columns = [
            column for column in required_columns
            if column not in normalized_header_map
        ]

        if missing_columns:
            st.error(
                "The uploaded file is missing these required columns: "
                + ", ".join(missing_columns)
            )
            st.caption(
                "Please check that the file contains all six required "
                "column headings."
            )
        else:
            source_columns = [
                normalized_header_map[column]
                for column in required_columns
            ]

            display_records = []

            for record in records:
                output_record = {
                    column: record.get(source_column)
                    for column, source_column in zip(
                        required_columns,
                        source_columns
                    )
                }

                # Calculate current price movement from the day's High
                # and Low. Keep invalid/zero values blank instead of
                # raising a division-by-zero error.
                try:
                    current_value = float(
                        str(output_record["CURRENT"]).replace(",", "").replace("%", "").strip()
                    )
                    high_value = float(
                        str(output_record["HIGH"]).replace(",", "").replace("%", "").strip()
                    )
                    low_value = float(
                        str(output_record["LOW"]).replace(",", "").replace("%", "").strip()
                    )

                    output_record["H2L%"] = (
                        round(((current_value - high_value) / high_value) * 100, 2)
                        if high_value != 0 else None
                    )
                    output_record["L2H%"] = (
                        round(((current_value - low_value) / low_value) * 100, 2)
                        if low_value != 0 else None
                    )
                except (TypeError, ValueError):
                    output_record["H2L%"] = None
                    output_record["L2H%"] = None

                display_records.append(output_record)

            # Summary calculations: compare the average CURRENT price
            # with the average HIGH and average LOW across all indices.
            def to_number(value):
                try:
                    cleaned = str(value).replace(",", "").replace("%", "").strip()
                    return float(cleaned)
                except (TypeError, ValueError):
                    return None

            average_current_values = [
                to_number(record.get("CURRENT"))
                for record in display_records
            ]
            average_high_values = [
                to_number(record.get("HIGH"))
                for record in display_records
            ]
            average_low_values = [
                to_number(record.get("LOW"))
                for record in display_records
            ]

            average_current_values = [v for v in average_current_values if v is not None]
            average_high_values = [v for v in average_high_values if v is not None]
            average_low_values = [v for v in average_low_values if v is not None]

            average_current = (
                sum(average_current_values) / len(average_current_values)
                if average_current_values else None
            )
            average_high = (
                sum(average_high_values) / len(average_high_values)
                if average_high_values else None
            )
            average_low = (
                sum(average_low_values) / len(average_low_values)
                if average_low_values else None
            )

            average_vs_high = (
                ((average_current - average_high) / average_high) * 100
                if average_current is not None and average_high not in (None, 0)
                else None
            )
            average_vs_low = (
                ((average_current - average_low) / average_low) * 100
                if average_current is not None and average_low not in (None, 0)
                else None
            )

            # Calculate absolute arithmetic averages across ALL uploaded indices.
            all_h2l_values = [
                value for value in (to_number(record.get("H2L%")) for record in display_records)
                if value is not None
            ]
            all_l2h_values = [
                value for value in (to_number(record.get("L2H%")) for record in display_records)
                if value is not None
            ]
            average_h2l = sum(all_h2l_values) / len(all_h2l_values) if all_h2l_values else None
            average_l2h = sum(all_l2h_values) / len(all_l2h_values) if all_l2h_values else None
            absolute_average_h2l = abs(average_h2l) if average_h2l is not None else None
            absolute_average_l2h = abs(average_l2h) if average_l2h is not None else None

            if (
                absolute_average_h2l is not None
                and absolute_average_l2h is not None
                and absolute_average_h2l > absolute_average_l2h
            ):
                dominant_market_side = "H2L"
            elif (
                absolute_average_h2l is not None
                and absolute_average_l2h is not None
                and absolute_average_l2h > absolute_average_h2l
            ):
                dominant_market_side = "L2H"
            else:
                dominant_market_side = None

            # Automatically identify the index with the lowest H2L%
            # among indices whose CURRENT value is above 10,000.
            eligible_indices = []
            highest_l2h_indices = []
            for record in display_records:
                current_value = to_number(record.get("CURRENT"))
                h2l_value = to_number(record.get("H2L%"))
                l2h_value = to_number(record.get("L2H%"))
                if (
                    current_value is not None
                    and current_value > 10000
                    and h2l_value is not None
                ):
                    eligible_indices.append((h2l_value, record))
                if (
                    current_value is not None
                    and current_value > 10000
                    and l2h_value is not None
                ):
                    highest_l2h_indices.append((l2h_value, record))

            eligible_indices.sort(key=lambda item: item[0])
            highest_l2h_indices.sort(key=lambda item: item[0], reverse=True)

            st.divider()
            st.subheader("Priority Stock Analysis")
            avg_h2l_col, avg_l2h_col, market_bias_col = st.columns(3)
            avg_h2l_col.metric(
                "| Average H2L% |",
                f"{absolute_average_h2l:.2f}%" if absolute_average_h2l is not None else "N/A"
            )
            avg_l2h_col.metric(
                "| Average L2H% |",
                f"{absolute_average_l2h:.2f}%" if absolute_average_l2h is not None else "N/A"
            )
            market_bias_col.metric("Dominant Side", dominant_market_side or "Balanced / N/A")
            st.caption("Absolute averages use all valid H2L% and L2H% values in the uploaded index file.")
            st.caption(
                "Your key watchlist comes first: H2L Top 4 gainers and L2H Bottom 4 "
                "stocks, with Master Excel metrics and panel-specific ranks. "
                "Full constituent and quote lists remain available below."
            )

            st.markdown("""
            <style>
            .st-key-h2l_panel {
                background: linear-gradient(145deg, #f3fbf5 0%, #e8f5eb 100%);
                border: 1px solid #b8d9c0;
                border-left: 5px solid #25834a;
                border-radius: 16px;
                padding: 1.25rem 1.35rem;
                margin: 0.8rem 0 1.2rem 0;
                box-shadow: 0 5px 18px rgba(27, 83, 48, 0.06);
            }
            .st-key-l2h_panel {
                background: linear-gradient(145deg, #fff9f1 0%, #fff0df 100%);
                border: 1px solid #efd1ad;
                border-left: 5px solid #c87925;
                border-radius: 16px;
                padding: 1.25rem 1.35rem;
                margin: 0.8rem 0 1.2rem 0;
                box-shadow: 0 5px 18px rgba(126, 76, 24, 0.06);
            }
            .st-key-h2l_panel [data-testid="stMetric"],
            .st-key-l2h_panel [data-testid="stMetric"] {
                background: rgba(255,255,255,0.78);
                border: 1px solid rgba(100,110,100,0.13);
                border-radius: 12px;
                padding: 0.75rem 0.9rem;
            }
            .st-key-h2l_panel [data-testid="stDataFrame"],
            .st-key-l2h_panel [data-testid="stDataFrame"] {
                border: 1px solid rgba(100,110,100,0.16);
                border-radius: 11px;
                overflow: hidden;
            }
            .st-key-h2l_panel button,
            .st-key-l2h_panel button {
                border-radius: 9px;
                font-weight: 600;
            }
            </style>
            """, unsafe_allow_html=True)
            with st.container(key="h2l_panel"):
                if eligible_indices:
                    selected_h2l, selected_index = eligible_indices[0]
                    selected_name = str(selected_index.get("INDEX", "")).strip().upper()
                    selected_current = to_number(selected_index.get("CURRENT"))
                    selected_high = to_number(selected_index.get("HIGH"))
                    selected_low = to_number(selected_index.get("LOW"))
    
                    select_col1, select_col2, select_col3, select_col4 = st.columns(4)
                    select_col1.metric("Selected INDEX", selected_name)
                    select_col2.metric(
                        "CURRENT",
                        f"{selected_current:,.2f}" if selected_current is not None else "N/A"
                    )
                    select_col3.metric(
                        "H2L%",
                        f"{selected_h2l:+.2f}%"
                    )
                    select_col4.metric(
                        "HIGH",
                        f"{selected_high:,.2f}" if selected_high is not None else "N/A"
                    )
    
                    # Fetch constituents for whichever index the uploaded file selects.
                    st.markdown(f"**{selected_name} — Constituent Stocks**")
                    import requests
                    import pandas as pd
                    import yfinance as yf
                    from datetime import datetime
                    from zoneinfo import ZoneInfo
    
                    @st.cache_data(ttl=3600, show_spinner=False)
                    def get_index_constituents(index_name):
                        """
                        Load current constituents for an index using a layered
                        fallback: India API archive, NSE's public index API,
                        then official Nifty Indices CSV files.
                        """
                        import io
                        import requests
                        import pandas as pd

                        def clean_name(value):
                            return " ".join(
                                re.sub(r"[^A-Z0-9]+", " ", str(value).upper()).split()
                            )

                        def build_members(rows):
                            members = []

                            def first_valid(row, keys, default=""):
                                for key in keys:
                                    value = row.get(key)
                                    if value is not None and not pd.isna(value):
                                        return value
                                return default

                            for row in rows:
                                raw_symbol = first_valid(
                                    row, ["symbol", "Symbol", "SYMBOL"]
                                )
                                symbol = str(raw_symbol).strip()
                                if (
                                    not symbol
                                    or symbol.lower() == "nan"
                                    or clean_name(symbol) == clean_name(index_name)
                                ):
                                    continue
                                meta = row.get("meta") or {}
                                company = first_valid(
                                    row,
                                    ["company", "Company Name", "Company"],
                                    first_valid(meta, ["companyName", "company"])
                                )
                                industry = first_valid(
                                    row,
                                    ["industry", "Industry"],
                                    first_valid(meta, ["industry"])
                                )
                                members.append({
                                    "No.": len(members) + 1,
                                    "NSE Symbol": symbol,
                                    "Company": str(company).strip(),
                                    "Industry": str(industry).strip()
                                })
                            return members

                        wanted_name = clean_name(index_name)
                        errors = []

                        # Source 1: India API archive (fast and already used by the app).
                        try:
                            api_headers = {"User-Agent": "Mozilla/5.0"}
                            base_url = "https://api.dharun.dev/v1/indices"
                            response = requests.get(
                                base_url, headers=api_headers, timeout=20
                            )
                            response.raise_for_status()
                            index_items = response.json().get("data", [])
                            matched = next(
                                (
                                    item for item in index_items
                                    if clean_name(item.get("name", "")) == wanted_name
                                ),
                                None
                            )
                            if matched is None:
                                aliases = {
                                    "NIFTY IT": "NIFTY IT",
                                    "NIFTY 50": "NIFTY 50",
                                    "NIFTY BANK": "NIFTY BANK",
                                    "NIFTY NEXT 50": "NIFTY NEXT 50",
                                    "NIFTY 100": "NIFTY 100",
                                    "NIFTY 200": "NIFTY 200",
                                    "NIFTY 500": "NIFTY 500",
                                    "NIFTY MIDCAP 50": "NIFTY MIDCAP 50",
                                    "NIFTY MIDCAP 100": "NIFTY MIDCAP 100",
                                    "NIFTY MIDCAP 150": "NIFTY MIDCAP 150",
                                    "NIFTY SMALLCAP 50": "NIFTY SMALLCAP 50",
                                    "NIFTY SMALLCAP 100": "NIFTY SMALLCAP 100",
                                    "NIFTY SMALLCAP 250": "NIFTY SMALLCAP 250"
                                }
                                alias_name = aliases.get(wanted_name)
                                if alias_name:
                                    matched = next(
                                        (
                                            item for item in index_items
                                            if clean_name(item.get("name", "")) ==
                                            clean_name(alias_name)
                                        ),
                                        None
                                    )
                            if matched is not None:
                                detail_response = requests.get(
                                    f"{base_url}/{matched['id']}",
                                    headers=api_headers,
                                    timeout=20
                                )
                                detail_response.raise_for_status()
                                payload = detail_response.json().get("data", {})
                                members = build_members(payload.get("members", []))
                                if members:
                                    return members
                                errors.append("India API returned no constituent rows.")
                            else:
                                errors.append("Index is not listed in the India API archive.")
                        except Exception as api_error:
                            errors.append(f"India API: {api_error}")

                        # Source 2: NSE public index constituents endpoint.
                        try:
                            session = requests.Session()
                            nse_headers = {
                                "User-Agent": (
                                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                                    "Chrome/131.0.0.0 Safari/537.36"
                                ),
                                "Accept": "application/json, text/plain, */*",
                                "Accept-Language": "en-US,en;q=0.9",
                                "Referer": "https://www.nseindia.com/"
                            }
                            session.get(
                                "https://www.nseindia.com/",
                                headers=nse_headers,
                                timeout=15
                            )
                            nse_name = " ".join(str(index_name).upper().split())
                            nse_response = session.get(
                                "https://www.nseindia.com/api/equity-stockIndices",
                                params={"index": nse_name},
                                headers=nse_headers,
                                timeout=20
                            )
                            nse_response.raise_for_status()
                            nse_rows = nse_response.json().get("data", [])
                            members = build_members(nse_rows)
                            if members:
                                return members
                            errors.append("NSE API returned no constituent rows.")
                        except Exception as nse_error:
                            errors.append(f"NSE API: {nse_error}")

                        # Source 3: official NSE Indices constituent CSV.
                        # Explicit filenames cover indices whose official naming
                        # does not follow a simple slug convention.
                        official_files = {
                            "NIFTY 50": "ind_nifty50list.csv",
                            "NIFTY NEXT 50": "ind_niftynext50list.csv",
                            "NIFTY NEXT 100": "ind_niftynext100list.csv",
                            "NIFTY 100": "ind_nifty100list.csv",
                            "NIFTY 200": "ind_nifty200list.csv",
                            "NIFTY 500": "ind_nifty500list.csv",
                            "NIFTY TOTAL MARKET": "ind_niftytotalmarket_list.csv",
                            "NIFTY MIDCAP 50": "ind_niftymidcap50list.csv",
                            "NIFTY MIDCAP 100": "ind_niftymidcap100list.csv",
                            "NIFTY MIDCAP 150": "ind_niftymidcap150list.csv",
                            "NIFTY MIDCAP SELECT": "ind_niftymidcapselect_list.csv",
                            "NIFTY SMALLCAP 50": "ind_niftysmallcap50list.csv",
                            "NIFTY SMALLCAP 100": "ind_niftysmallcap100list.csv",
                            "NIFTY SMALLCAP 250": "ind_niftysmallcap250list.csv",
                            "NIFTY MICROCAP 250": "ind_niftymicrocap250_list.csv",
                            "NIFTY LARGE MIDCAP 250": "ind_niftylargemidcap250list.csv",
                            "NIFTY MIDSMALLCAP 400": "ind_niftymidsmallcap400list.csv",
                            "NIFTY INDIA FPI 150": "ind_niftyIndiaFPI150_list.csv",
                            "NIFTY BANK": "ind_niftybanklist.csv",
                            "NIFTY IT": "ind_niftyitlist.csv",
                            "NIFTY PHARMA": "ind_niftypharmalist.csv",
                            "NIFTY FMCG": "ind_niftyfmcglist.csv",
                            "NIFTY AUTO": "ind_niftyautolist.csv",
                            "NIFTY CHEMICALS": "ind_niftyChemicals_list.csv",
                            "NIFTY METAL": "ind_niftymetallist.csv",
                            "NIFTY REALTY": "ind_niftyrealtylist.csv",
                            "NIFTY PSU BANK": "ind_niftypsubanklist.csv",
                            "NIFTY PRIVATE BANK": "ind_nifty_privatebanklist.csv",
                            "NIFTY FINANCIAL SERVICES": "ind_niftyfinancelist.csv",
                            "NIFTY CONSUMER DURABLES": "ind_niftyconsumerdurableslist.csv",
                            "NIFTY OIL AND GAS": "ind_niftyoilgaslist.csv",
                            "NIFTY OIL GAS": "ind_niftyoilgaslist.csv",
                            "NIFTY HEALTHCARE": "ind_niftyhealthcarelist.csv",
                            "NIFTY HOSPITALS": "ind_niftyhospitalslist.csv",
                            "NIFTY NBFC": "ind_niftynbfclist.csv",
                            "NIFTY INSURANCE": "ind_niftyinsurancelist.csv",
                            "NIFTY POWER": "ind_niftypowerlist.csv",
                            "NIFTY MEDIA": "ind_niftymedialist.csv",
                            "NIFTY500 HEALTHCARE": "ind_nifty500Healthcare_list.csv",
                            "NIFTY MIDSMALL FINANCIAL SERVICES": "ind_niftymidsmallfinancailservice_list.csv",
                            "NIFTY MIDSMALL HEALTHCARE": "ind_niftymidsmallhealthcare_list.csv",
                            "NIFTY MIDSMALL IT TELECOM": "ind_niftymidsmallitAndtelecom_list.csv",
                            "NIFTY100 EQUAL WEIGHT": "ind_nifty100list.csv",
                            "NIFTY100 LOW VOLATILITY 30": "ind_Nifty100LowVolatility30list.csv",
                            "NIFTY200 MOMENTUM 30": "ind_nifty200Momentum30_list.csv",
                            "NIFTY200 ALPHA 30": "ind_nifty200alpha30_list.csv",
                            "NIFTY100 ALPHA 30": "ind_nifty100Alpha30list.csv",
                            "NIFTY DIVIDEND OPPORTUNITIES 50": "ind_niftydivopp50list.csv",
                            "NIFTY GROWTH SECTORS 15": "ind_NiftyGrowth_Sectors15_Index.csv",
                            "NIFTY HIGH BETA 50": "nifty_High_Beta50_Index.csv",
                            "NIFTY LOW VOLATILITY 50": "nifty_low_Volatility50_Index.csv",
                            "NIFTY TOP 10 EQUAL WEIGHT": "ind_niftytop10EqualWeight_list.csv",
                            "NIFTY TOP 15 EQUAL WEIGHT": "ind_niftytop15EqualWeight_list.csv",
                            "NIFTY TOP 20 EQUAL WEIGHT": "ind_niftyTop20EqualWeight_list.csv",
                            "NIFTY100 QUALITY 30": "ind_nifty100Quality30list.csv",
                            "NIFTY MIDCAP150 MOMENTUM 50": "ind_niftymidcap150momentum50_list.csv",
                            "NIFTY MIDCAP150 QUALITY 50": "ind_niftymidcap150quality50list.csv",
                            "NIFTY ALPHA 50": "ind_nifty_Alpha_Index.csv",
                            "NIFTY50 EQUAL WEIGHT": "ind_Nifty50EqualWeight.csv",
                            "NIFTY50 VALUE 20": "ind_Nifty50_Value20.csv",
                            "NIFTY200 VALUE 30": "ind_nifty200Value30_list.csv",
                            "NIFTY500 FLEXICAP QUALITY 30": "ind_nifty500FlexicapQuality30_list.csv",
                            "NIFTY500 VALUE 50": "ind_nifty500Value50_list.csv",
                            "NIFTY500 EQUAL WEIGHT": "ind_nifty500EqualWeight_list.csv",
                            "NIFTY500 LOW VOLATILITY 50": "ind_nifty500LowVolatility50_list.csv",
                            "NIFTY500 MOMENTUM 50": "ind_nifty500Momentum50_list.csv",
                            "NIFTY500 QUALITY 50": "ind_nifty500Quality50_list.csv",
                            "NIFTY500 MULTIFACTOR MQVLV 50": "ind_nifty500MultifactorMQVLv50_list.csv",
                            "NIFTY200 QUALITY 30": "ind_nifty200Quality30_list.csv",
                            "NIFTY MIDSMALLCAP400 MOMENTUM QUALITY 100": "ind_niftyMidSmallcap400MomentumQuality100_list.csv",
                            "NIFTY SMALLCAP250 QUALITY 50": "ind_niftySmallcap250_Quality50_list.csv",
                            "NIFTY TOTAL MARKET MOMENTUM QUALITY 50": "ind_niftyTotalMarketMomentumQuality50_list.csv",
                            "NIFTY CAPITAL MARKETS": "ind_niftyCapitalMarkets_list.csv",
                            "NIFTY COMMODITIES": "ind_niftycommoditieslist.csv",
                            "NIFTY CONGLOMERATE 50": "ind_niftyConglomerate50_list.csv",
                            "NIFTY CORE HOUSING": "ind_niftyCoreHousing_list.csv",
                            "NIFTY CPSE": "ind_niftycpselist.csv",
                            "NIFTY EV NEW AGE AUTOMOTIVE": "ind_niftyEv_NewAgeAutomotive_list.csv",
                            "NIFTY ENERGY": "ind_niftyenergylist.csv",
                            "NIFTY HOUSING": "ind_niftyhousing_list.csv",
                            "NIFTY INDIA CONSUMPTION": "ind_niftyconsumptionlist.csv",
                            "NIFTY INDIA DEFENCE": "ind_niftyindiadefence_list.csv",
                            "NIFTY INDIA DIGITAL": "ind_niftyindiadigital_list.csv",
                            "NIFTY INDIA INFRASTRUCTURE LOGISTICS": "ind_niftyIndiaInfrastructure_Logistics_list.csv",
                            "NIFTY INDIA INTERNET": "ind_niftyIndiaInternet_list.csv",
                            "NIFTY INDIA MANUFACTURING": "ind_niftyindiamanufacturing_list.csv",
                            "NIFTY INDIA NEW AGE CONSUMPTION": "ind_niftyIndiaNewAgeConsumption_list.csv",
                            "NIFTY INDIA RAILWAYS PSU": "ind_niftyIndiaRailwaysPSU_list.csv",
                            "NIFTY INDIA TOURISM": "ind_niftyindiatourism_list.csv",
                            "NIFTY INFRASTRUCTURE": "ind_niftyinfralist.csv",
                            "NIFTY IPO": "ind_niftyIPO_list.csv",
                            "NIFTY MIDCAP LIQUID 15": "ind_Nifty_Midcap_Liquid15.csv",
                            "NIFTY MNC": "ind_niftymnclist.csv",
                            "NIFTY MOBILITY": "ind_niftymobility_list.csv",
                            "NIFTY PSE": "ind_niftypselist.csv",
                            "NIFTY REITS INVITS": "ind_niftyREITs_InvITs_list.csv",
                            "NIFTY RURAL": "ind_niftyRural_list.csv",
                            "NIFTY NON CYCLICAL CONSUMER": "ind_niftynon-cyclicalconsumer_list.csv",
                            "NIFTY SERVICES SECTOR": "ind_niftyservicelist.csv",
                            "NIFTY100 LIQUID 15": "ind_Nifty100_Liquid15.csv",
                            "NIFTY SME EMERGE": "ind_niftysmelist.csv",
                            "NIFTY WAVES": "ind_niftyWaves_list.csv",
                            "NIFTY INDIA CORPORATE GROUP ADITYA BIRLA GROUP": "ind_nifty_adityabirlalist.csv",
                            "NIFTY INDIA CORPORATE GROUP MAHINDRA GROUP": "ind_nifty_mahindralist.csv",
                            "NIFTY INDIA CORPORATE GROUP TATA GROUP": "ind_nifty_tatalist.csv",
                            "NIFTY INDIA SELECT 5 CORPORATE GROUPS MAATR": "ind_niftyIndiaSelect5CorporateGroupsMAATR_list.csv"
                        }
                        csv_filename = official_files.get(wanted_name)
                        if not csv_filename:
                            # Try common official filename patterns for other
                            # Nifty indices, including archive-host variants.
                            slug = re.sub(r"[^a-z0-9]+", "", wanted_name.lower())
                            slug = re.sub(r"^nifty", "nifty", slug)
                            csv_candidates = [
                                f"ind_{slug}list.csv",
                                f"ind_{slug}_list.csv"
                            ]
                        else:
                            csv_candidates = [csv_filename]

                        csv_hosts = [
                            "https://www.niftyindices.com/IndexConstituent/",
                            "https://niftyindices.com/IndexConstituent/",
                            "https://nsearchives.nseindia.com/content/indices/"
                        ]
                        csv_headers = {
                            "User-Agent": (
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/131.0.0.0 Safari/537.36"
                            ),
                            "Accept": "text/csv,application/csv,*/*;q=0.8",
                            "Referer": "https://www.niftyindices.com/"
                        }
                        for host in csv_hosts:
                            for filename in csv_candidates:
                                try:
                                    csv_response = requests.get(
                                        host + filename,
                                        headers=csv_headers,
                                        timeout=20
                                    )
                                    csv_response.raise_for_status()
                                    frame = pd.read_csv(
                                        io.BytesIO(csv_response.content)
                                    )
                                    frame.columns = [
                                        str(column).replace("\ufeff", "").strip()
                                        for column in frame.columns
                                    ]
                                    symbol_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) in
                                            {"SYMBOL", "NSE SYMBOL", "TRADING SYMBOL"}
                                        ),
                                        None
                                    )
                                    if symbol_column is None or frame.empty:
                                        continue
                                    company_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) in
                                            {"COMPANY NAME", "COMPANY", "NAME"}
                                        ),
                                        None
                                    )
                                    industry_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) == "INDUSTRY"
                                        ),
                                        None
                                    )
                                    csv_rows = []
                                    for _, csv_row in frame.iterrows():
                                        csv_rows.append({
                                            "symbol": csv_row.get(symbol_column, ""),
                                            "company": (
                                                csv_row.get(company_column, "")
                                                if company_column else ""
                                            ),
                                            "industry": (
                                                csv_row.get(industry_column, "")
                                                if industry_column else ""
                                            )
                                        })
                                    members = build_members(csv_rows)
                                    if members:
                                        return members
                                except Exception as csv_error:
                                    errors.append(
                                        f"{filename} from {host}: {csv_error}"
                                    )

                        raise ValueError(
                            f"Could not load constituents for {index_name} from "
                            "the available India API, NSE API, or official NSE Indices "
                            "CSV sources. " + " | ".join(errors[-4:])
                        )

                    try:
                        with st.spinner(
                            f"Loading NSE constituents for {selected_name}..."
                        ):
                            selected_constituents = get_index_constituents(
                                selected_name
                            )
    
                        with st.expander(
                            f"View all {len(selected_constituents)} H2L index constituents",
                            expanded=False
                        ):
                            st.dataframe(
                                selected_constituents,
                                use_container_width=True,
                                hide_index=True
                            )
                            st.caption(
                                "Constituent list sourced from NSE Indices published data "
                                "through the India API index archive. Membership lists may "
                                "update periodically."
                            )
    
                        st.markdown(f"**{selected_name} — Today's Stock Details**")
                        market_timezone = ZoneInfo("Asia/Kolkata")
                        market_date = datetime.now(market_timezone).date()
                        st.caption(
                            f"Market date: {market_date.strftime('%d %b %Y')} (India). "
                            "Quotes from Yahoo Finance may be delayed."
                        )
    
                        st.caption(
                            "Stock details refresh automatically every 60 seconds while "
                            "the dashboard is open. Use the button for an immediate refresh."
                        )
    
                        def render_today_quotes(
                            constituents,
                            index_name,
                            refresh_key,
                            selection_mode,
                            trade_signal_placeholder
                        ):
                            quote_symbols = [
                                item["NSE Symbol"] for item in constituents
                            ]
                            refresh_quotes = st.button(
                                "Refresh Today's Stock Details",
                                key=refresh_key
                            )
    
                            @st.cache_data(ttl=60, show_spinner=False)
                            def get_dynamic_index_quotes(symbols, requested_market_date):
                                from concurrent.futures import ThreadPoolExecutor, as_completed
    
                                # The caller passes an ISO date string; convert it to
                                # a date object before comparing with pandas index dates.
                                requested_date = datetime.fromisoformat(
                                    requested_market_date
                                ).date()
    
                                def get_one_symbol(symbol):
                                    ticker = f"{symbol}.NS"
                                    row = {
                                        "NSE Symbol": symbol,
                                        "Price": None,
                                        "Open": None,
                                        "High": None,
                                        "Low": None,
                                        "% Chg": None
                                    }
    
                                    try:
                                        stock = yf.Ticker(ticker)
                                        intraday = stock.history(
                                            period="5d",
                                            interval="5m",
                                            auto_adjust=False,
                                            actions=False,
                                            prepost=False
                                        )
                                        daily = stock.history(
                                            period="10d",
                                            interval="1d",
                                            auto_adjust=False,
                                            actions=False
                                        )
    
                                        if intraday is None or intraday.empty:
                                            return row, None
    
                                        intraday = intraday.dropna(how="all").copy()
                                        if intraday.index.tz is None:
                                            intraday.index = intraday.index.tz_localize(
                                                "Asia/Kolkata"
                                            )
                                        else:
                                            intraday.index = intraday.index.tz_convert(
                                                "Asia/Kolkata"
                                            )
    
                                        today = intraday[
                                            intraday.index.date == requested_date
                                        ].dropna(how="all")
    
                                        if today.empty:
                                            return row, None
    
                                        opens = today["Open"].dropna()
                                        highs = today["High"].dropna()
                                        lows = today["Low"].dropna()
                                        closes = today["Close"].dropna()
    
                                        if (
                                            opens.empty or highs.empty
                                            or lows.empty or closes.empty
                                        ):
                                            return row, None
    
                                        previous_close = None
                                        if daily is not None and not daily.empty:
                                            daily = daily.dropna(how="all").copy()
                                            if daily.index.tz is None:
                                                daily.index = daily.index.tz_localize(
                                                    "Asia/Kolkata"
                                                )
                                            else:
                                                daily.index = daily.index.tz_convert(
                                                    "Asia/Kolkata"
                                                )
                                            previous = daily[
                                                daily.index.date < requested_date
                                            ]["Close"].dropna()
                                            if not previous.empty:
                                                previous_close = float(previous.iloc[-1])
    
                                        price = float(closes.iloc[-1])
                                        row.update({
                                            "Price": price,
                                            "Open": float(opens.iloc[0]),
                                            "High": float(highs.max()),
                                            "Low": float(lows.min()),
                                            "% Chg": (
                                                ((price - previous_close) / previous_close) * 100
                                                if previous_close not in (None, 0) else None
                                            )
                                        })
                                        return row, closes.index[-1]
    
                                    except Exception:
                                        # Keep other constituents visible if one ticker
                                        # is unavailable or Yahoo temporarily rate-limits it.
                                        return row, None
    
                                rows = []
                                latest_quote_time = None
                                max_workers = min(8, max(1, len(symbols)))
    
                                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                                    futures = {
                                        executor.submit(get_one_symbol, symbol): symbol
                                        for symbol in symbols
                                    }
                                    results = {}
                                    for future in as_completed(futures):
                                        symbol = futures[future]
                                        try:
                                            results[symbol] = future.result()
                                        except Exception:
                                            results[symbol] = ({
                                                "NSE Symbol": symbol,
                                                "Price": None,
                                                "Open": None,
                                                "High": None,
                                                "Low": None,
                                                "% Chg": None
                                            }, None)
    
                                for symbol in symbols:
                                    row, quote_time = results[symbol]
                                    rows.append(row)
                                    if quote_time is not None and (
                                        latest_quote_time is None
                                        or quote_time > latest_quote_time
                                    ):
                                        latest_quote_time = quote_time
    
                                timestamp = (
                                    latest_quote_time.strftime("%d %b %Y, %I:%M %p IST")
                                    if latest_quote_time is not None else None
                                )
                                return rows, timestamp
    
                            if refresh_quotes:
                                get_dynamic_index_quotes.clear()
    
                            with st.spinner(
                                f"Loading today's {index_name} stock quotes..."
                            ):
                                quote_rows, quote_timestamp = get_dynamic_index_quotes(
                                    tuple(quote_symbols),
                                    market_date.isoformat()
                                )
    
                            if quote_timestamp:
                                st.caption(f"Latest quote timestamp: {quote_timestamp}")
                            else:
                                st.warning(
                                    "Today's intraday quotes have not been returned by "
                                    "the market-data provider. The table will not substitute "
                                    "yesterday's prices."
                                )
    
                            quote_df = pd.DataFrame(quote_rows)
                            for numeric_column in [
                                "Price", "Open", "High", "Low", "% Chg"
                            ]:
                                quote_df[numeric_column] = quote_df[numeric_column].map(
                                    lambda value: round(value, 2)
                                    if pd.notna(value) else None
                                )
    
                            with st.expander(
                                f"View all {len(quote_df)} live quote rows",
                                expanded=False
                            ):
                                st.dataframe(
                                    quote_df,
                                    use_container_width=True,
                                    hide_index=True,
                                    column_config={
                                        "NSE Symbol": st.column_config.TextColumn(
                                            "NSE Symbol",
                                            width="small"
                                        ),
                                        "Price": st.column_config.NumberColumn(
                                            "Price",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "Open": st.column_config.NumberColumn(
                                            "Open",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "High": st.column_config.NumberColumn(
                                            "High",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "Low": st.column_config.NumberColumn(
                                            "Low",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "% Chg": st.column_config.NumberColumn(
                                            "% Chg",
                                            format="%.2f%%",
                                            width="small"
                                        )
                                    }
                                )
    
                            available_count = int(quote_df["Price"].notna().sum())
                            st.caption(
                                f"Today's quotes available for {available_count} "
                                f"of {len(quote_df)} constituent stocks."
                            )
    
                            # Select the top or bottom four stocks by today's % Chg.
                            selected_four = (
                                quote_df.dropna(subset=["% Chg"])
                                .sort_values(
                                    "% Chg",
                                    ascending=(selection_mode == "bottom")
                                )
                                .head(4)
                                .copy()
                            )
    
                            selection_title = (
                                "Bottom 4 Stocks by % Chg"
                                if selection_mode == "bottom"
                                else "Top 4 Stocks by % Chg"
                            )
                            st.markdown(
                                f"#### {selection_title} · {index_name}"
                            )
                            st.caption(
                                "Top 3 / Bottom 3 are trade-eligible. The 4th stock is "
                                "confirmation-only: if its rank is best or tied-best, wait. "
                                "Master Excel values and ranks are shown for all four."
                            )
                            if selected_four.empty:
                                st.info("No stocks have a valid % Chg value yet.")
                            else:
                                with st.spinner("Matching selected stocks with Supabase Master Excel..."):
                                    master_bytes_for_selection = get_master()
                                    rank_source_column = (
                                        "Sum O2L.10"
                                        if selection_mode == "top"
                                        else "Sum O2H.10"
                                    )
                                    master_values, master_error = get_top_stocks_master_values(
                                        selected_four,
                                        master_bytes_for_selection,
                                        rank_source_column
                                    )
    
                                # Only the first three % Chg stocks on each side
                                # can qualify for a trade. The fourth is confirmation-only:
                                # if its rank is the best (or tied best), the setup must wait.
                                signal_side = "H2L" if selection_mode == "top" else "L2H"
                                rank_column = f"{rank_source_column} Rank"
                                quote_price_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("Price"))
                                    for _, row in quote_df.iterrows()
                                }
                                quote_high_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("High"))
                                    for _, row in quote_df.iterrows()
                                }
                                quote_low_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("Low"))
                                    for _, row in quote_df.iterrows()
                                }
                                side_candidates = []
                                if not master_error and master_values:
                                    for stock_position, item in enumerate(master_values):
                                        symbol = str(item.get("NSE Symbol", "")).strip()
                                        rank_value = to_number(item.get(rank_column))
                                        current_price = quote_price_by_symbol.get(symbol.upper())
                                        current_high = quote_high_by_symbol.get(symbol.upper())
                                        current_low = quote_low_by_symbol.get(symbol.upper())
                                        if (
                                            symbol
                                            and rank_value is not None
                                            and current_price is not None
                                        ):
                                            side_candidates.append({
                                                "side": signal_side,
                                                "symbol": symbol,
                                                "rank": rank_value,
                                                "price": current_price,
                                                "high": current_high,
                                                "low": current_low,
                                                "confirmation_only": stock_position >= 3
                                            })

                                st.session_state[f"trade_candidates_{signal_side}"] = side_candidates

                                h2l_candidates = st.session_state.get(
                                    "trade_candidates_H2L", []
                                )
                                l2h_candidates = st.session_state.get(
                                    "trade_candidates_L2H", []
                                )
                                # Include each fourth stock only as a blocker in the
                                # rank comparison; it can never become the signal stock.
                                all_ranked_candidates = h2l_candidates + l2h_candidates

                                winning_candidate = None
                                if all_ranked_candidates:
                                    lowest_rank_value = min(
                                        item["rank"] for item in all_ranked_candidates
                                    )
                                    lowest_rank_candidates = [
                                        item for item in all_ranked_candidates
                                        if item["rank"] == lowest_rank_value
                                    ]
                                    confirmation_blocks_signal = any(
                                        item.get("confirmation_only", False)
                                        for item in lowest_rank_candidates
                                    )
                                    # Wait if either fourth stock has the best rank,
                                    # including a tie at the best rank.
                                    # Also avoid choosing arbitrarily when eligible
                                    # H2L and L2H stocks share the best rank.
                                    winning_sides = {
                                        item["side"] for item in lowest_rank_candidates
                                        if not item.get("confirmation_only", False)
                                    }
                                    if not confirmation_blocks_signal and len(winning_sides) == 1:
                                        eligible_winners = [
                                            item for item in lowest_rank_candidates
                                            if not item.get("confirmation_only", False)
                                        ]
                                        if len(eligible_winners) == 1:
                                            winning_candidate = eligible_winners[0]

                                # Price-location confirmation:
                                # BUY only when current price is at least 0.20% below
                                # today's High; SELL only when current price is at
                                # least 0.20% above today's Low.
                                price_location_confirmed = False
                                if winning_candidate is not None:
                                    signal_price = winning_candidate.get("price")
                                    signal_high = winning_candidate.get("high")
                                    signal_low = winning_candidate.get("low")
                                    if signal_side == "H2L":
                                        price_location_confirmed = (
                                            signal_price is not None
                                            and signal_high is not None
                                            and signal_high > 0
                                            and ((signal_high - signal_price) / signal_high) >= 0.002
                                        )
                                    else:
                                        price_location_confirmed = (
                                            signal_price is not None
                                            and signal_low is not None
                                            and signal_low > 0
                                            and ((signal_price - signal_low) / signal_low) >= 0.002
                                        )

                                if (
                                    winning_candidate is not None
                                    and winning_candidate["side"] == signal_side
                                    and dominant_market_side == signal_side
                                    and price_location_confirmed
                                ):
                                    signal_symbol = winning_candidate["symbol"]
                                    entry_price = winning_candidate["price"]
                                    if signal_side == "H2L":
                                        trade_action = "BUY"
                                        target_price = entry_price * 1.0035
                                        signal_class = "buy-signal"
                                        signal_label = "H2L BUY SETUP"
                                        entry_label = "Buy Entry"
                                        target_label = "Target (+0.35%)"
                                    else:
                                        trade_action = "SELL"
                                        target_price = entry_price * 0.9965
                                        signal_class = "sell-signal"
                                        signal_label = "L2H SELL SETUP"
                                        entry_label = "Sell Entry"
                                        target_label = "Target (-0.35%)"

                                    trade_signal_placeholder.markdown(
                                        f"""
                                        <style>
                                        @keyframes prioritySignalPulse {{
                                            0%, 100% {{ transform: scale(1); filter: brightness(1); }}
                                            50% {{ transform: scale(1.008); filter: brightness(1.12); }}
                                        }}
                                        .priority-trade-signal {{
                                            animation: prioritySignalPulse 1.3s ease-in-out infinite;
                                            border-radius: 14px;
                                            padding: 1rem 1.15rem;
                                            margin: 0.7rem 0 1rem 0;
                                            color: white;
                                        }}
                                        .priority-trade-signal.buy-signal {{
                                            background: linear-gradient(115deg, #176b3a, #269653);
                                            border: 1px solid #58b77a;
                                        }}
                                        .priority-trade-signal.sell-signal {{
                                            background: linear-gradient(115deg, #a65312, #d17a27);
                                            border: 1px solid #edaa65;
                                        }}
                                        .priority-trade-signal .signal-title {{
                                            font-size: 1.05rem; font-weight: 800;
                                            letter-spacing: 0.04em; margin-bottom: 0.4rem;
                                        }}
                                        .priority-trade-signal .signal-symbol {{
                                            font-size: 1.4rem; font-weight: 800; margin-bottom: 0.65rem;
                                        }}
                                        .priority-trade-signal .signal-grid {{
                                            display: flex; flex-wrap: wrap; gap: 1.5rem;
                                        }}
                                        .priority-trade-signal .signal-label {{
                                            font-size: 0.78rem; opacity: 0.9;
                                        }}
                                        .priority-trade-signal .signal-value {{
                                            font-size: 1.1rem; font-weight: 750;
                                        }}
                                        </style>
                                        <div class="priority-trade-signal {signal_class}">
                                            <div class="signal-title">{signal_label} · {trade_action}</div>
                                            <div class="signal-symbol">{signal_symbol}</div>
                                            <div class="signal-grid">
                                                <div><div class="signal-label">{entry_label} (Current Price)</div>
                                                <div class="signal-value">₹{entry_price:,.2f}</div></div>
                                                <div><div class="signal-label">{target_label}</div>
                                                <div class="signal-value">₹{target_price:,.2f}</div></div>
                                                <div><div class="signal-label">Lowest Rank Across Both Lists</div>
                                                <div class="signal-value">{winning_candidate["rank"]:g}</div></div>
                                            </div>
                                        </div>
                                        """,
                                        unsafe_allow_html=True
                                    )
                                elif (
                                    winning_candidate is None
                                    and len(all_ranked_candidates) > 0
                                    and len({
                                        item["side"] for item in all_ranked_candidates
                                        if item["rank"] == min(
                                            candidate["rank"] for candidate in all_ranked_candidates
                                        )
                                    }) > 1
                                ):
                                    trade_signal_placeholder.info(
                                        "The lowest rank is tied between H2L and L2H lists. "
                                        "No trade signal is shown."
                                    )
                                # Do not let the other panel's independent 60-second
                                # fragment overwrite the shared signal card. Only the
                                # fragment for the winning side may render the signal.

                                if master_error:
                                    st.warning(master_error)
                                elif master_values:
                                    master_values_df = pd.DataFrame(master_values)
                                    st.dataframe(
                                        master_values_df,
                                        use_container_width=True,
                                        hide_index=True,
                                        height=230,
                                        column_config={
                                            "NSE Symbol": st.column_config.TextColumn(width="small"),
                                            "% Chg": st.column_config.NumberColumn(
                                                "% Chg",
                                                format="%.2f%%",
                                                width="small"
                                            ),
                                            f"{rank_source_column} Rank": st.column_config.NumberColumn(
                                                f"{rank_source_column} Rank",
                                                format="%d",
                                                width="small"
                                            ),
                                            "Sum O2L.10": st.column_config.TextColumn(width="small"),
                                            "Sum O2H.10": st.column_config.TextColumn(width="small"),
                                            "16< D-B / Avg.4": st.column_config.TextColumn(width="small"),
                                            "16> C-B / Avg.4": st.column_config.TextColumn(width="small"),
                                        }
                                    )
    
                        if hasattr(st, "fragment"):
                            st.fragment(run_every="60s")(
                                render_today_quotes
                            )(
                                selected_constituents,
                                selected_name,
                                "refresh_dynamic_index_quotes",
                                "top",
                                trade_signal_placeholder
                            )
                        else:
                            render_today_quotes(
                                selected_constituents,
                                selected_name,
                                "refresh_dynamic_index_quotes",
                                "top",
                                trade_signal_placeholder
                            )
    
                    except Exception as constituent_error:
                        st.error(
                            f"Could not load constituents or stock details for "
                            f"{selected_name}: {constituent_error}"
                        )
                        st.caption(
                            "The selected index remains visible above. Please try "
                            "again later if the constituent data provider is unavailable."
                        )
    
            if not eligible_indices:
                st.info("No index has CURRENT above 10,000 with a valid H2L% value.")

            # Second independent selection: highest L2H% index.
            with st.container(key="l2h_panel"):
                    if highest_l2h_indices:
                        selected_l2h, highest_l2h_index = highest_l2h_indices[0]
                        highest_l2h_name = str(
                            highest_l2h_index.get("INDEX", "")
                        ).strip().upper()
    
                        st.divider()
                        st.markdown("### L2H · Bottom 4 Stocks")
                        l2h_col1, l2h_col2, l2h_col3 = st.columns(3)
                        l2h_col1.metric("Selected INDEX", highest_l2h_name)
                        l2h_col2.metric("L2H%", f"{selected_l2h:+.2f}%")
                        l2h_current = to_number(highest_l2h_index.get("CURRENT"))
                        l2h_col3.metric(
                            "CURRENT",
                            f"{l2h_current:,.2f}" if l2h_current is not None else "N/A"
                        )
    
                        st.markdown(f"**{highest_l2h_name} — Constituent Stocks**")
                        try:
                            with st.spinner(
                                f"Loading NSE constituents for {highest_l2h_name}..."
                            ):
                                highest_l2h_constituents = get_index_constituents(
                                    highest_l2h_name
                                )
    
                            with st.expander(
                                f"View all {len(highest_l2h_constituents)} L2H index constituents",
                                expanded=False
                            ):
                                st.dataframe(
                                    highest_l2h_constituents,
                                    use_container_width=True,
                                    hide_index=True
                                )
                                st.caption(
                                    "Constituent list sourced from NSE Indices published data "
                                    "through the India API index archive."
                                )
                            st.markdown(
                                f"**{highest_l2h_name} — Today's Stock Details**"
                            )
                            if hasattr(st, "fragment"):
                                st.fragment(run_every="60s")(
                                    render_today_quotes
                                )(
                                    highest_l2h_constituents,
                                    highest_l2h_name,
                                    "refresh_highest_l2h_quotes",
                                    "bottom",
                                    trade_signal_placeholder
                                )
                            else:
                                render_today_quotes(
                                    highest_l2h_constituents,
                                    highest_l2h_name,
                                    "refresh_highest_l2h_quotes",
                                    "bottom",
                                    trade_signal_placeholder
                                )
                        except Exception as highest_l2h_error:
                            st.error(
                                f"Could not load constituents or stock details for "
                                f"{highest_l2h_name}: {highest_l2h_error}"
                            )


            st.subheader("Average Index Details")
            avg_col1, avg_col2, avg_col3, avg_col4, avg_col5 = st.columns(5)
            avg_col1.metric(
                "Average CURRENT",
                f"{average_current:,.2f}" if average_current is not None else "N/A"
            )
            avg_col2.metric(
                "Average HIGH",
                f"{average_high:,.2f}" if average_high is not None else "N/A"
            )
            avg_col3.metric(
                "Average LOW",
                f"{average_low:,.2f}" if average_low is not None else "N/A"
            )
            avg_col4.metric(
                "Average vs HIGH",
                f"{average_vs_high:+.2f}%" if average_vs_high is not None else "N/A"
            )
            avg_col5.metric(
                "Average vs LOW",
                f"{average_vs_low:+.2f}%" if average_vs_low is not None else "N/A"
            )

            search_text = st.text_input(
                "Search index",
                placeholder="Search by index name or displayed values..."
            ).strip().lower()

            if search_text:
                visible_records = [
                    record for record in display_records
                    if any(
                        search_text in str(value).lower()
                        for value in record.values()
                        if value is not None
                    )
                ]
            else:
                visible_records = display_records

            st.caption(
                f"Showing {len(visible_records)} of {len(display_records)} records."
            )

            if visible_records:
                st.dataframe(
                    visible_records,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No matching records found.")

        st.download_button(
            "Download Uploaded File",
            data=file_bytes,
            file_name=file_name,
            mime=(
                EXCEL_MIME
                if extension == "xlsx"
                else "text/csv"
            ),
            use_container_width=True
        )

        if index_wb is not None:
            index_wb.close()

    except Exception as e:
        st.error(f"Unable to read this file: {e}")

    st.stop()


elif page == "Index 2":

    import csv

    st.header("Index 2 Dashboard")
    st.caption(
        "Upload an Excel (.xlsx) or CSV (.csv) file. "
        "The uploaded file is processed in Streamlit memory and is not saved to Supabase."
    )

    # Each connected Streamlit session has its own session_state. The owner
    # marker also prevents a different account using the same browser session
    # from inheriting the previous account's uploaded Index file or signals.
    session_user_id = str(user.id)
    if st.session_state.get("index_file_owner_id") != session_user_id:
        clear_index_session_data()
        st.session_state["index_file_owner_id"] = session_user_id

    index_upload = st.file_uploader(
        "Upload Index File to replace the saved file",
        type=["xlsx", "csv"],
        key=f"index_excel_upload_{session_user_id}"
    )

    # The trade signal is rendered here so it stays directly below the uploader,
    # rather than inside either H2L or L2H analysis panel.
    trade_signal_placeholder = st.empty()

    try:
        # Keep the latest Index upload in this user's Streamlit session only.
        # It is never written to Supabase or shared with other users.
        if index_upload is not None:
            new_file_bytes = index_upload.getvalue()
            new_file_name = index_upload.name
            if (
                new_file_bytes != st.session_state.get("index_file_bytes")
                or new_file_name != st.session_state.get("index_file_name")
            ):
                # A new upload invalidates old H2L/L2H candidates so signals
                # cannot briefly use stocks calculated from the previous file.
                st.session_state.pop("trade_candidates_H2L", None)
                st.session_state.pop("trade_candidates_L2H", None)
                st.session_state.pop("trade_candidates_index2_H2L", None)
                st.session_state.pop("trade_candidates_index2_L2H", None)
            st.session_state["index_file_bytes"] = new_file_bytes
            st.session_state["index_file_name"] = new_file_name
            st.session_state["index_file_owner_id"] = session_user_id

        file_bytes = st.session_state.get("index_file_bytes")
        file_name = st.session_state.get("index_file_name")

        if file_bytes is None or file_name is None:
            st.info("No Index file in this session yet. Upload an Excel or CSV file above.")
            st.stop()
        extension = file_name.lower().rsplit(".", 1)[-1]

        if extension == "csv":
            decoded = None
            for encoding in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
                try:
                    decoded = file_bytes.decode(encoding)
                    break
                except UnicodeDecodeError:
                    continue

            if decoded is None:
                raise ValueError("Could not decode this CSV file.")

            sample = decoded[:8192]
            try:
                dialect = csv.Sniffer().sniff(sample, delimiters=",;\\t|")
            except csv.Error:
                dialect = csv.excel

            csv_reader = csv.reader(io.StringIO(decoded), dialect)
            csv_rows = list(csv_reader)

            # Remove completely blank lines.
            csv_rows = [
                row for row in csv_rows
                if any(str(value).strip() for value in row)
            ]

            if not csv_rows:
                raise ValueError("The CSV file is empty.")

            raw_headers = csv_rows[0]
            data_values = csv_rows[1:]

            index_wb = None
            sheet_names = ["CSV Data"]
            selected_sheet = "CSV Data"

        else:
            index_wb = load_workbook(
                io.BytesIO(file_bytes),
                read_only=True,
                data_only=True
            )
            sheet_names = index_wb.sheetnames
            selected_sheet = st.selectbox(
                "Select worksheet",
                sheet_names,
                key="index_sheet_selector"
            )
            index_ws = index_wb[selected_sheet]
            all_rows = list(
                index_ws.iter_rows(values_only=True)
            )
            all_rows = [
                list(row) for row in all_rows
                if any(value is not None and str(value).strip() for value in row)
            ]

            if not all_rows:
                raise ValueError("The selected worksheet is empty.")

            raw_headers = all_rows[0]
            data_values = all_rows[1:]

        headers = []
        used_headers = set()

        for position, value in enumerate(raw_headers, start=1):
            header = str(value).strip().replace("\\n", " ") if value is not None else ""
            header = " ".join(header.split())
            if not header:
                header = f"Column {position}"

            base_header = header
            suffix = 2
            while header in used_headers:
                header = f"{base_header}_{suffix}"
                suffix += 1

            used_headers.add(header)
            headers.append(header)

        records = []
        for row in data_values:
            values = list(row)
            if len(values) < len(headers):
                values.extend([None] * (len(headers) - len(values)))
            records.append({
                headers[i]: values[i]
                for i in range(len(headers))
            })

        st.success(f"Loaded: {file_name}")

        metric1, metric2, metric3 = st.columns(3)
        metric1.metric("Worksheets", len(sheet_names))
        metric2.metric("Data Rows", len(records))
        metric3.metric("Columns in Output", 6)

        st.divider()
        st.subheader("Index Data")

        # Display only the six requested columns, regardless of
        # extra spaces or line breaks in the uploaded file headers.
        required_columns = [
            "INDEX",
            "CURRENT",
            "%CHANGE",
            "OPEN",
            "HIGH",
            "LOW"
        ]

        normalized_header_map = {
            " ".join(str(header).upper().split()): header
            for header in headers
        }

        missing_columns = [
            column for column in required_columns
            if column not in normalized_header_map
        ]

        if missing_columns:
            st.error(
                "The uploaded file is missing these required columns: "
                + ", ".join(missing_columns)
            )
            st.caption(
                "Please check that the file contains all six required "
                "column headings."
            )
        else:
            source_columns = [
                normalized_header_map[column]
                for column in required_columns
            ]

            display_records = []

            for record in records:
                output_record = {
                    column: record.get(source_column)
                    for column, source_column in zip(
                        required_columns,
                        source_columns
                    )
                }

                # Calculate current price movement from the day's High
                # and Low. Keep invalid/zero values blank instead of
                # raising a division-by-zero error.
                try:
                    current_value = float(
                        str(output_record["CURRENT"]).replace(",", "").replace("%", "").strip()
                    )
                    high_value = float(
                        str(output_record["HIGH"]).replace(",", "").replace("%", "").strip()
                    )
                    low_value = float(
                        str(output_record["LOW"]).replace(",", "").replace("%", "").strip()
                    )

                    output_record["H2L%"] = (
                        round(((current_value - high_value) / high_value) * 100, 2)
                        if high_value != 0 else None
                    )
                    output_record["L2H%"] = (
                        round(((current_value - low_value) / low_value) * 100, 2)
                        if low_value != 0 else None
                    )
                except (TypeError, ValueError):
                    output_record["H2L%"] = None
                    output_record["L2H%"] = None

                display_records.append(output_record)

            # Summary calculations: compare the average CURRENT price
            # with the average HIGH and average LOW across all indices.
            def to_number(value):
                try:
                    cleaned = str(value).replace(",", "").replace("%", "").strip()
                    return float(cleaned)
                except (TypeError, ValueError):
                    return None

            average_current_values = [
                to_number(record.get("CURRENT"))
                for record in display_records
            ]
            average_high_values = [
                to_number(record.get("HIGH"))
                for record in display_records
            ]
            average_low_values = [
                to_number(record.get("LOW"))
                for record in display_records
            ]

            average_current_values = [v for v in average_current_values if v is not None]
            average_high_values = [v for v in average_high_values if v is not None]
            average_low_values = [v for v in average_low_values if v is not None]

            average_current = (
                sum(average_current_values) / len(average_current_values)
                if average_current_values else None
            )
            average_high = (
                sum(average_high_values) / len(average_high_values)
                if average_high_values else None
            )
            average_low = (
                sum(average_low_values) / len(average_low_values)
                if average_low_values else None
            )

            average_vs_high = (
                ((average_current - average_high) / average_high) * 100
                if average_current is not None and average_high not in (None, 0)
                else None
            )
            average_vs_low = (
                ((average_current - average_low) / average_low) * 100
                if average_current is not None and average_low not in (None, 0)
                else None
            )

            # Calculate absolute arithmetic averages across ALL uploaded indices.
            all_h2l_values = [
                value for value in (to_number(record.get("H2L%")) for record in display_records)
                if value is not None
            ]
            all_l2h_values = [
                value for value in (to_number(record.get("L2H%")) for record in display_records)
                if value is not None
            ]
            average_h2l = sum(all_h2l_values) / len(all_h2l_values) if all_h2l_values else None
            average_l2h = sum(all_l2h_values) / len(all_l2h_values) if all_l2h_values else None
            absolute_average_h2l = abs(average_h2l) if average_h2l is not None else None
            absolute_average_l2h = abs(average_l2h) if average_l2h is not None else None

            if (
                absolute_average_h2l is not None
                and absolute_average_l2h is not None
                and absolute_average_h2l > absolute_average_l2h
            ):
                dominant_market_side = "H2L"
            elif (
                absolute_average_h2l is not None
                and absolute_average_l2h is not None
                and absolute_average_l2h > absolute_average_h2l
            ):
                dominant_market_side = "L2H"
            else:
                dominant_market_side = None

            # Automatically identify the index with the lowest H2L%
            # among indices whose CURRENT value is above 10,000.
            eligible_indices = []
            highest_l2h_indices = []
            for record in display_records:
                current_value = to_number(record.get("CURRENT"))
                h2l_value = to_number(record.get("H2L%"))
                l2h_value = to_number(record.get("L2H%"))
                if (
                    current_value is not None
                    and current_value > 10000
                    and h2l_value is not None
                ):
                    eligible_indices.append((h2l_value, record))
                if (
                    current_value is not None
                    and current_value > 10000
                    and l2h_value is not None
                ):
                    highest_l2h_indices.append((l2h_value, record))

            eligible_indices.sort(key=lambda item: item[0])
            highest_l2h_indices.sort(key=lambda item: item[0], reverse=True)

            st.divider()
            st.subheader("Priority Stock Analysis")
            avg_h2l_col, avg_l2h_col, market_bias_col = st.columns(3)
            avg_h2l_col.metric(
                "| Average H2L% |",
                f"{absolute_average_h2l:.2f}%" if absolute_average_h2l is not None else "N/A"
            )
            avg_l2h_col.metric(
                "| Average L2H% |",
                f"{absolute_average_l2h:.2f}%" if absolute_average_l2h is not None else "N/A"
            )
            market_bias_col.metric("Dominant Side", dominant_market_side or "Balanced / N/A")
            st.caption("Absolute averages use all valid H2L% and L2H% values in the uploaded index file.")
            st.caption(
                "Your key watchlist comes first: H2L Top 4 gainers and L2H Bottom 4 "
                "stocks, with Master Excel metrics and panel-specific ranks. "
                "Full constituent and quote lists remain available below."
            )

            st.markdown("""
            <style>
            .st-key-h2l_panel {
                background: linear-gradient(145deg, #e8f7f2 0%, #d4eee5 52%, #c5e5d9 100%);
                border: 1px solid #9acdbb;
                border-left: 5px solid #167b68;
                border-radius: 16px;
                padding: 1.25rem 1.35rem;
                margin: 0.8rem 0 1.2rem 0;
                box-shadow: 0 7px 22px rgba(20, 105, 87, 0.11);
            }
            .st-key-l2h_panel {
                background: linear-gradient(145deg, #fff4eb 0%, #fbe2d5 52%, #f5d2c2 100%);
                border: 1px solid #e8b9a3;
                border-left: 5px solid #bd6546;
                border-radius: 16px;
                padding: 1.25rem 1.35rem;
                margin: 0.8rem 0 1.2rem 0;
                box-shadow: 0 7px 22px rgba(153, 78, 51, 0.11);
            }
            .st-key-h2l_panel [data-testid="stMetric"],
            .st-key-l2h_panel [data-testid="stMetric"] {
                background: rgba(255,255,255,0.78);
                border: 1px solid rgba(100,110,100,0.13);
                border-radius: 12px;
                padding: 0.75rem 0.9rem;
            }
            .st-key-h2l_panel [data-testid="stDataFrame"],
            .st-key-l2h_panel [data-testid="stDataFrame"] {
                border: 1px solid rgba(100,110,100,0.16);
                border-radius: 11px;
                overflow: hidden;
            }
            .st-key-h2l_panel button,
            .st-key-l2h_panel button {
                border-radius: 9px;
                font-weight: 600;
            }
            </style>
            """, unsafe_allow_html=True)
            with st.container(key="h2l_panel"):
                if len(eligible_indices) >= 2:
                    selected_h2l, selected_index = eligible_indices[1]
                    selected_name = str(selected_index.get("INDEX", "")).strip().upper()
                    selected_current = to_number(selected_index.get("CURRENT"))
                    selected_high = to_number(selected_index.get("HIGH"))
                    selected_low = to_number(selected_index.get("LOW"))
    
                    select_col1, select_col2, select_col3, select_col4 = st.columns(4)
                    select_col1.metric("Selected INDEX", selected_name)
                    select_col2.metric(
                        "CURRENT",
                        f"{selected_current:,.2f}" if selected_current is not None else "N/A"
                    )
                    select_col3.metric(
                        "H2L%",
                        f"{selected_h2l:+.2f}%"
                    )
                    select_col4.metric(
                        "HIGH",
                        f"{selected_high:,.2f}" if selected_high is not None else "N/A"
                    )
    
                    # Fetch constituents for whichever index the uploaded file selects.
                    st.markdown(f"**{selected_name} — Constituent Stocks**")
                    import requests
                    import pandas as pd
                    import yfinance as yf
                    from datetime import datetime
                    from zoneinfo import ZoneInfo
    
                    @st.cache_data(ttl=3600, show_spinner=False)
                    def get_index_constituents(index_name):
                        """
                        Load current constituents for an index using a layered
                        fallback: India API archive, NSE's public index API,
                        then official Nifty Indices CSV files.
                        """
                        import io
                        import requests
                        import pandas as pd

                        def clean_name(value):
                            return " ".join(
                                re.sub(r"[^A-Z0-9]+", " ", str(value).upper()).split()
                            )

                        def build_members(rows):
                            members = []

                            def first_valid(row, keys, default=""):
                                for key in keys:
                                    value = row.get(key)
                                    if value is not None and not pd.isna(value):
                                        return value
                                return default

                            for row in rows:
                                raw_symbol = first_valid(
                                    row, ["symbol", "Symbol", "SYMBOL"]
                                )
                                symbol = str(raw_symbol).strip()
                                if (
                                    not symbol
                                    or symbol.lower() == "nan"
                                    or clean_name(symbol) == clean_name(index_name)
                                ):
                                    continue
                                meta = row.get("meta") or {}
                                company = first_valid(
                                    row,
                                    ["company", "Company Name", "Company"],
                                    first_valid(meta, ["companyName", "company"])
                                )
                                industry = first_valid(
                                    row,
                                    ["industry", "Industry"],
                                    first_valid(meta, ["industry"])
                                )
                                members.append({
                                    "No.": len(members) + 1,
                                    "NSE Symbol": symbol,
                                    "Company": str(company).strip(),
                                    "Industry": str(industry).strip()
                                })
                            return members

                        wanted_name = clean_name(index_name)
                        errors = []

                        # Source 1: India API archive (fast and already used by the app).
                        try:
                            api_headers = {"User-Agent": "Mozilla/5.0"}
                            base_url = "https://api.dharun.dev/v1/indices"
                            response = requests.get(
                                base_url, headers=api_headers, timeout=20
                            )
                            response.raise_for_status()
                            index_items = response.json().get("data", [])
                            matched = next(
                                (
                                    item for item in index_items
                                    if clean_name(item.get("name", "")) == wanted_name
                                ),
                                None
                            )
                            if matched is None:
                                aliases = {
                                    "NIFTY IT": "NIFTY IT",
                                    "NIFTY 50": "NIFTY 50",
                                    "NIFTY BANK": "NIFTY BANK",
                                    "NIFTY NEXT 50": "NIFTY NEXT 50",
                                    "NIFTY 100": "NIFTY 100",
                                    "NIFTY 200": "NIFTY 200",
                                    "NIFTY 500": "NIFTY 500",
                                    "NIFTY MIDCAP 50": "NIFTY MIDCAP 50",
                                    "NIFTY MIDCAP 100": "NIFTY MIDCAP 100",
                                    "NIFTY MIDCAP 150": "NIFTY MIDCAP 150",
                                    "NIFTY SMALLCAP 50": "NIFTY SMALLCAP 50",
                                    "NIFTY SMALLCAP 100": "NIFTY SMALLCAP 100",
                                    "NIFTY SMALLCAP 250": "NIFTY SMALLCAP 250"
                                }
                                alias_name = aliases.get(wanted_name)
                                if alias_name:
                                    matched = next(
                                        (
                                            item for item in index_items
                                            if clean_name(item.get("name", "")) ==
                                            clean_name(alias_name)
                                        ),
                                        None
                                    )
                            if matched is not None:
                                detail_response = requests.get(
                                    f"{base_url}/{matched['id']}",
                                    headers=api_headers,
                                    timeout=20
                                )
                                detail_response.raise_for_status()
                                payload = detail_response.json().get("data", {})
                                members = build_members(payload.get("members", []))
                                if members:
                                    return members
                                errors.append("India API returned no constituent rows.")
                            else:
                                errors.append("Index is not listed in the India API archive.")
                        except Exception as api_error:
                            errors.append(f"India API: {api_error}")

                        # Source 2: NSE public index constituents endpoint.
                        try:
                            session = requests.Session()
                            nse_headers = {
                                "User-Agent": (
                                    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                                    "Chrome/131.0.0.0 Safari/537.36"
                                ),
                                "Accept": "application/json, text/plain, */*",
                                "Accept-Language": "en-US,en;q=0.9",
                                "Referer": "https://www.nseindia.com/"
                            }
                            session.get(
                                "https://www.nseindia.com/",
                                headers=nse_headers,
                                timeout=15
                            )
                            nse_name = " ".join(str(index_name).upper().split())
                            nse_response = session.get(
                                "https://www.nseindia.com/api/equity-stockIndices",
                                params={"index": nse_name},
                                headers=nse_headers,
                                timeout=20
                            )
                            nse_response.raise_for_status()
                            nse_rows = nse_response.json().get("data", [])
                            members = build_members(nse_rows)
                            if members:
                                return members
                            errors.append("NSE API returned no constituent rows.")
                        except Exception as nse_error:
                            errors.append(f"NSE API: {nse_error}")

                        # Source 3: official NSE Indices constituent CSV.
                        # Explicit filenames cover indices whose official naming
                        # does not follow a simple slug convention.
                        official_files = {
                            "NIFTY 50": "ind_nifty50list.csv",
                            "NIFTY NEXT 50": "ind_niftynext50list.csv",
                            "NIFTY NEXT 100": "ind_niftynext100list.csv",
                            "NIFTY 100": "ind_nifty100list.csv",
                            "NIFTY 200": "ind_nifty200list.csv",
                            "NIFTY 500": "ind_nifty500list.csv",
                            "NIFTY TOTAL MARKET": "ind_niftytotalmarket_list.csv",
                            "NIFTY MIDCAP 50": "ind_niftymidcap50list.csv",
                            "NIFTY MIDCAP 100": "ind_niftymidcap100list.csv",
                            "NIFTY MIDCAP 150": "ind_niftymidcap150list.csv",
                            "NIFTY MIDCAP SELECT": "ind_niftymidcapselect_list.csv",
                            "NIFTY SMALLCAP 50": "ind_niftysmallcap50list.csv",
                            "NIFTY SMALLCAP 100": "ind_niftysmallcap100list.csv",
                            "NIFTY SMALLCAP 250": "ind_niftysmallcap250list.csv",
                            "NIFTY MICROCAP 250": "ind_niftymicrocap250_list.csv",
                            "NIFTY LARGE MIDCAP 250": "ind_niftylargemidcap250list.csv",
                            "NIFTY MIDSMALLCAP 400": "ind_niftymidsmallcap400list.csv",
                            "NIFTY INDIA FPI 150": "ind_niftyIndiaFPI150_list.csv",
                            "NIFTY BANK": "ind_niftybanklist.csv",
                            "NIFTY IT": "ind_niftyitlist.csv",
                            "NIFTY PHARMA": "ind_niftypharmalist.csv",
                            "NIFTY FMCG": "ind_niftyfmcglist.csv",
                            "NIFTY AUTO": "ind_niftyautolist.csv",
                            "NIFTY CHEMICALS": "ind_niftyChemicals_list.csv",
                            "NIFTY METAL": "ind_niftymetallist.csv",
                            "NIFTY REALTY": "ind_niftyrealtylist.csv",
                            "NIFTY PSU BANK": "ind_niftypsubanklist.csv",
                            "NIFTY PRIVATE BANK": "ind_nifty_privatebanklist.csv",
                            "NIFTY FINANCIAL SERVICES": "ind_niftyfinancelist.csv",
                            "NIFTY CONSUMER DURABLES": "ind_niftyconsumerdurableslist.csv",
                            "NIFTY OIL AND GAS": "ind_niftyoilgaslist.csv",
                            "NIFTY OIL GAS": "ind_niftyoilgaslist.csv",
                            "NIFTY HEALTHCARE": "ind_niftyhealthcarelist.csv",
                            "NIFTY HOSPITALS": "ind_niftyhospitalslist.csv",
                            "NIFTY NBFC": "ind_niftynbfclist.csv",
                            "NIFTY INSURANCE": "ind_niftyinsurancelist.csv",
                            "NIFTY POWER": "ind_niftypowerlist.csv",
                            "NIFTY MEDIA": "ind_niftymedialist.csv",
                            "NIFTY500 HEALTHCARE": "ind_nifty500Healthcare_list.csv",
                            "NIFTY MIDSMALL FINANCIAL SERVICES": "ind_niftymidsmallfinancailservice_list.csv",
                            "NIFTY MIDSMALL HEALTHCARE": "ind_niftymidsmallhealthcare_list.csv",
                            "NIFTY MIDSMALL IT TELECOM": "ind_niftymidsmallitAndtelecom_list.csv",
                            "NIFTY100 EQUAL WEIGHT": "ind_nifty100list.csv",
                            "NIFTY100 LOW VOLATILITY 30": "ind_Nifty100LowVolatility30list.csv",
                            "NIFTY200 MOMENTUM 30": "ind_nifty200Momentum30_list.csv",
                            "NIFTY200 ALPHA 30": "ind_nifty200alpha30_list.csv",
                            "NIFTY100 ALPHA 30": "ind_nifty100Alpha30list.csv",
                            "NIFTY DIVIDEND OPPORTUNITIES 50": "ind_niftydivopp50list.csv",
                            "NIFTY GROWTH SECTORS 15": "ind_NiftyGrowth_Sectors15_Index.csv",
                            "NIFTY HIGH BETA 50": "nifty_High_Beta50_Index.csv",
                            "NIFTY LOW VOLATILITY 50": "nifty_low_Volatility50_Index.csv",
                            "NIFTY TOP 10 EQUAL WEIGHT": "ind_niftytop10EqualWeight_list.csv",
                            "NIFTY TOP 15 EQUAL WEIGHT": "ind_niftytop15EqualWeight_list.csv",
                            "NIFTY TOP 20 EQUAL WEIGHT": "ind_niftyTop20EqualWeight_list.csv",
                            "NIFTY100 QUALITY 30": "ind_nifty100Quality30list.csv",
                            "NIFTY MIDCAP150 MOMENTUM 50": "ind_niftymidcap150momentum50_list.csv",
                            "NIFTY MIDCAP150 QUALITY 50": "ind_niftymidcap150quality50list.csv",
                            "NIFTY ALPHA 50": "ind_nifty_Alpha_Index.csv",
                            "NIFTY50 EQUAL WEIGHT": "ind_Nifty50EqualWeight.csv",
                            "NIFTY50 VALUE 20": "ind_Nifty50_Value20.csv",
                            "NIFTY200 VALUE 30": "ind_nifty200Value30_list.csv",
                            "NIFTY500 FLEXICAP QUALITY 30": "ind_nifty500FlexicapQuality30_list.csv",
                            "NIFTY500 VALUE 50": "ind_nifty500Value50_list.csv",
                            "NIFTY500 EQUAL WEIGHT": "ind_nifty500EqualWeight_list.csv",
                            "NIFTY500 LOW VOLATILITY 50": "ind_nifty500LowVolatility50_list.csv",
                            "NIFTY500 MOMENTUM 50": "ind_nifty500Momentum50_list.csv",
                            "NIFTY500 QUALITY 50": "ind_nifty500Quality50_list.csv",
                            "NIFTY500 MULTIFACTOR MQVLV 50": "ind_nifty500MultifactorMQVLv50_list.csv",
                            "NIFTY200 QUALITY 30": "ind_nifty200Quality30_list.csv",
                            "NIFTY MIDSMALLCAP400 MOMENTUM QUALITY 100": "ind_niftyMidSmallcap400MomentumQuality100_list.csv",
                            "NIFTY SMALLCAP250 QUALITY 50": "ind_niftySmallcap250_Quality50_list.csv",
                            "NIFTY TOTAL MARKET MOMENTUM QUALITY 50": "ind_niftyTotalMarketMomentumQuality50_list.csv",
                            "NIFTY CAPITAL MARKETS": "ind_niftyCapitalMarkets_list.csv",
                            "NIFTY COMMODITIES": "ind_niftycommoditieslist.csv",
                            "NIFTY CONGLOMERATE 50": "ind_niftyConglomerate50_list.csv",
                            "NIFTY CORE HOUSING": "ind_niftyCoreHousing_list.csv",
                            "NIFTY CPSE": "ind_niftycpselist.csv",
                            "NIFTY EV NEW AGE AUTOMOTIVE": "ind_niftyEv_NewAgeAutomotive_list.csv",
                            "NIFTY ENERGY": "ind_niftyenergylist.csv",
                            "NIFTY HOUSING": "ind_niftyhousing_list.csv",
                            "NIFTY INDIA CONSUMPTION": "ind_niftyconsumptionlist.csv",
                            "NIFTY INDIA DEFENCE": "ind_niftyindiadefence_list.csv",
                            "NIFTY INDIA DIGITAL": "ind_niftyindiadigital_list.csv",
                            "NIFTY INDIA INFRASTRUCTURE LOGISTICS": "ind_niftyIndiaInfrastructure_Logistics_list.csv",
                            "NIFTY INDIA INTERNET": "ind_niftyIndiaInternet_list.csv",
                            "NIFTY INDIA MANUFACTURING": "ind_niftyindiamanufacturing_list.csv",
                            "NIFTY INDIA NEW AGE CONSUMPTION": "ind_niftyIndiaNewAgeConsumption_list.csv",
                            "NIFTY INDIA RAILWAYS PSU": "ind_niftyIndiaRailwaysPSU_list.csv",
                            "NIFTY INDIA TOURISM": "ind_niftyindiatourism_list.csv",
                            "NIFTY INFRASTRUCTURE": "ind_niftyinfralist.csv",
                            "NIFTY IPO": "ind_niftyIPO_list.csv",
                            "NIFTY MIDCAP LIQUID 15": "ind_Nifty_Midcap_Liquid15.csv",
                            "NIFTY MNC": "ind_niftymnclist.csv",
                            "NIFTY MOBILITY": "ind_niftymobility_list.csv",
                            "NIFTY PSE": "ind_niftypselist.csv",
                            "NIFTY REITS INVITS": "ind_niftyREITs_InvITs_list.csv",
                            "NIFTY RURAL": "ind_niftyRural_list.csv",
                            "NIFTY NON CYCLICAL CONSUMER": "ind_niftynon-cyclicalconsumer_list.csv",
                            "NIFTY SERVICES SECTOR": "ind_niftyservicelist.csv",
                            "NIFTY100 LIQUID 15": "ind_Nifty100_Liquid15.csv",
                            "NIFTY SME EMERGE": "ind_niftysmelist.csv",
                            "NIFTY WAVES": "ind_niftyWaves_list.csv",
                            "NIFTY INDIA CORPORATE GROUP ADITYA BIRLA GROUP": "ind_nifty_adityabirlalist.csv",
                            "NIFTY INDIA CORPORATE GROUP MAHINDRA GROUP": "ind_nifty_mahindralist.csv",
                            "NIFTY INDIA CORPORATE GROUP TATA GROUP": "ind_nifty_tatalist.csv",
                            "NIFTY INDIA SELECT 5 CORPORATE GROUPS MAATR": "ind_niftyIndiaSelect5CorporateGroupsMAATR_list.csv"
                        }
                        csv_filename = official_files.get(wanted_name)
                        if not csv_filename:
                            # Try common official filename patterns for other
                            # Nifty indices, including archive-host variants.
                            slug = re.sub(r"[^a-z0-9]+", "", wanted_name.lower())
                            slug = re.sub(r"^nifty", "nifty", slug)
                            csv_candidates = [
                                f"ind_{slug}list.csv",
                                f"ind_{slug}_list.csv"
                            ]
                        else:
                            csv_candidates = [csv_filename]

                        csv_hosts = [
                            "https://www.niftyindices.com/IndexConstituent/",
                            "https://niftyindices.com/IndexConstituent/",
                            "https://nsearchives.nseindia.com/content/indices/"
                        ]
                        csv_headers = {
                            "User-Agent": (
                                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                "AppleWebKit/537.36 (KHTML, like Gecko) "
                                "Chrome/131.0.0.0 Safari/537.36"
                            ),
                            "Accept": "text/csv,application/csv,*/*;q=0.8",
                            "Referer": "https://www.niftyindices.com/"
                        }
                        for host in csv_hosts:
                            for filename in csv_candidates:
                                try:
                                    csv_response = requests.get(
                                        host + filename,
                                        headers=csv_headers,
                                        timeout=20
                                    )
                                    csv_response.raise_for_status()
                                    frame = pd.read_csv(
                                        io.BytesIO(csv_response.content)
                                    )
                                    frame.columns = [
                                        str(column).replace("\ufeff", "").strip()
                                        for column in frame.columns
                                    ]
                                    symbol_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) in
                                            {"SYMBOL", "NSE SYMBOL", "TRADING SYMBOL"}
                                        ),
                                        None
                                    )
                                    if symbol_column is None or frame.empty:
                                        continue
                                    company_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) in
                                            {"COMPANY NAME", "COMPANY", "NAME"}
                                        ),
                                        None
                                    )
                                    industry_column = next(
                                        (
                                            column for column in frame.columns
                                            if clean_name(column) == "INDUSTRY"
                                        ),
                                        None
                                    )
                                    csv_rows = []
                                    for _, csv_row in frame.iterrows():
                                        csv_rows.append({
                                            "symbol": csv_row.get(symbol_column, ""),
                                            "company": (
                                                csv_row.get(company_column, "")
                                                if company_column else ""
                                            ),
                                            "industry": (
                                                csv_row.get(industry_column, "")
                                                if industry_column else ""
                                            )
                                        })
                                    members = build_members(csv_rows)
                                    if members:
                                        return members
                                except Exception as csv_error:
                                    errors.append(
                                        f"{filename} from {host}: {csv_error}"
                                    )

                        raise ValueError(
                            f"Could not load constituents for {index_name} from "
                            "the available India API, NSE API, or official NSE Indices "
                            "CSV sources. " + " | ".join(errors[-4:])
                        )

                    try:
                        with st.spinner(
                            f"Loading NSE constituents for {selected_name}..."
                        ):
                            selected_constituents = get_index_constituents(
                                selected_name
                            )
    
                        with st.expander(
                            f"View all {len(selected_constituents)} H2L index constituents",
                            expanded=False
                        ):
                            st.dataframe(
                                selected_constituents,
                                use_container_width=True,
                                hide_index=True
                            )
                            st.caption(
                                "Constituent list sourced from NSE Indices published data "
                                "through the India API index archive. Membership lists may "
                                "update periodically."
                            )
    
                        st.markdown(f"**{selected_name} — Today's Stock Details**")
                        market_timezone = ZoneInfo("Asia/Kolkata")
                        market_date = datetime.now(market_timezone).date()
                        st.caption(
                            f"Market date: {market_date.strftime('%d %b %Y')} (India). "
                            "Quotes from Yahoo Finance may be delayed."
                        )
    
                        st.caption(
                            "Stock details refresh automatically every 60 seconds while "
                            "the dashboard is open. Use the button for an immediate refresh."
                        )
    
                        def render_today_quotes(
                            constituents,
                            index_name,
                            refresh_key,
                            selection_mode,
                            trade_signal_placeholder
                        ):
                            quote_symbols = [
                                item["NSE Symbol"] for item in constituents
                            ]
                            refresh_quotes = st.button(
                                "Refresh Today's Stock Details",
                                key=refresh_key
                            )
    
                            @st.cache_data(ttl=60, show_spinner=False)
                            def get_dynamic_index_quotes(symbols, requested_market_date):
                                from concurrent.futures import ThreadPoolExecutor, as_completed
    
                                # The caller passes an ISO date string; convert it to
                                # a date object before comparing with pandas index dates.
                                requested_date = datetime.fromisoformat(
                                    requested_market_date
                                ).date()
    
                                def get_one_symbol(symbol):
                                    ticker = f"{symbol}.NS"
                                    row = {
                                        "NSE Symbol": symbol,
                                        "Price": None,
                                        "Open": None,
                                        "High": None,
                                        "Low": None,
                                        "% Chg": None
                                    }
    
                                    try:
                                        stock = yf.Ticker(ticker)
                                        intraday = stock.history(
                                            period="5d",
                                            interval="5m",
                                            auto_adjust=False,
                                            actions=False,
                                            prepost=False
                                        )
                                        daily = stock.history(
                                            period="10d",
                                            interval="1d",
                                            auto_adjust=False,
                                            actions=False
                                        )
    
                                        if intraday is None or intraday.empty:
                                            return row, None
    
                                        intraday = intraday.dropna(how="all").copy()
                                        if intraday.index.tz is None:
                                            intraday.index = intraday.index.tz_localize(
                                                "Asia/Kolkata"
                                            )
                                        else:
                                            intraday.index = intraday.index.tz_convert(
                                                "Asia/Kolkata"
                                            )
    
                                        today = intraday[
                                            intraday.index.date == requested_date
                                        ].dropna(how="all")
    
                                        if today.empty:
                                            return row, None
    
                                        opens = today["Open"].dropna()
                                        highs = today["High"].dropna()
                                        lows = today["Low"].dropna()
                                        closes = today["Close"].dropna()
    
                                        if (
                                            opens.empty or highs.empty
                                            or lows.empty or closes.empty
                                        ):
                                            return row, None
    
                                        previous_close = None
                                        if daily is not None and not daily.empty:
                                            daily = daily.dropna(how="all").copy()
                                            if daily.index.tz is None:
                                                daily.index = daily.index.tz_localize(
                                                    "Asia/Kolkata"
                                                )
                                            else:
                                                daily.index = daily.index.tz_convert(
                                                    "Asia/Kolkata"
                                                )
                                            previous = daily[
                                                daily.index.date < requested_date
                                            ]["Close"].dropna()
                                            if not previous.empty:
                                                previous_close = float(previous.iloc[-1])
    
                                        price = float(closes.iloc[-1])
                                        row.update({
                                            "Price": price,
                                            "Open": float(opens.iloc[0]),
                                            "High": float(highs.max()),
                                            "Low": float(lows.min()),
                                            "% Chg": (
                                                ((price - previous_close) / previous_close) * 100
                                                if previous_close not in (None, 0) else None
                                            )
                                        })
                                        return row, closes.index[-1]
    
                                    except Exception:
                                        # Keep other constituents visible if one ticker
                                        # is unavailable or Yahoo temporarily rate-limits it.
                                        return row, None
    
                                rows = []
                                latest_quote_time = None
                                max_workers = min(8, max(1, len(symbols)))
    
                                with ThreadPoolExecutor(max_workers=max_workers) as executor:
                                    futures = {
                                        executor.submit(get_one_symbol, symbol): symbol
                                        for symbol in symbols
                                    }
                                    results = {}
                                    for future in as_completed(futures):
                                        symbol = futures[future]
                                        try:
                                            results[symbol] = future.result()
                                        except Exception:
                                            results[symbol] = ({
                                                "NSE Symbol": symbol,
                                                "Price": None,
                                                "Open": None,
                                                "High": None,
                                                "Low": None,
                                                "% Chg": None
                                            }, None)
    
                                for symbol in symbols:
                                    row, quote_time = results[symbol]
                                    rows.append(row)
                                    if quote_time is not None and (
                                        latest_quote_time is None
                                        or quote_time > latest_quote_time
                                    ):
                                        latest_quote_time = quote_time
    
                                timestamp = (
                                    latest_quote_time.strftime("%d %b %Y, %I:%M %p IST")
                                    if latest_quote_time is not None else None
                                )
                                return rows, timestamp
    
                            if refresh_quotes:
                                get_dynamic_index_quotes.clear()
    
                            with st.spinner(
                                f"Loading today's {index_name} stock quotes..."
                            ):
                                quote_rows, quote_timestamp = get_dynamic_index_quotes(
                                    tuple(quote_symbols),
                                    market_date.isoformat()
                                )
    
                            if quote_timestamp:
                                st.caption(f"Latest quote timestamp: {quote_timestamp}")
                            else:
                                st.warning(
                                    "Today's intraday quotes have not been returned by "
                                    "the market-data provider. The table will not substitute "
                                    "yesterday's prices."
                                )
    
                            quote_df = pd.DataFrame(quote_rows)
                            for numeric_column in [
                                "Price", "Open", "High", "Low", "% Chg"
                            ]:
                                quote_df[numeric_column] = quote_df[numeric_column].map(
                                    lambda value: round(value, 2)
                                    if pd.notna(value) else None
                                )
    
                            with st.expander(
                                f"View all {len(quote_df)} live quote rows",
                                expanded=False
                            ):
                                st.dataframe(
                                    quote_df,
                                    use_container_width=True,
                                    hide_index=True,
                                    column_config={
                                        "NSE Symbol": st.column_config.TextColumn(
                                            "NSE Symbol",
                                            width="small"
                                        ),
                                        "Price": st.column_config.NumberColumn(
                                            "Price",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "Open": st.column_config.NumberColumn(
                                            "Open",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "High": st.column_config.NumberColumn(
                                            "High",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "Low": st.column_config.NumberColumn(
                                            "Low",
                                            format="₹%.2f",
                                            width="small"
                                        ),
                                        "% Chg": st.column_config.NumberColumn(
                                            "% Chg",
                                            format="%.2f%%",
                                            width="small"
                                        )
                                    }
                                )
    
                            available_count = int(quote_df["Price"].notna().sum())
                            st.caption(
                                f"Today's quotes available for {available_count} "
                                f"of {len(quote_df)} constituent stocks."
                            )
    
                            # Select the top or bottom four stocks by today's % Chg.
                            selected_four = (
                                quote_df.dropna(subset=["% Chg"])
                                .sort_values(
                                    "% Chg",
                                    ascending=(selection_mode == "bottom")
                                )
                                .head(4)
                                .copy()
                            )
    
                            selection_title = (
                                "Bottom 4 Stocks by % Chg"
                                if selection_mode == "bottom"
                                else "Top 4 Stocks by % Chg"
                            )
                            st.markdown(
                                f"#### {selection_title} · {index_name}"
                            )
                            st.caption(
                                "Top 3 / Bottom 3 are trade-eligible. The 4th stock is "
                                "confirmation-only: if its rank is best or tied-best, wait. "
                                "Master Excel values and ranks are shown for all four."
                            )
                            if selected_four.empty:
                                st.info("No stocks have a valid % Chg value yet.")
                            else:
                                with st.spinner("Matching selected stocks with Supabase Master Excel..."):
                                    master_bytes_for_selection = get_master()
                                    rank_source_column = (
                                        "Sum O2L.10"
                                        if selection_mode == "top"
                                        else "Sum O2H.10"
                                    )
                                    master_values, master_error = get_top_stocks_master_values(
                                        selected_four,
                                        master_bytes_for_selection,
                                        rank_source_column
                                    )
    
                                # Only the first three % Chg stocks on each side
                                # can qualify for a trade. The fourth is confirmation-only:
                                # if its rank is the best (or tied best), the setup must wait.
                                signal_side = "H2L" if selection_mode == "top" else "L2H"
                                rank_column = f"{rank_source_column} Rank"
                                quote_price_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("Price"))
                                    for _, row in quote_df.iterrows()
                                }
                                quote_high_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("High"))
                                    for _, row in quote_df.iterrows()
                                }
                                quote_low_by_symbol = {
                                    str(row.get("NSE Symbol", "")).strip().upper():
                                    to_number(row.get("Low"))
                                    for _, row in quote_df.iterrows()
                                }
                                side_candidates = []
                                if not master_error and master_values:
                                    for stock_position, item in enumerate(master_values):
                                        symbol = str(item.get("NSE Symbol", "")).strip()
                                        rank_value = to_number(item.get(rank_column))
                                        current_price = quote_price_by_symbol.get(symbol.upper())
                                        current_high = quote_high_by_symbol.get(symbol.upper())
                                        current_low = quote_low_by_symbol.get(symbol.upper())
                                        if (
                                            symbol
                                            and rank_value is not None
                                            and current_price is not None
                                        ):
                                            side_candidates.append({
                                                "side": signal_side,
                                                "symbol": symbol,
                                                "rank": rank_value,
                                                "price": current_price,
                                                "high": current_high,
                                                "low": current_low,
                                                "confirmation_only": stock_position >= 3
                                            })

                                st.session_state[f"trade_candidates_index2_{signal_side}"] = side_candidates

                                h2l_candidates = st.session_state.get(
                                    "trade_candidates_index2_H2L", []
                                )
                                l2h_candidates = st.session_state.get(
                                    "trade_candidates_index2_L2H", []
                                )
                                # Include each fourth stock only as a blocker in the
                                # rank comparison; it can never become the signal stock.
                                all_ranked_candidates = h2l_candidates + l2h_candidates

                                winning_candidate = None
                                if all_ranked_candidates:
                                    lowest_rank_value = min(
                                        item["rank"] for item in all_ranked_candidates
                                    )
                                    lowest_rank_candidates = [
                                        item for item in all_ranked_candidates
                                        if item["rank"] == lowest_rank_value
                                    ]
                                    confirmation_blocks_signal = any(
                                        item.get("confirmation_only", False)
                                        for item in lowest_rank_candidates
                                    )
                                    # Wait if either fourth stock has the best rank,
                                    # including a tie at the best rank.
                                    # Also avoid choosing arbitrarily when eligible
                                    # H2L and L2H stocks share the best rank.
                                    winning_sides = {
                                        item["side"] for item in lowest_rank_candidates
                                        if not item.get("confirmation_only", False)
                                    }
                                    if not confirmation_blocks_signal and len(winning_sides) == 1:
                                        eligible_winners = [
                                            item for item in lowest_rank_candidates
                                            if not item.get("confirmation_only", False)
                                        ]
                                        if len(eligible_winners) == 1:
                                            winning_candidate = eligible_winners[0]

                                # Price-location confirmation:
                                # BUY only when current price is at least 0.20% below
                                # today's High; SELL only when current price is at
                                # least 0.20% above today's Low.
                                price_location_confirmed = False
                                if winning_candidate is not None:
                                    signal_price = winning_candidate.get("price")
                                    signal_high = winning_candidate.get("high")
                                    signal_low = winning_candidate.get("low")
                                    if signal_side == "H2L":
                                        price_location_confirmed = (
                                            signal_price is not None
                                            and signal_high is not None
                                            and signal_high > 0
                                            and ((signal_high - signal_price) / signal_high) >= 0.002
                                        )
                                    else:
                                        price_location_confirmed = (
                                            signal_price is not None
                                            and signal_low is not None
                                            and signal_low > 0
                                            and ((signal_price - signal_low) / signal_low) >= 0.002
                                        )

                                if (
                                    winning_candidate is not None
                                    and winning_candidate["side"] == signal_side
                                    and dominant_market_side == signal_side
                                    and price_location_confirmed
                                ):
                                    signal_symbol = winning_candidate["symbol"]
                                    entry_price = winning_candidate["price"]
                                    if signal_side == "H2L":
                                        trade_action = "BUY"
                                        target_price = entry_price * 1.0035
                                        signal_class = "buy-signal"
                                        signal_label = "H2L BUY SETUP"
                                        entry_label = "Buy Entry"
                                        target_label = "Target (+0.35%)"
                                    else:
                                        trade_action = "SELL"
                                        target_price = entry_price * 0.9965
                                        signal_class = "sell-signal"
                                        signal_label = "L2H SELL SETUP"
                                        entry_label = "Sell Entry"
                                        target_label = "Target (-0.35%)"

                                    trade_signal_placeholder.markdown(
                                        f"""
                                        <style>
                                        @keyframes prioritySignalPulse {{
                                            0%, 100% {{ transform: scale(1); filter: brightness(1); }}
                                            50% {{ transform: scale(1.008); filter: brightness(1.12); }}
                                        }}
                                        .priority-trade-signal {{
                                            animation: prioritySignalPulse 1.3s ease-in-out infinite;
                                            border-radius: 14px;
                                            padding: 1rem 1.15rem;
                                            margin: 0.7rem 0 1rem 0;
                                            color: white;
                                        }}
                                        .priority-trade-signal.buy-signal {{
                                            background: linear-gradient(115deg, #126b5c, #269b80);
                                            border: 1px solid #63bda5;
                                        }}
                                        .priority-trade-signal.sell-signal {{
                                            background: linear-gradient(115deg, #a84f3b, #d17d5d);
                                            border: 1px solid #e5a18b;
                                        }}
                                        .priority-trade-signal .signal-title {{
                                            font-size: 1.05rem; font-weight: 800;
                                            letter-spacing: 0.04em; margin-bottom: 0.4rem;
                                        }}
                                        .priority-trade-signal .signal-symbol {{
                                            font-size: 1.4rem; font-weight: 800; margin-bottom: 0.65rem;
                                        }}
                                        .priority-trade-signal .signal-grid {{
                                            display: flex; flex-wrap: wrap; gap: 1.5rem;
                                        }}
                                        .priority-trade-signal .signal-label {{
                                            font-size: 0.78rem; opacity: 0.9;
                                        }}
                                        .priority-trade-signal .signal-value {{
                                            font-size: 1.1rem; font-weight: 750;
                                        }}
                                        </style>
                                        <div class="priority-trade-signal {signal_class}">
                                            <div class="signal-title">{signal_label} · {trade_action}</div>
                                            <div class="signal-symbol">{signal_symbol}</div>
                                            <div class="signal-grid">
                                                <div><div class="signal-label">{entry_label} (Current Price)</div>
                                                <div class="signal-value">₹{entry_price:,.2f}</div></div>
                                                <div><div class="signal-label">{target_label}</div>
                                                <div class="signal-value">₹{target_price:,.2f}</div></div>
                                                <div><div class="signal-label">Lowest Rank Across Both Lists</div>
                                                <div class="signal-value">{winning_candidate["rank"]:g}</div></div>
                                            </div>
                                        </div>
                                        """,
                                        unsafe_allow_html=True
                                    )
                                elif (
                                    winning_candidate is None
                                    and len(all_ranked_candidates) > 0
                                    and len({
                                        item["side"] for item in all_ranked_candidates
                                        if item["rank"] == min(
                                            candidate["rank"] for candidate in all_ranked_candidates
                                        )
                                    }) > 1
                                ):
                                    trade_signal_placeholder.info(
                                        "The lowest rank is tied between H2L and L2H lists. "
                                        "No trade signal is shown."
                                    )
                                # Do not let the other panel's independent 60-second
                                # fragment overwrite the shared signal card. Only the
                                # fragment for the winning side may render the signal.

                                if master_error:
                                    st.warning(master_error)
                                elif master_values:
                                    master_values_df = pd.DataFrame(master_values)
                                    st.dataframe(
                                        master_values_df,
                                        use_container_width=True,
                                        hide_index=True,
                                        height=230,
                                        column_config={
                                            "NSE Symbol": st.column_config.TextColumn(width="small"),
                                            "% Chg": st.column_config.NumberColumn(
                                                "% Chg",
                                                format="%.2f%%",
                                                width="small"
                                            ),
                                            f"{rank_source_column} Rank": st.column_config.NumberColumn(
                                                f"{rank_source_column} Rank",
                                                format="%d",
                                                width="small"
                                            ),
                                            "Sum O2L.10": st.column_config.TextColumn(width="small"),
                                            "Sum O2H.10": st.column_config.TextColumn(width="small"),
                                            "16< D-B / Avg.4": st.column_config.TextColumn(width="small"),
                                            "16> C-B / Avg.4": st.column_config.TextColumn(width="small"),
                                        }
                                    )
    
                        if hasattr(st, "fragment"):
                            st.fragment(run_every="60s")(
                                render_today_quotes
                            )(
                                selected_constituents,
                                selected_name,
                                "refresh_dynamic_index_quotes",
                                "top",
                                trade_signal_placeholder
                            )
                        else:
                            render_today_quotes(
                                selected_constituents,
                                selected_name,
                                "refresh_dynamic_index_quotes",
                                "top",
                                trade_signal_placeholder
                            )
    
                    except Exception as constituent_error:
                        st.error(
                            f"Could not load constituents or stock details for "
                            f"{selected_name}: {constituent_error}"
                        )
                        st.caption(
                            "The selected index remains visible above. Please try "
                            "again later if the constituent data provider is unavailable."
                        )
    
            if not eligible_indices:
                st.info("No index has CURRENT above 10,000 with a valid H2L% value.")
            elif len(eligible_indices) == 1:
                st.info("Index 2 H2L needs at least two eligible indices with CURRENT above 10,000.")

            # Second independent selection: highest L2H% index.
            with st.container(key="l2h_panel"):
                    if len(highest_l2h_indices) >= 2:
                        selected_l2h, highest_l2h_index = highest_l2h_indices[1]
                        highest_l2h_name = str(
                            highest_l2h_index.get("INDEX", "")
                        ).strip().upper()
    
                        st.divider()
                        st.markdown("### L2H · Bottom 4 Stocks")
                        l2h_col1, l2h_col2, l2h_col3 = st.columns(3)
                        l2h_col1.metric("Selected INDEX", highest_l2h_name)
                        l2h_col2.metric("L2H%", f"{selected_l2h:+.2f}%")
                        l2h_current = to_number(highest_l2h_index.get("CURRENT"))
                        l2h_col3.metric(
                            "CURRENT",
                            f"{l2h_current:,.2f}" if l2h_current is not None else "N/A"
                        )
    
                        st.markdown(f"**{highest_l2h_name} — Constituent Stocks**")
                        try:
                            with st.spinner(
                                f"Loading NSE constituents for {highest_l2h_name}..."
                            ):
                                highest_l2h_constituents = get_index_constituents(
                                    highest_l2h_name
                                )
    
                            with st.expander(
                                f"View all {len(highest_l2h_constituents)} L2H index constituents",
                                expanded=False
                            ):
                                st.dataframe(
                                    highest_l2h_constituents,
                                    use_container_width=True,
                                    hide_index=True
                                )
                                st.caption(
                                    "Constituent list sourced from NSE Indices published data "
                                    "through the India API index archive."
                                )
                            st.markdown(
                                f"**{highest_l2h_name} — Today's Stock Details**"
                            )
                            if hasattr(st, "fragment"):
                                st.fragment(run_every="60s")(
                                    render_today_quotes
                                )(
                                    highest_l2h_constituents,
                                    highest_l2h_name,
                                    "refresh_highest_l2h_quotes",
                                    "bottom",
                                    trade_signal_placeholder
                                )
                            else:
                                render_today_quotes(
                                    highest_l2h_constituents,
                                    highest_l2h_name,
                                    "refresh_highest_l2h_quotes",
                                    "bottom",
                                    trade_signal_placeholder
                                )
                        except Exception as highest_l2h_error:
                            st.error(
                                f"Could not load constituents or stock details for "
                                f"{highest_l2h_name}: {highest_l2h_error}"
                            )
                    elif len(highest_l2h_indices) == 1:
                        st.info("Index 2 L2H needs at least two eligible indices with CURRENT above 10,000.")
                    else:
                        st.info("No index has CURRENT above 10,000 with a valid L2H% value.")


            st.subheader("Average Index Details")
            avg_col1, avg_col2, avg_col3, avg_col4, avg_col5 = st.columns(5)
            avg_col1.metric(
                "Average CURRENT",
                f"{average_current:,.2f}" if average_current is not None else "N/A"
            )
            avg_col2.metric(
                "Average HIGH",
                f"{average_high:,.2f}" if average_high is not None else "N/A"
            )
            avg_col3.metric(
                "Average LOW",
                f"{average_low:,.2f}" if average_low is not None else "N/A"
            )
            avg_col4.metric(
                "Average vs HIGH",
                f"{average_vs_high:+.2f}%" if average_vs_high is not None else "N/A"
            )
            avg_col5.metric(
                "Average vs LOW",
                f"{average_vs_low:+.2f}%" if average_vs_low is not None else "N/A"
            )

            search_text = st.text_input(
                "Search index",
                placeholder="Search by index name or displayed values..."
            ).strip().lower()

            if search_text:
                visible_records = [
                    record for record in display_records
                    if any(
                        search_text in str(value).lower()
                        for value in record.values()
                        if value is not None
                    )
                ]
            else:
                visible_records = display_records

            st.caption(
                f"Showing {len(visible_records)} of {len(display_records)} records."
            )

            if visible_records:
                st.dataframe(
                    visible_records,
                    use_container_width=True,
                    hide_index=True
                )
            else:
                st.warning("No matching records found.")

        st.download_button(
            "Download Uploaded File",
            data=file_bytes,
            file_name=file_name,
            mime=(
                EXCEL_MIME
                if extension == "xlsx"
                else "text/csv"
            ),
            use_container_width=True
        )

        if index_wb is not None:
            index_wb.close()

    except Exception as e:
        st.error(f"Unable to read this file: {e}")

    st.stop()

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
