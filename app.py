import io
import re
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
            result = {"NSE Symbol": display_symbol}

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


def logout():
    try:
        sb.auth.sign_out()
    except Exception:
        pass

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
        ["Master Excel", "Index"],
        key="main_dashboard"
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

    index_upload = st.file_uploader(
        "Upload Index File to replace the saved file",
        type=["xlsx", "csv"],
        key="index_excel_upload"
    )

    try:
        # Keep the latest Index upload in Streamlit session memory only.
        # It remains available during reruns in this session and is replaced
        # when a new CSV/Excel file is uploaded. It is not sent to Supabase.
        if index_upload is not None:
            st.session_state["index_file_bytes"] = index_upload.getvalue()
            st.session_state["index_file_name"] = index_upload.name

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
            st.subheader("Index Selection & Constituents")
            st.caption(
                "Automatically selects the lowest H2L% index where CURRENT is above 10,000."
            )

            st.markdown("""
            <style>
            .st-key-h2l_panel { background-color: #eaf7ec; border: 1px solid #b9dfc0; border-radius: 14px; padding: 1rem; }
            .st-key-l2h_panel { background-color: #fff2e2; border: 1px solid #f2d0a8; border-radius: 14px; padding: 1rem; }
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
                        api_headers = {"User-Agent": "Mozilla/5.0"}
                        base_url = "https://api.dharun.dev/v1/indices"
                        response = requests.get(
                            base_url,
                            headers=api_headers,
                            timeout=20
                        )
                        response.raise_for_status()
                        index_items = response.json().get("data", [])
    
                        def clean_name(value):
                            return " ".join(
                                str(value).upper().replace("-", " ").split()
                            )
    
                        wanted_name = clean_name(index_name)
                        matched = next(
                            (
                                item for item in index_items
                                if clean_name(item.get("name", "")) == wanted_name
                            ),
                            None
                        )
    
                        if matched is None:
                            # Handle common NIFTY label variations.
                            aliases = {
                                "NIFTY IT": "Nifty IT",
                                "NIFTY 50": "Nifty 50",
                                "NIFTY BANK": "Nifty Bank",
                                "NIFTY NEXT 50": "Nifty Next 50",
                                "NIFTY 100": "Nifty 100",
                                "NIFTY 200": "Nifty 200",
                                "NIFTY 500": "Nifty 500",
                                "NIFTY MIDCAP 50": "Nifty Midcap 50",
                                "NIFTY MIDCAP 100": "Nifty Midcap 100",
                                "NIFTY MIDCAP 150": "Nifty Midcap 150",
                                "NIFTY SMALLCAP 50": "Nifty Smallcap 50",
                                "NIFTY SMALLCAP 100": "Nifty Smallcap 100",
                                "NIFTY SMALLCAP 250": "Nifty Smallcap 250"
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
    
                        if matched is None:
                            raise ValueError(
                                f"No constituent archive match was found for {index_name}."
                            )
    
                        detail_response = requests.get(
                            f"{base_url}/{matched['id']}",
                            headers=api_headers,
                            timeout=20
                        )
                        detail_response.raise_for_status()
                        payload = detail_response.json().get("data", {})
                        members = payload.get("members", [])
    
                        if not members:
                            raise ValueError(
                                f"The constituent list for {index_name} is empty."
                            )
    
                        return [
                            {
                                "No.": number,
                                "NSE Symbol": member.get("symbol"),
                                "Company": member.get("company"),
                                "Industry": member.get("industry", "")
                            }
                            for number, member in enumerate(members, start=1)
                            if member.get("symbol")
                        ]
    
                    try:
                        with st.spinner(
                            f"Loading NSE constituents for {selected_name}..."
                        ):
                            selected_constituents = get_index_constituents(
                                selected_name
                            )
    
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
    
                        def render_today_quotes(constituents, index_name, refresh_key, selection_mode):
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
                            st.subheader(
                                f"{index_name} — {selection_title} — Master Excel Details"
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
    
                                if master_error:
                                    st.warning(master_error)
                                elif master_values:
                                    master_values_df = pd.DataFrame(master_values)
                                    st.dataframe(
                                        master_values_df,
                                        use_container_width=True,
                                        hide_index=True,
                                        column_config={
                                            "NSE Symbol": st.column_config.TextColumn(width="small"),
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
                                "top"
                            )
                        else:
                            render_today_quotes(
                                selected_constituents,
                                selected_name,
                                "refresh_dynamic_index_quotes",
                                "top"
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
    
                    # Second independent selection: highest L2H% index.
                with st.container(key="l2h_panel"):
                        if highest_l2h_indices:
                            selected_l2h, highest_l2h_index = highest_l2h_indices[0]
                            highest_l2h_name = str(
                                highest_l2h_index.get("INDEX", "")
                            ).strip().upper()
        
                            st.divider()
                            st.subheader("Highest L2H% Index & Bottom 4 Stocks")
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
                                        "bottom"
                                    )
                                else:
                                    render_today_quotes(
                                        highest_l2h_constituents,
                                        highest_l2h_name,
                                        "refresh_highest_l2h_quotes",
                                        "bottom"
                                    )
                            except Exception as highest_l2h_error:
                                st.error(
                                    f"Could not load constituents or stock details for "
                                    f"{highest_l2h_name}: {highest_l2h_error}"
                                )
                else:
                st.info(
                    "No index has CURRENT above 10,000 with a valid H2L% value."
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
