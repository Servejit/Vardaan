import streamlit as st
import pandas as pd
import base64
from io import BytesIO
from datetime import datetime, timezone
from supabase import create_client


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="6thSense Master Excel",
    page_icon="📊",
    layout="wide"
)


# ============================================================
# SUPABASE
# ============================================================

@st.cache_resource
def get_supabase():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_ANON_KEY"]
    )


supabase = get_supabase()


# ============================================================
# SESSION STATE
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False


# ============================================================
# ADMIN CHECK
# ============================================================

def check_admin(user_id):
    try:
        result = (
            supabase
            .table("admin_users")
            .select("user_id")
            .eq("user_id", user_id)
            .limit(1)
            .execute()
        )

        return bool(result.data)

    except Exception:
        return False


# ============================================================
# GET MASTER FILE
# ============================================================

def get_master_file():
    try:
        result = (
            supabase
            .table("master_file")
            .select(
                "id,file_name,file_data,uploaded_by,"
                "uploaded_at,updated_at"
            )
            .order("uploaded_at", desc=True)
            .limit(1)
            .execute()
        )

        if result.data:
            return result.data[0]

        return None

    except Exception as e:
        st.error("Unable to load Master Excel.")
        st.error(str(e))
        return None


# ============================================================
# UPLOAD / REPLACE MASTER FILE
# ============================================================

def save_master_file(uploaded_file, user_id):

    try:
        file_bytes = uploaded_file.getvalue()

        encoded_data = base64.b64encode(
            file_bytes
        ).decode("utf-8")

        now = datetime.now(
            timezone.utc
        ).isoformat()

        existing = get_master_file()

        data = {
            "file_name": uploaded_file.name,
            "file_data": encoded_data,
            "uploaded_by": user_id,
            "uploaded_at": now,
            "updated_at": now
        }

        # ----------------------------------------------------
        # REPLACE EXISTING FILE
        # ----------------------------------------------------

        if existing:

            result = (
                supabase
                .table("master_file")
                .update(data)
                .eq("id", existing["id"])
                .execute()
            )

        # ----------------------------------------------------
        # FIRST FILE
        # ----------------------------------------------------

        else:

            result = (
                supabase
                .table("master_file")
                .insert(data)
                .execute()
            )

        return True, result

    except Exception as e:
        return False, str(e)


# ============================================================
# DECODE MASTER FILE
# ============================================================

def get_file_bytes(master):

    try:

        return base64.b64decode(
            master["file_data"]
        )

    except Exception:

        return None


# ============================================================
# LOGOUT
# ============================================================

def logout():

    try:
        supabase.auth.sign_out()
    except Exception:
        pass

    st.session_state.user = None
    st.session_state.is_admin = False

    st.rerun()


# ============================================================
# LOGIN
# ============================================================

if st.session_state.user is None:

    st.title("📊 6thSense Master Excel")

    st.subheader("Login")

    with st.form("login_form"):

        email = st.text_input(
            "Email",
            placeholder="Enter your email"
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter your password"
        )

        login = st.form_submit_button(
            "Login",
            use_container_width=True
        )

    if login:

        if not email or not password:

            st.error(
                "Please enter Email and Password."
            )

        else:

            try:

                result = (
                    supabase
                    .auth
                    .sign_in_with_password(
                        {
                            "email": email,
                            "password": password
                        }
                    )
                )

                if result.user:

                    st.session_state.user = result.user

                    st.session_state.is_admin = (
                        check_admin(result.user.id)
                    )

                    st.rerun()

                else:

                    st.error(
                        "Login failed."
                    )

            except Exception as e:

                st.error(
                    f"Login failed: {e}"
                )

    st.info(
        "Login using your Supabase Authentication account."
    )

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user
is_admin = st.session_state.is_admin


# ============================================================
# HEADER
# ============================================================

header_col1, header_col2 = st.columns(
    [5, 1]
)

with header_col1:

    st.title("📊 6thSense Master Excel")

    if is_admin:

        st.success(
            "🔐 Administrator"
        )

    else:

        st.info(
            "👤 User"
        )


with header_col2:

    if st.button(
        "Logout",
        use_container_width=True
    ):

        logout()


st.divider()


# ============================================================
# ADMIN PANEL
# ============================================================

if is_admin:

    st.header(
        "🔐 Administrator Panel"
    )

    st.write(
        "Only the administrator can upload or replace "
        "the Master Excel file."
    )

    uploaded_file = st.file_uploader(
        "Select Master Excel file",
        type=["xlsx"],
        key="master_excel_upload"
    )

    if uploaded_file is not None:

        st.write(
            f"**Selected file:** {uploaded_file.name}"
        )

        size_kb = uploaded_file.size / 1024

        st.write(
            f"**File size:** {size_kb:,.1f} KB"
        )

        if st.button(
            "⬆️ Upload / Replace Master Excel",
            type="primary",
            use_container_width=True
        ):

            with st.spinner(
                "Uploading Master Excel..."
            ):

                success, result = save_master_file(
                    uploaded_file,
                    user.id
                )

            if success:

                st.success(
                    "Master Excel uploaded successfully."
                )

                st.rerun()

            else:

                st.error(
                    "Upload failed."
                )

                st.code(
                    str(result)
                )

    st.divider()


# ============================================================
# CURRENT MASTER FILE
# ============================================================

st.header(
    "📁 Current Master File"
)

master = get_master_file()


if master is None:

    st.warning(
        "No Master Excel file has been uploaded yet."
    )

    if is_admin:

        st.info(
            "Please upload the first Master Excel file."
        )

    else:

        st.info(
            "Please contact the administrator."
        )

    st.stop()


# ============================================================
# FILE INFORMATION
# ============================================================

file_name = master.get(
    "file_name",
    "summary_output.xlsx"
)

uploaded_at = master.get(
    "uploaded_at"
)

st.write(
    f"**File:** {file_name}"
)


if uploaded_at:

    try:

        dt = datetime.fromisoformat(
            uploaded_at.replace(
                "Z",
                "+00:00"
            )
        )

        st.write(
            "**Last updated:** "
            + dt.strftime(
                "%d %b %Y, %I:%M:%S %p"
            )
        )

    except Exception:

        st.write(
            f"**Last updated:** {uploaded_at}"
        )


# ============================================================
# GET EXCEL BYTES
# ============================================================

file_bytes = get_file_bytes(master)


if file_bytes is None:

    st.error(
        "Unable to decode the Master Excel file."
    )

    st.stop()


# ============================================================
# DOWNLOAD
# ============================================================

st.download_button(
    label="⬇️ Download Current Master Excel",
    data=file_bytes,
    file_name=file_name,
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)


# ============================================================
# EXCEL PREVIEW
# ============================================================

st.divider()

st.header(
    "📋 Excel Preview"
)

try:

    excel_buffer = BytesIO(
        file_bytes
    )

    excel = pd.ExcelFile(
        excel_buffer,
        engine="openpyxl"
    )

    sheets = excel.sheet_names

    st.write(
        f"**Total Sheets:** {len(sheets)}"
    )

    selected_sheet = st.selectbox(
        "Select Sheet",
        sheets
    )

    preview_buffer = BytesIO(
        file_bytes
    )

    df = pd.read_excel(
        preview_buffer,
        sheet_name=selected_sheet,
        engine="openpyxl"
    )

    col1, col2 = st.columns(2)

    with col1:

        st.metric(
            "Rows",
            f"{len(df):,}"
        )

    with col2:

        st.metric(
            "Columns",
            f"{len(df.columns):,}"
        )

    st.dataframe(
        df,
        use_container_width=True,
        height=600
    )

except Exception as e:

    st.error(
        "Unable to preview the Excel file."
    )

    st.code(
        str(e)
    )


# ============================================================
# ACCESS INFORMATION
# ============================================================

st.divider()

if is_admin:

    st.success(
        "🔐 Administrator access — "
        "You can upload and replace the Master Excel."
    )

else:

    st.info(
        "👤 User access — "
        "Master Excel is read-only."
    )
