import io
import streamlit as st
from supabase import create_client

# ============================================================
# SETTINGS
# ============================================================

st.set_page_config(
    page_title="6thSense Excel Manager",
    page_icon="📊",
    layout="wide"
)

BUCKET = "excel-files"
FILE = "master.xlsx"


# ============================================================
# SUPABASE
# ============================================================

@st.cache_resource
def get_supabase():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_KEY"]
    )


supabase = get_supabase()


# ============================================================
# SESSION
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "profile" not in st.session_state:
    st.session_state.profile = None


# ============================================================
# PROFILE
# ============================================================

def get_profile(user_id):

    try:
        r = (
            supabase.table("profiles")
            .select("id,email,role,is_active")
            .eq("id", user_id)
            .eq("is_active", True)
            .limit(1)
            .execute()
        )

        if not r or not r.data:
            return None

        return r.data[0]

    except Exception:
        return None


def is_admin():

    p = st.session_state.profile

    return bool(
        p
        and p.get("role") == "admin"
        and p.get("is_active") is True
    )


# ============================================================
# LOGIN
# ============================================================

def login_user(email, password):

    try:

        r = supabase.auth.sign_in_with_password({
            "email": email.strip(),
            "password": password
        })

        if not r or not r.user:
            return False, "Login failed."

        p = get_profile(r.user.id)

        if not p:
            supabase.auth.sign_out()
            return False, "Account is not active."

        st.session_state.user = r.user
        st.session_state.profile = p

        return True, "Login successful."

    except Exception as e:

        return False, str(e)


# ============================================================
# LOGOUT
# ============================================================

def logout_user():

    try:
        supabase.auth.sign_out()
    except Exception:
        pass

    st.session_state.user = None
    st.session_state.profile = None

    st.rerun()


# ============================================================
# DOWNLOAD EXCEL
# ============================================================

def get_excel():

    try:

        return supabase.storage.from_(BUCKET).download(
            FILE
        )

    except Exception:

        return None


# ============================================================
# UPLOAD / REPLACE EXCEL
# ============================================================

def save_excel(data):

    try:

        supabase.storage.from_(BUCKET).upload(
            FILE,
            data,
            {
                "content-type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upsert": True
            }
        )

        return True, "Success"

    except Exception as e:

        return False, str(e)


# ============================================================
# DELETE EXCEL
# ============================================================

def delete_excel():

    try:

        supabase.storage.from_(BUCKET).remove(
            [FILE]
        )

        return True, "Deleted"

    except Exception as e:

        return False, str(e)


# ============================================================
# LOGIN SCREEN
# ============================================================

if st.session_state.user is None:

    st.title("6thSense Excel Manager")

    st.subheader("Login")

    with st.form("login_form"):

        email = st.text_input(
            "Email"
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

        if not email or not password:

            st.error(
                "Please enter email and password."
            )

        else:

            success, message = login_user(
                email,
                password
            )

            if success:

                st.rerun()

            else:

                st.error(message)

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user

st.title("6thSense Excel Manager")

st.caption(
    f"Logged in as: {user.email}"
)

if is_admin():

    st.success("ADMIN ACCESS")

else:

    st.info("USER ACCESS — READ ONLY")


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.write(
        f"**User:** {user.email}"
    )

    st.write(
        f"**Role:** {'Admin' if is_admin() else 'User'}"
    )

    st.divider()

    if st.button(
        "Logout",
        use_container_width=True
    ):

        logout_user()


# ============================================================
# MASTER EXCEL
# ============================================================

st.header("Master Excel File")

excel_data = get_excel()


# ============================================================
# NO FILE
# ============================================================

if excel_data is None:

    st.warning(
        "No master Excel file is currently available."
    )

    if is_admin():

        st.subheader(
            "Upload Master Excel"
        )

        uploaded = st.file_uploader(
            "Choose Excel file",
            type=["xlsx"]
        )

        if uploaded:

            if st.button(
                "Upload Master File",
                use_container_width=True
            ):

                success, message = save_excel(
                    uploaded.getvalue()
                )

                if success:

                    st.success(
                        "Master Excel uploaded successfully."
                    )

                    st.rerun()

                else:

                    st.error(
                        f"Upload failed: {message}"
                    )

    st.stop()


# ============================================================
# DOWNLOAD
# ============================================================

st.success(
    "Master Excel file is available."
)

st.download_button(
    label="⬇️ Download Master Excel",
    data=excel_data,
    file_name="master.xlsx",
    mime=(
        "application/vnd.openxmlformats-officedocument."
        "spreadsheetml.sheet"
    ),
    use_container_width=True
)


# ============================================================
# ADMIN CONTROLS
# ============================================================

if is_admin():

    st.divider()

    st.header("Admin Controls")


    # ========================================================
    # REPLACE
    # ========================================================

    st.subheader(
        "Replace Master Excel"
    )

    replacement = st.file_uploader(
        "Choose new Excel file",
        type=["xlsx"],
        key="replacement"
    )

    if replacement:

        if st.button(
            "Replace Master File",
            use_container_width=True
        ):

            success, message = save_excel(
                replacement.getvalue()
            )

            if success:

                st.success(
                    "Master Excel replaced successfully."
                )

                st.rerun()

            else:

                st.error(
                    f"Replacement failed: {message}"
                )


    # ========================================================
    # DELETE
    # ========================================================

    st.divider()

    st.subheader(
        "Delete Master File"
    )

    st.warning(
        "Deleting this file will remove it for all users."
    )

    if st.button(
        "🗑️ Delete Master File",
        use_container_width=True
    ):

        success, message = delete_excel()

        if success:

            st.success(
                "Master Excel deleted."
            )

            st.rerun()

        else:

            st.error(
                f"Delete failed: {message}"
            )


    # ========================================================
    # MODIFY
    # ========================================================

    st.divider()

    st.subheader(
        "Modify Excel"
    )

    try:

        import pandas as pd

        workbook = pd.ExcelFile(
            io.BytesIO(excel_data),
            engine="openpyxl"
        )

        sheets = workbook.sheet_names

        selected_sheet = st.selectbox(
            "Select worksheet",
            sheets
        )

        df = pd.read_excel(
            io.BytesIO(excel_data),
            sheet_name=selected_sheet,
            engine="openpyxl"
        )

        st.write(
            f"Editing: **{selected_sheet}**"
        )

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            num_rows="dynamic"
        )

        if st.button(
            "💾 Save Modified Excel",
            use_container_width=True
        ):

            output = io.BytesIO()

            with pd.ExcelWriter(
                output,
                engine="openpyxl"
            ) as writer:

                edited_df.to_excel(
                    writer,
                    sheet_name=selected_sheet,
                    index=False
                )

                for sheet in sheets:

                    if sheet == selected_sheet:
                        continue

                    other_df = pd.read_excel(
                        io.BytesIO(excel_data),
                        sheet_name=sheet,
                        engine="openpyxl"
                    )

                    other_df.to_excel(
                        writer,
                        sheet_name=sheet,
                        index=False
                    )

            success, message = save_excel(
                output.getvalue()
            )

            if success:

                st.success(
                    "Modified Excel saved successfully."
                )

                st.rerun()

            else:

                st.error(
                    f"Save failed: {message}"
                )

    except Exception as e:

        st.error(
            f"Excel error: {e}"
        )


# ============================================================
# NORMAL USER
# ============================================================

else:

    st.info(
        "You have read-only access. "
        "You can view and download the master Excel file."
    )
