import io
import streamlit as st
from supabase import create_client
import pandas as pd


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
        SUPABASE_SERVICE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]
    )


supabase = get_supabase()


# ============================================================
# SESSION STATE
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "profile" not in st.session_state:
    st.session_state.profile = None


# ============================================================
# GET USER PROFILE
# ============================================================

def get_profile(user_id):

    try:

        response = (
            supabase
            .table("profiles")
            .select(
                "id,email,role,is_active"
            )
            .eq(
                "id",
                user_id
            )
            .eq(
                "is_active",
                True
            )
            .limit(1)
            .execute()
        )

        if not response:
            return None

        if not response.data:
            return None

        return response.data[0]

    except Exception:

        return None


# ============================================================
# CHECK ADMIN
# ============================================================

def is_admin():

    profile = st.session_state.profile

    if not profile:
        return False

    return (
        profile.get("role") == "admin"
        and profile.get("is_active") is True
    )


# ============================================================
# LOGIN
# ============================================================

def login_user(email, password):

    try:

        response = (
            supabase
            .auth
            .sign_in_with_password(
                {
                    "email": email.strip(),
                    "password": password
                }
            )
        )

        if not response:
            return False, "Login failed."

        if not response.user:
            return False, "Login failed."

        user_id = response.user.id

        profile = get_profile(
            user_id
        )

        if not profile:

            try:
                supabase.auth.sign_out()
            except Exception:
                pass

            return (
                False,
                "Account is not active."
            )

        st.session_state.user = response.user
        st.session_state.profile = profile

        return (
            True,
            "Login successful."
        )

    except Exception as e:

        return (
            False,
            str(e)
        )


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
# DOWNLOAD MASTER EXCEL
# ============================================================

def get_excel():

    try:

        data = (
            supabase
            .storage
            .from_(BUCKET)
            .download(FILE)
        )

        return data

    except Exception:

        return None


# ============================================================
# UPLOAD / REPLACE MASTER EXCEL
# ============================================================

def save_excel(data):

    try:

        supabase.storage.from_(BUCKET).upload(
            FILE,
            data,
            {
                "content-type":
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet",

                "upsert":
                    "true"
            }
        )

        return True, "Success"

    except Exception as e:

        return False, str(e)


# ============================================================
# DELETE MASTER EXCEL
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

    st.title(
        "6thSense Excel Manager"
    )

    st.subheader(
        "Login"
    )

    with st.form(
        "login_form"
    ):

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

                st.error(
                    message
                )

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user

st.title(
    "6thSense Excel Manager"
)

st.caption(
    f"Logged in as: {user.email}"
)


# ============================================================
# ACCESS STATUS
# ============================================================

if is_admin():

    st.success(
        "ADMIN ACCESS"
    )

else:

    st.info(
        "USER ACCESS — READ ONLY"
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.write(
        f"**User:** {user.email}"
    )

    if is_admin():

        st.write(
            "**Role:** Admin"
        )

    else:

        st.write(
            "**Role:** User"
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

st.header(
    "Master Excel File"
)

excel_data = get_excel()


# ============================================================
# NO MASTER FILE
# ============================================================

if excel_data is None:

    st.warning(
        "No master Excel file is currently available."
    )

    # --------------------------------------------------------
    # ADMIN UPLOAD
    # --------------------------------------------------------

    if is_admin():

        st.subheader(
            "Upload Master Excel"
        )

        uploaded = st.file_uploader(
            "Choose Excel file",
            type=["xlsx"],
            key="initial_upload"
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
# MASTER FILE AVAILABLE
# ============================================================

st.success(
    "Master Excel file is available."
)


# ============================================================
# DOWNLOAD
# ============================================================

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

    st.header(
        "Admin Controls"
    )


    # ========================================================
    # REPLACE MASTER EXCEL
    # ========================================================

    st.subheader(
        "Replace Master Excel"
    )

    replacement = st.file_uploader(
        "Choose new Excel file",
        type=["xlsx"],
        key="replacement_upload"
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
    # DELETE MASTER FILE
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
    # MODIFY EXCEL
    # ========================================================

    st.divider()

    st.subheader(
        "Modify Excel"
    )

    try:

        # ----------------------------------------------------
        # LOAD WORKBOOK
        # ----------------------------------------------------

        workbook = pd.ExcelFile(
            io.BytesIO(excel_data),
            engine="openpyxl"
        )

        sheets = workbook.sheet_names

        if not sheets:

            st.warning(
                "The Excel file contains no worksheets."
            )

        else:

            # ------------------------------------------------
            # SELECT WORKSHEET
            # ------------------------------------------------

            selected_sheet = st.selectbox(
                "Select worksheet",
                sheets
            )

            # ------------------------------------------------
            # READ SELECTED SHEET
            # ------------------------------------------------

            df = pd.read_excel(
                io.BytesIO(excel_data),
                sheet_name=selected_sheet,
                engine="openpyxl"
            )

            st.write(
                f"Editing: **{selected_sheet}**"
            )

            # ------------------------------------------------
            # EDIT DATA
            # ------------------------------------------------

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic"
            )

            # ------------------------------------------------
            # SAVE MODIFIED EXCEL
            # ------------------------------------------------

            if st.button(
                "💾 Save Modified Excel",
                use_container_width=True
            ):

                output = io.BytesIO()

                with pd.ExcelWriter(
                    output,
                    engine="openpyxl"
                ) as writer:

                    # ----------------------------------------
                    # SAVE EDITED SHEET
                    # ----------------------------------------

                    edited_df.to_excel(
                        writer,
                        sheet_name=selected_sheet,
                        index=False
                    )

                    # ----------------------------------------
                    # SAVE OTHER SHEETS
                    # ----------------------------------------

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

                # --------------------------------------------
                # UPLOAD MODIFIED WORKBOOK
                # --------------------------------------------

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
