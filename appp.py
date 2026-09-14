import io
import pandas as pd
import streamlit as st
from supabase import create_client, ClientOptions

# ============================================================
# SETTINGS
# ============================================================

st.set_page_config(
    page_title="Master Excel Manager",
    page_icon="📊",
    layout="wide"
)

SUPABASE_URL = st.secrets["SUPABASE_URL"]
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
SUPABASE_SERVICE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]

BUCKET = "excel-files"
FILE = "master.xlsx"


# ============================================================
# SUPABASE CLIENT
# ============================================================

@st.cache_resource
def get_supabase():

    options = ClientOptions(
        flow_type="pkce"
    )

    return create_client(
        SUPABASE_URL,
        SUPABASE_KEY,
        options=options
    )


@st.cache_resource
def get_admin_supabase():

    return create_client(
        SUPABASE_URL,
        SUPABASE_SERVICE_KEY
    )


supabase = get_supabase()
admin_supabase = get_admin_supabase()


# ============================================================
# SESSION
# ============================================================

if "user" not in st.session_state:
    st.session_state.user = None

if "profile" not in st.session_state:
    st.session_state.profile = None

if "is_admin" not in st.session_state:
    st.session_state.is_admin = False


# ============================================================
# APP URL
# ============================================================

try:

    host = st.context.headers.get("host", "")

    if host:
        APP_URL = "https://" + host
    else:
        APP_URL = ""

except:

    APP_URL = ""


# ============================================================
# PROFILE
# ============================================================

def get_profile(user_id):

    try:

        result = (
            supabase
            .table("profiles")
            .select("*")
            .eq("id", user_id)
            .limit(1)
            .execute()
        )

        if result.data:
            return result.data[0]

    except Exception:
        pass

    return None


# ============================================================
# LOGIN
# ============================================================

def login_user(email, password):

    try:

        result = supabase.auth.sign_in_with_password(
            {
                "email": email,
                "password": password
            }
        )

        user = result.user

        if not user:
            return False, "Login failed."

        profile = get_profile(user.id)

        if not profile:
            supabase.auth.sign_out()
            return False, "User profile not found."

        if not profile.get("is_active", False):
            supabase.auth.sign_out()
            return False, "Your account is inactive."

        st.session_state.user = user
        st.session_state.profile = profile
        st.session_state.is_admin = (
            profile.get("role") == "admin"
        )

        return True, "Login successful."

    except Exception as e:

        return False, str(e)


# ============================================================
# LOGOUT
# ============================================================

def logout_user():

    try:
        supabase.auth.sign_out()
    except:
        pass

    st.session_state.user = None
    st.session_state.profile = None
    st.session_state.is_admin = False

    st.rerun()


# ============================================================
# EXCEL FUNCTIONS
# ============================================================

def get_excel():

    try:

        data = (
            supabase
            .storage
            .from_(BUCKET)
            .download(FILE)
        )

        if data:
            return data

    except:
        pass

    return None


def save_excel(data):

    try:

        supabase.storage.from_(BUCKET).upload(
            FILE,
            data,
            {
                "content-type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upsert": "true"
            }
        )

        return True, "Success"

    except Exception as e:

        return False, str(e)


def delete_excel():

    try:

        supabase.storage.from_(BUCKET).remove(
            [FILE]
        )

        return True

    except Exception as e:

        st.error(str(e))
        return False


# ============================================================
# PASSWORD RESET
# ============================================================

def send_password_reset(email):

    try:

        if not APP_URL:
            return False, "App URL could not be detected."

        supabase.auth.reset_password_for_email(
            email,
            {
                "redirect_to": APP_URL
            }
        )

        return True, "Password reset email sent."

    except Exception as e:

        return False, str(e)


# ============================================================
# PASSWORD RECOVERY
# ============================================================

query_params = st.query_params

auth_code = query_params.get("code")

if auth_code and st.session_state.user is None:

    try:

        supabase.auth.exchange_code_for_session(
            {
                "auth_code": auth_code
            }
        )

        st.session_state.recovery_mode = True

        st.query_params.clear()

        st.rerun()

    except Exception as e:

        st.error(
            "Password reset link is invalid or expired."
        )

        st.stop()


# ============================================================
# NEW PASSWORD SCREEN
# ============================================================

if st.session_state.get("recovery_mode", False):

    st.title("🔐 Set New Password")

    st.write(
        "Enter your new password below."
    )

    with st.form("new_password_form"):

        new_password = st.text_input(
            "New Password",
            type="password"
        )

        confirm_password = st.text_input(
            "Confirm New Password",
            type="password"
        )

        change_password = st.form_submit_button(
            "Change Password",
            use_container_width=True
        )

    if change_password:

        if not new_password:
            st.error("Please enter a password.")

        elif len(new_password) < 6:
            st.error(
                "Password must be at least 6 characters."
            )

        elif new_password != confirm_password:
            st.error(
                "Passwords do not match."
            )

        else:

            try:

                supabase.auth.update_user(
                    {
                        "password": new_password
                    }
                )

                st.session_state.recovery_mode = False

                st.success(
                    "Password changed successfully."
                )

                st.info(
                    "You can now login with your new password."
                )

                st.stop()

            except Exception as e:

                st.error(str(e))

    st.stop()


# ============================================================
# LOGIN SCREEN
# ============================================================

if st.session_state.user is None:

    st.title(
        "📊 Master Excel Manager"
    )

    st.subheader(
        "Login"
    )

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

        ok, message = login_user(
            email,
            password
        )

        if ok:

            st.success(message)

            st.rerun()

        else:

            st.error(message)


    st.divider()

    st.subheader(
        "Forgot Password?"
    )

    with st.form("forgot_password_form"):

        reset_email = st.text_input(
            "Enter your email"
        )

        reset_button = st.form_submit_button(
            "Send Reset Email",
            use_container_width=True
        )

    if reset_button:

        if not reset_email:

            st.error(
                "Please enter your email."
            )

        else:

            ok, message = send_password_reset(
                reset_email
            )

            if ok:

                st.success(
                    "Password reset email sent. Check your email."
                )

            else:

                st.error(message)

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user
is_admin = st.session_state.is_admin

st.title(
    "📊 Master Excel Manager"
)

st.caption(
    f"Logged in as: {user.email}"
)

if is_admin:

    st.success(
        "ADMIN ACCESS"
    )

else:

    st.info(
        "USER ACCESS — View and download only."
    )


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.write(
        f"**User:** {user.email}"
    )

    if is_admin:

        st.write(
            "**Role:** Admin"
        )

    else:

        st.write(
            "**Role:** User"
        )

    if st.button(
        "Logout",
        use_container_width=True
    ):

        logout_user()


# ============================================================
# MASTER EXCEL
# ============================================================

st.header(
    "Master Excel"
)

excel_data = get_excel()

if excel_data:

    st.success(
        "Master Excel file available."
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

else:

    st.warning(
        "No master Excel file uploaded yet."
    )


# ============================================================
# ADMIN EXCEL CONTROLS
# ============================================================

if is_admin:

    st.divider()

    st.header(
        "🔐 Admin Controls"
    )

    st.subheader(
        "Upload / Replace Master Excel"
    )

    uploaded_file = st.file_uploader(
        "Choose Excel file",
        type=["xlsx"]
    )

    if uploaded_file:

        if st.button(
            "⬆️ Upload / Replace",
            use_container_width=True
        ):

            ok, message = save_excel(
                uploaded_file.getvalue()
            )

            if ok:

                st.success(
                    "Master Excel uploaded successfully."
                )

                st.rerun()

            else:

                st.error(message)


    st.subheader(
        "Modify Excel"
    )

    if excel_data:

        try:

            df = pd.read_excel(
                io.BytesIO(excel_data)
            )

            edited_df = st.data_editor(
                df,
                use_container_width=True,
                num_rows="dynamic"
            )

            if st.button(
                "💾 Save Changes",
                use_container_width=True
            ):

                output = io.BytesIO()

                edited_df.to_excel(
                    output,
                    index=False
                )

                output.seek(0)

                ok, message = save_excel(
                    output.getvalue()
                )

                if ok:

                    st.success(
                        "Excel updated successfully."
                    )

                    st.rerun()

                else:

                    st.error(message)

        except Exception as e:

            st.error(
                f"Excel error: {e}"
            )


    st.subheader(
        "Delete Master Excel"
    )

    if st.button(
        "🗑️ Delete Master Excel",
        use_container_width=True
    ):

        if delete_excel():

            st.success(
                "Master Excel deleted."
            )

            st.rerun()


# ============================================================
# USER MANAGEMENT
# ============================================================

if is_admin:

    st.divider()

    st.header(
        "👥 User Management"
    )


    # --------------------------------------------------------
    # ADD USER
    # --------------------------------------------------------

    st.subheader(
        "Add New User"
    )

    with st.form("add_user_form"):

        new_email = st.text_input(
            "User Email"
        )

        new_password = st.text_input(
            "User Password",
            type="password"
        )

        new_role = st.selectbox(
            "Role",
            ["user", "admin"]
        )

        add_user_button = st.form_submit_button(
            "Add User",
            use_container_width=True
        )

    if add_user_button:

        if not new_email or not new_password:

            st.error(
                "Email and password are required."
            )

        elif len(new_password) < 6:

            st.error(
                "Password must be at least 6 characters."
            )

        else:

            try:

                result = admin_supabase.auth.admin.create_user(
                    {
                        "email": new_email,
                        "password": new_password,
                        "email_confirm": True
                    }
                )

                new_user = result.user

                admin_supabase.table(
                    "profiles"
                ).insert(
                    {
                        "id": new_user.id,
                        "email": new_email,
                        "role": new_role,
                        "is_active": True
                    }
                ).execute()

                st.success(
                    "User created successfully."
                )

                st.rerun()

            except Exception as e:

                st.error(str(e))


    # --------------------------------------------------------
    # CURRENT USERS
    # --------------------------------------------------------

    st.subheader(
        "Current Users"
    )

    try:

        users = (
            admin_supabase
            .table("profiles")
            .select("*")
            .order("email")
            .execute()
            .data
        )

        for current_user in users:

            uid = current_user["id"]
            email = current_user["email"]
            role = current_user.get(
                "role",
                "user"
            )

            active = current_user.get(
                "is_active",
                False
            )

            st.write(
                f"**{email}** — "
                f"{role.upper()} — "
                f"{'ACTIVE' if active else 'BANNED'}"
            )

            c1, c2, c3 = st.columns(3)


            # ROLE
            with c1:

                if role == "admin":

                    if st.button(
                        "Make User",
                        key=f"make_user_{uid}"
                    ):

                        admin_supabase.table(
                            "profiles"
                        ).update(
                            {
                                "role": "user"
                            }
                        ).eq(
                            "id",
                            uid
                        ).execute()

                        st.rerun()

                else:

                    if st.button(
                        "Make Admin",
                        key=f"make_admin_{uid}"
                    ):

                        admin_supabase.table(
                            "profiles"
                        ).update(
                            {
                                "role": "admin"
                            }
                        ).eq(
                            "id",
                            uid
                        ).execute()

                        st.rerun()


            # ACTIVE / BAN
            with c2:

                if active:

                    if st.button(
                        "Ban",
                        key=f"ban_{uid}"
                    ):

                        admin_supabase.table(
                            "profiles"
                        ).update(
                            {
                                "is_active": False
                            }
                        ).eq(
                            "id",
                            uid
                        ).execute()

                        st.rerun()

                else:

                    if st.button(
                        "Activate",
                        key=f"activate_{uid}"
                    ):

                        admin_supabase.table(
                            "profiles"
                        ).update(
                            {
                                "is_active": True
                            }
                        ).eq(
                            "id",
                            uid
                        ).execute()

                        st.rerun()


            # DELETE
            with c3:

                if uid != user.id:

                    if st.button(
                        "Delete",
                        key=f"delete_{uid}"
                    ):

                        try:

                            admin_supabase.auth.admin.delete_user(
                                uid
                            )

                            admin_supabase.table(
                                "profiles"
                            ).delete().eq(
                                "id",
                                uid
                            ).execute()

                            st.success(
                                "User deleted."
                            )

                            st.rerun()

                        except Exception as e:

                            st.error(str(e))

    except Exception as e:

        st.error(str(e))


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Master Excel Manager • Supabase Authentication"
)

