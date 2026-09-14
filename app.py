import io
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


# ============================================================
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
