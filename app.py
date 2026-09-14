import io
import streamlit as st
from openpyxl import load_workbook
from supabase import create_client, ClientOptions

# ============================================================
# SETTINGS
# ============================================================

URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
SERVICE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]

BUCKET = "excel-files"
FILE = "master.xlsx"
MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

sb = create_client(URL, KEY, options=ClientOptions(flow_type="pkce"))
admin = create_client(URL, SERVICE_KEY)

# ============================================================
# SESSION
# ============================================================

for k, v in {
    "user": None,
    "profile": None,
    "is_admin": False,
    "recovery": False
}.items():
    if k not in st.session_state:
        st.session_state[k] = v

# ============================================================
# FUNCTIONS
# ============================================================

def profile(uid):
    r = sb.table("profiles").select("*").eq("id", uid).execute()
    return r.data[0] if r.data else None


def get_excel():
    try:
        return sb.storage.from_(BUCKET).download(FILE)
    except Exception:
        return None


def save_excel(data):
    admin.storage.from_(BUCKET).upload(
        FILE,
        data,
        {"content-type": MIME, "upsert": "true"}
    )


def remove_excel():
    admin.storage.from_(BUCKET).remove([FILE])


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
            {"auth_code": st.query_params["code"]}
        )
        st.session_state.recovery = True
        st.query_params.clear()
        st.rerun()
    except Exception as e:
        st.error(f"Recovery error: {e}")
        st.stop()


if st.session_state.recovery:
    st.title("Set New Password")

    p1 = st.text_input("New Password", type="password")
    p2 = st.text_input("Confirm Password", type="password")

    if st.button("Update Password", use_container_width=True):
        if len(p1) < 6:
            st.error("Password must be at least 6 characters.")
        elif p1 != p2:
            st.error("Passwords do not match.")
        else:
            try:
                sb.auth.update_user({"password": p1})
                st.success("Password updated successfully.")
                st.session_state.recovery = False
                st.session_state.user = None
                st.rerun()
            except Exception as e:
                st.error(str(e))

    st.stop()


# ============================================================
# LOGIN
# ============================================================

if st.session_state.user is None:

    st.title("6thSense Excel Manager")
    st.subheader("Login")

    with st.form("login"):
        email = st.text_input("Email / User ID")
        password = st.text_input("Password", type="password")
        login = st.form_submit_button(
            "Login",
            use_container_width=True
        )

    if login:
        try:
            r = sb.auth.sign_in_with_password({
                "email": email.strip(),
                "password": password
            })

            u = r.user
            p = profile(u.id)

            if not p:
                st.error("Profile not found.")
            elif not p["is_active"]:
                st.error("Your account is inactive.")
            else:
                st.session_state.user = u
                st.session_state.profile = p
                st.session_state.is_admin = p["role"] == "admin"
                st.rerun()

        except Exception:
            st.error("Invalid email or password.")

    st.divider()
    st.subheader("Forgot Password")

    reset_email = st.text_input("Email")

    if st.button("Send Reset Link", use_container_width=True):
        try:
            app_url = "https://" + st.context.headers.get("host", "")
            sb.auth.reset_password_for_email(
                reset_email.strip(),
                {"redirect_to": app_url}
            )
            st.success("Password reset link sent.")
        except Exception as e:
            st.error(str(e))

    st.stop()


# ============================================================
# LOGGED-IN USER
# ============================================================

user = st.session_state.user
is_admin = st.session_state.is_admin

st.title("6thSense Excel Manager")

with st.sidebar:
    st.write(f"**User:** {user.email}")
    st.write(f"**Role:** {'Admin' if is_admin else 'User'}")

    if st.button("Logout", use_container_width=True):
        logout()


# ============================================================
# MASTER EXCEL
# ============================================================

data = get_excel()

if data:

    st.success("Master Excel available.")

    st.download_button(
        "Download Master Excel",
        data=data,
        file_name=FILE,
        mime=MIME,
        use_container_width=True
    )

else:
    st.warning("Master Excel has not been uploaded.")


# ============================================================
# ADMIN PANEL
# ============================================================

if is_admin:

    st.divider()
    st.header("Admin Panel")

    # --------------------------------------------------------
    # UPLOAD / REPLACE
    # --------------------------------------------------------

    st.subheader("Upload / Replace Master Excel")

    upload = st.file_uploader(
        "Choose Excel file",
        type=["xlsx"]
    )

    if upload and st.button(
        "Save Master Excel",
        use_container_width=True
    ):
        try:
            # Original uploaded bytes are saved directly.
            # No pandas = formatting is preserved.
            save_excel(upload.getvalue())
            st.success("Master Excel saved.")
            st.rerun()
        except Exception as e:
            st.error(str(e))

    # --------------------------------------------------------
    # MODIFY EXCEL
    # --------------------------------------------------------

    if data:

        st.divider()
        st.subheader("Modify Excel")

        try:
            wb = load_workbook(
                io.BytesIO(data),
                data_only=False
            )

            sheet = st.selectbox(
                "Select Sheet",
                wb.sheetnames
            )

            ws = wb[sheet]

            cell = st.text_input(
                "Cell Address",
                "A1"
            ).upper().strip()

            try:
                old = ws[cell].value
                st.write(f"Current value: `{old}`")
            except Exception:
                st.error("Invalid cell address.")
                old = ""

            new = st.text_input(
                "New Value",
                "" if old is None else str(old)
            )

            if st.button(
                "Save Cell Change",
                use_container_width=True
            ):

                ws[cell].value = new

                out = io.BytesIO()
                wb.save(out)

                save_excel(out.getvalue())

                st.success(
                    "Cell changed. Excel formatting preserved."
                )
                st.rerun()

        except Exception as e:
            st.error(f"Excel error: {e}")

    # --------------------------------------------------------
    # DELETE EXCEL
    # --------------------------------------------------------

    st.divider()
    st.subheader("Delete Master File")

    if data:
        if st.button(
            "Delete Master Excel",
            use_container_width=True
        ):
            try:
                remove_excel()
                st.success("Master Excel deleted.")
                st.rerun()
            except Exception as e:
                st.error(str(e))

    # --------------------------------------------------------
    # ADD USER
    # --------------------------------------------------------

    st.divider()
    st.subheader("Add User")

    email = st.text_input("New User Email")
    password = st.text_input(
        "New User Password",
        type="password"
    )
    role = st.selectbox(
        "Role",
        ["user", "admin"]
    )

    if st.button(
        "Create User",
        use_container_width=True
    ):

        try:
            # Create Auth user using SERVICE KEY
            r = admin.auth.admin.create_user({
                "email": email.strip(),
                "password": password,
                "email_confirm": True
            })

            uid = r.user.id

            # Create matching profile using SERVICE KEY
            admin.table("profiles").insert({
                "id": uid,
                "email": email.strip(),
                "role": role,
                "is_active": True
            }).execute()

            st.success("User created successfully.")
            st.rerun()

        except Exception as e:
            st.error(f"Create user error: {e}")

    # --------------------------------------------------------
    # USER MANAGEMENT
    # --------------------------------------------------------

    st.divider()
    st.subheader("User Management")

    try:

        users = admin.table(
            "profiles"
        ).select("*").order(
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

            # ROLE
            if c2.button(
                "Role",
                key=f"role_{uid}"
            ):

                new_role = (
                    "user"
                    if u["role"] == "admin"
                    else "admin"
                )

                admin.table("profiles").update({
                    "role": new_role
                }).eq("id", uid).execute()

                st.rerun()

            # BAN / ACTIVATE
            if c3.button(
                "Ban" if u["is_active"] else "Activate",
                key=f"active_{uid}"
            ):

                admin.table("profiles").update({
                    "is_active": not u["is_active"]
                }).eq("id", uid).execute()

                st.rerun()

            # DELETE
            if c4.button(
                "Delete",
                key=f"delete_{uid}"
            ):

                if uid == user.id:
                    st.error("You cannot delete yourself.")
                else:

                    admin.table(
                        "profiles"
                    ).delete().eq(
                        "id", uid
                    ).execute()

                    try:
                        admin.auth.admin.delete_user(uid)
                    except Exception:
                        pass

                    st.rerun()

    except Exception as e:
        st.error(f"User management error: {e}")


# ============================================================
# NORMAL USER
# ============================================================

else:

    st.info(
        "USER ACCESS â€” Read-only"
    )
