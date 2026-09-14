import io
import pandas as pd
import streamlit as st
from supabase import create_client

st.set_page_config(page_title="Master Excel Manager", page_icon="📊", layout="wide")

URL = st.secrets["SUPABASE_URL"]
KEY = st.secrets["SUPABASE_KEY"]
SERVICE_KEY = st.secrets["SUPABASE_SERVICE_KEY"]

BUCKET = "excel-files"
FILE = "master.xlsx"

@st.cache_resource
def client():
    return create_client(URL, KEY)

@st.cache_resource
def admin_client():
    return create_client(URL, SERVICE_KEY)

sb = client()
admin = admin_client()

for k, v in {
    "user": None,
    "profile": None,
    "is_admin": False
}.items():
    if k not in st.session_state:
        st.session_state[k] = v


def profile(uid):
    try:
        r = sb.table("profiles").select("*").eq("id", uid).limit(1).execute()
        return r.data[0] if r.data else None
    except:
        return None


def login(email, password):
    try:
        r = sb.auth.sign_in_with_password({
            "email": email,
            "password": password
        })
        p = profile(r.user.id)

        if not p or not p.get("is_active", False):
            sb.auth.sign_out()
            return False, "Account is inactive or not authorized."

        st.session_state.user = r.user
        st.session_state.profile = p
        st.session_state.is_admin = p.get("role") == "admin"
        return True, "Login successful."
    except Exception as e:
        return False, str(e)


def logout():
    try:
        sb.auth.sign_out()
    except:
        pass
    st.session_state.user = None
    st.session_state.profile = None
    st.session_state.is_admin = False
    st.rerun()


def get_file():
    try:
        return sb.storage.from_(BUCKET).download(FILE)
    except:
        return None


def save_file(data):
    try:
        sb.storage.from_(BUCKET).upload(
            FILE,
            data,
            {
                "content-type":
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                "upsert": "true"
            }
        )
        return True
    except Exception as e:
        st.error(str(e))
        return False


def delete_file():
    try:
        sb.storage.from_(BUCKET).remove([FILE])
        return True
    except Exception as e:
        st.error(str(e))
        return False


# LOGIN
if st.session_state.user is None:

    st.title("📊 Master Excel Manager")
    st.subheader("Login")

    with st.form("login"):
        email = st.text_input("Email / User ID")
        password = st.text_input("Password", type="password")
        submit = st.form_submit_button("Login", use_container_width=True)

    if submit:
        ok, msg = login(email, password)
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    st.stop()


# LOGGED IN
user = st.session_state.user
is_admin = st.session_state.is_admin

st.title("📊 Master Excel Manager")
st.caption(f"Logged in as: {user.email}")

with st.sidebar:
    st.write(f"**User:** {user.email}")
    st.write(f"**Role:** {'Admin' if is_admin else 'User'}")

    if st.button("Logout", use_container_width=True):
        logout()


# EXCEL
st.header("Master Excel")

data = get_file()

if data:
    st.success("Master Excel file available.")

    st.download_button(
        "⬇️ Download Master Excel",
        data=data,
        file_name="master.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

    if is_admin:
        st.subheader("Modify Excel")

        try:
            df = pd.read_excel(io.BytesIO(data))
            edited = st.data_editor(df, use_container_width=True)

            if st.button("💾 Save Changes", use_container_width=True):
                out = io.BytesIO()
                edited.to_excel(out, index=False)
                out.seek(0)

                if save_file(out.getvalue()):
                    st.success("Excel updated successfully.")
                    st.rerun()

        except Exception as e:
            st.error(f"Excel error: {e}")

else:
    st.warning("No master Excel file uploaded yet.")


# ADMIN
if is_admin:

    st.divider()
    st.header("🔐 Admin Controls")

    st.subheader("Upload / Replace Master Excel")

    uploaded = st.file_uploader(
        "Choose Excel file",
        type=["xlsx"]
    )

    if uploaded and st.button("⬆️ Upload / Replace", use_container_width=True):
        if save_file(uploaded.getvalue()):
            st.success("Master Excel uploaded successfully.")
            st.rerun()

    st.subheader("Delete Master Excel")

    if st.button("🗑️ Delete Master Excel", use_container_width=True):
        if delete_file():
            st.success("Master Excel deleted.")
            st.rerun()


    # USER MANAGEMENT
    st.divider()
    st.header("👥 User Management")

    with st.form("add_user"):
        new_email = st.text_input("New User Email")
        new_password = st.text_input("New User Password", type="password")
        new_role = st.selectbox("Role", ["user", "admin"])
        add = st.form_submit_button("Add User", use_container_width=True)

    if add:
        try:
            r = admin.auth.admin.create_user({
                "email": new_email,
                "password": new_password,
                "email_confirm": True
            })

            uid = r.user.id

            admin.table("profiles").insert({
                "id": uid,
                "email": new_email,
                "role": new_role,
                "is_active": True
            }).execute()

            st.success("User created successfully.")
            st.rerun()

        except Exception as e:
            st.error(str(e))


    st.subheader("Current Users")

    try:
        users = admin.table("profiles").select("*").order("email").execute().data

        for u in users:

            uid = u["id"]
            email = u["email"]
            role = u.get("role", "user")
            active = u.get("is_active", False)

            st.write(
                f"**{email}** — "
                f"{'ADMIN' if role == 'admin' else 'USER'} — "
                f"{'ACTIVE' if active else 'BANNED'}"
            )

            c1, c2, c3 = st.columns(3)

            with c1:
                if role == "admin":
                    if st.button("Make User", key=f"user_{uid}"):
                        admin.table("profiles").update(
                            {"role": "user"}
                        ).eq("id", uid).execute()
                        st.rerun()
                else:
                    if st.button("Make Admin", key=f"admin_{uid}"):
                        admin.table("profiles").update(
                            {"role": "admin"}
                        ).eq("id", uid).execute()
                        st.rerun()

            with c2:
                if active:
                    if st.button("Ban", key=f"ban_{uid}"):
                        admin.table("profiles").update(
                            {"is_active": False}
                        ).eq("id", uid).execute()
                        st.rerun()
                else:
                    if st.button("Activate", key=f"act_{uid}"):
                        admin.table("profiles").update(
                            {"is_active": True}
                        ).eq("id", uid).execute()
                        st.rerun()

            with c3:
                if uid != user.id:
                    if st.button("Delete", key=f"del_{uid}"):
                        try:
                            admin.auth.admin.delete_user(uid)
                            admin.table("profiles").delete().eq(
                                "id", uid
                            ).execute()
                            st.success("User deleted.")
                            st.rerun()
                        except Exception as e:
                            st.error(str(e))

    except Exception as e:
        st.error(str(e))

else:
    st.info("USER ACCESS — View and download only.")
