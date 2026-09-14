import io
import streamlit as st
from supabase import create_client

st.set_page_config(
    page_title="6thSense Excel Manager",
    page_icon="📊",
    layout="wide"
)

BUCKET = "excel-files"
FILE = "master.xlsx"


@st.cache_resource
def supabase_client():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_KEY"]
    )


supabase = supabase_client()


if "user" not in st.session_state:
    st.session_state.user = None

if "profile" not in st.session_state:
    st.session_state.profile = None


def profile(user_id):
    r = (
        supabase.table("profiles")
        .select("id,email,role,is_active")
        .eq("id", user_id)
        .eq("is_active", True)
        .maybe_single()
        .execute()
    )
    return r.data


def admin():
    p = st.session_state.profile
    return bool(
        p and
        p.get("role") == "admin" and
        p.get("is_active") is True
    )


def login(email, password):
    try:
        r = supabase.auth.sign_in_with_password({
            "email": email.strip(),
            "password": password
        })

        if not r.user:
            return False, "Login failed."

        p = profile(r.user.id)

        if not p:
            supabase.auth.sign_out()
            return False, "Account is not active."

        st.session_state.user = r.user
        st.session_state.profile = p
        return True, "Login successful."

    except Exception as e:
        return False, str(e)


def logout():
    try:
        supabase.auth.sign_out()
    except Exception:
        pass

    st.session_state.user = None
    st.session_state.profile = None
    st.rerun()


def get_file():
    try:
        return supabase.storage.from_(BUCKET).download(FILE)
    except Exception:
        return None


def save_file(data):
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
        return True, ""
    except Exception as e:
        return False, str(e)


def delete_file():
    try:
        supabase.storage.from_(BUCKET).remove([FILE])
        return True, ""
    except Exception as e:
        return False, str(e)


# ============================================================
# LOGIN
# ============================================================

if st.session_state.user is None:

    st.title("6thSense Excel Manager")
    st.subheader("Login")

    with st.form("login"):

        email = st.text_input("Email")
        password = st.text_input(
            "Password",
            type="password"
        )

        submit = st.form_submit_button(
            "Login",
            use_container_width=True
        )

    if submit:

        if not email or not password:
            st.error("Enter email and password.")

        else:
            ok, msg = login(email, password)

            if ok:
                st.rerun()
            else:
                st.error(msg)

    st.stop()


# ============================================================
# USER
# ============================================================

user = st.session_state.user

st.title("6thSense Excel Manager")
st.caption(f"Logged in as: {user.email}")

if admin():
    st.success("ADMIN ACCESS")
else:
    st.info("USER ACCESS — Read only")


with st.sidebar:

    st.write(f"**User:** {user.email}")
    st.write(f"**Role:** {'Admin' if admin() else 'User'}")

    if st.button(
        "Logout",
        use_container_width=True
    ):
        logout()


# ============================================================
# MASTER FILE
# ============================================================

st.header("Master Excel File")

data = get_file()


# ============================================================
# NO FILE
# ============================================================

if data is None:

    st.warning("No master Excel file available.")

    if admin():

        upload = st.file_uploader(
            "Upload Master Excel",
            type=["xlsx"]
        )

        if upload and st.button(
            "Upload Master File",
            use_container_width=True
        ):

            ok, msg = save_file(upload.getvalue())

            if ok:
                st.success("File uploaded.")
                st.rerun()
            else:
                st.error(msg)

    st.stop()


# ============================================================
# DOWNLOAD
# ============================================================

st.success("Master Excel file available.")

st.download_button(
    "⬇️ Download Master Excel",
    data=data,
    file_name="master.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    use_container_width=True
)


# ============================================================
# ADMIN
# ============================================================

if admin():

    st.divider()
    st.header("Admin Controls")

    # Replace
    st.subheader("Replace Master Excel")

    new_file = st.file_uploader(
        "Choose new Excel file",
        type=["xlsx"],
        key="replace"
    )

    if new_file and st.button(
        "Replace Master File",
        use_container_width=True
    ):

        ok, msg = save_file(new_file.getvalue())

        if ok:
            st.success("Master file replaced.")
            st.rerun()
        else:
            st.error(msg)

    # Delete
    st.divider()
    st.subheader("Delete Master File")

    st.warning(
        "This will remove the file for all users."
    )

    if st.button(
        "🗑️ Delete Master File",
        use_container_width=True
    ):

        ok, msg = delete_file()

        if ok:
            st.success("Master file deleted.")
            st.rerun()
        else:
            st.error(msg)

    # Modify
    st.divider()
    st.subheader("Modify Excel")

    try:

        import pandas as pd

        xls = pd.ExcelFile(
            io.BytesIO(data),
            engine="openpyxl"
        )

        sheet = st.selectbox(
            "Worksheet",
            xls.sheet_names
        )

        df = pd.read_excel(
            io.BytesIO(data),
            sheet_name=sheet,
            engine="openpyxl"
        )

        edited = st.data_editor(
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

                edited.to_excel(
                    writer,
                    sheet_name=sheet,
                    index=False
                )

                for s in xls.sheet_names:

                    if s == sheet:
                        continue

                    other = pd.read_excel(
                        io.BytesIO(data),
                        sheet_name=s,
                        engine="openpyxl"
                    )

                    other.to_excel(
                        writer,
                        sheet_name=s,
                        index=False
                    )

            ok, msg = save_file(
                output.getvalue()
            )

            if ok:
                st.success("Excel saved.")
                st.rerun()
            else:
                st.error(msg)

    except Exception as e:
        st.error(f"Excel error: {e}")

else:

    st.info(
        "You have read-only access. "
        "You can view and download the master Excel file."
      )
