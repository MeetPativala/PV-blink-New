import os
import requests
import streamlit as st


# ============================================================
# PV-BLINK LOGIN FORM
# ============================================================
# Confirmed PV-Blink Login API. It can still be overridden for testing.
#
# PowerShell example:
#   $env:PVBLINK_LOGIN_API_URL="https://solar.e-khata.com/api/acct-crm/user/login"
#
# IMPORTANT:
# Do not hard-code real passwords inside this file.
# ============================================================

LOGIN_API_URL = (
    os.getenv("PVBLINK_LOGIN_API_URL", "").strip()
    or "https://solar.e-khata.com/api/acct-crm/user/login"
)


def login_with_api(username: str, password: str):
    """
    Call your Login API.

    Request JSON:
        {
            "email": "...",
            "password": "..."
        }

    The confirmed API returns the authenticated user in ``data`` with
    ``error: false`` and the token in ``data.accessToken``.
    """
    payload = {
        "email": username,
        "password": password,
    }

    headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://solar.e-khata.com",
        "Referer": "https://solar.e-khata.com/login",
    }

    try:
        response = requests.post(
            LOGIN_API_URL,
            json=payload,
            headers=headers,
            timeout=20,
        )

    except requests.exceptions.ConnectionError:
        return False, "Login server is not reachable.", None

    except requests.exceptions.Timeout:
        return False, "Login request timed out.", None

    except requests.exceptions.RequestException:
        return False, "Login request failed. Check the login API configuration.", None

    try:
        data = response.json()
    except ValueError:
        return False, "Login API returned an invalid response.", None

    if not isinstance(data, dict):
        return False, "Login API returned an unexpected response.", data

    response_data = data.get("data")
    if response.ok and isinstance(response_data, dict) and response_data.get("error") is False:
        access_token = response_data.get("accessToken")
        if access_token:
            return True, "Login successful.", data

    error_data = data.get("error")
    if isinstance(error_data, dict):
        message = (
            error_data.get("message")
            or error_data.get("debugMessage")
            or "Login failed."
        )

    else:
        message = (
            data.get("message")
            or data.get("detail")
            or "Invalid Username or Password."
        )

    if not response.ok:
        message = f"{message} (HTTP {response.status_code})"

    return False, message, data


def logout():
    """
    Clear only login-related session values.
    """
    for key in [
        "pv_authenticated",
        "pv_auth_token",
        "pv_user",
        "pv_login_response",
    ]:
        st.session_state.pop(key, None)

    st.rerun()


def login_form():
    """
    Render the login page.

    Returns:
        True  -> logged in
        False -> not logged in
    """

    if st.session_state.get("pv_authenticated", False):
        return True

    st.markdown(
        """
        <style>
        .block-container {
            max-width: 480px;
            padding-top: 7vh;
        }

        .pv-login-title {
            text-align: center;
            font-size: 2rem;
            font-weight: 800;
            margin-bottom: 0.25rem;
        }

        .pv-login-subtitle {
            text-align: center;
            opacity: 0.72;
            margin-bottom: 1.5rem;
        }

        div[data-testid="stForm"] {
            border: 1px solid rgba(255, 125, 0, 0.35);
            border-radius: 18px;
            padding: 1.4rem 1.35rem 1.2rem 1.35rem;
            box-shadow: 0 10px 35px rgba(0, 0, 0, 0.12);
        }

        div.stButton > button,
        div[data-testid="stFormSubmitButton"] > button {
            width: 100%;
            border-radius: 10px;
            font-weight: 700;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="pv-login-title">PV-BLINK</div>',
        unsafe_allow_html=True,
    )

    st.markdown(
        '<div class="pv-login-subtitle">Analytics Dashboard Login</div>',
        unsafe_allow_html=True,
    )

    with st.form(
        "pv_login_form",
        clear_on_submit=False,
    ):
        username = st.text_input(
            "Email",
            placeholder="Enter email",
            autocomplete="username",
        )

        password = st.text_input(
            "Password",
            type="password",
            placeholder="Enter password",
            autocomplete="current-password",
        )

        submitted = st.form_submit_button(
            "Login",
            width="stretch",
        )

    if submitted:
        username = (username or "").strip()

        if not username or not password:
            st.error("Please enter Username and Password.")
            return False

        with st.spinner("Signing in..."):
            ok, message, data = login_with_api(
                username,
                password,
            )

        if not ok:
            st.error(message)
            return False

        response_data = data.get("data", {})
        token = response_data.get("accessToken", "")
        user_data = response_data or {"email": username}

        st.session_state["pv_authenticated"] = True
        st.session_state["pv_auth_token"] = token
        st.session_state["pv_user"] = user_data
        st.session_state["pv_login_response"] = data

        st.rerun()

    return False


def require_login():
    """
    Put this near the top of your main Streamlit app:

        if not require_login():
            st.stop()

    Then your normal dashboard code continues below it.
    """
    if not login_form():
        st.stop()

    return True


def login_sidebar_user():
    """
    Optional helper after successful login.
    Shows current user + Logout button in sidebar.
    """
    if not st.session_state.get("pv_authenticated", False):
        return

    user = st.session_state.get("pv_user", {})

    if isinstance(user, dict):
        display_name = (
            user.get("name")
            or user.get("username")
            or user.get("email")
            or "User"
        )
    else:
        display_name = "User"

    with st.sidebar:
        st.caption(f"Logged in as: {display_name}")

        if st.button(
            "Logout",
            width="stretch",
            key="pv_logout_button",
        ):
            logout()


# ============================================================
# STANDALONE TEST
# ============================================================
if __name__ == "__main__":
    st.set_page_config(
        page_title="PV-BLINK Login",
        page_icon="🔐",
        layout="centered",
    )

    require_login()

    login_sidebar_user()

    st.success("Login successful.")
