import io
import re
import time
from collections import defaultdict
from email.header import Header
from email.utils import formataddr, parseaddr

import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

st.set_page_config(page_title="Email Automation Tool")

# --- SMTP Settings (Gmail by default) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
USE_TLS = True

# ---------------- Helpers ----------------
def clean_value(val):
    """Clean individual cell values (remove invisible characters)."""
    if isinstance(val, str):
        return (
            val.replace("\xa0", " ")      # non-breaking space
               .replace("\u200b", "")     # zero-width space
               .strip()
        )
    return val

def clean_email_address(raw_email: str) -> str | None:
    """Parse and sanitize an email address string."""
    if not raw_email:
        return None
    raw_email = clean_value(raw_email)
    _, addr = parseaddr(raw_email)
    if not addr:
        addr = re.sub(r"[<>\s\"']", "", raw_email)
    addr = addr.strip()
    if "@" not in addr:
        return None
    try:
        local, domain = addr.rsplit("@", 1)
    except ValueError:
        return None
    try:
        domain_ascii = domain.encode("idna").decode("ascii")
    except Exception:
        domain_ascii = "".join(ch for ch in domain if ord(ch) < 128)
    return f"{local}@{domain_ascii}"

def safe_format(template: str, mapping: dict) -> str:
    """Format template safely with missing keys allowed."""
    return template.format_map(defaultdict(str, mapping))

def strip_non_ascii(s: str) -> str:
    """Remove non-ASCII characters safely."""
    if not isinstance(s, str):
        return s
    return ''.join(ch if ord(ch) < 128 else ' ' for ch in s)

# Initialize session state flags
if "stop_sending" not in st.session_state:
    st.session_state.stop_sending = False
if "is_sending" not in st.session_state:
    st.session_state.is_sending = False

# ---------------- Upload & Sample CSV ----------------
st.title("Email Automation Tool")
st.subheader("Upload recipient list")
uploaded_file = st.file_uploader("Upload CSV file", type=["csv"], key="csv_uploader")

# Sample CSV
sample_df = pd.DataFrame({
    "email": ["john.doe@example.com", "jane.smith@example.com"],
    "name": ["John Doe", "Jane Smith"],
    "company": ["Acme Corp", "Globex Inc"]
})
buf = io.StringIO()
sample_df.to_csv(buf, index=False)
st.download_button("📥 Download sample CSV", data=buf.getvalue(),
                   file_name="sample_recipients.csv", mime="text/csv",
                   key="download_sample_csv")

df = None
if uploaded_file:
    try:
        df = pd.read_csv(uploaded_file)
    except Exception:
        uploaded_file.seek(0)
        try:
            df = pd.read_csv(uploaded_file, encoding="latin1")
        except Exception as e:
            st.error(f"Couldn't read CSV: {e}")
            df = None
    if df is not None:
        # Clean entire DataFrame
        df = df.applymap(clean_value)
        df.columns = [clean_value(c) for c in df.columns]
        st.success("CSV uploaded and cleaned successfully ✅")
        st.dataframe(df)

# ---------------- Email Config ----------------
st.subheader("Email configuration")

from_email = st.text_input(
    "Your email address", key="from_email", disabled=st.session_state.is_sending
)
app_password = st.text_input(
    "App password", type="password", key="app_password", disabled=st.session_state.is_sending
)
from_name = st.text_input(
    "Your name (optional)", key="from_name", disabled=st.session_state.is_sending
)

# ---------------- Compose Message ----------------
subject_tpl = st.text_input(
    "Enter subject line template",
    placeholder="Enter your subject line (you can use placeholders like {name}, {company}, {sender})",
    value="",
    disabled=st.session_state.is_sending
)

body_tpl = st.text_area(
    "Enter body template",
    placeholder="Enter your email body (you can use placeholders like {name}, {company}, {sender})",
    height=250,
    value="",
    disabled=st.session_state.is_sending
)

# ---------------- Send & Stop Buttons ----------------
col1, col2 = st.columns(2)

with col1:
    send_clicked = st.button(
        "🚀 Send Emails", key="send_emails_btn", disabled=st.session_state.is_sending
    )

with col2:
    stop_clicked = st.button(
        "🛑 Stop Sending", key="stop_sending_btn", disabled=not st.session_state.is_sending
    )

# Handle button clicks and update sending state
if send_clicked:
    st.session_state.is_sending = True
    st.session_state.stop_sending = False  # reset stop flag

if stop_clicked:
    st.session_state.stop_sending = True

# Email sending loop
if st.session_state.is_sending:
    if not from_email or not app_password:
        st.error("Please provide your email and app password.")
        st.session_state.is_sending = False
        st.stop()
    if df is None:
        st.error("Please upload a CSV file with recipients.")
        st.session_state.is_sending = False
        st.stop()

    if not subject_tpl.strip():
        st.error("Subject line template cannot be empty.")
        st.session_state.is_sending = False
        st.stop()
    if not body_tpl.strip():
        st.error("Body template cannot be empty.")
        st.session_state.is_sending = False
        st.stop()

    progress = st.progress(0)
    total = len(df)
    sent = 0
    skipped_rows = []
    failed_rows = []

    for idx, row in df.iterrows():
        if st.session_state.get("stop_sending", False):
            st.warning("🛑 Email sending stopped by user.")
            break

        rowd = {str(k): clean_value(v) for k, v in row.to_dict().items()}

        # Validate recipient email
        recip_addr = clean_email_address(rowd.get("email", ""))
        if not recip_addr:
            skipped_rows.append({**rowd, "__reason": "missing/invalid email"})
            progress.progress((idx + 1) / total)
            continue

        # Defaults
        rowd.setdefault("sender", from_name)
        rowd.setdefault("company", "")
        rowd.setdefault("name", "")

        subj_text = strip_non_ascii(safe_format(subject_tpl, rowd))
        body_text = strip_non_ascii(safe_format(body_tpl, rowd))

        # Build message
        msg = MIMEMultipart()
        from_display = clean_value(from_name or "")
        to_display = clean_value(rowd.get("name", "") or "")

        from_display = strip_non_ascii(from_display)
        to_display = strip_non_ascii(to_display)

        from_header = formataddr((str(Header(from_display, "utf-8")), from_email))
        to_header = formataddr((str(Header(to_display, "utf-8")), recip_addr))

        msg["From"] = from_header
        msg["To"] = to_header
        msg["Subject"] = str(Header(subj_text, "utf-8"))
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        try:
            if USE_TLS:
                with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                    server.starttls()
                    server.login(from_email, app_password)
                    server.send_message(msg)
            else:
                with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as server:
                    server.login(from_email, app_password)
                    server.send_message(msg)

            sent += 1
            st.success(f"✅ Sent to {recip_addr}")
        except Exception as e:
            st.error(f"❌ Failed to send to {recip_addr}: {e}")
            failed_rows.append({**rowd, "__reason": str(e)})

        progress.progress((idx + 1) / total)

        # --- ⏳ Wait 20s before next email ---
        if idx < total - 1:
            wait_time = 20
            countdown_placeholder = st.empty()
            for remaining in range(wait_time, 0, -1):
                countdown_placeholder.info(f"⏳ Waiting {remaining} seconds before next email...")
                time.sleep(1)
            countdown_placeholder.empty()

    st.info(f"✅ Done — attempted {total}, sent {sent}, skipped {len(skipped_rows)}, failed {len(failed_rows)}")
    st.info(f"📧 Total Emails Successfully Sent: {sent}")

    # Download skipped/failed rows if any
    if skipped_rows:
        skipped_df = pd.DataFrame(skipped_rows)
        buf_skipped = io.StringIO()
        skipped_df.to_csv(buf_skipped, index=False)
        st.download_button("📥 Download skipped rows", data=buf_skipped.getvalue(),
                           file_name="skipped_recipients.csv", mime="text/csv",
                           key="download_skipped")
    if failed_rows:
        failed_df = pd.DataFrame(failed_rows)
        buf_failed = io.StringIO()
        failed_df.to_csv(buf_failed, index=False)
        st.download_button("📥 Download failed rows", data=buf_failed.getvalue(),
                           file_name="failed_recipients.csv", mime="text/csv",
                           key="download_failed")

    # Reset sending state after done
    st.session_state.is_sending = False
