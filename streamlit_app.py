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

st.set_page_config(page_title="Team Niwrutti")

# --- SMTP Settings (Gmail by default) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
USE_TLS = True

# ---------------- Helpers ----------------
def clean_value(val):
    if isinstance(val, str):
        return (
            val.replace("\xa0", " ")
               .replace("\u200b", "")
               .strip()
        )
    return val

def clean_email_address(raw_email: str) -> str | None:
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
    return template.format_map(defaultdict(str, mapping))

def clean_display_name(name: str) -> str:
    if not name:
        return ""
    name = name.replace("\xa0", " ").replace("\u200b", "")
    return name.strip()

def clean_invisible_unicode(s: str) -> str:
    if not isinstance(s, str):
        return s
    return s.replace('\xa0', '').replace('\u200b', '').strip()

def format_first_name(full_name: str) -> str:
    if not full_name:
        return ""

    prefixes = {"dr", "dr.", "mr", "mr.", "mrs", "mrs.", "ms", "ms.", "prof", "prof."}
    parts = full_name.strip().split()

    if not parts:
        return ""

    parts = [p.capitalize() for p in parts]
    first_word = parts[0].lower()

    if first_word in prefixes and len(parts) > 1:
        prefix = parts[0].capitalize().replace(".", "")
        name = parts[1].capitalize()
        return f"{prefix} {name}"

    if len(parts[0]) == 1 and len(parts) > 1:
        last_name = parts[1].capitalize()
        return f"Mr {last_name}"

    return parts[0].capitalize()

# ---------------- Upload ----------------
st.title("Team Niwrutti")
st.subheader("Upload recipient list")
uploaded_file = st.file_uploader("Upload CSV file", type=["csv"])

df = None
if uploaded_file:
    df = pd.read_csv(uploaded_file).applymap(clean_value)
    st.dataframe(df)

# ---------------- UI ----------------
col1, col2, col3 = st.columns(3)

with col1:
    send_clicked = st.button("Send Emails")

with col2:
    stop_clicked = st.button("Stop Sending")

with col3:
    cooling_timer_placeholder = st.empty()

# ---------------- Config ----------------
from_email = clean_invisible_unicode(st.text_input("Your email"))
app_password = clean_invisible_unicode(st.text_input("App password", type="password"))
from_name = st.text_input("Your name")

cc_raw = st.text_input("CC (optional)")
cc_email = clean_email_address(cc_raw) if cc_raw else None

subject_tpl_1 = st.text_input("Subject 1")
subject_tpl_2 = st.text_input("Subject 2")

body_tpl_1 = st.text_area("Body 1", height=300)
body_tpl_2 = st.text_area("Body 2", height=300)

sent_count = 0
progress = st.progress(0)

# ---------------- Send ----------------
if send_clicked and df is not None:

    total = len(df)
    sent = 0
    skipped_rows = []
    failed_rows = []

    for idx, row in df.iterrows():

        rowd = row.to_dict()
        recip_addr = clean_email_address(rowd.get("email", ""))

        if not recip_addr:
            skipped_rows.append(rowd)
            continue

        full_name = rowd.get("name", "")
        first_name = format_first_name(full_name)

        subject = subject_tpl_1 if sent % 2 == 0 else subject_tpl_2
        body = body_tpl_1 if sent % 2 == 0 else body_tpl_2

        body = safe_format(body, {"name": first_name})

        html_body = f"""
        <html><body><pre>{body}</pre></body></html>
        """

        msg = MIMEMultipart()
        msg["From"] = formataddr((from_name, from_email))
        msg["To"] = recip_addr
        msg["Subject"] = subject
        msg.attach(MIMEText(html_body, "html"))

        try:
            with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
                server.starttls()
                server.login(from_email, app_password)
                server.send_message(msg)

            sent += 1
            st.success(f"Sent to {recip_addr}")

        except Exception as e:
            failed_rows.append(rowd)
            st.error(f"Failed: {recip_addr}")

        progress.progress((idx + 1) / total)
        time.sleep(2)

    st.success(f"Done. Sent: {sent}")
