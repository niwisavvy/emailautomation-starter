import io
import re
import unicodedata
import time
from collections import defaultdict
from email.header import Header
from email.utils import formataddr, parseaddr
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import smtplib
import ssl

import streamlit as st
import pandas as pd

# --- Streamlit page config ---
st.set_page_config(page_title="Email Automation 360", page_icon="📧")
st.title("📧 Email Automation 360 — Hardcoded SMTP")

# --- Hardcoded SMTP config ---
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT = 465

st.subheader("Email ID")
FROM_EMAIL = st.text_input("Your Email ID")  # <-- replace with your email

# --- SMTP password input only ---
st.subheader("SMTP password / App Password")
smtp_pass = st.text_input("SMTP password / App Password", type="password")

# --- Pause configuration ---
st.write("")
pause = 200.0  # 200 seconds pause between emails

# --- Upload recipients ---
st.subheader("Recipients CSV")
st.caption("Upload a CSV with at least an `email` column; extra columns (name, company) become placeholders.")
file = st.file_uploader("Upload CSV", type=["csv"])
df = None
if file:
    try:
        df = pd.read_csv(file)
        st.write(df.head())
    except Exception as e:
        st.error(f"Could not read CSV: {e}")

# --- Compose message ---
st.subheader("Compose message")
subject_options = [
    "Special proposal for {company}",
    "Collaboration opportunity with {company}",
    "Exclusive offer for {name}",
    "Your personalized proposal from {sender}"
]
subject_tpl = st.selectbox("Choose a subject line", subject_options)

# Cost input
st.subheader("Proposal details")
currency = st.selectbox("Currency", ["USD", "AED"])
cost = st.number_input(f"Cost in {currency}", min_value=0.0, step=10.0, value=1000.0)

# Body templates
body_templates = {
    "Proposal (standard)": (
        "Hi {name},\n\n"
        "I’m reaching out with a tailored proposal for {company}. "
        "Our solution is designed to add real value, and we can offer this at "
        "{cost} {currency}.\n\n"
        "Let me know if this works for you, and I’d be happy to discuss further.\n\n"
        "Best regards,\n{sender}"
    ),
    "Follow-up (gentle reminder)": (
        "Hi {name},\n\n"
        "I just wanted to follow up on my earlier message about {company}. "
        "This opportunity is still available for {cost} {currency}, "
        "and I’d love to hear your thoughts.\n\n"
        "Best regards,\n{sender}"
    ),
    "Short intro (very concise)": (
        "Hi {name},\n\n"
        "Quick note to share a proposal for {company}: {cost} {currency}. "
        "Would you like to discuss?\n\n"
        "Cheers,\n{sender}"
    )
}
st.subheader("Message body")
body_choice = st.selectbox("Choose a body template", list(body_templates.keys()))
body_tpl = st.text_area("Body", value=body_templates[body_choice], height=250)

# ---------------- Helpers ----------------
def clean_value(val):
    """Clean cell values: remove hidden chars and non-ASCII."""
    if isinstance(val, str):
        val = val.replace("\xa0", " ").replace("\u200b", "").strip()
        val = unicodedata.normalize("NFKD", val)
        val = "".join(ch for ch in val if ord(ch) < 128)
        return val
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

# ---------------- Preview ----------------
if df is not None and not df.empty:
    st.subheader("Preview (first row)")
    first = df.iloc[0].to_dict()
    first.setdefault("sender", FROM_EMAIL)
    st.markdown("**Subject preview:**")
    st.write(safe_format(subject_tpl, first))
    st.markdown("**Body preview:**")
    st.write(safe_format(body_tpl, first))

# ---------------- Send Emails ----------------
if st.button("Send Emails"):
    if not smtp_pass:
        st.error("Please provide your SMTP password / App password.")
        st.stop()
    if df is None or df.empty:
        st.error("Please upload a CSV with recipients.")
        st.stop()

    progress = st.progress(0)
    total = len(df)
    logs = []

    for idx, row in df.iterrows():
        rowd = {str(k): clean_value(v) for k, v in row.to_dict().items()}
        rowd.setdefault("sender", FROM_EMAIL)
        rowd.setdefault("cost", str(cost))
        rowd.setdefault("currency", currency)
        rowd.setdefault("company", "")
        rowd.setdefault("name", "")

        recip_addr = clean_email_address(rowd.get("email", ""))
        if not recip_addr:
            logs.append({**rowd, "__status": "skipped", "__reason": "invalid email"})
            progress.progress((idx + 1) / total)
            continue

        subj_text = safe_format(subject_tpl, rowd)
        body_text = safe_format(body_tpl, rowd)

        # Compose email
        msg = MIMEMultipart()
        msg["From"] = FROM_EMAIL
        to_name_clean = clean_value(rowd.get("name", ""))
        if to_name_clean:
            msg["To"] = formataddr((str(Header(to_name_clean, "utf-8")), recip_addr))
        else:
            msg["To"] = recip_addr
        msg["Subject"] = str(Header(subj_text, "utf-8"))
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        try:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, context=context) as server:
                server.login(FROM_EMAIL, smtp_pass)
                server.send_message(msg)
            st.success(f"Sent to {recip_addr}")
            logs.append({**rowd, "__status": "sent"})
        except Exception as e:
            st.error(f"Failed to send to {recip_addr}: {e}")
            logs.append({**rowd, "__status": "failed", "__reason": str(e)})

        progress.progress((idx + 1) / total)
        time.sleep(pause)

    st.success("Done sending emails.")
    logs_df = pd.DataFrame(logs)
    st.dataframe(logs_df)
    csv = logs_df.to_csv(index=False).encode("utf-8")
    st.download_button("Download send log (CSV)", data=csv, file_name="send_log.csv", mime="text/csv")
