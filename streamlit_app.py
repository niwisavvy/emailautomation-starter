import io
import unicodedata
from collections import defaultdict

import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.header import Header
from email.utils import formataddr

# --- Email (SMTP) settings (hard-coded) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587  # use 465 if SSL
USE_TLS = True   # set False if using SSL instead of TLS

st.set_page_config(page_title="Email Automation Tool")

# --- Helpers ---------------------------------------------------------------
def clean_text(value):
    """
    Convert value to str, replace NBSP with normal space, normalize unicode, strip.
    Returns '' for NaN/None.
    """
    if pd.isna(value):
        return ""
    if not isinstance(value, str):
        value = str(value)
    # remove non-breaking spaces and normalize
    value = value.replace("\xa0", " ")
    value = unicodedata.normalize("NFKC", value)
    return value.strip()

def safe_format(template, mapping):
    """
    Format using format_map with a default dict so missing keys become empty strings.
    """
    return template.format_map(defaultdict(str, mapping))

# --- UI: title & upload ----------------------------------------------------
st.title("Email Automation Tool")

st.subheader("Upload recipient list")
uploaded_file = st.file_uploader("Upload CSV file", type=["csv"], key="csv_uploader")

# Provide a downloadable sample CSV so users know the required headers
sample_data = pd.DataFrame({
    "email": ["john.doe@example.com", "jane.smith@example.com"],
    "name": ["John Doe", "Jane Smith"],
    "company": ["Acme Corp", "Globex Inc"]
})
csv_buffer = io.StringIO()
sample_data.to_csv(csv_buffer, index=False)
st.download_button(
    label="📥 Download sample CSV",
    data=csv_buffer.getvalue(),
    file_name="sample_recipients.csv",
    mime="text/csv",
    key="download_sample_csv"
)

df = None
if uploaded_file:
    try:
        # try common encodings if user uploaded a file with non-utf8
        df = pd.read_csv(uploaded_file)
    except Exception:
        try:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding="latin1")
        except Exception as e:
            st.error(f"Couldn't read CSV: {e}")
            df = None

    if df is not None:
        st.dataframe(df)

# --- Email configuration (frontend only) -----------------------------------
st.subheader("Email configuration")
from_email = clean_text(st.text_input("Your email address", key="from_email"))
app_password = st.text_input("App password", type="password", key="app_password")
from_name = clean_text(st.text_input("Your name (optional)", key="from_name"))

# --- Proposal details (cost/currency) --------------------------------------
st.subheader("Cost Associated")
currency = st.selectbox("Currency", ["USD", "AED"], key="currency_select")
cost = st.number_input(f"Cost in {currency}", min_value=0.0, step=50.0, value=1000.0, key="cost_input")

# --- Compose message -------------------------------------------------------
st.subheader("Compose message")

subject_options = [
    "Special proposal for {company}",
    "Collaboration opportunity with {company}",
    "Exclusive offer for {name}",
    "Your personalized proposal from {sender}"
]
subject_tpl = st.selectbox("Choose a subject line", subject_options, key="subject_select")

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
body_choice = st.selectbox("Choose a body template", list(body_templates.keys()), key="body_template_select")
body_tpl = st.text_area("Body", value=body_templates[body_choice], height=250, key="body_text")

# --- Send emails -----------------------------------------------------------
if st.button("Send Emails", key="send_emails_btn"):
    # basic checks
    if not from_email or not app_password:
        st.error("Please provide your email and app password.")
        st.stop()
    if df is None:
        st.error("Please upload a CSV file with recipients.")
        st.stop()

    # sanitize column names to be safe strings
    df.columns = [clean_text(c) for c in df.columns]

    progress = st.progress(0)
    skipped_rows = []   # store dicts for skipped rows so we can offer a download
    total = len(df)
    sent = 0

    for idx, row in df.iterrows():
        # convert row to dict and clean every value
        raw_rowd = row.to_dict()
        rowd = {str(k): clean_text(v) for k, v in raw_rowd.items()}

        # skip if no email
        recipient_email = rowd.get("email", "")
        if not recipient_email:
            skipped_rows.append(rowd)
            continue

        # set safe defaults
        rowd.setdefault("sender", from_name or from_email)
        rowd.setdefault("cost", str(cost))
        rowd.setdefault("currency", currency)
        rowd.setdefault("company", "")
        rowd.setdefault("name", "")

        # safe format (missing keys become empty strings)
        subj_text = safe_format(subject_tpl, rowd)
        body_text = safe_format(body_tpl, rowd)

        # encode headers & set From/To properly (handles non-ascii display names)
        msg = MIMEMultipart()
        # From: use display name encoded if provided
        display_from_name = rowd.get("sender") or from_email
        msg["From"] = formataddr((str(Header(display_from_name, "utf-8")), from_email))

        # To: include recipient name if provided, encoded
        recipient_name = rowd.get("name", "")
        if recipient_name:
            msg["To"] = formataddr((str(Header(recipient_name, "utf-8")), recipient_email))
        else:
            msg["To"] = recipient_email

        # Subject encoded
        msg["Subject"] = str(Header(subj_text, "utf-8"))

        # attach body with utf-8 encoding
        msg.attach(MIMEText(body_text, "plain", "utf-8"))

        # send
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
            st.success(f"Sent to {recipient_email}")
        except Exception as e:
            # show error but continue sending others
            st.error(f"Failed to send to {recipient_email}: {e}")

        # update progress
        progress.progress((idx + 1) / total)

    # summary
    st.info(f"Done — attempted {total} rows, sent {sent} messages, skipped {len(skipped_rows)} rows (no email).")

    # offer skipped rows as CSV for download so user can fix and reupload
    if skipped_rows:
        skipped_df = pd.DataFrame(skipped_rows)
        buf = io.StringIO()
        skipped_df.to_csv(buf, index=False)
        st.download_button(
            label="📥 Download skipped rows (fix and re-upload)",
            data=buf.getvalue(),
            file_name="skipped_recipients.csv",
            mime="text/csv",
            key="download_skipped"
        )
