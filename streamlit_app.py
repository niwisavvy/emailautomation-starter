import streamlit as st
import pandas as pd
import time
import smtplib
import ssl
from email.message import EmailMessage

st.set_page_config(page_title="Email Automation 360", page_icon="📧")

st.title("📧 Email Automation 360 — Starter (SMTP)")

# --- safe defaults from Streamlit secrets (when deployed) ---
smtp_host_default = st.secrets.get("SMTP_HOST", "smtp.gmail.com") if hasattr(st, "secrets") else "smtp.gmail.com"
smtp_port_default = int(st.secrets.get("SMTP_PORT", 465)) if hasattr(st, "secrets") else 465
smtp_user_default = st.secrets.get("SMTP_USER", "") if hasattr(st, "secrets") else ""
smtp_from_default = st.secrets.get("SMTP_FROM", smtp_user_default) if hasattr(st, "secrets") else smtp_user_default
smtp_pass_default = st.secrets.get("SMTP_PASS", "") if hasattr(st, "secrets") else ""

# --- Email account setup ---
st.subheader("SMTP settings (use Streamlit Secrets in cloud for safety)")
smtp_host = st.text_input("SMTP host", value=smtp_host_default)
smtp_port = st.number_input("SMTP port", value=smtp_port_default, step=1)
smtp_user = st.text_input("SMTP username (your email)", value=smtp_user_default)
smtp_pass = st.text_input("SMTP password / App Password", type="password", value=smtp_pass_default)
from_name = st.text_input("From name", value="")
from_email = st.text_input("From email", value=smtp_from_default)

# --- Test mode & throttle ---
st.write("")
test_mode = st.checkbox("TEST MODE — send all messages to this address", value=True)
test_email = st.text_input("Test recipient email (used in TEST MODE)", value=smtp_user or from_email)
pause = st.slider("Pause between emails (seconds)", 0.0, 60.0, 10.0)

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

# Subject line options
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
cost = st.number_input(f"Cost in {currency}", min_value=0.0, step=10.0, value=100.0)

# Body template options (predefined)
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

# Choose body template
st.subheader("Message body")
body_choice = st.selectbox("Choose a body template", list(body_templates.keys()))
body_tpl = st.text_area("Body", value=body_templates[body_choice], height=250)


# --- helpers ---
def render(tpl: str, row: dict) -> str:
    try:
        return tpl.format(**row)
    except Exception:
        return tpl

def compose_message(to_email, subject, body):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = f"{from_name} <{from_email}>" if from_name else from_email
    msg["To"] = to_email
    msg.set_content(body)
    return msg

def send_email(msg):
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL(smtp_host, smtp_port, context=context) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)

# --- Preview ---
if df is not None and not df.empty:
    st.subheader("Preview (first row)")
    first = df.iloc[0].to_dict()
    first.setdefault("sender", from_name or from_email)
    st.markdown("**Subject preview:**")
    st.write(render(subject_tpl, first))
    st.markdown("**Body preview:**")
    st.write(render(body_tpl, first))

# --- Send action ---
if st.button("🚀 Send emails"):
    if df is None or df.empty:
        st.error("Please upload a CSV with at least one recipient.")
    elif not smtp_user or not smtp_pass:
        st.error("SMTP username and password are required.")
    else:
        logs = []
        total = len(df)
        progress = st.progress(0)
        for i, row in df.iterrows():
            rowd = row.to_dict()
            rowd.setdefault("sender", from_name or from_email)
            rowd.setdefault("cost", cost)
            rowd.setdefault("currency", currency)

            to_addr = test_email if test_mode else rowd.get("email")
            subject = render(subject_tpl, rowd)
            body = render(body_tpl, rowd)
            msg = compose_message(to_addr, subject, body)
            try:
                send_email(msg)
                logs.append({"email": to_addr, "status": "SENT"})
                st.write(f"Sent to: {to_addr}")
            except Exception as e:
                logs.append({"email": to_addr, "status": f"ERROR: {e}"})
                st.error(f"Error sending to {to_addr}: {e}")
            time.sleep(pause)
            progress.progress(int((i+1)/total*100))
        st.success("Done sending.")
        logs_df = pd.DataFrame(logs)
        st.dataframe(logs_df)
        csv = logs_df.to_csv(index=False).encode("utf-8")
        st.download_button("Download send log (CSV)", data=csv, file_name="send_log.csv", mime="text/csv")
