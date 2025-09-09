import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# --- Email (SMTP) settings (hard-coded) ---
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587  # use 465 if SSL
USE_TLS = True   # set False if using SSL instead of TLS

# --- App title ---
st.title("Email Automation Tool")

# --- Upload CSV ---
st.subheader("Upload recipient list")
uploaded_file = st.file_uploader("Upload CSV file", type=["csv"], key="csv_uploader")
df = None
if uploaded_file:
    df = pd.read_csv(uploaded_file)
    st.dataframe(df)

# --- Email configuration (frontend only) ---
st.subheader("Email configuration")
from_email = st.text_input("Your email address", key="from_email")
app_password = st.text_input("App password", type="password", key="app_password")
from_name = st.text_input("Your name (optional)", key="from_name")

# --- Proposal details ---
st.subheader("Cost Associated")
currency = st.selectbox("Currency", ["USD", "AED"], key="currency_select")
cost = st.number_input(f"Cost in {currency}", min_value=0.0, step=50.0, value=1000.0, key="cost_input")

# --- Compose message ---
st.subheader("Compose message")

# Subject line options
subject_options = [
    "Special proposal for {company}",
    "Collaboration opportunity with {company}",
    "Exclusive offer for {name}",
    "Your personalized proposal from {sender}"
]
subject_tpl = st.selectbox("Choose a subject line", subject_options, key="subject_select")

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
body_choice = st.selectbox("Choose a body template", list(body_templates.keys()), key="body_template_select")
body_tpl = st.text_area("Body", value=body_templates[body_choice], height=250, key="body_text")

# --- Send emails ---
if st.button("Send Emails", key="send_emails_btn"):
    if not from_email or not app_password:
        st.error("Please provide your email and app password.")
    elif df is None:
        st.error("Please upload a CSV file with recipients.")
    else:
        progress = st.progress(0)
        skipped = []  # collect skipped recipients

        for idx, row in df.iterrows():
            rowd = row.to_dict()

            # Skip rows without email
            if not rowd.get("email"):
                skipped.append(rowd.get("name", f"Row {idx}"))
                continue

            # Fill placeholders with defaults
            rowd.setdefault("sender", from_name or from_email)
            rowd.setdefault("cost", cost)
            rowd.setdefault("currency", currency)
            rowd.setdefault("company", "your company")
            rowd.setdefault("name", "there")

            subj = subject_tpl.format(**rowd)
            body = body_tpl.format(**rowd)

            msg = MIMEMultipart()
            msg["From"] = f"{from_name or from_email} <{from_email}>"
            msg["To"] = rowd["email"]
            msg["Subject"] = subj
            msg.attach(MIMEText(body, "plain"))

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

                st.success(f"Sent to {rowd['email']}")
            except Exception as e:
                st.error(f"Failed to send to {rowd['email']}: {e}")
            
            progress.progress((idx + 1) / len(df))

        # Show skipped recipients
        if skipped:
            st.warning("Skipped recipients (no email found):")
            st.write(", ".join(skipped))
