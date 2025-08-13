import streamlit as st
import pandas as pd
import re
import smtplib
import pdfplumber
import docx
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from io import StringIO

# -------------------- EMAIL EXTRACTION FUNCTION --------------------
def extract_emails(text):
    pattern = r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
    return re.findall(pattern, text)

# -------------------- READ FILE AND GET EMAILS --------------------
def process_file(file):
    emails = set()

    if file.name.endswith('.csv'):
        df = pd.read_csv(file)
        text = ' '.join(df.astype(str).fillna('').values.flatten())
        emails.update(extract_emails(text))

    elif file.name.endswith(('.xls', '.xlsx')):
        df = pd.read_excel(file)
        text = ' '.join(df.astype(str).fillna('').values.flatten())
        emails.update(extract_emails(text))

    elif file.name.endswith('.txt'):
        stringio = StringIO(file.getvalue().decode("utf-8"))
        text = stringio.read()
        emails.update(extract_emails(text))

    elif file.name.endswith('.pdf'):
        with pdfplumber.open(file) as pdf:
            for page in pdf.pages:
                text = page.extract_text() or ""
                emails.update(extract_emails(text))

    elif file.name.endswith('.docx'):
        doc = docx.Document(file)
        text = "\n".join([para.text for para in doc.paragraphs])
        emails.update(extract_emails(text))

    return list(emails)

# -------------------- SEND EMAIL FUNCTION WITH ATTACHMENT --------------------
def send_bulk_email_with_attachment(sender_email, sender_password, recipients, subject, message, resume_file):
    try:
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Gmail SMTP
        server.starttls()
        server.login(sender_email, sender_password)

        # Read resume file content
        resume_data = resume_file.read()
        resume_file.seek(0)  # Reset pointer so it can be reused

        for recipient in recipients:
            msg = MIMEMultipart()
            msg['From'] = sender_email
            msg['To'] = recipient
            msg['Subject'] = subject

            # Email body
            msg.attach(MIMEText(message, 'plain'))

            # Attach resume
            part = MIMEBase('application', 'octet-stream')
            part.set_payload(resume_data)
            encoders.encode_base64(part)
            part.add_header('Content-Disposition', f'attachment; filename="{resume_file.name}"')
            msg.attach(part)

            server.sendmail(sender_email, recipient, msg.as_string())

        server.quit()
        return True
    except Exception as e:
        st.error(f"Error sending emails: {e}")
        return False

# -------------------- STREAMLIT FRONTEND --------------------
st.title("📧 HR Bulk Email Sender + Resume Attachment")
st.write("Upload dataset → Extract Emails → Send Bulk Emails with Resume")

uploaded_files = st.file_uploader("Upload files", type=["csv", "xls", "xlsx", "pdf", "txt", "docx"], accept_multiple_files=True)
resume_file = st.file_uploader("Upload your Resume", type=["pdf", "docx", "txt"])

all_emails = set()

if uploaded_files:
    for file in uploaded_files:
        extracted = process_file(file)
        all_emails.update(extracted)
    st.success(f"Found {len(all_emails)} email(s).")
    st.write(all_emails)

st.subheader("✉ Email Details")
sender_email = st.text_input("Your Email (Gmail recommended)")
sender_password = st.text_input("Your Gmail App Password", type="password")
subject = st.text_input("Subject")
message = st.text_area("Message")

if st.button("Send Bulk Emails"):
    if not all_emails:
        st.error("No emails found to send.")
    elif not sender_email or not sender_password or not subject or not message or not resume_file:
        st.error("Please fill all fields and upload a resume.")
    else:
        success = send_bulk_email_with_attachment(sender_email, sender_password, list(all_emails), subject, message, resume_file)
        if success:
            st.success(f"Emails with resume sent successfully to {len(all_emails)} recipients!")
