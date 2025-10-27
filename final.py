import streamlit as st
import smtplib
import pandas as pd
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

def create_sample_excel_dynamic():
    sample_data = {
        "Name": ["Anushaa", "Abi"],
        "Email": ["anushaa@example.com", "abi@example.com"],
        "Attachment1": ["", "C:/users/yourname/file1.pdf"],
        "Attachment2": ["", "C:/users/yourname/file2.docx"]
    }
    df = pd.DataFrame(sample_data)
    df.to_excel("sample_dynamic_recipients.xlsx", index=False)

def create_sample_excel_single():
    sample_data = {
        "Name": ["Anushaa", "Abi"],
        "Email": ["anushaa@example.com", "abi@example.com"],
    }
    df = pd.DataFrame(sample_data)
    df.to_excel("sample_single_recipients.xlsx", index=False)

def send_email_single(to_email, name, attachment_paths, email_body, email_subject, outlook_user, outlook_password):
    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(outlook_user, outlook_password)
        
        msg = MIMEMultipart()
        msg["From"] = outlook_user
        msg["To"] = to_email
        msg["Subject"] = email_subject if email_subject else "No Subject"
        personalized_body = email_body.replace("[Name]", name) if email_body else ""
        msg.attach(MIMEText(personalized_body, "plain"))
        
        for file_path in attachment_paths:
            if os.path.exists(file_path):
                with open(file_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                encoders.encode_base64(part)
                part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
                msg.attach(part)
        
        server.sendmail(outlook_user, to_email, msg.as_string())
        server.quit()
        return f"✅ Email sent to {to_email} with {len(attachment_paths)} attachments."
    except Exception as e:
        return f"❌ Failed to send email to {to_email}. Error: {e}"
    
def send_email_dynamic(to_email, name, attachment_paths):
    try:
        server = smtplib.SMTP("smtp.office365.com", 587)
        server.starttls()
        server.login(outlook_user, outlook_password)
        
        msg = MIMEMultipart()
        msg["From"] = outlook_user
        msg["To"] = to_email
        msg["Subject"] = email_subject if email_subject else "No Subject"
        
        personalized_body = email_body.replace("[Name]", name) if email_body else ""
        msg.attach(MIMEText(personalized_body, "plain"))
        
        for file_path in attachment_paths:
            full_path = os.path.join(os.getcwd(), file_path)  # Convert to full path
            if os.path.exists(full_path):
                with open(full_path, "rb") as attachment:
                    part = MIMEBase("application", "octet-stream")
                    part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(file_path)}")
                    msg.attach(part)
            else:
                st.warning(f"⚠️ Attachment not found: {file_path}")
        
        server.sendmail(outlook_user, to_email, msg.as_string())
        server.quit()
        return f"✅ Email sent to {to_email} with {len(attachment_paths)} attachments."
    except Exception as e:
        return f"❌ Failed to send email to {to_email}. Error: {e}"

st.sidebar.title("📌 Email Sender Options")
app_mode = st.sidebar.radio("Select Mode", ["Same Attachments for All", "Dynamic Attachments"])

outlook_user = st.text_input("📧 Enter your Outlook Email Address")
outlook_password = st.text_input("🔒 Enter your Email Password", type="password")
email_subject = st.text_input("📌 Enter Email Subject (Optional)")
email_body = st.text_area("✉️ Enter Email Body (Optional)")

if st.sidebar.button("📥 Download Sample Excel File for Same Attachment"):
    create_sample_excel_single()
    with open("sample_single_recipients.xlsx", "rb") as file:
        st.sidebar.download_button("📂 Download sample_single_recipients.xlsx", file, file_name="sample_single_recipients.xlsx")

if st.sidebar.button("📥 Download Sample Excel File for Dynamic Attachment"):
    create_sample_excel_dynamic()
    with open("sample_dynamic_recipients.xlsx", "rb") as file:
        st.sidebar.download_button("📂 Download sample_dynamic_recipients.xlsx", file, file_name="sample_dynamic_recipients.xlsx")

if app_mode == "Same Attachments for All":
    uploaded_excel = st.file_uploader("📂 Upload Excel File", type=["xlsx"])
    uploaded_attachments = st.file_uploader("📎 Upload Attachments (Optional)", accept_multiple_files=True)
    
    if st.button("🚀 Send Emails"):
        if not outlook_user or not outlook_password:
            st.error("⚠️ Please enter your Outlook credentials.")
        elif not uploaded_excel:
            st.error("⚠️ Please upload an Excel file.")
        else:
            df = pd.read_excel(uploaded_excel)
            recipients = df.dropna(subset=["Email"]).to_dict(orient="records")
            
            attachment_paths = []
            if uploaded_attachments:
                for file in uploaded_attachments:
                    temp_path = os.path.join(os.getcwd(), file.name)
                    with open(temp_path, "wb") as f:
                        f.write(file.getbuffer())  # Save the file temporarily
                    attachment_paths.append(temp_path)

            results = []
            for recipient in recipients:
                results.append(send_email_single(recipient["Email"], recipient.get("Name", "User"), attachment_paths, email_body, email_subject, outlook_user, outlook_password))
            for res in results:
                st.write(res)

elif app_mode == "Dynamic Attachments":
    uploaded_excel = st.file_uploader("📂 Upload Excel File", type=["xlsx"])
    
    if st.button("🚀 Send Emails"):
        if not outlook_user or not outlook_password:
            st.error("⚠️ Please enter your Outlook credentials.")
        elif not uploaded_excel:
            st.error("⚠️ Please upload an Excel file.")
        else:
            recipients = pd.read_excel(uploaded_excel).dropna(subset=["Email"])
            results = []
            progress_bar = st.progress(0)
            total_recipients = len(recipients)
        
            for index, row in recipients.iterrows():
                name = row.get("Name", "User")
                email = row["Email"]
            
                attachment_paths = []
                for col in recipients.columns:
                    if "Attachment" in col and isinstance(row[col], str) and row[col].strip():
                        file_path = os.path.join(os.getcwd(), row[col])  # Convert to full path
                        if os.path.exists(file_path):
                            attachment_paths.append(row[col])
                        else:
                            st.warning(f"⚠️ Attachment not found: {row[col]}")

                if attachment_paths:
                    result = send_email_dynamic(email, name, attachment_paths)
                else:
                    result = f"⚠️ No valid attachments found for {email}, skipping."
                results.append(result)
                progress_bar.progress((index + 1) / total_recipients)
        
            for res in results:
                st.write(res)
