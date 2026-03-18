import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

df_global = None


#  Load Excel Data
def load_server_requests(path):
    global df_global

    df_global = pd.read_excel(path, dtype=str).fillna("")

    rows = []

    for idx, row in df_global.iterrows():

        apps = [a.strip() for a in row["Applications"].split(",") if a.strip()]

        rows.append({
            "RequestID": row["RequestID"],
            "OS": row["OS"],
            "RAM": row["RAM"],
            "HDD": row["HDD"],
            "Applications": apps,
            "Email": row.iloc[5]   # Column F (0-based index)
        })

    return rows


#  Send Email Function
def send_email_report(to_email, request_id, result):

    #  CONFIGURE THESE (VERY IMPORTANT)
    sender_email = "zulqarnainzia8@gmail.com"
    sender_password = "rmhk pvdg gxlw imjk"   # Use App Password (NOT normal password)

    subject = f"Server Creation Status - {request_id}"

    body = f"""
Hello,

Your server request ({request_id}) has been processed.

Result:
{result}

Regards,
RPA Bot
"""

    msg = MIMEMultipart()
    msg["From"] = sender_email
    msg["To"] = to_email
    msg["Subject"] = subject

    msg.attach(MIMEText(body, "plain"))

    try:
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(sender_email, sender_password)

        server.send_message(msg)
        server.quit()

        print(f"Email sent to {to_email}")

    except Exception as e:
        print(f"Failed to send email: {str(e)}")