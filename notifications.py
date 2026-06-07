import smtplib
from email.mime.text import MIMEText

def send_email(devices, to_email, smtp_server="smtp.gmail.com", smtp_port=587, smtp_user="", smtp_pass=""):
    vuln_count = sum(1 for d in devices if d["vulnerable"])
    msg = MIMEText(f"Siziň toruňyzda {vuln_count} sany howpsuz däl gurluş tapyldy.")
    msg["Subject"] = "IoT Howpsuzlyk Skanirleme Duýduryşy"
    msg["From"] = smtp_user or "iot.security@example.com"
    msg["To"] = to_email

    try:
        with smtplib.SMTP(smtp_server, smtp_port) as server:
            server.starttls()
            if smtp_user and smtp_pass:
                server.login(smtp_user, smtp_pass)
            server.send_message(msg)
    except Exception as e:
        print(f"Email ibermek ýalňyşlygy: {e}")
