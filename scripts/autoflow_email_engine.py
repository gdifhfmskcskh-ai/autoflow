"""
AutoFlow Studio - Bulk Email Campaign Engine
Uses environment variables for secure SMTP credentials.
"""

import smtplib
import ssl
from email.message import EmailMessage
import pandas as pd
import os

def send_bulk_campaign(
    csv_recipients_file: str,
    subject_template: str,
    body_template: str,
    simulate_only: bool = True
):
    # Load credentials safely from environment variables
    smtp_server = os.getenv("SMTP_SERVER", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "465"))
    sender_email = os.getenv("SMTP_USER", "your-email@example.com")
    sender_password = os.getenv("SMTP_PASSWORD", "")

    df = pd.read_csv(csv_recipients_file)
    print(f"Loaded {len(df)} recipients from {csv_recipients_file}")

    for idx, row in df.iterrows():
        # Tag substitution
        recipient_name = row.get("Name", "Customer")
        recipient_email = row.get("Email", "")
        
        personalized_subject = subject_template.replace("{{Name}}", str(recipient_name))
        personalized_body = body_template.replace("{{Name}}", str(recipient_name))
        for col in df.columns:
            personalized_body = personalized_body.replace(f"{{{{{col}}}}}", str(row[col]))

        msg = EmailMessage()
        msg["Subject"] = personalized_subject
        msg["From"] = sender_email
        msg["To"] = recipient_email
        msg.set_content(personalized_body)

        if simulate_only or not sender_password:
            print(f"[SIMULATION] Ready to send to {recipient_email} | Subject: {personalized_subject}")
        else:
            context = ssl.create_default_context()
            with smtplib.SMTP_SSL(smtp_server, smtp_port, context=context) as server:
                server.login(sender_email, sender_password)
                server.send_message(msg)
            print(f"[SENT] Dispatched to {recipient_email}")

if __name__ == "__main__":
    send_bulk_campaign("fake_data.csv", "Hello {{Name}}!", "Hi {{Name}}, welcome to AutoFlow Studio.", simulate_only=True)
