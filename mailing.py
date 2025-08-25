import pandas as pd
import smtplib
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email import encoders
import os

# === CONFIGURATION ===
MASTER_FILE = r"main.xlsx"  # your main Excel file
EMAILS_FILE = r"agent-email.xlsx"  # mapping file with Agent Cd + Email
SENDER_EMAIL = "my-mail"
SENDER_PASSWORD = "app_password"  # generated 16-char app password

# === Load data ===
df = pd.read_excel(MASTER_FILE, engine="openpyxl")
emails_df = pd.read_excel(EMAILS_FILE, engine="openpyxl")

# Convert email mapping to dictionary {AgentCd: Email}
agent_email_map = dict(zip(emails_df["Agent Cd"], emails_df["Email"]))

# === Email sending function ===
def send_email(receiver_email, subject, body, attachment_path):
    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = receiver_email
    msg["Subject"] = subject

    # Body
    msg.attach(MIMEText(body, "plain"))

    # Attachment
    with open(attachment_path, "rb") as f:
        part = MIMEBase("application", "octet-stream")
        part.set_payload(f.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(attachment_path)}")
        msg.attach(part)

    # Send via Gmail
    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)

    print(f"Sent to {receiver_email} with {os.path.basename(attachment_path)}")

# === Generate & send per-agent sheets ===
for agent_cd in df["Agent Cd"].unique():
    subset = df[df["Agent Cd"] == agent_cd]

    # Skip if agent not in email map
    if agent_cd not in agent_email_map:
        print(f"No email found for Agent {agent_cd}, skipping.")
        continue

    receiver_email = agent_email_map[agent_cd]

    # Save subset to temp Excel file
    temp_file = f"Agent_{agent_cd}.xlsx"
    subset.to_excel(temp_file, index=False)

    # Send email
    subject = f"Policy Report for Agent {agent_cd}"
    body = f"Dear Agent {agent_cd},\n\nPlease find attached your policy report.\n\nRegards,\nTeam"
    send_email(receiver_email, subject, body, temp_file)

    # Remove temp file after sending
    os.remove(temp_file)
