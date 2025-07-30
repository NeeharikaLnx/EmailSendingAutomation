import smtplib
import pandas as pd
from email.message import EmailMessage
import os
from datetime import datetime

# -------- CONFIGURATION --------
SENDER_EMAIL = "niharikaa972@gmail.com"
SENDER_PASSWORD = "upzv snky qbhg itmu"  # Use App Password, not your Gmail password
CSV_FILE = r"D:\EmailAutomation\Deduplicated_Vendor_Emails.csv"
RESUME_FILE = r"D:\EmailAutomation\Windows&LinuxAdmin_Neeharika1 (4).docx"

SUBJECT = "Experienced Linux/Windows VMware Admin | OpenShift, AWS, Backup & Storage Expert"

# This is the email template with a placeholder for the recipient's first name
BODY = """
Dear {name},

I hope this message finds you well.

I am reaching out to express my interest in new opportunities and to introduce myself as a seasoned Systems Administrator with over 11 years of experience supporting and securing enterprise-grade infrastructure across diverse Linux and Unix environments.

My core areas of expertise include:
- Automating and administering Linux/Unix systems in enterprise data center environments
- Deploying and managing containerized workloads using OpenShift
- Administering VMware vSphere for scalable and resilient virtualization solutions
- Implementing robust storage architectures and enterprise backup strategies
- Utilizing AWS services such as EC2, S3, IAM, and VPC in hybrid cloud deployments

I have attached my resume for your review and would welcome the opportunity to discuss any current or upcoming positions that align with my background.

Thank you for your time and consideration. I look forward to connecting.

Best regards,  
Neeharika Adella  
Phone: (760)-483-4972  
Email: niharikaa972@gmail.com  
LinkedIn: https://www.linkedin.com/in/neeharika-a-947b9583
"""

def send_bulk_emails():
    print(f"\nStarting email job at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")

    # Check if the resume file exists
    if not os.path.exists(RESUME_FILE):
        print(f"Resume file not found: {RESUME_FILE}")
        return

    # Try to read the CSV
    try:
        df = pd.read_csv(CSV_FILE)
    except Exception as e:
        print(f"Failed to read CSV: {e}")
        return

    # Counters for success/failure
    success_count = 0
    fail_count = 0

    with smtplib.SMTP('smtp.gmail.com', 587) as smtp:
        smtp.starttls()
        try:
            smtp.login(SENDER_EMAIL, SENDER_PASSWORD)
        except Exception as e:
            print(f"Login failed: {e}")
            return

        for _, row in df.iterrows():
            # Extract email and name from each row
            recipient = row['Email'].strip()
            name = row['Name'].strip().split()[0].capitalize() if 'Name' in row and pd.notna(row['Name']) else "Recruiter"

            # Personalize the email body
            personalized_body = BODY.format(name=name)

            # Compose the email
            msg = EmailMessage()
            msg['Subject'] = SUBJECT
            msg['From'] = SENDER_EMAIL
            msg['To'] = recipient
            msg.set_content(personalized_body)

            # Attach the resume file
            try:
                with open(RESUME_FILE, 'rb') as f:
                    msg.add_attachment(
                        f.read(),
                        maintype='application',
                        subtype='octet-stream',
                        filename=os.path.basename(RESUME_FILE)
                    )

                smtp.send_message(msg)
                print(f"Email sent to: {name} <{recipient}>")
                success_count += 1

            except Exception as e:
                print(f" Failed to send to {recipient}: {e}")
                fail_count += 1

    print(f"\n Email job completed at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f" Total Emails Sent: {success_count}")
    print(f" Failed Emails: {fail_count}")
    print("Script finished.\n")

# -------- EXECUTE ONCE AND EXIT --------
send_bulk_emails()
