import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import os
import pandas as pd

import time

def get_name_email_dict():
    file_path = "Mails.xlsx"
    
    try:
        df = pd.read_excel(file_path, engine='openpyxl')
        
        # Check for required columns
        if {'Name', 'Mails'}.issubset(df.columns):
            return dict(zip(df['Name'], df['Mails']))
        else:
            print(f"Error: Missing required columns. Found columns: {df.columns.tolist()}")
            return {}
    
    except FileNotFoundError:
        print(f"Error: '{file_path}' not found.")
        return {}
    except Exception as e:
        print(f"Error reading '{file_path}': {e}")
        return {}

# Email Configuration
smtp_server = "smtp.gmail.com"
smtp_port = 587
sender_email = "Abc"
sender_password = "token"
cc_recipients = ["abc"]  
# File to attach
attachment_path = "attach";  

    # Connect to SMTP Server
server = smtplib.SMTP(smtp_server, smtp_port)
server.starttls()
server.login(sender_email, sender_password)
email_limit = 80 # Number of emails to send

email_count = 0


# List of recipients
for (name ,email) in get_name_email_dict().items():
    if email_count >= 8:
        if email_count >= email_limit:
            print(f"Email limit of {email_limit} reached. Stopping.")
            break
        to_recipients = [email]

        # Email Content
        subject = f"Seeking Job Opportunities at with 9+ Years of Expertise"
        body = body = f"""
Greeting {name},

I hope you're doing well. I wanted to reach out regarding potential job opportunities. With over 9 years of experience in software development, I have expertise in:
- Java, Spring Boot, and Microservices
- React and Angular for Frontend Development

I would appreciate the chance to explore any roles that might align with my experience.

Please find my resume attached for your consideration. I look forward to hearing from you.

Best regards,  
Damodhar Reddy
"""



        # Create message
        msg = MIMEMultipart()
        msg["From"] = sender_email
        msg["To"] = ", ".join(to_recipients)
        msg["Cc"] = ", ".join(cc_recipients)  # Add CC recipients
        msg["Subject"] = subject
        msg.attach(MIMEText(body,"plain"))

        
        try:
            with open(attachment_path, "rb") as file:
                part = MIMEApplication(file.read(), Name=attachment_path)
                part["Content-Disposition"] = f'attachment; filename="{attachment_path}"'
                msg.attach(part)
        except FileNotFoundError:
            print("Attachment file not found. Email will be sent without an attachment.")

        # Combine To and CC recipients for sending
        all_recipients = to_recipients + cc_recipients

        # Send email
        server.sendmail(sender_email, all_recipients, msg.as_string())
      
        del msg
        all_recipients.clear()
        time.sleep(300)
    email_count += 1
# Close Connection
server.quit()

print("Email sent successfully with attachment and CC!")

