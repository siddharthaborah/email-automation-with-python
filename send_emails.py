import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Load Excel file
email_data = pd.read_excel('path_to_your_excel_file.xlsx')

# Extract emails from the column (replace 'Email' with the actual column name)
email_list = email_data['Email'].tolist()

# Email sender credentials
sender_email = "your_email@gmail.com"
sender_password = "your_password"

# Setup the SMTP server
server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(sender_email, sender_password)

# Email content
subject = "Your Email Subject"
body = "This is the body of your email."

# Loop through email list and send emails
for recipient in email_list:
    msg = MIMEMultipart()
    msg['From'] = sender_email
    msg['To'] = recipient
    msg['Subject'] = subject

    msg.attach(MIMEText(body, 'plain'))

    text = msg.as_string()
    server.sendmail(sender_email, recipient, text)

# Close the SMTP server
server.quit()

print("Emails sent successfully!")
