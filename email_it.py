# from email.mime.image import MIMEImage
# from pathlib import Path
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import smtplib


def send_excel_file(path_to_file, recipient_email="hemmarsbach@gmail.com"):
    msg = MIMEMultipart()
    msg["From"] = "mejlbox@yahoo.com"
    msg["To"] = recipient_email
    msg["Subject"] = "Excel file for checking in assets is attached."

    # attach the file to the message
    with open(path_to_file, 'rb') as f:
        attachment = MIMEApplication(f.read(), _subtype='xlsx')
        attachment.add_header('Content-Disposition', 'attachment', filename=path_to_file)
        msg.attach(attachment)

    msg.attach(MIMEText("Please see the attached file."))
    # msg.attach(MIMEImage(Path("dilshad.png").read_bytes()))

    with smtplib.SMTP(host="smtp.gmail.com", port=587) as server:
        server.ehlo()  # Shake had with smtp server by saying hello. Part of smtp protocol.
        server.starttls()  # ttls stands for Transport Layer Security. Everything sent to smtp server will be encrypted.
        server.login("anoldwebsite@gmail.com",
                     "ycphwubxbtfqrigt")  # Your email that you want to send email from and password
        server.send_message(msg)
        print("Email sent successfully!")