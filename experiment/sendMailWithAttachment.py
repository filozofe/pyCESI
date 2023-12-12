# test  procedure to send email with attachment
#
# from https://medium.com/@neonforge/how-to-send-emails-with-attachments-with-python-by-using-microsoft-outlook-or-office365-smtp-b20405c9e63a

import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import os.path


def send_email(email_recipient,
               email_subject,
               email_message,
               attachment_location = ''):

    email_sender = 'phofmann@cesi.fr'

    msg = MIMEMultipart()
    msg['From'] = email_sender
    msg['To'] = email_recipient
    msg['Subject'] = email_subject

    msg.attach(MIMEText(email_message, 'plain'))

    if attachment_location != '':
        filename = os.path.basename(attachment_location)
        attachment = open(attachment_location, "rb")
        part = MIMEBase('application', 'vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header('Content-Description', os.path.basename(attachment_location))
        part.add_header('Content-Disposition',
                        'attachment', filename=os.path.basename(attachment_location))
        msg.attach(part)

    try:
        server = smtplib.SMTP('smtp.office365.com', 587)
        server.ehlo()
        server.starttls()
        server.login('phofmann@cesi.fr', '@Cesi2308')
        text = msg.as_string()
        server.sendmail(email_sender, email_recipient, text)
        print('email sent')
        server.quit()
    except:
        print("SMPT server connection error")
    return True

send_email('phofmannh@gmail.com',
           'test envoi email avec attachement',
           'Ceci est un test envoi email avec attachement',
           r"C:\Users\phofmann\Downloads\Grille mémoire, soutenance, projet pro_ASR (1).xlsx")

