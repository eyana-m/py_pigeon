import smtplib
import csv
from string import Template

from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
import datetime
from email import encoders
import imaplib
import time


CREDENTIALS_USER ="eyana.mallari@rcoffice.ringcentral.com"
CREDENTIALS_PASS = "<password>"
EMAIL_FROM_DEFAULT = "eyana.mallari@ringcentral.com"

EMAIL_SUBJECT = "Hello There!"
EMAIL_CC_DEFAULT = ""
EMAIL_BCC_DEFAULT = ""

CONTACTS_FILE = "/Users/eyana.mallari/Projects-Local/py_teamlink/Input/CONTACTS SAMPLE2.csv"


def getEmailContent(first_name):

    email_content ="""
Hi """+first_name+""",
<br><br>
Email content goes here
<br><br>
Regards,<br>
Eyana

    """

    return email_content

def loop_contacts(filename):
    print('looping contacts')

    # set up the SMTP server
    print('Sending emails')
    s = smtplib.SMTP(host='smtp.office365.com', port=587)
    s.starttls()
    s.login(CREDENTIALS_USER, CREDENTIALS_PASS)


    count = 1

    with open(filename, mode='r') as contacts_file:
        reader = csv.reader(contacts_file)
        next(reader)
        for contact in reader:
            contact_full_name = contact[0]
            company = contact[1]
            title = contact[2]
            email = contact[3]
            first_name= contact[4]

            msg = MIMEMultipart()

            print(count)
            count = count +1
            print("Sending email to", contact_full_name)

            msg['From']=EMAIL_FROM_DEFAULT
            msg['To']=email
            msg['Bcc'] = EMAIL_BCC_DEFAULT
            msg['Subject']=  EMAIL_SUBJECT

            msg.attach(MIMEText(getEmailContent(first_name), 'html'))
            print(type(msg))
            s.send_message(msg)

            del msg

    s.quit()

loop_contacts(CONTACTS_FILE)
