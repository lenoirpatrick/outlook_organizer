#!python3
# -*- coding: utf-8 -*-

from outlook_organizer import OutlookOrganizer
from constants import *
from rich import print
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

if __name__ == "__main__":
    OO = OutlookOrganizer("appsettings_demo.json")
    OO.empty_trash()
    OO.delete_notifs_invits()
    OO.notifs_mails()
    OO.mails_projets()
    OO.mails_emails()
    OO.mails_from_me()
    OO.notifs_divers()
    OO.recap_email()
    OO.mails_sans_reponse()

    if OO.send_mail_recap is True:
        message = MIMEMultipart()

        message['From'] = OO.sender_address
        message['To'] = OO.user_email
        message['Subject'] = "Recap journalier " + str(date_du_jour)
        message.attach(MIMEText(OO.body, 'plain'))
        session = smtplib.SMTP(OO.email_smtp, OO.email_port)  # use gmail with port
        session.starttls()  # enable security
        session.login(OO.sender_address, OO.sender_pass)  # login with mail_id and password
        text = message.as_string()
        session.sendmail(OO.sender_address, OO.user_email, text)
        session.quit()
        print("[green3]Message recap envoyé")
