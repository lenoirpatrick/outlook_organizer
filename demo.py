#!python3
# -*- coding: utf-8 -*-

from outlook_organizer import OutlookOrganizer
from constants import *
from rich import print
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os

if __name__ == "__main__":
    os.system('cls')

    # Création d'une instance de la classe OutlookOrganizer en utilisant le fichier "appsettings_demo.json"
    OO = OutlookOrganizer("appsettings_demo.json")

    # Suppression des éléments dans la corbeille Outlook
    OO.empty_trash()

    # Suppression des notifications et invitations
    OO.delete_notifs_invits()

    # Gestion des notifications de mails
    OO.notifs_mails()

    # Gestion des mails de projets
    OO.mails_projets()

    # Gestion des mails d'e-mails
    OO.mails_emails()

    # Gestion des mails provenant de l'utilisateur actuel
    OO.mails_from_me()

    # Gestion des notifications diverses
    OO.notifs_divers()

    # Récapitulatif des e-mails
    OO.recap_email()

    # Gestion des e-mails sans réponse
    OO.mails_sans_reponse()

    # Envoi d'un récapitulatif par e-mail si la condition est vraie
    if OO.send_mail_recap is True:
        message = MIMEMultipart()

        # Construction du message e-mail
        message['From'] = OO.sender_address
        message['To'] = OO.user_email
        message['Subject'] = "Recap journalier " + str(date_du_jour)
        message.attach(MIMEText(OO.body, 'plain'))

        # Connexion au serveur SMTP et envoi du message
        session = smtplib.SMTP(OO.email_smtp, OO.email_port)  # utilisation de Gmail avec le port spécifié
        session.starttls()  # activation de la sécurité
        session.login(OO.sender_address, OO.sender_pass)  # connexion avec l'adresse e-mail et le mot de passe
        text = message.as_string()
        session.sendmail(OO.sender_address, OO.user_email, text)
        session.quit()
        print("[green3]Message recap envoyé")