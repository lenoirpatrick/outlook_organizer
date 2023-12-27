#!python3
# -*- coding: utf-8 -*-

from outlook_organizer import OutlookOrganizer
from constants import *
from console_log import print_prologue
from rich import print
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import os
import sys

if __name__ == "__main__":
    os.system('cls')

    # Création d'une instance de la classe OutlookOrganizer en utilisant le fichier "appsettings_demo.json"
    OO = OutlookOrganizer("appsettings_demo.json")

    print_prologue(OO.config)

    # Suppression des éléments dans la corbeille Outlook
    OO.empty_trash()

    # Suppression des notifications et invitations
    if OO.config["step_notifs_invits"] is True:
        OO.delete_notifs_invits()

    # Gestion des notifications de mails
    if OO.config["step_notifs_mails"] is True:
        OO.notifs_mails()

    # Gestion des mails de projets
    if OO.config["step_mails_projets"] is True:
        OO.mails_projets()

    # Gestion des mails partenaires et interne
    for config_params in ["step_mails_partenaires", "step_mails_internes"]:
        if OO.config[config_params] is True:
            OO.mails_emails(config_params.split("_")[2])

    # Gestion des mails provenant de l'utilisateur actuel
    if OO.config["step_from_me"] is True:
        OO.mails_from_me()

    # Gestion des notifications diverses
    if OO.config["step_notifs_invits"] is True:
        OO.notifs_divers()

    # On ne fait pas ces étapes l'après midi
    if OO.config["now"].hour > OO.time_short_version:
        sys.exit(0)

    # Récapitulatif des rdvs
    if OO.config["step_email_appointments"] is True:
        OO.recap_email()

    # Gestion des e-mails sans réponse
    if OO.config["step_unread_mails"] is True:
        OO.mails_sans_reponse()

    print()

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
