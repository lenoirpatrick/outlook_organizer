#!python3
# -*- coding: utf-8 -*-
import git
import sys
import json
import os
from datetime import timedelta

# Accès Outlook
import win32com.client
from win32com.client import Dispatch

# Rich text dans la console
from rich import print
from rich.progress import Progress, TimeElapsedColumn, SpinnerColumn

# Envoie d'email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from constants import *
from console_log import (table_recap, print_titre, print_check, print_no_response, print_deplace, print_supprime,
                         press_any_key, print_exception)
from tools_dir import parse_dir, set_indir, set_archive_dir
from tools_message import set_subject, move_message, is_archivable, get_nb_old_days, move_mail, archivemails


class OutlookOrganizer:
    def __init__(self, jsonfile="appsettings.json"):
        self.jsonfile = jsonfile
        self.index = 0
        self.body = ""

        with open(os.path.join(os.path.dirname(__file__), self.jsonfile)) as json_data:
            self.__config = json.load(json_data)

        self.send_mail_recap = self.__config["parameters"]["send_mail_recap"]

        self.__outofinboxdays = self.__config['parameters']['outofinboxdays']
        self.__archivabledays = self.__config['parameters']['archivabledays']

        # utilisateur
        self.user_email = self.__config['user']['user_email']
        self.user_name = self.__config['user']['user_name']
        self.proxy_address = self.__config["user"]["proxy_address"]

        # Sharepoint config
        self.sharepoint_sp_site = self.__config["user"]["sharepoint_sp_site"]
        self.sharepoint_login = self.__config["user"]["sharepoint_login"]
        self.sharepoint_pass = self.__config["user"]["sharepoint_pass"]

        # email config sender
        self.sender_address = self.__config["email"]["sender_address"]
        self.sender_pass = self.__config["email"]["sender_pass"]

        self.time_short_version = self.__config['parameters']['time_short_version']

        self.email_smtp = self.__config['email']['smtp']
        self.email_port = self.__config['email']['port']

        datequarter = now - timedelta(self.__archivabledays)
        annee = datequarter.year
        mois = datequarter.month
        trimestre = (mois - 1) / 3 + 1
        self.quarter = str(annee) + "_Q" + str(int(trimestre))

        self.config = {
            "outofinboxdays": int(self.__outofinboxdays),
            "archivabledays": int(self.__archivabledays),
            "tmpdir": self.__config["parameters"]["tmpdir"],
            "sharepoint_sp_site": self.__config["user"]["sharepoint_sp_site"],
            "sharepoint_login": self.__config["user"]["sharepoint_login"],
            "sharepoint_pass": self.__config["user"]["sharepoint_pass"]
                       }

    def empty_trash(self):
        self.index = self.index + 1
        print_titre(str(self.index) + " - Suppression des anciens mails de la Corbeille")
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = ("      Deleted Items (" + str(len(deleteditems.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(deleteditems.Items))
            for item in deleteditems.Items:
                progress.advance(task)
                if is_archivable(item, self.config) is True:
                    print_supprime(item)

    def delete_notifs_invits(self):
        if self.__config["etapes"]["notifs_invits"] is True:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Suppression des notifications (invitations, AR, absences)")
            lenitem = len(inbox.Items) + len(sentitems.Items)
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = ("      Inbox/Sent Items (" + str(lenitem) + ")").ljust(30)
                task = progress.add_task(libelle, total=lenitem)
                for liste in [inbox.Items, sentitems.Items]:
                    """ https://learn.microsoft.com/en-us/office/vba/api/outlook.olobjectclass
                           Class :  43 -> message
                                    46 -> Non remis
                                    56 -> Réunion acceptée
                                    55 -> Réunion refusée (Réception)
                                    54 -> Réunion refusée (Envoie)
                                    53 -> Invitation
                                    57 -> Acceptation provisoire d'une demande de réunion
                                    181 -> Transfert d'email """

                    for item in liste:
                        progress.advance(task)
                        try:
                            if not item.Unread:
                                if item.Class in [46, 53, 54, 55, 57, 181] or (
                                        item.Class == 43 and str(item.Subject).startswith("Réponse automatique")):
                                    print_supprime(item)
                            elif item.Class == 56:
                                print_supprime(item)
                        except Exception as ex:
                            print("Err 005 : " + str(ex) + " / " + str(item.Subject))
                            print_exception()

    def notifs_mails(self):
        if self.__config["etapes"]["notifs_mails"] is True:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Traitement des notifications")
            for item in self.__config["notifications"]:
                if item["active"] is True:
                    if "deletenotif" not in item:
                        item["deletenotif"] = ""

                    # Vérification de la présence des répertoires
                    indir = set_indir(inbox, item)
                    archivedir = set_archive_dir(indir, self.quarter, item["deletearchive"])

                    move_mail(item, indir, self.config, lookup_type="Sender")

                    archivemails(indir, archivedir, self.config, item["deletearchive"])

    def mails_projets(self):
        if self.__config["etapes"]["mails_projets"] is True:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Traitement des Projets")
            for item in self.__config["projects"]:
                if item["active"] is True:
                    # Vérification de la présence des répertoires
                    indir = set_indir(inbox, item)
                    archivedir = set_archive_dir(indir, self.quarter)

                    move_mail(item, indir, self.config)
                    # move_mail(item["name"], item["keywords"], indir, item["keepInInbox"])

                    archivemails(indir, archivedir, self.config)
                    nb_msg_indir = len(indir.Items)
                    if nb_msg_indir > 0:
                        with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                                      TimeElapsedColumn(), ) as progress:
                            libelle = (INDIR + str(nb_msg_indir) + ")").ljust(30)
                            task = progress.add_task(libelle, total=len(indir.Items))
                            for message in indir.Items:
                                progress.advance(task)
                                subject = set_subject(message.Subject)
                                for liste in [inbox.Items, sentitems.Items]:
                                    parse_dir(liste, indir, subject, self.config)

    def mails_emails(self):
        # Parcours des mails externe et interne
        for global_item in ["mails_partenaires", "mails_internes"]:
            if self.__config["etapes"][global_item] is True:
                self.index = self.index + 1
                print_titre(str(self.index) + " - Traitement des messages " + global_item)

                json_section = global_item.split("_")[1]
                for item in self.__config[json_section]:
                    if item["active"] is True:
                        title = (item["team"].encode("latin-1").decode("utf-8") + " > " +
                                 item["dir"].encode("latin-1").decode("utf-8") + "/" +
                                 item["subdir"].encode("latin-1").decode("utf-8"))
                        print_check(title)

                        # Vérification de la présence des répertoires
                        indir = set_indir(inbox, item)
                        archivedir = set_archive_dir(indir, self.quarter)

                        # Parcours de la boite de réception pour déplacer les messages vers le rep Utilisateur
                        with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                                      TimeElapsedColumn(), ) as progress:
                            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                            task = progress.add_task(libelle, total=len(inbox.Items))
                            for mail in inbox.Items:
                                progress.advance(task)

                                for user in item["users"]:
                                    try:
                                        if user.encode("latin-1").decode("utf-8").upper() in str(mail.Sender).upper():
                                            move_message(mail, indir, self.config, keep_in_inbox=True)
                                    except AttributeError:
                                        pass
                                    except (Exception,):

                                        print_exception()

                        # Parcours rep utilisateur pour retrouver des messages dans la inbox et Send Items
                        nb_msg_indir = len(indir.Items)
                        if nb_msg_indir > 0:
                            with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                                          TimeElapsedColumn(), ) as progress:
                                libelle = (INDIR + str(nb_msg_indir) + ")").ljust(30)
                                task = progress.add_task(libelle, total=len(indir.Items))
                                for message in indir.Items:
                                    progress.advance(task)
                                    subject = set_subject(message.Subject)
                                    # Parcours de Inbox & Send Items
                                    for liste in [inbox.Items, sentitems.Items]:
                                        for message2 in liste:
                                            subject2 = set_subject(message2.Subject)
                                            if subject in subject2:
                                                print_deplace(subject2[0:80])
                                                try:
                                                    message2.Move(indir)
                                                except (Exception,):
                                                    print_exception()
                                                break

                        archivemails(indir, archivedir, self.config)

    def mails_from_me(self):
        # From Me
        if self.__config["etapes"]["from_me"] is True:
            self.index = self.index + 1
            print_titre(
                str(self.index) + " - Suppression des messages en inbox de moi vers " + self.user_email +
                " / " + self.sender_address)
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(inbox.Items))
                for item in inbox.Items:
                    try:
                        progress.advance(task)
                        if str(item.Sender) in [self.user_name, self.user_email, self.sender_address]:
                            print_supprime(item)
                    except (AttributeError,):
                        pass

    def notifs_divers(self):
        # Notification diverses
        try:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Suppression des mails de notifs divers")
            emails = self.__config["safetodelete"]["emails"]
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(inbox.Items))
                for item in inbox.Items:
                    try:
                        for mail in emails:
                            if mail in str(item.Sender):
                                if is_archivable(item, self.config) is True:
                                    print_supprime(item)
                                break
                    except (AttributeError,):
                        pass
                    progress.advance(task)
        except Exception as ex:
            print("[[bright_red]KO[white]]     Erreur lors du traitement des notification diverses.")
            print("[[bright_red]KO[white]]     " + str(ex))
            press_any_key()

    def recap_email(self):
        # Recap RDV
        if self.__config["etapes"]["email_appointments"] is True and self.send_mail_recap is True:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Récap des RDV du jour")
            appointments.Sort("[Start]")

            begin = date.today()
            end = begin + timedelta(days=1)
            restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" + end.strftime(
                "%d/%m/%Y") + "'"
            restricted_items = appointments.Restrict(restriction)

            joursemaine = now.weekday() + 1
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = ("      RDV (" + str(len(restricted_items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(restricted_items))
                for appointment_item in restricted_items:
                    progress.advance(task)
                    if joursemaine == appointment_item.StartInStartTimeZone.isoweekday():
                        if str(appointment_item.StartInStartTimeZone)[8:10] == str(jour_en_cours):
                            # print(appointmentItem.StartInStartTimeZone.isoweekday())
                            # print("{0} \n  Start: {1}, \n  End: {2}, \n  Organizer: {3}".format(
                            #    appointmentIbody + "   tem.Subject, appointmentItem.StartInStartTimeZone,
                            #    appointmentItem.End, appointmentItem.Organizer))
                            self.body += f"{str(appointment_item.Subject)}\n"
                            self.body += f"     Début : {str(appointment_item.StartInStartTimeZone)[:19]}\n"
                            self.body += f"     Durée : {str(appointment_item.Duration)}mn\n"
                            self.body += f"     Orga. : {str(appointment_item.Organizer)}\n"
                            self.body += f"     Recu. : {str(appointment_item.IsRecurring)}\n\n"

    def mails_sans_reponse(self):
        # Mails sans réponse
        if self.__config["etapes"]["unread_mails"] is True:
            self.index = self.index + 1
            if self.send_mail_recap is True:
                self.body = self.body + "\n"
                self.body = self.body + "Check des mails non lus : \n"
            print_titre(str(self.index) + " - Check des mails non lus")
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(inbox.Items))
                for item in inbox.Items:
                    progress.advance(task)
                    if item.Unread is True:
                        nbolddays = get_nb_old_days(item)
                        if self.send_mail_recap is True:
                            self.body = self.body + "      " + item.Subject + " / " + str(item.Sender) + " [" + str(
                                nbolddays.days) + "j]\n"
                        print_no_response(item.Subject, unread=True)

            self.index = self.index + 1
            if self.send_mail_recap is True:
                self.body += "\nCheck des mails sans réponses :\n"

            print_titre(str(self.index) + " - Check des mails sans réponses")
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = ("      Sent Items (" + str(len(sentitems.Items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(sentitems.Items))
                for item in sentitems.Items:
                    progress.advance(task)
                    repondu = False
                    if item.Subject[0:3] not in ["RE:", "TR:", ""]:
                        repondu = False

                    for inboxitems in inbox.Items:
                        if item.Subject in inboxitems.Subject:
                            repondu = True

                    # Le mail n'a pas de réponses
                    if repondu is False:
                        # Le mail est-il archivable ?
                        if is_archivable(item, self.config) is True:
                            print_no_response(item.Subject)
                            nbolddays = get_nb_old_days(item)

                            if self.send_mail_recap is True:
                                self.body = self.body + "      " + item.Subject + " / " + str(item.Sender) + " [" + str(
                                    nbolddays.days) + "j]\n"


if __name__ == "__main__":
    os.system('cls')
    OO = OutlookOrganizer()

    # Git version
    repo = git.Repo(search_parent_directories=True)
    sha = repo.head.object.hexsha

    print("[[green3]OK[white]]     Initialisation du programme : " + os.path.basename(__file__))
    print("[[green3]OK[white]]         Librairies python")
    print("[[green3]OK[white]]         Version git : " + sha)

    ''' STATIC '''
    print("[[green3]OK[white]]         Date : " + str(now))

    print("[[green3]OK[white]]     Chargement du fichier de paramétrage")

    print("[[green3]OK[white]]         outofinboxdays: " + str(OO.config["outofinboxdays"]))
    print("[[green3]OK[white]]         archivabledays: " + str(OO.config["archivabledays"]))

    # Accès via Proxy
    proxy_address = OO.proxy_address
    os.environ["HTTP_PROXY"] = proxy_address
    print("[[green3]OK[white]]     Chargement du proxy")
    print("[[green3]OK[white]]         " + proxy_address)

    # Ouverture d'Outlook
    try:
        print("[[green3]OK[white]]     Chargement du fichier Outlook")
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        sentitems = outlook.GetDefaultFolder(5)
        deleteditems = outlook.GetDefaultFolder(3)
        appointments = outlook.GetDefaultFolder(9).Items
        table_recap(inbox, sentitems, deleteditems)
    except (Exception, ):
        print("[[bright_red]KO[white]]         L'accès Outlook n'est pas disponible, le programme va s'arrêter.")
        press_any_key()
        sys.exit(1)

    print("[[green3]OK[white]]     Définition des variables globales")

    print("[[green3]OK[white]]     Démarrage de l'application")
    print()

    OO.empty_trash()
    OO.delete_notifs_invits()
    OO.notifs_mails()
    OO.mails_projets()
    OO.mails_emails()
    OO.mails_from_me()
    OO.notifs_divers()

    # On ne fait pas ces étapes l'après midi
    if now.hour > OO.time_short_version:
        sys.exit(0)

    OO.recap_email()
    OO.mails_sans_reponse()

    print()
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

    print("[green3]Traitement terminé")
    table_recap(inbox, sentitems, deleteditems)
    press_any_key()
