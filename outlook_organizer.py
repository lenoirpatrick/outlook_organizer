#!python3
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
import git
import json
import os
import sys
from datetime import datetime, date, timedelta

# Rich text dans la console
from rich.progress import Progress, TimeElapsedColumn, SpinnerColumn
from rich import print
from rich.table import Table
from rich.console import Console

# Envoie d'email
import smtplib
import win32com.client
from win32com.client import Dispatch
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Connexion à SharePoint
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

INBOX = "      Inbox ("
INDIR = "      Indir ("


def table_recap():
    """ Affiche un tableau récapitulatif des principaux répertoires Outlook """
    table = Table(title="Répertoires Outlook")

    table.add_column("Répertoire", justify="left", style="cyan", no_wrap=True)
    table.add_column("Nb Messages", style="magenta")

    table.add_row("Inbox", str(len(inbox.Items)))
    table.add_row("Send Items", str(len(sentitems.Items)))
    table.add_row("Deleted Items", str(len(deleteditems.Items)))
    print(table)


def set_subject(subject) -> str:
    """ Retourne le nom du mail sans les prefixes de réponses et transferts.

     :param str subject: Sujet du mail à vérifier

     :return: Sujet corrigé
     :rtype str:
     """
    for item in ["RE: ", "TR: "]:
        if item in subject:
            subject = subject.replace(item, "")
    return subject


def parse_dir(outlookdirin, outlookdirout, subject):
    """ Permet de déplacer avec le même sujet dans un même répertoire

    :param win32com.client.CDispatch outlookdirin: Répertoire d'entrée à vérifier
    :param win32com.client.CDispatch outlookdirout: Répertoire de copie en cas d'occurence trouvée
    :param str subject: Sujet du mail à vérifier
    """
    move = True
    while move is True:
        move = False
        for message in outlookdirin:
            try:
                subject2 = set_subject(message.Subject)
                if len(subject) > 10:
                    if (subject2 in subject or subject in subject2) and subject2 != "" and subject != "":
                        move_message(message, outlookdirout, keep_in_inbox=False)
                        move = True
                        break
            except (Exception,):
                # print_exception()
                print_erreur("Err 004_2 : " + str(ex) + " / " + str(item.Subject))
                pass


def move_message(message, outlookdir, keep_in_inbox=False, mark_as_read=False):
    """
    Utilisé pour parcours message par message un répertoire

    :param win32com.client.CDispatch message: message à déplacer
    :param win32com.client.CDispatch outlookdir: répertoire où déplacer
    :param boolean keep_in_inbox:
    :param boolean mark_as_read:
    """
    move = False
    try:
        if keep_in_inbox is True:
            if is_archivable(message) is True:
                move = True
        else:
            move = True

        if move is True:
            print_deplace(message.Subject[0:80])
            if mark_as_read is True:
                item.Unread = False
            message.Move(outlookdir)
    except Exception as ex:
        print_erreur("Err 004 : " + str(ex) + " / " + str(item.Subject))
        # print_exception()


def is_archivable(mail):
    """ Détermine si un mail peut être déplacé du répertoire source

        :param win32com.client.CDispatch mail: Mail à vérifier

        :return: Le mail doit-il être archivé
        :rtyp bool:
    """
    try:
        nbolddays = get_nb_old_days(mail)
        if outofinboxdays > nbolddays.days:
            # return False, outofinboxdays - nbolddays.days
            return False
        else:
            return True
    except Exception as ex:
        print_erreur("Err 005 : " + str(ex) + " / " + str(item.Subject))
        print_exception()
        return False


def get_nb_old_days(mail) -> timedelta:
    """ Calcule le nombre de jours d'ancienneté du mail

        :param win32com.client.CDispatch mail: Mail à vérifier

        :return: Le nombre de jours du mail
    """
    try:
        d2 = mail.receivedtime
    except (Exception,):
        d2 = mail.creationtime
    d2 = date(d2.year, d2.month, d2.day)

    nbolddays = (date_du_jour - d2)
    return nbolddays


def move_mail(title, kw, folder, keep_in_inbox=False, mark_as_read=False,
              lookup_type="Subject", deletionexception=None):
    """ Deplacement de mail en fonction d'une liste de keywords (sujets)

        :param str title: Mail à vérifier
        :param list kw: Mail à vérifier
        :param win32com.client.CDispatch folder: Mail à vérifier
        :param bool keep_in_inbox: Mail à vérifier
        :param bool mark_as_read: Mail à vérifier
        :param str lookup_type: Mail à vérifier
        :param list deletionexception: Mail à vérifier
        """
    print_check(title)

    move = True
    while move is True:
        move = False
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                progress.advance(task)
                try:

                    # Suppression de certains mails par défaut
                    deletemail = False
                    if deletionexception is not None:
                        for subject_name in deletionexception:
                            # print(deletionexception)
                            try:
                                if subject_name in item.Subject:
                                    print_supprime(item)
                                    deletemail = True
                            except (Exception, ):
                                pass

                    # Mail 'urgent' et non lu à ne pas traiter
                    nepastraiter = False
                    try:
                        if item.Importance == 2 and item.Unread is True:
                            nepastraiter = True
                    except (Exception, ):
                        pass

                    # Traitement du mail
                    if deletemail is False and nepastraiter is False:
                        if keep_in_inbox is True:
                            for keyword in kw:
                                # print(keyword)
                                if lookup_type == "Sender":
                                    lookup_field = item.Sender
                                else:
                                    lookup_field = item.Subject

                                if keyword.encode("latin-1").decode("utf-8").lower() in str(lookup_field).lower():
                                    # print("ici")
                                    if is_archivable(item) is True:
                                        # Affichage du mail
                                        print_deplace(item.Subject)

                                        if mark_as_read is True:
                                            item.Unread = False
                                        item.Move(folder)
                                        move = True
                                        break
                                    break
                        else:
                            for keyword in kw:
                                if lookup_type == "Sender":
                                    lookup_field = item.Sender
                                else:
                                    lookup_field = item.Subject
                                # print(keyword.encode("latin-1").decode("utf-8"))
                                if keyword.encode("latin-1").decode("utf-8").lower() in str(lookup_field).lower():
                                    print_deplace(item.Subject)
                                    if mark_as_read is True:
                                        item.Unread = False
                                    item.Move(folder)
                                    move = True
                                    break
                except (AttributeError, ):
                    pass

                except (Exception,):
                    print_exception()

    move = True
    while move is True:
        move = False
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = ("      Sent Items (" + str(len(sentitems.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(sentitems.Items))
            for item in sentitems.Items:
                progress.advance(task)
                if keep_in_inbox is True:
                    if is_archivable(item) is True:
                        for keyword in kw:
                            if keyword.lower() in item.Subject.lower():
                                print_deplace(item.Subject)
                                item.Move(folder)
                                move = True
                                break
                else:
                    for keyword in kw:
                        if keyword.lower() in item.Subject.lower():
                            print_deplace(item.Subject)
                            item.Move(folder)
                            move = True
                            break


def set_archive_dir(indir, deletion_dir=False) -> win32com.client.CDispatch:
    """ Validation du répertoire d'archive de mail. Si il n'existe pas, le répertoire est créé

    :param win32com.client.CDispatch indir: Mail à vérifier
    :param bool deletion_dir: Mail à vérifier

    :return: Le nom du répertoire d'archive à utiliser
    """

    if deletion_dir is False:
        try:
            archivedir = indir.Folders[quarter]
        except (Exception,):
            archivedir = indir.Folders.Add(quarter)
    else:
        archivedir = indir
        try:
            indir.Folders[quarter].delete()
        except (Exception,):
            # nothing to do here
            pass
    return archivedir


def archivemails(indir, archivedir=None, deletearchive=False):
    """ Archive les mails du répertoire en paramètre

    :param win32com.client.CDispatch indir: Répertoire d'entrée
    :param win32com.client.CDispatch archivedir: Répertoire d'archive
    :param bool deletearchive: Suppresion ou archivage du mail
    """
    for mail in indir.Items:
        try:
            nbolddays = get_nb_old_days(mail)
            if archivabledays < nbolddays.days:
                if deletearchive is False:
                    print_archive(mail.Subject)
                    mail.Move(archivedir)
                else:
                    print_supprime(mail)
        except (Exception,):
            print_exception()


def set_indir(inbox, item):
    check_dir(inbox, item["dir"])
    check_dir(inbox.Folders[item["dir"]], item["subdir"])
    indir = inbox.Folders[item["dir"]].Folders[item["subdir"]]
    return indir

def check_dir(dir, subdir):
    """ Détermine si un mail peut être déplacé du répertoire source

    :param win32com.client.CDispatch dir: Répertoire source
    :param str subdir: Sous répertoire à tester
    """
    try:
        dir.Folders[subdir]
    except (Exception,):
        dir.Folders.Add(subdir)


def save_attachment(attachments, attach_ext, dir, prefix_name=None):
    """Sauvegarde la pièce joint sur SharePoint

    Args:
        attachments: Liste de pièces jointes du mail
        attach_ext: Extension à sélectionner
        dir: Répertoire où copier les fichiers
        prefix_name: Préfixe des fichiers à utiliser
    """
    for i in range(1, len(attachments) + 1):
        attachment = attachments.Item(i)
        # the name of attachment file
        if prefix_name is None:
            attachment_name = str(attachment).upper()
        else:
            attachment_name = prefix_name + "_" + str(attachment).upper()

        if attachment_name.endswith(attach_ext.upper()):
            try:
                attach_file = os.path.join(config["parameters"]["tmpdir"], attachment_name)
                attachment.SaveASFile(attach_file)

                sp_site = config["sharepoint"]["sp_site"]
                relative_url = dir
                client_credentials = UserCredential(config["sharepoint"]["login"], config["sharepoint"]["pass"])
                ctx = ClientContext(sp_site).with_credentials(client_credentials)

                remotepath = relative_url + "/" + attachment_name  # "  # existing folder path under sharepoint site.
                print_fichier(remotepath)
                with open(attach_file, 'rb') as content_file:
                    file_content = content_file.read()

                dir_name, name = os.path.split(remotepath)

                file = ctx.web.get_folder_by_server_relative_url(dir_name).upload_file(name,
                                                                                  file_content).execute_query()
                os.remove(attach_file)
            except Exception as ex:
                print("[[bright_red]KO[white]]     ctx.web.get_folder_by_server_relative_path(relative_url) : "
                      + relative_url)


def print_titre(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[blue]" + str(texte))


def print_check(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[deep_sky_blue3]    " + str(texte))


def print_fichier(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[deep_sky_blue3]            Enregistrement du document : " + str(texte))


def print_no_response(texte, unread=False):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    :param bool unread: Distinction sur la cas de gestion
    """
    if unread is False:
        print("[deep_sky_blue3]            Mail sans réponse à traiter manuellement : " + str(texte))
    else:
        print("[deep_sky_blue3]            Mail non lu : " + str(texte))


def print_archive(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[blue_violet]            Archivage du message : " + str(texte))


def print_deplace(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[green3]            Déplacement du message : " + str(texte))


def print_supprime(mail):
    """ Affiche du texte mis en forme dans la console

    :param win32com.client.CDispatch mail: Mail à supprimer
    """
    print("[dark_orange3]            Suppression du message : " + str(mail.Subject))
    try:
        mail.delete()
    except (Exception,):
        print_exception()


def print_erreur(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[bright_red]            " + str(texte))


def press_any_key():
    """ Demande un action utilisateur
    """
    input("Pressez une touche pour quitter")


def print_exception():
    """ Affiche la pile sur une erreur via le package rich
    """
    console.print_exception(show_locals=True)


if __name__ == "__main__":

    os.system('cls')

    # Git version
    repo = git.Repo(search_parent_directories=True)
    sha = repo.head.object.hexsha

    print("[[green3]OK[white]]     Initialisation du programme : " + os.path.basename(__file__))
    print("[[green3]OK[white]]         Librairies python")
    print("[[green3]OK[white]]         Version git : " + sha)

    ''' STATIC '''
    now = datetime.now()  # current date and time
    print("[[green3]OK[white]]         Date : " + str(now))

    print("[[green3]OK[white]]     Chargement du fichier de paramétrage")
    with open(os.path.join(os.path.dirname(__file__), 'appsettings.json')) as json_data:
        config = json.load(json_data)
    outofinboxdays = config['parameters']['outofinboxdays']
    archivabledays = config['parameters']['archivabledays']
    print("[[green3]OK[white]]         outofinboxdays: " + str(outofinboxdays))
    print("[[green3]OK[white]]         archivabledays: " + str(archivabledays))

    # Accès via Proxy
    os.environ["HTTP_PROXY"] = config["proxy"]["address"]
    print("[[green3]OK[white]]     Chargement du proxy")
    print("[[green3]OK[white]]         " + config["proxy"]["address"])

    # Ouverture d'Outlook
    try:
        print("[[green3]OK[white]]     Chargement du fichier Outlook")
        outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        sentitems = outlook.GetDefaultFolder(5)
        deleteditems = outlook.GetDefaultFolder(3)
        appointments = outlook.GetDefaultFolder(9).Items
        table_recap()
    except (Exception, ):
        print("[[bright_red]KO[white]]         L'accès Outlook n'est pas disponible, le programme va s'arrêter.")
        press_any_key()
        sys.exit(1)

    print("[[green3]OK[white]]     Définition des variables globales")
    # Déterminer le trimestre pour les sous dirs d'archivage
    datequarter = now - timedelta(archivabledays)
    annee = datequarter.year
    mois = datequarter.month
    trimestre = (mois - 1) / 3 + 1
    quarter = str(annee) + "_Q" + str(int(trimestre))

    # Date du jour
    date_du_jour = date(now.year, now.month, now.day)
    timenow = now.time()

    console = Console()

    print("[[green3]OK[white]]     Démarrage de l'application")
    print()
    print()

    index = 0

    # PARTIE AVEC EMAIL
    body = ""
    send_mail_recap = config["parameters"]["send_mail_recap"]

    index = index + 1
    print_titre(str(index) + " - Suppression des anciens mails de la Corbeille")
    with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
        libelle = ("      Deleted Items (" + str(len(deleteditems.Items)) + ")").ljust(30)
        task = progress.add_task(libelle, total=len(deleteditems.Items))
        for item in deleteditems.Items:
            progress.advance(task)
            if is_archivable(item) is True:
                print_supprime(item)

    if config["etapes"]["notifs_invits"] is True:
        index = index + 1
        print_titre(str(index) + " - Suppression des notifications d'invitations")
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
                        if item.Unread is False:
                            if item.Class in [46, 53, 54, 55, 56, 57, 181]:
                                print_supprime(item)
                            elif item.Class == 43 and str(item.Subject).startswith("Réponse automatique"):
                                print_supprime(item)
                        elif item.Class in [56]:
                            print_supprime(item)
                    except Exception as ex:
                        print("Err 005 : " + str(ex) + " / " + str(item.Subject))
                        print_exception()

    # Notifications
    if config["etapes"]["notifs_mails"] is True:
        index = index + 1
        print_titre(str(index) + " - Traitement des notifications")
        for item in config["notifications"]:
            if item["active"] is True:
                if "deletenotif" not in item:
                    item["deletenotif"] = ""

                # Vérification de la présence des répertoires
                indir = set_indir(inbox, item)
                archivedir = set_archive_dir(indir, item["deletearchive"])

                move_mail(item["name"], item["keywords"], indir, item["keepInInbox"],
                          lookup_type="Sender", mark_as_read=item["markAsRead"], deletionexception=item["deletenotif"])

                archivemails(indir, archivedir, item["deletearchive"])

    # Projets
    if config["etapes"]["mails_projets"] is True:
        index = index + 1
        print_titre(str(index) + " - Traitement des Projets")
        for item in config["projects"]:
            if item["active"] is True:
                # Vérification de la présence des répertoires
                indir = set_indir(inbox, item)
                archivedir = set_archive_dir(indir)

                move_mail(item["name"], item["keywords"], indir, item["keepInInbox"])

                archivemails(indir, archivedir)
                with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                    libelle = (INDIR + str(len(indir.Items)) + ")").ljust(30)
                    task = progress.add_task(libelle, total=len(indir.Items))
                    for message in indir.Items:
                        progress.advance(task)
                        subject = set_subject(message.Subject)
                        for liste in [inbox.Items, sentitems.Items]:
                            parse_dir(liste, indir, subject)

    # Parcours des mails externe et interne
    for global_item in ["mails_partenaires", "mails_internes"]:
        if config["etapes"][global_item] is True:
            index = index + 1
            print_titre(str(index) + " - Traitement des messages " + global_item)

            json_section = global_item.split("_")[1]
            for item in config[json_section]:
                if item["active"] is True:
                    print_check(
                        item["team"].encode("latin-1").decode("utf-8") + " > " + item["dir"].encode("latin-1").decode(
                            "utf-8") + "/" + item["subdir"].encode("latin-1").decode("utf-8"))

                    # Vérification de la présence des répertoires
                    indir = set_indir(inbox, item)
                    archivedir = set_archive_dir(indir)

                    # Parcours de la boite de réception pour déplacer les messages vers le rep Utilisateur
                    with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                        libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                        task = progress.add_task(libelle, total=len(inbox.Items))
                        for mail in inbox.Items:
                            progress.advance(task)

                            for user in item["users"]:
                                try:
                                    if user.encode("latin-1").decode("utf-8").upper() in str(mail.Sender).upper():
                                        move_message(mail, indir, keep_in_inbox=True)
                                except AttributeError:
                                    pass
                                except (Exception,):

                                    print_exception()

                    # Parcours rep utilisateur pour retrouver des messages dans la inbox et Send Items
                    with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                        libelle = (INDIR + str(len(indir.Items)) + ")").ljust(30)
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

                    archivemails(indir, archivedir)

    # From Me
    if config["etapes"]["from_me"] is True:
        index = index + 1
        print_titre(
            str(index) + " - Suppression des messages en inbox de moi vers @editique-ccm / plenoir.sefas@gmail.com")
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                try:
                    progress.advance(task)
                    if str(item.Sender) == "Patrick LENOIR" or str(item.Sender) == "plenoir.sefas@gmail.com":
                        print_supprime(item)
                except (AttributeError, ):
                    pass

    # Notification diverses
    try:
        index = index + 1
        print_titre(str(index) + " - Suppression des mails de notifs divers")
        emails = config["safetodelete"]["emails"]
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                try:
                    for mail in emails:
                        if mail in str(item.Sender):
                            if is_archivable(item) is True:
                                print_supprime(item)
                            break
                except (AttributeError, ):
                    pass
                progress.advance(task)
    except Exception as ex:
        print("[[bright_red]KO[white]]     Erreur lors du traitement des notification diverses.")
        print("[[bright_red]KO[white]]     " + str(ex))
        press_any_key()

    # On ne fait pas ces étapes l'après midi
    if now.hour > config['parameters']['time_short_version']:
        sys.exit(0)

    # Recap RDV
    if config["etapes"]["email_appointments"] is True and send_mail_recap is True:
        index = index + 1
        print_titre(str(index) + " - Récap des RDV du jour")
        appointments.Sort("[Start]")

        begin = date.today()
        end = begin + timedelta(days=1)
        restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" + end.strftime("%d/%m/%Y") + "'"
        restrictedItems = appointments.Restrict(restriction)

        # Détermination du jour de la semaine
        days = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi"]
        now = datetime.today()
        jour = now.day
        joursemaine = now.weekday() + 1
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = ("      RDV (" + str(len(restrictedItems)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(restrictedItems))
            for appointmentItem in restrictedItems:
                progress.advance(task)
                if joursemaine == appointmentItem.StartInStartTimeZone.isoweekday():
                    if str(appointmentItem.StartInStartTimeZone)[8:10] == str(jour):
                        # print(appointmentItem.StartInStartTimeZone.isoweekday())
                        # print("{0} \n  Start: {1}, \n  End: {2}, \n  Organizer: {3}".format(
                        #    appointmentIbody + "   tem.Subject, appointmentItem.StartInStartTimeZone,
                        #    appointmentItem.End, appointmentItem.Organizer))
                        body = body + str(appointmentItem.Subject) + "\n"
                        body = body + "     Début : " + str(appointmentItem.StartInStartTimeZone)[0:19] + "\n"
                        body = body + "     Durée : " + str(appointmentItem.Duration) + "mn" + "\n"
                        body = body + "     Orga. : " + str(appointmentItem.Organizer) + "\n"
                        body = body + "     Recu. : " + str(appointmentItem.IsRecurring) + "\n"
                        part = 0
                        body = body + "\n"

    # Mails sans réponse
    if config["etapes"]["unread_mails"] is True:
        index = index + 1
        if send_mail_recap is True:
            body = body + "\n"
            body = body + "Check des mails non lus : \n"
        print_titre(str(index) + " - Check des mails non lus")
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                progress.advance(task)
                if item.Unread is True:
                    nbolddays = get_nb_old_days(item)
                    if send_mail_recap is True:
                        body = body + "      " + item.Subject + " / " + str(item.Sender) + " [" + str(
                            nbolddays.days) + "j]\n"
                    print_no_response(item.Subject, unread=True)

        index = index + 1
        if send_mail_recap is True:
            body = body + "\n"
            body = body + "Check des mails sans réponses : \n"
        print_titre(str(index) + " - Check des mails sans réponses")
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
                    if is_archivable(item) is True:
                        print_no_response(item.Subject)
                        nbolddays = get_nb_old_days(item)

                        if send_mail_recap is True:
                            body = body + "      " + item.Subject + " / " + str(item.Sender) + " [" + str(
                                nbolddays.days) + "j]\n"

    if send_mail_recap is True:
        sender_address = config["email"]["sender_address"]
        sender_pass = config["email"]["sender_pass"]
        receiver_address = config["email"]["receiver_address"]

        Outlook = win32com.client.Dispatch("Outlook.Application")
        ns = Outlook.GetNamespace("MAPI")
        message = MIMEMultipart()

        message['From'] = sender_address
        message['To'] = receiver_address
        message['Subject'] = "Recap journalier " + str(date_du_jour)
        message.attach(MIMEText(body, 'plain'))
        session = smtplib.SMTP(config["email"]["smtp"], config["email"]["port"])  # use gmail with port
        session.starttls()  # enable security
        session.login(sender_address, sender_pass)  # login with mail_id and password
        text = message.as_string()
        session.sendmail(sender_address, receiver_address, text)
        session.quit()
    else:
        table_recap()
        print("[green3]Traitement terminé")
        press_any_key()
