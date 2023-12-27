#!python3
# -*- coding: utf-8 -*-
from rich.progress import Progress, TimeElapsedColumn, SpinnerColumn
from datetime import timedelta
import os

# Connexion à SharePoint
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

from constants import *
from console_log import (print_check, print_fichier, print_archive, print_deplace, print_supprime, print_erreur,
                         print_exception)


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


def move_message(message, outlookdir, config, keep_in_inbox=False, mark_as_read=False):
    """
    Utilisé pour parcours message par message un répertoire

    :param config:
    :param win32com.client.CDispatch message: message à déplacer
    :param win32com.client.CDispatch outlookdir: répertoire où déplacer
    :param boolean keep_in_inbox:
    :param boolean mark_as_read:
    """
    move = False
    try:
        if keep_in_inbox is True:
            if is_archivable(message, config) is True:
                move = True
        else:
            move = True

        if move is True:
            print_deplace(message.Subject[0:80])
            if mark_as_read is True:
                message.Unread = False
            message.Move(outlookdir)
    except Exception as ex:
        if message.Subject is not None:
            print_erreur("Err 004 : " + str(ex) + " / " + str(message.Subject))
        # print_exception()


def is_archivable(mail, config):
    """ Détermine si un mail peut être déplacé du répertoire source

        :param config:
        :param win32com.client.CDispatch mail: Mail à vérifier

        :return: Le mail doit-il être archivé
        :rtyp bool:
    """
    try:
        nbolddays = get_nb_old_days(mail)
        if int(config["outofinboxdays"]) > int(nbolddays.days):
            return False
        else:
            return True
    except Exception as ex:
        print_erreur("Err tools_message 005 : " + str(ex) + " / " + str(mail.Subject))
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


def move_mail(item, folder, config, lookup_type="Subject"):
    """ Deplacement de mail en fonction d'une liste de keywords (sujets)
        :param config:
        :param item:
        :param win32com.client.CDispatch folder: Mail à vérifier
        :param str lookup_type: Mail à vérifier
        """
    title = item["name"]
    kw = item["keywords"]
    keep_in_inbox = item["keepInInbox"]

    mark_as_read = False
    try:
        mark_as_read = item["markAsRead"]
    except (Exception,):
        pass

    deletionexception = None
    try:
        deletionexception = item["deletenotif"]
    except (Exception, ):
        pass

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
                                if lookup_type == "Sender":
                                    lookup_field = item.Sender
                                else:
                                    lookup_field = item.Subject

                                if keyword.encode("latin-1").decode("utf-8").lower() in str(lookup_field).lower():
                                    if is_archivable(item, config) is True:
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
                    if is_archivable(item, config) is True:
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


def archivemails(indir, archivedir, config, deletearchive=False):
    """ Archive les mails du répertoire en paramètre
    :param config:
    :param win32com.client.CDispatch indir: Répertoire d'entrée
    :param win32com.client.CDispatch archivedir: Répertoire d'archive
    :param bool deletearchive: Suppresion ou archivage du mail
    """
    for mail in indir.Items:
        try:
            nbolddays = get_nb_old_days(mail)
            if config["archivabledays"] < nbolddays.days:
                if deletearchive is False:
                    print_archive(mail.Subject)
                    mail.Move(archivedir)
                else:
                    print_supprime(mail)
        except (Exception,):
            print_exception()


def save_attachment(attachments, attach_ext, dir, config, prefix_name=None):
    """Sauvegarde la pièce joint sur SharePoint

    Args:
        attachments: Liste de pièces jointes du mail
        attach_ext: Extension à sélectionner
        dir: Répertoire où copier les fichiers
        prefix_name: Préfixe des fichiers à utiliser
    """
    relative_url = dir

    for i in range(1, len(attachments) + 1):
        attachment = attachments.Item(i)
        # the name of attachment file
        if prefix_name is None:
            attachment_name = str(attachment).upper()
        else:
            attachment_name = prefix_name + "_" + str(attachment).upper()

        if attachment_name.endswith(attach_ext.upper()):
            try:
                attach_file = os.path.join(config["sharepoint_sp_site"], attachment_name)
                attachment.SaveASFile(attach_file)

                client_credentials = UserCredential(config["sharepoint_login"], config["sharepoint_pass"])
                ctx = ClientContext(config["sharepoint_sp_site"]).with_credentials(client_credentials)

                remotepath = relative_url + "/" + attachment_name  # existing folder path under sharepoint site.
                print_fichier(remotepath)
                with open(attach_file, 'rb') as content_file:
                    file_content = content_file.read()

                dir_name, name = os.path.split(remotepath)

                ctx.web.get_folder_by_server_relative_url(dir_name).upload_file(name, file_content).execute_query()
                os.remove(attach_file)
            except (Exception, ):
                print("[[bright_red]KO[white]]     ctx.web.get_folder_by_server_relative_path(relative_url) : "
                      + relative_url)
