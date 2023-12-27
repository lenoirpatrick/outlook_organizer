#!python3
# -*- coding: utf-8 -*-
import win32com.client

from tools_message import set_subject, move_message
from console_log import print_erreur


def parse_dir(outlookdirin, outlookdirout, subject, config):
    """Analyse les messages dans un répertoire spécifique pour déplacer ceux correspondant à un sujet donné.

    Args:
        outlookdirin: Répertoire d'entrée contenant les messages à analyser.
        outlookdirout: Répertoire de sortie où les messages correspondants seront déplacés.
        subject (str): Sujet à rechercher dans les messages.
        config: Configuration pour le déplacement des messages.

    Returns:
        None
    """
    move = True
    while move is True:
        move = False
        for message in outlookdirin:
            try:
                subject2 = set_subject(message.Subject)
                if len(subject) > 10:
                    # Vérifie si le sujet est présent dans le message et déplace le message
                    if (subject2 in subject or subject in subject2) and subject2 != "" and subject != "":
                        move_message(message, outlookdirout, config, keep_in_inbox=False)
                        move = True
                        break
            except Exception as ex:
                # En cas d'erreur, affiche l'erreur avec le sujet du message
                print_erreur("Err 004_2 : " + str(ex) + " / " + str(message.Subject))


def set_indir(inbox, item):
    """Définit le répertoire d'entrée pour un élément.

    Args:
        inbox: Boîte de réception principale.
        item (dict): Élément avec les clés 'dir' et 'subdir'.

    Returns:
        folder: Répertoire d'entrée spécifié par 'dir' et 'subdir'.
    """
    # Vérifie l'existence du répertoire principal
    check_dir(inbox, item["dir"])

    # Vérifie l'existence du sous-répertoire dans le répertoire principal
    check_dir(inbox.Folders[item["dir"]], item["subdir"])

    # Accède au répertoire d'entrée spécifié par 'dir' et 'subdir'
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


def set_archive_dir(indir, quarter, deletion_dir=False) -> win32com.client.CDispatch:
    """ Validation du répertoire d'archive de mail. S'il n'existe pas, le répertoire est créé

    :param quarter: Numéro du trimestre
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
