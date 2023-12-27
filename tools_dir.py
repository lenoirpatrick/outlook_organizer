#!python3
# -*- coding: utf-8 -*-
import win32com.client

from constants import *
from tools_message import set_subject, move_message


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
            except Exception as ex:
                # print_exception()
                print_erreur("Err 004_2 : " + str(ex) + " / " + str(message.Subject))


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


def set_archive_dir(indir, deletion_dir=False) -> win32com.client.CDispatch:
    """ Validation du répertoire d'archive de mail. S'il n'existe pas, le répertoire est créé

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
