#!python3
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
from datetime import datetime, date


INBOX = "      Inbox ("
INDIR = "      Indir ("

# Déterminer le trimestre pour les sous-dirs d'archivage
now = datetime.now()  # date et heure actuelles
jour_en_cours = now.day  # Jour en cours

# Date du jour
date_du_jour = date(now.year, now.month, now.day)  # Date du jour (année, mois, jour)
timenow = now.time()  # Heure actuelle

# Accès aux répertoires Outlook par défaut
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Répertoire Inbox par défaut (6 pour Inbox)
sentitems = outlook.GetDefaultFolder(5)  # Répertoire Sent Items par défaut (5 pour Sent Items)
deleteditems = outlook.GetDefaultFolder(3)  # Répertoire Deleted Items par défaut (3 pour Deleted Items)
appointments = outlook.GetDefaultFolder(9).Items  # Éléments de calendrier par défaut (9 pour Calendar)
