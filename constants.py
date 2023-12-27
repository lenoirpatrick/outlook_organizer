#!python3
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
from datetime import datetime, date


INBOX = "      Inbox ("
INDIR = "      Indir ("

# Configuration de l'application

# Déterminer le trimestre pour les sous dirs d'archivage
now = datetime.now()  # current date and time
jour_en_cours = now.day

# Date du jour
date_du_jour = date(now.year, now.month, now.day)
timenow = now.time()

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
sentitems = outlook.GetDefaultFolder(5)
deleteditems = outlook.GetDefaultFolder(3)
appointments = outlook.GetDefaultFolder(9).Items
