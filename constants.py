#!python3
# -*- coding: utf-8 -*-
from win32com.client import Dispatch
import json
import os
from datetime import datetime, date, timedelta


INBOX = "      Inbox ("
INDIR = "      Indir ("

with open(os.path.join(os.path.dirname(__file__), 'appsettings.json')) as json_data:
    config = json.load(json_data)
outofinboxdays = config['parameters']['outofinboxdays']
archivabledays = config['parameters']['archivabledays']

# Déterminer le trimestre pour les sous dirs d'archivage
now = datetime.now()  # current date and time
datequarter = now - timedelta(archivabledays)
annee = datequarter.year
mois = datequarter.month
trimestre = (mois - 1) / 3 + 1
quarter = str(annee) + "_Q" + str(int(trimestre))

# Date du jour
date_du_jour = date(now.year, now.month, now.day)
timenow = now.time()

outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)
sentitems = outlook.GetDefaultFolder(5)
deleteditems = outlook.GetDefaultFolder(3)
appointments = outlook.GetDefaultFolder(9).Items
