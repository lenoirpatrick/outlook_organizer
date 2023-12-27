#!python3
# -*- coding: utf-8 -*-
import os
import sys
import git

# Rich text dans la console
from rich import print
from rich.table import Table
from rich.console import Console
from rich.progress import Progress, TimeElapsedColumn, SpinnerColumn

# Envoie d'email
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

# Connexion à SharePoint
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# Accès Outlook
import win32com.client
from win32com.client import Dispatch

from datetime import datetime, date, timedelta

INBOX = "      Inbox ("
INDIR = "      Indir ("

# Accès aux répertoires Outlook par défaut
outlook = Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)  # Répertoire Inbox par défaut (6 pour Inbox)
sentitems = outlook.GetDefaultFolder(5)  # Répertoire Sent Items par défaut (5 pour Sent Items)
deleteditems = outlook.GetDefaultFolder(3)  # Répertoire Deleted Items par défaut (3 pour Deleted Items)
appointments = outlook.GetDefaultFolder(9).Items  # Éléments de calendrier par défaut (9 pour Calendar)
