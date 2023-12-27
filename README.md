# Répertoire Python : C:\Users\plenoir\Documents\AOG-AzureDevops\outlook_organizer

## .\console_log.py

> Description : 

### Imports utilisés

- rich
- rich.console
- rich.table

### Fonctions et classes

- Function: **table_recap** (Classe: ***['inbox', 'sentitems', 'deleteditems']***)
- Function: **print_titre** (Classe: ***['texte']***)
- Function: **print_check** (Classe: ***['texte']***)
- Function: **print_fichier** (Classe: ***['texte']***)
- Function: **print_no_response** (Classe: ***['texte', 'unread']***)
- Function: **print_archive** (Classe: ***['texte']***)
- Function: **print_deplace** (Classe: ***['texte']***)
- Function: **print_supprime** (Classe: ***['mail']***)
- Function: **print_erreur** (Classe: ***['texte']***)
- Function: **press_any_key** (Classe: ***[]***)
- Function: **print_exception** (Classe: ***[]***)

## .\constants.py

> Description : 

### Imports utilisés

- datetime
- win32com.client

## .\demo.py

> Description : 

### Imports utilisés

- constants
- email.mime.multipart
- email.mime.text
- os
- outlook_organizer
- rich
- smtplib

## .\outlook_organizer.py

> Description : 

### Imports utilisés

- console_log
- constants
- datetime
- email.mime.multipart
- email.mime.text
- git
- json
- os
- rich
- rich.progress
- smtplib
- sys
- tools_dir
- tools_message
- win32com.client

### Fonctions et classes

- Class: **OutlookOrganizer**
- Function: **__init__** (Classe: ***['self', 'jsonfile']***)
- Function: **empty_trash** (Classe: ***['self']***)
- Function: **delete_notifs_invits** (Classe: ***['self']***)
- Function: **process_notif** (Classe: ***['item']***)
- Function: **notifs_mails** (Classe: ***['self']***)
- Function: **mails_projets** (Classe: ***['self']***)
- Function: **mails_emails** (Classe: ***['self']***)
- Function: **process_mails_in_userdir** (Classe: ***['nb_msg_indir', 'indir']***)
- Function: **process_mails_in_indir** (Classe: ***['self', 'indir', 'item']***)
- Function: **mails_from_me** (Classe: ***['self']***)
- Function: **notifs_divers** (Classe: ***['self']***)
- Function: **recap_email** (Classe: ***['self']***)
- Function: **mails_sans_reponse_in_recap** (Classe: ***['self', 'item']***)
- Function: **mails_sans_reponse** (Classe: ***['self']***)
- Function: **process_sans_reponse** (Classe: ***['self', 'item']***)

## .\tools_dir.py

> Description : 

### Imports utilisés

- console_log
- tools_message
- win32com.client

### Fonctions et classes

- Function: **parse_dir** (Classe: ***['outlookdirin', 'outlookdirout', 'subject', 'config']***)
- Function: **set_indir** (Classe: ***['inbox', 'item']***)
- Function: **check_dir** (Classe: ***['dir', 'subdir']***)
- Function: **set_archive_dir** (Classe: ***['indir', 'quarter', 'deletion_dir']***)

## .\tools_message.py

> Description : 

### Imports utilisés

- console_log
- constants
- datetime
- office365.runtime.auth.user_credential
- office365.sharepoint.client_context
- os
- rich.progress

### Fonctions et classes

- Function: **set_subject** (Classe: ***['subject']***)
- Function: **move_message** (Classe: ***['message', 'outlookdir', 'config', 'keep_in_inbox', 'mark_as_read']***)
- Function: **is_archivable** (Classe: ***['mail', 'config']***)
- Function: **get_nb_old_days** (Classe: ***['mail']***)
- Function: **move_mail** (Classe: ***['item', 'folder', 'config', 'lookup_type']***)
- Function: **archivemails** (Classe: ***['indir', 'archivedir', 'config', 'deletearchive']***)
- Function: **save_attachment** (Classe: ***['attachments', 'attach_ext', 'dir', 'config', 'prefix_name']***)

