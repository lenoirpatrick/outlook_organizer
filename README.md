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
- json
- os
- win32com.client

## .\outlook_organizer.py

> Description : 

### Imports utilisés

- console_log
- constants
- datetime
- email.mime.multipart
- email.mime.text
- git
- rich
- rich.progress
- smtplib
- sys
- tools_dir
- tools_message
- win32com.client

## .\tools_dir.py

> Description : 

### Imports utilisés

- constants
- tools_message
- win32com.client

### Fonctions et classes

- Function: **parse_dir** (Classe: ***['outlookdirin', 'outlookdirout', 'subject']***)
- Function: **set_indir** (Classe: ***['inbox', 'item']***)
- Function: **check_dir** (Classe: ***['dir', 'subdir']***)
- Function: **set_archive_dir** (Classe: ***['indir', 'deletion_dir']***)

## .\tools_message.py

> Description : 

### Imports utilisés

- console_log
- constants
- office365.runtime.auth.user_credential
- office365.sharepoint.client_context
- rich.progress

### Fonctions et classes

- Function: **set_subject** (Classe: ***['subject']***)
- Function: **move_message** (Classe: ***['message', 'outlookdir', 'keep_in_inbox', 'mark_as_read']***)
- Function: **is_archivable** (Classe: ***['mail']***)
- Function: **get_nb_old_days** (Classe: ***['mail']***)
- Function: **move_mail** (Classe: ***['title', 'kw', 'folder', 'keep_in_inbox', 'mark_as_read', 'lookup_type', 'deletionexception']***)
- Function: **archivemails** (Classe: ***['indir', 'archivedir', 'deletearchive']***)
- Function: **save_attachment** (Classe: ***['attachments', 'attach_ext', 'dir', 'prefix_name']***)

