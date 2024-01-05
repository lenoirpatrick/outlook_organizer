#!python3
# -*- coding: utf-8 -*-
from constants import *
from console_log import (table_recap, print_titre, print_check, print_no_response, print_deplace, print_supprime,
                         press_any_key, print_exception, print_prologue)
from tools_dir import parse_dir, set_indir, set_archive_dir
from tools_message import set_subject, move_message, is_archivable, get_nb_old_days, move_mail, archivemails


class OutlookOrganizer:
    def __init__(self, jsonfile="appsettings.json"):
        self.jsonfile = jsonfile
        self.index = 0
        self.body = ""

        with open(os.path.join(os.path.dirname(__file__), self.jsonfile)) as json_data:
            self.__config = json.load(json_data)

        self.send_mail_recap = self.__config["parameters"]["send_mail_recap"]

        self.__outofinboxdays = self.__config['parameters']['outofinboxdays']
        self.__archivabledays = self.__config['parameters']['archivabledays']

        # utilisateur
        self.user_email = self.__config['user']['user_email']
        self.user_name = self.__config['user']['user_name']
        self.__proxy_active = self.__config["user"]["proxy_active"]
        self.__proxy_address = self.__config["user"]["proxy_address"]

        # Sharepoint config
        self.sharepoint_sp_site = self.__config["user"]["sharepoint_sp_site"]
        self.sharepoint_login = self.__config["user"]["sharepoint_login"]
        self.sharepoint_pass = self.__config["user"]["sharepoint_pass"]

        # email config sender
        self.sender_address = self.__config["email"]["sender_address"]
        self.sender_pass = self.__config["email"]["sender_pass"]

        self.time_short_version = self.__config['parameters']['time_short_version']

        self.email_smtp = self.__config['email']['smtp']
        self.email_port = self.__config['email']['port']

        # Déterminer le trimestre pour les sous dirs d'archivage
        now = datetime.now()  # date et heure actuelles
        jour_en_cours = now.day  # Jour en cours

        # Date du jour
        date_du_jour = date(now.year, now.month, now.day)  # Date du jour (année, mois, jour)
        # timenow = now.time()  # Heure actuelle

        datequarter = now - timedelta(self.__archivabledays)
        annee = datequarter.year
        mois = datequarter.month
        trimestre = (mois - 1) / 3 + 1
        self.quarter = str(annee) + "_Q" + str(int(trimestre))

        self.config = {
            "outofinboxdays": int(self.__outofinboxdays),
            "archivabledays": int(self.__archivabledays),
            "tmpdir": self.__config["parameters"]["tmpdir"],
            "sharepoint_sp_site": self.__config["user"]["sharepoint_sp_site"],
            "sharepoint_login": self.__config["user"]["sharepoint_login"],
            "sharepoint_pass": self.__config["user"]["sharepoint_pass"],
            "step_notifs_invits": self.__config["etapes"]["notifs_invits"],
            "step_notifs_mails": self.__config["etapes"]["notifs_mails"],
            "step_mails_projets": self.__config["etapes"]["mails_projets"],
            "step_mails_partenaires": self.__config["etapes"]["mails_partenaires"],
            "step_mails_internes": self.__config["etapes"]["mails_internes"],
            "step_from_me": self.__config["etapes"]["from_me"],
            "step_email_appointments": self.__config["etapes"]["email_appointments"],
            "step_unread_mails": self.__config["etapes"]["unread_mails"],
            "proxy_active": self.__proxy_active,
            "proxy_address": self.__proxy_address,
            "now": now,
            "jour_en_cours": jour_en_cours,
            "date_du_jour": date_du_jour
                       }

    def empty_trash(self):
        """Supprime les anciens mails du dossier Corbeille Outlook.

        La méthode parcourt les éléments du dossier Corbeille et supprime les mails
        qui répondent aux critères d'archivage définis dans la configuration.

        Returns:
            None
        """
        self.index = self.index + 1  # Incrémente l'index interne de la classe
        print_titre(str(self.index) + " - Suppression des anciens mails de la Corbeille")
        # Affiche une barre de progression pour indiquer l'avancement de la suppression
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            # Crée une étiquette pour afficher le nombre d'éléments dans le dossier Corbeille
            libelle = ("      Deleted Items (" + str(len(deleteditems.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(deleteditems.Items))
            for item in deleteditems.Items:  # Parcours des éléments dans la Corbeille
                progress.advance(task)  # Fait avancer la barre de progression
                if is_archivable(item, self.config) is True:
                    print_supprime(item)  # Supprime l'élément selon les critères définis

    def delete_notifs_invits(self):
        """Supprime les notifications telles que les invitations, les réponses automatiques et autres."""
        self.index = self.index + 1
        print_titre(str(self.index) + " - Suppression des notifications (invitations, AR, absences)")
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
                    self.process_notif(item)

    @staticmethod
    def process_notif(item):
        try:
            if not item.Unread:
                # Vérifie et supprime certains types de messages
                if item.Class in [46, 53, 54, 55, 57, 181] or (
                        item.Class == 43 and str(item.Subject).startswith("Réponse automatique")):
                    print_supprime(item)
            elif item.Class == 56:
                # Supprime les messages de réunions acceptées
                print_supprime(item)
        except Exception as ex:
            # Gère les exceptions et affiche les erreurs rencontrées
            print("Err 005 : " + str(ex) + " / " + str(item.Subject))
            print_exception()

    def notifs_mails(self):
        """Traite les notifications pour certains mails selon la configuration."""
        self.index = self.index + 1
        print_titre(str(self.index) + " - Traitement des notifications")
        for item in self.__config["notifications"]:
            if item["active"] is True:
                if "deletenotif" not in item:
                    item["deletenotif"] = ""

                # Vérification des répertoires et déplacement des mails
                indir = set_indir(inbox, item)  # Détermine le répertoire de destination pour les mails
                archivedir = set_archive_dir(indir, self.quarter,
                                             item["deletearchive"])  # Détermine le répertoire d'archivage

                # Déplace les mails vers le répertoire spécifié
                move_mail(item, indir, self.config, lookup_type="Sender")

                # Archive les mails déplacés
                archivemails(indir, archivedir, self.config,
                             item["deletearchive"])  # Archive les mails selon la configuration

    def mails_projets(self):
        """Traite les e-mails liés aux projets selon la configuration définie."""
        self.index = self.index + 1
        print_titre(str(self.index) + " - Traitement des Projets")

        for item in [item for item in self.__config["projects"] if item["active"]]:
            # Vérification des répertoires pour le projet et l'archivage
            indir = set_indir(inbox, item)  # Répertoire de destination pour les e-mails du projet
            archivedir = set_archive_dir(indir, self.quarter)  # Répertoire d'archivage pour les e-mails

            # Déplacement des e-mails du projet et archivage
            move_mail(item, indir, self.config)  # Déplace les e-mails vers le répertoire spécifié
            archivemails(indir, archivedir, self.config)  # Archive les e-mails déplacés

            nb_msg_indir = len(indir.Items)  # Nombre de messages dans le répertoire du projet
            if nb_msg_indir > 0:
                # Affiche une barre de progression pour le traitement des e-mails
                with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                              TimeElapsedColumn(), ) as progress:
                    libelle = (INDIR + str(nb_msg_indir) + ")").ljust(30)  # Étiquette pour la barre de progression
                    task = progress.add_task(libelle, total=len(indir.Items))  # Tâche pour la progression
                    for message in indir.Items:
                        progress.advance(task)  # Avance dans la barre de progression
                        subject = set_subject(message.Subject)  # Récupère le sujet du message
                        combined_lists = [inbox.Items, sentitems.Items]

                        for liste in combined_lists:
                            # Analyse des e-mails dans Inbox et Sent Items en fonction du sujet
                            parse_dir(liste, indir, subject, self.config)  # Analyse les e-mails pour le projet

    def mails_emails(self, json_section):
        # Parcours des mails d'une personne spécifique
        self.index = self.index + 1
        print_titre(str(self.index) + " - Traitement des messages " + json_section)

        for item in self.__config[json_section]:
            if item["active"] is True:
                title = (item["team"].encode("latin-1").decode("utf-8") + " > " +
                         item["dir"].encode("latin-1").decode("utf-8") + "/" +
                         item["subdir"].encode("latin-1").decode("utf-8"))
                print_check(title)

                # Vérification de la présence des répertoires
                indir = set_indir(inbox, item)
                archivedir = set_archive_dir(indir, self.quarter)

                self.process_mails_in_indir(indir, item)

                # Parcours rep utilisateur pour retrouver des messages dans la inbox et Send Items
                nb_msg_indir = len(indir.Items)
                if nb_msg_indir > 0:
                    self.process_mails_in_userdir(nb_msg_indir, indir)

                archivemails(indir, archivedir, self.config)

    @staticmethod
    def process_mails_in_userdir(nb_msg_indir, indir):
        with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                      TimeElapsedColumn(), ) as progress:
            libelle = (INDIR + str(nb_msg_indir) + ")").ljust(30)
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

    def process_mails_in_indir(self, indir, item):
        # Parcours de la boite de réception pour déplacer les messages vers le rep Utilisateur
        with Progress(SpinnerColumn(), *Progress.get_default_columns(),
                      TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for mail in inbox.Items:
                progress.advance(task)

                for user in item["users"]:
                    try:
                        if user.encode("latin-1").decode("utf-8").upper() in str(mail.Sender).upper():
                            move_message(mail, indir, self.config, keep_in_inbox=True)
                    except AttributeError:
                        pass
                    except (Exception,):
                        print_exception()

    def mails_from_me(self):
        """Supprime les messages de la boîte de réception provenant de l'utilisateur actuel."""
        self.index = self.index + 1
        print_titre(
            str(self.index) + " - Suppression des messages en inbox de moi vers " + self.user_email +
            " / " + self.sender_address)
        # Affiche une barre de progression pour indiquer l'avancement de la suppression
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                try:
                    progress.advance(task)
                    # Vérifie si l'expéditeur est l'utilisateur actuel et supprime le message
                    if str(item.Sender) in [self.user_name, self.user_email, self.sender_address]:
                        print_supprime(item)  # Supprime le message de la boîte de réception
                except (AttributeError,):
                    pass  # Ignore les erreurs liées à l'attribut Sender (si inexistant)

    def notifs_divers(self):
        """Supprime les e-mails de notifications diverses selon la configuration."""
        try:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Suppression des mails de notifs divers")
            emails = self.__config["safetodelete"]["emails"]
            # Barre de progression pour indiquer l'avancement de la suppression
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(inbox.Items))
                for item in inbox.Items:
                    try:
                        for mail in emails:
                            # Vérifie si l'e-mail de l'expéditeur correspond à une notification à supprimer
                            if mail in str(item.Sender):
                                if is_archivable(item, self.config) is True:
                                    print_supprime(item)  # Supprime l'e-mail s'il est archivable
                                break
                    except (AttributeError,):
                        pass  # Ignore les erreurs liées à l'attribut Sender (si inexistant)
                    progress.advance(task)  # Fait avancer la barre de progression
        except Exception as ex:
            # Gère les exceptions et affiche l'erreur en cas de problème
            print("[[bright_red]KO[white]]     Erreur lors du traitement des notification diverses.")
            print("[[bright_red]KO[white]]     " + str(ex))
            press_any_key()  # Attend une action de l'utilisateur

    def recap_email(self):
        """Crée un récapitulatif des rendez-vous du jour et l'envoie par e-mail si activé dans la configuration."""
        # Recap RDV
        if self.send_mail_recap is True:
            self.index = self.index + 1
            print_titre(str(self.index) + " - Récap des RDV du jour")
            appointments.Sort("[Start]")

            begin = date.today()
            end = begin + timedelta(days=1)
            restriction = "[Start] >= '" + begin.strftime("%d/%m/%Y") + "' AND [End] <= '" + end.strftime(
                "%d/%m/%Y") + "'"
            restricted_items = appointments.Restrict(restriction)

            joursemaine = self.config["now"].weekday() + 1
            with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
                libelle = ("      RDV (" + str(len(restricted_items)) + ")").ljust(30)
                task = progress.add_task(libelle, total=len(restricted_items))
                for appointment_item in restricted_items:
                    progress.advance(task)
                    if joursemaine == appointment_item.StartInStartTimeZone.isoweekday():
                        if str(appointment_item.StartInStartTimeZone)[8:10] == str(self.config["jour_en_cours"]):
                            # Construit le contenu pour le récapitulatif des rendez-vous
                            self.body += f"{str(appointment_item.Subject)}\n"
                            self.body += f"     Début : {str(appointment_item.StartInStartTimeZone)[:19]}\n"
                            self.body += f"     Durée : {str(appointment_item.Duration)}mn\n"
                            self.body += f"     Orga. : {str(appointment_item.Organizer)}\n"
                            self.body += f"     Recu. : {str(appointment_item.IsRecurring)}\n\n"

    def mails_sans_reponse_in_recap(self, item):
        if self.send_mail_recap is True:
            nbolddays = get_nb_old_days(item, self.config)
            self.body = self.body + "      " + item.Subject + " / " + str(item.Sender) + " [" + str(
                nbolddays.days) + "j]\n"

    def mails_sans_reponse(self):
        # Mails sans réponse
        self.index = self.index + 1
        if self.send_mail_recap is True:
            self.body = self.body + "\n"
            self.body = self.body + "Check des mails non lus : \n"
        print_titre(str(self.index) + " - Check des mails non lus")
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = (INBOX + str(len(inbox.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(inbox.Items))
            for item in inbox.Items:
                progress.advance(task)
                if item.Unread is True:
                    self.mails_sans_reponse_in_recap(item)
                    print_no_response(item.Subject, unread=True)

        self.index = self.index + 1
        if self.send_mail_recap is True:
            self.body += "\nCheck des mails sans réponses :\n"

        print_titre(str(self.index) + " - Check des mails sans réponses")
        with Progress(SpinnerColumn(), *Progress.get_default_columns(), TimeElapsedColumn(), ) as progress:
            libelle = ("      Sent Items (" + str(len(sentitems.Items)) + ")").ljust(30)
            task = progress.add_task(libelle, total=len(sentitems.Items))
            for item in sentitems.Items:
                progress.advance(task)
                self.process_sans_reponse(item)

    def process_sans_reponse(self, item):
        repondu = False

        for inbox_items in inbox.Items:
            if item.Subject in inbox_items.Subject:
                repondu = True

        # Le mail n'a pas de réponses et est archivable
        if repondu is False and is_archivable(item, self.config) is True:
            self.mails_sans_reponse_in_recap(item)
            print_no_response(item.Subject)


if __name__ == "__main__":
    os.system('cls')

    # Création d'une instance de la classe OutlookOrganizer en utilisant le fichier "appsettings_demo.json"
    OO = OutlookOrganizer()

    # Affichage prologue
    print_prologue(OO.config)

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

    print("[[green3]OK[white]]     Démarrage de l'application")
    print()

    # Suppression des éléments dans la corbeille Outlook
    OO.empty_trash()

    # Suppression des notifications et invitations
    if OO.config["step_notifs_invits"] is True:
        OO.delete_notifs_invits()

    # Gestion des notifications de mails
    if OO.config["step_notifs_mails"] is True:
        OO.notifs_mails()

    # Gestion des mails de projets
    if OO.config["step_mails_projets"] is True:
        OO.mails_projets()

    # Gestion des mails partenaires et interne
    for config_params in ["step_mails_partenaires", "step_mails_internes"]:
        if OO.config[config_params] is True:
            OO.mails_emails(config_params.split("_")[2])

    # Gestion des mails provenant de l'utilisateur actuel
    if OO.config["step_from_me"] is True:
        OO.mails_from_me()

    # Gestion des notifications diverses
    if OO.config["step_notifs_invits"] is True:
        OO.notifs_divers()

    # On ne fait pas ces étapes l'après midi
    if OO.config["now"].hour > OO.time_short_version:
        sys.exit(0)

    # Récapitulatif des rdvs
    if OO.config["step_email_appointments"] is True:
        OO.recap_email()

    # Gestion des e-mails sans réponse
    if OO.config["step_unread_mails"] is True:
        OO.mails_sans_reponse()

    print()

    # Envoi d'un récapitulatif par e-mail si la condition est vraie
    if OO.send_mail_recap is True:
        message = MIMEMultipart()

        # Construction du message e-mail
        message['From'] = OO.sender_address
        message['To'] = OO.user_email
        message['Subject'] = "Recap journalier " + str(OO.config["date_du_jour"])
        message.attach(MIMEText(OO.body, 'plain'))

        # Connexion au serveur SMTP et envoi du message
        session = smtplib.SMTP(OO.email_smtp, OO.email_port)  # utilisation de Gmail avec le port spécifié
        session.starttls()  # activation de la sécurité
        session.login(OO.sender_address, OO.sender_pass)  # connexion avec l'adresse e-mail et le mot de passe
        text = message.as_string()
        session.sendmail(OO.sender_address, OO.user_email, text)
        session.quit()
        print("[green3]Message recap envoyé")

    print("[green3]Traitement terminé")
    table_recap()
    press_any_key()
