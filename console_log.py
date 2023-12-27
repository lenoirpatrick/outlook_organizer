#!python3
# -*- coding: utf-8 -*-
from rich import print
from rich.table import Table
from rich.console import Console


def table_recap(inbox, sentitems, deleteditems):
    """Affiche un tableau récapitulatif des principaux répertoires Outlook.

    Args:
        inbox: Répertoire Inbox.
        sentitems: Répertoire Sent Items.
        deleteditems: Répertoire Deleted Items.

    Returns:
        None
    """
    table = Table(title="Répertoires Outlook")

    # Ajoute des colonnes au tableau pour les répertoires et le nombre de messages
    table.add_column("Répertoire", justify="left", style="cyan", no_wrap=True)
    table.add_column("Nb Messages", style="magenta")

    # Ajoute des lignes au tableau pour les répertoires Inbox, Sent Items et Deleted Items
    table.add_row("Inbox", str(len(inbox.Items)))
    table.add_row("Send Items", str(len(sentitems.Items)))
    table.add_row("Deleted Items", str(len(deleteditems.Items)))

    # Affiche le tableau récapitulatif
    print(table)


def print_titre(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print()
    print("[blue]" + str(texte))


def print_check(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[deep_sky_blue3]    " + str(texte))


def print_fichier(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[deep_sky_blue3]            Enregistrement du document : " + str(texte))


def print_no_response(texte, unread=False):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    :param bool unread: Distinction sur le cas de gestion
    """
    if unread is False:
        print("[deep_sky_blue3]            Mail sans réponse à traiter manuellement : " + str(texte))
    else:
        print("[deep_sky_blue3]            Mail non lu : " + str(texte))


def print_archive(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[blue_violet]            Archivage du message : " + str(texte))


def print_deplace(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[green3]            Déplacement du message : " + str(texte))


def print_supprime(mail):
    """ Affiche du texte mis en forme dans la console

    :param win32com.client.CDispatch mail: Mail à supprimer
    """
    print("[dark_orange3]            Suppression du message : " + str(mail.Subject))
    try:
        mail.delete()
    except (Exception,):
        print_exception()


def print_erreur(texte):
    """ Affiche du texte mis en forme dans la console

    :param str texte: Texte à afficher
    """
    print("[bright_red]            " + str(texte))


def press_any_key():
    """ Demande une action utilisateur
    """
    input("Pressez une touche pour quitter")


def print_exception():
    """ Affiche la pile sur une erreur via le package rich
    """
    console = Console()
    console.print_exception(show_locals=True)
