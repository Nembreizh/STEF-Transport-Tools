#-------------------------------------------------------------------------------
# Name:        RecTiersExt
# Purpose:
#
# Author:      Dreumont.b
#
# Created:     27/10/2023
# Copyright:   (c) Dreumont.b 2023
# Licence:     STEF TRANSPORT SAS
# Version:     1.0
#-------------------------------------------------------------------------------

import subprocess
import tkinter as tk
from tkinter import ttk

def est_gln_valide(reference):
    return reference.isdigit() and len(reference) == 13

def executer_requete(ntiers_ext_saisi):
     # Vérification de la validité du GLN
    if not est_gln_valide(ntiers_ext_saisi):
        return "Erreur : La référence externe doit être un GLN numérique sur 13 caractères."

    # Paramètres de connexion à la base de données Oracle
    nom_service = 'GTIADSYN'
    utilisateur = 'UREG_NIF'
    mot_de_passe = '*********'

    # Requête SQL
    requete_sql = f"""
        SELECT NTIERS, NTIERS_EXT, COUNT(*) as NombreOccurences
        FROM ADR_REM
        WHERE NTIERS_EXT = '{ntiers_ext_saisi}'
        GROUP BY NTIERS, NTIERS_EXT;
    """

    # Création du script SQL
    script_sql = f"""\
    SET PAGESIZE 0
    SET FEEDBACK OFF
    SET HEADING OFF
    SET ECHO OFF
    SET VERIFY OFF
    CONNECT {utilisateur}/{mot_de_passe}@{nom_service}
    {requete_sql}
    EXIT;
    """

    # Écriture du script SQL dans un fichier temporaire
    with open('script.sql', 'w') as fichier_script:
        fichier_script.write(script_sql)

    # Appel de SQL*Plus depuis Python avec redirection de fichiers
    commande_sqlplus = ['sqlplus', '-S', f'{utilisateur}/{mot_de_passe}@{nom_service}', '@script.sql']
    resultat = subprocess.run(commande_sqlplus, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

    # Suppression du fichier temporaire
    subprocess.run(['del', 'script.sql'], shell=True)

    return resultat.stdout.strip()

def on_bouton_clic():
    ntiers_ext_saisi = champ_saisie.get()
    # Vérification de la validité du GLN
    if not est_gln_valide(ntiers_ext_saisi):
        texte_resultat.delete(1.0, tk.END)
        texte_resultat.insert(tk.END, "Erreur : La référence externe doit être un GLN numérique sur 13 caractères.")
        return

    resultat_requete = executer_requete(ntiers_ext_saisi)

    # Efface le contenu précédent
    texte_resultat.delete(1.0, tk.END)

    # Affiche les noms de colonnes
    texte_resultat.insert(tk.END, "NTIERS\tNTIERS_EXT\t    NB_CONV\n")

    # Affiche les résultats sous forme tabulaire
    for ligne in resultat_requete.split('\n'):
        texte_resultat.insert(tk.END, ligne + '\n')

# Création de la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Recherche de tiers dans les autres CA")

# Création des widgets
etiquette = tk.Label(fenetre, text="Entrez le référence externe du client (GLN numérique sur 13 caractères) :")
champ_saisie = tk.Entry(fenetre)
bouton_executer = tk.Button(fenetre, text="Exécuter Requête", command=on_bouton_clic)
texte_resultat = tk.Text(fenetre, height=10, width=50)

# Message en bas de la fenêtre
message_label = tk.Label(fenetre, text="Attention : Bien vérifier la conversion dans le Carnet d'adresse en vérifiant les champs adresse (Nom, Rue, Code Postal, Ville...)")

# Placement des widgets dans la fenêtre
etiquette.grid(row=0, column=0, padx=10, pady=10, columnspan=3)
champ_saisie.grid(row=1, column=0, padx=10, pady=10, columnspan=2)
bouton_executer.grid(row=1, column=2, padx=10, pady=10)
texte_resultat.grid(row=2, column=0, columnspan=3, padx=10, pady=10)
message_label.grid(row=3, column=0, columnspan=3, padx=10, pady=10)

# Lancement de la boucle principale
fenetre.mainloop()