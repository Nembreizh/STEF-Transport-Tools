#-------------------------------------------------------------------------------
# Name:        RepriseCAGroupage
# Purpose:
#
# Author:      Dreumont.b
#
# Created:     14/11/2023
# Copyright:   (c) Dreumont.b 2023
# Licence:     STEF TRANSPORT SAS
#-------------------------------------------------------------------------------

import subprocess
import tkinter as tk
from tkinter import filedialog

def executer_requete_sql(nagence, ndordre, nagence_dest, inclure_destinataires, inclure_expediteurs):
    # Paramètres de connexion à la base de données Oracle
    nom_service = 'GTIADSYN'
    utilisateur = 'UREG_NIF'
    mot_de_passe = '*********'

    # Initialisation des requêtes SQL
    requete_sql_destinataires = ""
    requete_sql_expediteurs = ""

    # Construction des requêtes en fonction des cases à cocher
    if inclure_destinataires:
        requete_sql_destinataires = f"""
select 'AD;DE;'||adr_rem.ntiers||';;'||adr_rem.ntiers_ext||';'||adr_rem.nom||';'||adr_rem.nom2||';'||
adr_rem.rue||';'||adr_rem.rue2||';'||adr_rem.ville||';'||adr_rem.cpost||';'||adr_rem.pays||';'||
adr_rem.nloc3||';'||adr_rem.nident_cee||';'||adr_rem.telephone||';'||adr_rem.fax||';'||
adr_rem.e_mail||';'||adr_rem.pdep||';;'
FROM adr_rem, tiers
where adr_rem.ntiers = tiers.ntiers
and nagence = '{nagence}'
and ndordre = '{ndordre}'
and t_affich = 'A'
and annul = 'N'
and typ_expe = 'N';
"""

    if inclure_expediteurs:
        requete_sql_expediteurs = f"""
select 'AD;EX;'||adr_rem.ntiers||';;'||adr_rem.ntiers_ext||';'||adr_rem.nom||';'||adr_rem.nom2||';'||
adr_rem.rue||';'||adr_rem.rue2||';'||adr_rem.ville||';'||adr_rem.cpost||';'||adr_rem.pays||';'||
adr_rem.nloc3||';'||adr_rem.nident_cee||';'||adr_rem.telephone||';'||adr_rem.fax||';'||
adr_rem.e_mail||';'||adr_rem.pdep||';;'
FROM adr_rem, tiers
where adr_rem.ntiers = tiers.ntiers
and nagence = '{nagence}'
and ndordre = '{ndordre}'
and t_affich = 'A'
and annul = 'N'
and typ_expe = 'O';
"""

    # Création du script SQL
    script_sqlplus = f"""\
    SET PAGESIZE 0
    SET LINESIZE 32767
    SET FEEDBACK OFF
    SET HEADING OFF
    SET ECHO OFF
    SET VERIFY OFF
    SET NEWPAGE NONE
    CONNECT {utilisateur}/{mot_de_passe}@{nom_service}
    {requete_sql_destinataires}
    {requete_sql_expediteurs}
    EXIT;
    """

    # Écriture du script SQL dans un fichier temporaire
    with open('script.sql', 'w') as fichier_script:
        fichier_script.write(script_sqlplus)

    # Appel de SQL*Plus depuis Python avec redirection de fichiers
    commande_sqlplus = ['sqlplus', '-S', f'{utilisateur}/{mot_de_passe}@{nom_service}', '@script.sql']
    try:
        # Exécution de SQL*Plus avec la requête SQL
        resultat = subprocess.run(commande_sqlplus, stdout=subprocess.PIPE, stderr=subprocess.PIPE)

        # Vérification si la sortie est None
        if resultat.stdout is None:
            raise subprocess.CalledProcessError(resultat.returncode, commande_sqlplus)

        # Décodage de la sortie en utilisant l'encodage approprié
        resultat_stdout = resultat.stdout.decode('utf-8', errors='replace')

        if resultat.returncode != 0:
            raise subprocess.CalledProcessError(resultat.returncode, commande_sqlplus)
    except subprocess.CalledProcessError as e:
        print(f"Erreur lors de l'exécution de la commande. Code de retour : {e.returncode}")
        return f"Erreur lors de l'exécution de la commande. Code de retour : {e.returncode}"
    finally:
        # Suppression du fichier temporaire
        subprocess.run(['del', 'script.sql'], shell=True)

    # Concaténation avec le reste du résultat
    resultat_formate = resultat_stdout.strip()

    return resultat_formate

def enregistrer_resultat(resultat):
    fichier = filedialog.asksaveasfile(defaultextension=".txt", filetypes=[("Fichiers texte", "*.txt")])
    if fichier is not None:
        with open(fichier.name, 'w', encoding="utf-8") as f:
            f.write(resultat)


def on_bouton_clic():
    global texte_resultat
    nagence = champ_nagence.get()
    ndordre = champ_ndordre.get()
    nagence_dest = champ_nagence_dest.get()  # Ajout pour récupérer l'agence destinataire
    inclure_destinataires = check_destinataires_var.get()
    inclure_expediteurs = check_expediteurs_var.get()

    # Vérification qu'au moins une case est cochée
    if not inclure_destinataires and not inclure_expediteurs:
        print("Veuillez cocher au moins une option (Destinataires ou Expéditeurs)")
        return

    resultat_requete = executer_requete_sql(nagence, ndordre, nagence_dest, inclure_destinataires, inclure_expediteurs)

    # Efface le contenu précédent
    texte_resultat.delete(1.0, tk.END)

    # Affiche les résultats sous forme tabulaire
    resultat_requete = resultat_requete.replace('\n', '')
    texte_resultat.insert(tk.END, resultat_requete)

    # Ajout des deux premières lignes si au moins une case est cochée
    if inclure_destinataires or inclure_expediteurs:
        lignes_intro = f"""\
01;OT;02;;3010671600109;201106151052;;;;{nagence_dest[-3:]};{ndordre};
PO;C;{nagence_dest[-3:]};{ndordre};;XX4;;S;D;TM;;;;;;;;36011;;;O;;;;3010671600109;;;3010671600109;201310241100;201310241100;;;;;;3010671600109;FR17;;O;;;3010671600109;;;;FR;;;;;;;;11;000016;;0;;;;;;;;;;;;PP;ME;;;;;;;;;;FR;;
"""
        resultat_requete = lignes_intro.strip() + resultat_requete

    # Enregistre le résultat dans un fichier texte
    if inclure_destinataires or inclure_expediteurs:
        enregistrer_resultat(resultat_requete)

# Création de la fenêtre principale
fenetre = tk.Tk()
fenetre.title("Exécution de requête SQL avec enregistrement du résultat")

# Création des widgets
etiquette_nagence = tk.Label(fenetre, text="N° Agence :")
champ_nagence = tk.Entry(fenetre)

etiquette_ndordre = tk.Label(fenetre, text="N° Ordre :")
champ_ndordre = tk.Entry(fenetre)

etiquette_nagence_dest = tk.Label(fenetre, text="N° Agence Destinataire :")  # Ajout pour l'agence destinataire
champ_nagence_dest = tk.Entry(fenetre)  # Ajout pour l'agence destinataire

check_destinataires_var = tk.BooleanVar()
check_destinataires = tk.Checkbutton(fenetre, text="Destinataires", variable=check_destinataires_var)

check_expediteurs_var = tk.BooleanVar()
check_expediteurs = tk.Checkbutton(fenetre, text="Expéditeurs", variable=check_expediteurs_var)

bouton_executer = tk.Button(fenetre, text="Exécuter Requête et Enregistrer", command=on_bouton_clic)

# Placement des widgets dans la fenêtre
etiquette_nagence.grid(row=0, column=0, padx=10, pady=10)
champ_nagence.grid(row=0, column=1, padx=10, pady=10)

etiquette_ndordre.grid(row=1, column=0, padx=10, pady=10)
champ_ndordre.grid(row=1, column=1, padx=10, pady=10)

etiquette_nagence_dest.grid(row=2, column=0, padx=10, pady=10)  # Ajout pour l'agence destinataire
champ_nagence_dest.grid(row=2, column=1, padx=10, pady=10)  # Ajout pour l'agence destinataire

check_destinataires.grid(row=3, column=0, padx=10, pady=10)
check_expediteurs.grid(row=3, column=1, padx=10, pady=10)

texte_resultat = tk.Text(fenetre, height=10, width=50, wrap=tk.WORD)
texte_resultat.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

bouton_executer.grid(row=4, column=0, columnspan=2, padx=10, pady=10)

# Lancement de la boucle principale
fenetre.mainloop()