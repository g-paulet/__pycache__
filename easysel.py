"""Module de traitement des fichiers Easysel"""

import streamlit as st
import pandas as pd

def copier_donnees_easysel(sheet):
    """Simule le comportement du VBA CopierDonneesEasysel"""
    # Recherche des mots-clés
    prix_conditions_index = sheet[sheet.iloc[:, 0] == "PREZZI E CONDIZIONI"].index
    total_index = sheet[sheet.iloc[:, 0] == "TOTAL"].index

    if prix_conditions_index.empty or total_index.empty:
        st.error("Les sections 'PREZZI E CONDIZIONI' ou 'TOTAL' sont introuvables.")
        return None

    prix_conditions_row = prix_conditions_index[0] + 1
    total_row = total_index[0]

    # Recherche des colonnes nécessaires
    headers = sheet.iloc[prix_conditions_row].tolist()
    try:
        ref_col = headers.index("Ref.")
        code_col = headers.index("Code")
        qte_col = headers.index("Q.té")
    except ValueError:
        st.error("Les colonnes 'Ref.', 'Code', ou 'Q.té' sont introuvables.")
        return None

    # Extraction des données
    data = sheet.iloc[prix_conditions_row + 1 : total_row, [ref_col, code_col, qte_col]]
    data.columns = ["Ref.", "Code", "Q.té"]
    return data

def modifier_tableau_easysel(df):
    """Mise en forme du tableau, traitement des correspondances plénum"""

    # 1. S'assurer que le DataFrame a au moins 9 colonnes
    while df.shape[1] < 9:
        df[f'Col_{df.shape[1]+1}'] = ""

    # 2. Déplacer la colonne A vers la colonne I (index 8)
    #    # Renommer la colonne A en "Sous total"
    df.insert(8, 'Sous total', df.iloc[:, 0])  # Déplacer la colonne A vers la colonne I
    df.drop(df.columns[0], axis=1, inplace=True)  # Supprimer la colonne A d'origine

    # 3. Insérer une colonne vide entre A et B (index 1)
    df.insert(1, 'Libellé', "")  # Nouvelle colonne vide en B

    # 4. Renommer correctement les colonnes
    titres = ["Code", "Libellé", "Qté"] + [f"Col_{i+4}" for \
        i in range(df.shape[1] - 4)] + ["Sous total"]
    df.columns = titres

    # 5. Traitement des valeurs dans la colonne "Sous total" (ancienne colonne I)
    for i in df.index:
        # Vérifier si la valeur dans "Sous total" est une chaîne de caractères
        valeur_col_I = df.at[i, "Col_9"]
        if pd.notna(valeur_col_I) and not str(valeur_col_I).isnumeric():
            # Si c'est une chaîne de caractères, copier en "Libellé" et mettre "T" en "Sous total"
            df.at[i, "Libellé"] = valeur_col_I  # Copier dans la colonne B ("Libellé")
            df.at[i, "Sous total"] = "T"        # Remplacer par "T" dans la colonne I

    #Suppression de la colonne I
    df.drop(df.columns[8], axis=1, inplace=True)

    # 6. Remplir les cellules vides avec des chaînes vides pour éviter les NaN
    df.fillna("", inplace=True)

    # Intégration de la correspondance codes plénums
        # Ajout 19/03/25
        # Définition des codes "BEL"

    codes_bel = {
        9066613, 9066603, 9066593, 9066615, 9066605, 9066595,
        9066617, 9066607, 9066597, 9038037, 9038038, 9038039, 9038047
    }

    # Table de correspondance plénums
    table_correspondance = {
        9069180: (9069570, 9069570),
        9069190: (9069560, 9069190),
        9069181: (9069571, 9069571),
        9069191: (9069561, 9069191),
        9038050: (9069572, 9069572),
        9069222: (9069562, 9069222),
        9066468: (9069573, 9069573),
        9066368: (9069563, 9066368),
        9069185: (9069575, 9069575),
        9069195: (9069565, 9069195),
        9069186: (9069576, 9069576),
        9069196: (9069566, 9069196),
        9069188: (9069578, 9069578),
        9069198: (9069568, 9069198),
    }

    # Réinitialiser l'index pour éviter les erreurs d'accès aux lignes
    df.reset_index(drop=True, inplace=True)

    # Fonction pour traiter un bloc complet
    def traiter_bloc(df, indices_bloc):
        code_col_idx = df.columns.get_loc("Code")  # Récupérer l'index de la colonne "Code"
        # Récupérer les valeurs du bloc et vérifier si au moins un code est dans la liste BEL
        bloc_values = df.iloc[indices_bloc, code_col_idx].dropna().astype(str)

        # Remplacement des valeurs du bloc
        for idx in indices_bloc:
            valeur_actuelle = df.iloc[idx, code_col_idx]

            # Vérifier si AU MOINS UN code BEL est présent dans le bloc
            contient_bel = any(int(val) in codes_bel for val in bloc_values if val.isdigit())

            if pd.notna(valeur_actuelle) and str(valeur_actuelle).isdigit():
                valeur_actuelle = int(valeur_actuelle)

                if valeur_actuelle in table_correspondance:
                    nouvelle_valeur = (
                        table_correspondance[valeur_actuelle][1] if contient_bel
                        else table_correspondance[valeur_actuelle][0]
                    )
                    df.iloc[idx, code_col_idx] = nouvelle_valeur

    # Identification des blocs de lignes consécutives non vides
    indices_bloc = []
    for i, valeur in enumerate(df["Code"]):
        if pd.isna(valeur) or str(valeur).strip() == "":  
            # Si on rencontre une ligne vide, on traite le bloc en cours
            if indices_bloc:  
                traiter_bloc(df, indices_bloc)  
                indices_bloc = []  # Réinitialiser pour le prochain bloc
        else:
            indices_bloc.append(i)  # Ajouter l'index de la ligne au bloc actuel

    # Traiter le dernier bloc s'il existe
    if indices_bloc:
        traiter_bloc(df, indices_bloc)


    return df
