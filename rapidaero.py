"""Module dev Rapidaero"""

import pandas as pd
import streamlit as st


def copier_tableaux_rapid_aero(excel_file):
    """ Copie les données des feuilles se terminant par '°C' """
    liste_feuilles = excel_file.sheet_names  # Récupérer la liste des feuilles
    resultats = []
    ligne_destination = 0

    for sheet_name in liste_feuilles:
        if sheet_name.endswith("°C"):
            st.write(f"Traitement de la feuille : {sheet_name}")  # Debugging
            # ajout dtype=str pour éviter le problème de référence 0008314
            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None, dtype=str)

            # Vérifier que le fichier contient au moins 18 colonnes (jusqu'à la colonne R)
            if df.shape[1] > 17:
                # Trouver la dernière ligne utilisée dans la colonne P (colonne index 15)
                dern_ligne = df.iloc[:, 15].last_valid_index()

                if dern_ligne and dern_ligne >= 1:  # Vérifier qu'il y a bien des données
                    print("Dernière ligne remplie en colonne P :", dern_ligne)

                # Créer une liste vide pour stocker les données formatées
                extrait_liste = []

                #  **Ajouter le titre dans la première colonne (ligne actuelle)**
                # Colonnes P, Q, R vides
                extrait_liste.append([f"Aérotherme - {sheet_name}"] + [""] * 2)

                #  **Ajouter les données à partir de la ligne suivante**
                extrait_liste.extend(df.iloc[1:dern_ligne + 1, [15, 16, 17]].values.tolist())

                # Convertir la liste en DataFrame pandas
                extrait = pd.DataFrame(extrait_liste, columns=["P", "Q", "R"])

                # Ajouter au résultat final
                resultats.append(extrait)

                # Mise à jour de la ligne de destination (incrémentation)
                ligne_destination += len(extrait)

    if resultats:
        return pd.concat(resultats, ignore_index=True)
    else:
        return None

def mise_en_forme_rapid_aero(df):
    """
    Met en forme le tableau extrait :
    - Déplace une colonne,
    - Supprime les colonnes inutiles,
    - Ajoute des titres,
    - Supprime les lignes vides.
    """

    # S'assurer que le DataFrame a au moins 11 colonnes
    while df.shape[1] < 11:
        df[f'Col_{df.shape[1] + 1}'] = ""

    # Ajouter les titres
    titres = ["Code", "Libellé", "Qté", "Col_4", "Col_5", "Col_6", "Col_7", "Col_8", "Sous total", "Col_10", "Col_11"]
    df.columns = titres  # On applique les titres au DataFrame

    # Swap colonnes A et B
    df[["Code", "Libellé", "Qté"]] = df[["Libellé","Code", "Qté"]]

    print("Colonnes dans le df :", titres)

    # Ajouter "T" dans "sous-total" si "Aérotherme" est trouvé dans "Libellé"
    df.loc[df["Libellé"].str.contains("Aérotherme", na=False), "Sous total"] = "T"

    # Supprimer les lignes où les colonnes A et I sont vides
    df = df[~(df["Code"].isna() | (df["Code"] == "")) |
            ~(df["Sous total"].isna() | (df["Sous total"] == ""))]

    # Déplacement des libellés en colonne K si ce n'est pas un titre
    # Vérifier si "Sous total" existe avant de commencer
    if "Sous total" in df.columns:
        for i, row in df.iterrows():
            if row["Sous total"] != "T":
                df.at[i, "Col_11"] = row["Libellé"]  # Déplacement vers Col_11
                df.at[i, "Libellé"] = ""  # Vider la colonne "Libellé"

    return df
# End-of-file (EOF)
