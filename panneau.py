"""Module de traitement des fichiers panneau"""

import streamlit as st
import pandas as pd


def lettre_en_index(lettre):
    """
    Convertit une référence de colonne sous forme de lettres (ex: 'A', 'Z', 'AA') 
    en un index numérique pour Pandas.
    """
    index = 0
    for char in lettre:
        index *= 26
        index += ord(char.upper()) - ord('A') + 1
    return index - 1  # Ajuste l'index pour correspondre aux indices Python (commence à 0)



def traiter_pulsar(df_source, df_export):
    """Fonction pour traiter les données de l'onglet PULSAR"""
    # 🔹 Définir les colonnes cibles (doivent être identiques à df_export)
    colonnes_cibles = ["Code Produit", "Libellé", "Quantité"]

    # 🔹 Initialisation d'une liste pour stocker les données
    lignes_pulsar = []

    titres_exportes = set()
    dernier_titre = ""

    # Utilisation des indices numériques pour éviter les erreurs de colonnes
    col_titre = 0   # Colonne A
    col_quantite = 1 # Colonne B
    col_type = 2     # Colonne C
    col_position = 3 # Colonne D
    col_ref = 14     # Colonne O

    # 🔹 1. Traitement par titre (Plage A15:A41)
    for i, row in df_source.iloc[14:41].iterrows():
        if pd.notna(row.iloc[col_titre]):  # Vérifier si la cellule A contient un titre
            dernier_titre = row.iloc[col_titre]

            if dernier_titre not in titres_exportes:
                titres_exportes.add(dernier_titre)
                lignes_pulsar.append(["", dernier_titre, ""])  # Ajout du titre

        if dernier_titre != "":
            quantite = row.iloc[col_quantite]
            if pd.notna(quantite):
                type_valeur = row.iloc[col_type]
                position = row.iloc[col_position]
                reference_valeur = row.iloc[col_ref]

                libelle = f"{type_valeur} {position}" if pd.notna(type_valeur) \
                    and pd.notna(position) else ""
                lignes_pulsar.append([reference_valeur, libelle, quantite])

    # 🔹 2. Cas sans titre en colonne A (Plage O15:O41)
    if dernier_titre == "":
        for i, row in df_source.iloc[14:41].iterrows():
            reference_valeur = row.iloc[col_ref]
            if pd.notna(reference_valeur):
                quantite = row.iloc[col_quantite]
                if pd.notna(quantite):
                    type_valeur = row.iloc[col_type]
                    position = row.iloc[col_position]

                    libelle = f"{type_valeur} {position}" if pd.notna(type_valeur) \
                        and pd.notna(position) else ""
                    lignes_pulsar.append([reference_valeur, libelle, quantite])

    # 🔹 3. Traitement des options/accessoires (Plage A47:O69)
    if lignes_pulsar:
        lignes_pulsar.append(["", "", ""])  # Ligne vide pour séparer
        lignes_pulsar.append(["", "Accessoires PULSAR", ""])

        for i, row in df_source.iloc[45:69].iterrows():
            if pd.notna(row.iloc[col_titre]):  # Vérifier si la cellule contient un code produit
                # Code produit + quantité
                lignes_pulsar.append([row.iloc[col_titre], "", row.iloc[col_ref]])

    # 🔹 Convertir la liste en DataFrame avec les bonnes colonnes
    df_pulsar = pd.DataFrame(lignes_pulsar, columns=colonnes_cibles)

    # 🔹 **Concaténer les résultats** sans décaler les colonnes
    df_export = pd.concat([df_export, df_pulsar], ignore_index=True)

    return df_export



def traiter_ds18(df_source, df_export):
    """Extrait et formate les données de l'onglet DS18 pour les ajouter au fichier de sortie"""

    if df_source is None:
        st.error("Erreur : L'onglet 'DS18' n'existe pas dans le fichier source.")
        return df_export

    # Liste des types, colonnes de quantités et colonnes codes correspondantes
    types = ["PanneauComplet", "PanneauFirst", "PanneauIntermédiaire", "PanneauFinal",
             "CapotEntree", "CapotInter", "CapotFinal", "Jonction", "Travail", "CacheTube"]
    cols_quantite = ["AR", "AZ", "BH", "BP", "BY", "CA", "CC", "CE", "CH", "CJ"]
    cols_code = ["AS", "BA", "BI", "BQ", "BZ", "CB", "CD", "CF", "CI", "CK"]

    # Conversion en indices numériques
    cols_quantite_indices = [lettre_en_index(col) for col in cols_quantite]
    cols_code_indices = [lettre_en_index(col) for col in cols_code]

    # Liste pour stocker les nouvelles lignes avant concaténation
    nouvelles_lignes = []

    # Traitement des panneaux DS18 (lignes 15 à 44)
    for i in range(13, 44):  # Indices commencent à 0 en Python
        # Vérifier que les cellules B, C et D ne sont pas vides
        if pd.notna(df_source.iloc[i, 1]) and pd.notna(df_source.iloc[i, 2]) \
            and pd.notna(df_source.iloc[i, 3]):
            titre = f"{df_source.iloc[i, 1]}x {df_source.iloc[i, 2]} de {df_source.iloc[i, 3]}m"
            nouvelles_lignes.append({"Code Produit": "", "Libellé": titre, "Quantité": ""})

            # Parcourir les types et copier les données associées
            for j, libelle in enumerate(types):
                # Accès à la colonne de quantités
                quantite = df_source.iloc[i, cols_quantite_indices[j]]
                # Accès à la colonne de codes
                code = df_source.iloc[i, cols_code_indices[j]]

                # Vérifier si la quantité et le code sont valides
                if pd.notna(quantite) and quantite != 0 and pd.notna(code):
                    nouvelles_lignes.append({"Code Produit": code, \
                    "Libellé": libelle, "Quantité": quantite})

    # Traitement des accessoires DS18 (lignes 50 à 79)

    if nouvelles_lignes: #19/03/25 ajout pour éviter que les accessoires DS18 apparaissent si uniquement produit PULSAR
        nouvelles_lignes.append({"Code Produit": "", "Libellé": "Accessoires DS18", "Quantité": ""})

        for i in range(48, 79):
            if pd.notna(df_source.iloc[i, 0]):
                code = df_source.iloc[i, 0]
                libelle = df_source.iloc[i, 1] if pd.notna(df_source.iloc[i, 1]) else ""
                quantite = df_source.iloc[i, 15] if pd.notna(df_source.iloc[i, 15]) else ""

                nouvelles_lignes.append({"Code Produit": code, \
                "Libellé": libelle, "Quantité": quantite})

    # Concaténation avec le DataFrame export
    df_export = pd.concat([df_export, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    return df_export


def modifier_tableau_panneau(df_export):
    """Mise en forme du dataframe."""

    # 1. S'assurer que le DataFrame a au moins 11 colonnes
    while df_export.shape[1] < 11:
        df_export[f'Col_{df_export.shape[1] + 1}'] = ""

    # 2. Ajouter les titres
    titres = ["Code", "Libellé", "Qté", "Col_4", "Col_5", "Col_6", "Col_7", "Col_8", "Sous total", "Col_10", "Col_11"]
    df_export.columns = titres  # On applique les titres au DataFrame

    # 3a Supprimer les lignes où les colonnes A et B sont vides
    df_export = df_export[~(df_export["Code"].isna() | (df_export["Code"] == "")) |
            ~(df_export["Libellé"].isna() | (df_export["Libellé"] == ""))]
    # 3b Supprimer les lignes où les colonnes B et C sont vides
    df_export = df_export[~(df_export["Qté"].isna() | (df_export["Qté"] == "")) |
            ~(df_export["Libellé"].isna() | (df_export["Libellé"] == ""))]

    # 4. Ajout de "T" en colonne "Sous total" si SEUL "Libellé" contient une donnée
    if "Sous total" in df_export.columns:
        for i, row in df_export.iterrows():
            if pd.notna(row["Libellé"]) and (pd.isna(row["Code"]) or row["Code"] == "") \
                and (pd.isna(row["Qté"]) or row["Qté"] == ""):
                df_export.at[i, "Sous total"] = "T"

    # Remplacement des NaN par ""
    df_export.fillna("", inplace=True)

    # Déplacement des libellés en colonne K si ce n'est pas un titre
    # Vérifier si "Sous total" existe avant de commencer
    if "Sous total" in df_export.columns:
        for i, row in df_export.iterrows():
            if row["Sous total"] != "T":
                df_export.at[i, "Col_11"] = row["Libellé"]  # Déplacement vers Col_11
                df_export.at[i, "Libellé"] = ""  # Vider la colonne "Libellé" après déplacement

    return df_export
