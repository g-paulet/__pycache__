import streamlit as st
import pandas as pd
import io


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


# Fonction pour traiter les données de l'onglet PULSAR
def traiter_pulsar(df_source, df_export):
    """Applique les transformations définies pour l'onglet PULSAR et ajoute les résultats à df_export"""
    
    # Vérification de l'existence des colonnes attendues
    if "I" not in df_source.columns or "Sous total" not in df_source.columns:
        st.error("Erreur : L'onglet 'PULSAR' ne contient pas les colonnes attendues.")
        return df_export

    # Copier les valeurs de la colonne I en colonne B si ce sont des chaînes de caractères
    df_source["B"] = df_source["I"].where(df_source["I"].apply(lambda x: isinstance(x, str)), df_source["B"])

    # Déplacer la colonne "Sous total" en colonne I (au lieu de J)
    if "Sous total" in df_source.columns:
        df_source.insert(8, "I", df_source.pop("Sous total"))

    # Ajouter la lettre "T" en colonne I uniquement si une chaîne de caractères est en colonne B
    df_source["I"] = df_source.apply(lambda row: "T" if isinstance(row["B"], str) else row["I"], axis=1)

    # Ajouter les données transformées au fichier d'export
    df_export = pd.concat([df_export, df_source], ignore_index=True)

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
    for i in range(14, 44):  # Indices commencent à 0 en Python
        # Vérifier que les cellules B, C et D ne sont pas vides
        if pd.notna(df_source.iloc[i, 1]) and pd.notna(df_source.iloc[i, 2]) and pd.notna(df_source.iloc[i, 3]):
            titre = f"{df_source.iloc[i, 1]}x {df_source.iloc[i, 2]} de {df_source.iloc[i, 3]}m"
            nouvelles_lignes.append({"Code Produit": "", "Libellé": titre, "Quantité": ""})
            
            # Parcourir les types et copier les données associées
            for j in range(len(types)):
                quantite = df_source.iloc[i, cols_quantite_indices[j]]  # Accès à la colonne de quantités
                code = df_source.iloc[i, cols_code_indices[j]]  # Accès à la colonne de codes
                libelle = types[j]

                # Vérifier si la quantité et le code sont valides
                if pd.notna(quantite) and quantite != 0 and pd.notna(code):
                    nouvelles_lignes.append({"Code Produit": code, "Libellé": libelle, "Quantité": quantite})

    # Traitement des accessoires DS18 (lignes 50 à 79)
    nouvelles_lignes.append({"Code Produit": "", "Libellé": "Accessoires DS18", "Quantité": "T"})

    for i in range(49, 79):
        if pd.notna(df_source.iloc[i, 0]):
            code = df_source.iloc[i, 0]
            libelle = df_source.iloc[i, 1] if pd.notna(df_source.iloc[i, 1]) else ""
            quantite = df_source.iloc[i, 15] if pd.notna(df_source.iloc[i, 15]) else ""

            nouvelles_lignes.append({"Code Produit": code, "Libellé": libelle, "Quantité": quantite})

    # Concaténation avec le DataFrame export
    df_export = pd.concat([df_export, pd.DataFrame(nouvelles_lignes)], ignore_index=True)

    return df_export




# Interface Streamlit
st.title("Exportation des données PULSAR & DS18")

# Upload du fichier source
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xlsm"])
if uploaded_file:
    # Lecture du fichier source
    xl = pd.ExcelFile(uploaded_file)
    
    # Chargement des onglets PULSAR et DS18 s'ils existent
    df_pulsar = xl.parse("PULSAR") if "PULSAR" in xl.sheet_names else None
    df_ds18 = xl.parse("DS18") if "DS18" in xl.sheet_names else None

    # Création du DataFrame de sortie
    df_export = pd.DataFrame(columns=["Code Produit", "Libellé", "Quantité"])

    # Traitement des données de PULSAR
    if df_pulsar is not None:
        df_export = traiter_pulsar(df_pulsar, df_export)

    # Traitement des données de DS18
    if df_ds18 is not None:
        df_export = traiter_ds18(df_ds18, df_export)

    # Génération du fichier Excel de sortie
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df_export.to_excel(writer, sheet_name="Données Exportées", index=False)

    st.success("Traitement terminé ! Téléchargez votre fichier ci-dessous.")

    # Bouton de téléchargement
    st.download_button(
        label="Télécharger le fichier Excel",
        data=output.getvalue(),
        file_name="export_donnees.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
