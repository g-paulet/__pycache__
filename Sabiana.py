import streamlit as st
import pandas as pd
import io

def copier_donnees_easysel(sheet):
    """Simule le comportement du VBA CopierDonneesEasysel."""
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
    """Corrige fidèlement le traitement selon la macro VBA."""

    # 1. S'assurer que le DataFrame a au moins 9 colonnes
    while df.shape[1] < 9:
        df[f'Col_{df.shape[1]+1}'] = ""

    # 2. Déplacer la colonne A vers la colonne I (index 8) 
    # Renommer la colonne A en "Sous total"
    df.insert(8, 'Sous total', df.iloc[:, 0])  # Déplacer la colonne A vers la colonne I
    df.drop(df.columns[0], axis=1, inplace=True)  # Supprimer la colonne A d'origine

    # 3. Insérer une colonne vide entre A et B (index 1)
    df.insert(1, 'Libellé', "")  # Nouvelle colonne vide en B

    # 4. Renommer correctement les colonnes
    titres = ["Code", "Libellé", "Qté"] + [f"Col_{i+4}" for i in range(df.shape[1] - 4)] + ["Sous total"]
    df.columns = titres

    # 5. Traitement des valeurs dans la colonne "Sous total" (ancienne colonne I)
    for i in df.index:
        # Vérifier si la valeur dans "Sous total" est une chaîne de caractères
        valeur_col_I = df.at[i, "Sous total"]

        if pd.notna(valeur_col_I) and not str(valeur_col_I).isnumeric():
            # Si c'est une chaîne de caractères, copier en "Libellé" et mettre "T" en "Sous total"
            df.at[i, "Libellé"] = valeur_col_I  # Copier dans la colonne B ("Libellé")
            df.at[i, "Sous total"] = "T"        # Remplacer par "T" dans la colonne I

    # 6. Remplir les cellules vides avec des chaînes vides pour éviter les NaN
    df.fillna("", inplace=True)

    return df

# Interface web avec Streamlit
st.image("sabiana-logo.png", use_container_width=True)
st.title("Traitement des fichiers Excel - Easysel à Magenta")
st.subheader("Déposez votre fichier Excel ci-dessous :")

uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx"])

if uploaded_file:
    try:
        # Charger le fichier Excel
        sheet = pd.read_excel(uploaded_file, sheet_name=0, header=None)

        # Étape 1 : Copier les données
        st.info("Traitement des données en cours...")
        donnees = copier_donnees_easysel(sheet)

        if donnees is not None:
            # Étape 2 : Modifier le tableau
            resultat = modifier_tableau_easysel(donnees)

            # Télécharger le fichier traité
            st.success("Traitement terminé ! Téléchargez le fichier ci-dessous :")
            # Création d'un tampon en mémoire
            buffer = io.BytesIO()
            resultat.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)  # Remet le curseur au début du fichier

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=buffer,
                file_name="fichier_traite.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
