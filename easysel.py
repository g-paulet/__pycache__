import streamlit as st
import pandas as pd
import io

def copier_donnees_easysel(sheet):
    """Simule le comportement du VBA CopierDonneesEasysel."""
    prix_conditions_index = sheet[sheet.iloc[:, 0] == "PREZZI E CONDIZIONI"].index
    total_index = sheet[sheet.iloc[:, 0] == "TOTAL"].index

    if prix_conditions_index.empty or total_index.empty:
        st.error("Les sections 'PREZZI E CONDIZIONI' ou 'TOTAL' sont introuvables.")
        return None

    prix_conditions_row = prix_conditions_index[0] + 1
    total_row = total_index[0]

    headers = sheet.iloc[prix_conditions_row].tolist()
    try:
        ref_col = headers.index("Ref.")
        code_col = headers.index("Code")
        qte_col = headers.index("Q.té")
    except ValueError:
        st.error("Les colonnes 'Ref.', 'Code', ou 'Q.té' sont introuvables.")
        return None

    data = sheet.iloc[prix_conditions_row + 1 : total_row, [ref_col, code_col, qte_col]]
    data.columns = ["Ref.", "Code", "Q.té"]
    return data

def modifier_tableau_easysel(df):
    """Adaptation fidèle de la macro VBA en Python."""
    
    # 1. Déplacer la première colonne vers la colonne I (index 8)
    df.insert(8, 'Sous total', df.iloc[:, 0])
    df.drop(df.columns[0], axis=1, inplace=True)

    # 2. Insérer une colonne vide entre A et B
    df.insert(1, 'Libellé', "")

    # 3. Ajouter les titres
    titres = ["Code", "Libellé", "Qté"] + [f"Col_{i+4}" for i in range(df.shape[1] - 4)] + ["Sous total"]
    df.columns = titres

    # 4. Traitement conditionnel de la colonne "Sous total"
    for i in df.index:
        valeur_col_I = df.at[i, "Sous total"]
        if pd.notna(valeur_col_I) and not str(valeur_col_I).isnumeric():
            df.at[i, "Libellé"] = valeur_col_I
            df.at[i, "Sous total"] = "T"

    df.fillna("", inplace=True)
    return df

# Interface Streamlit
st.image("sabiana-logo.png", use_container_width=True)
st.title("Traitement des fichiers Excel - Easysel à Magenta")
st.subheader("Déposez votre fichier Excel ci-dessous :")

uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx"])

if uploaded_file:
    try:
        sheet = pd.read_excel(uploaded_file, sheet_name=0, header=None)
        st.info("Traitement des données en cours...")
        donnees = copier_donnees_easysel(sheet)

        if donnees is not None:
            resultat = modifier_tableau_easysel(donnees)
            st.success("Traitement terminé ! Téléchargez le fichier ci-dessous :")

            buffer = io.BytesIO()
            resultat.to_excel(buffer, index=False, engine="openpyxl")
            buffer.seek(0)

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=buffer,
                file_name="fichier_traite.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
    except Exception as e:
        st.error(f"Une erreur est survenue : {e}")
