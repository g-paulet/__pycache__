import io
import pandas as pd
import streamlit as st
import traceback

def copier_tableaux_rapid_aero(xls):
    """ Copie les données des feuilles se terminant par '°C' """
    liste_feuilles = xls.sheet_names  # Récupérer la liste des feuilles
    resultats = []

    for sheet_name in liste_feuilles:
        if sheet_name.endswith("°C"):
            df = pd.read_excel(uploaded_file, sheet_name=0, header=0, engine="openpyxl")  # Lire avec un header
            st.write(f"Traitement de la feuille : {sheet_name}")  # Debugging

            # Récupération des colonnes P, Q, R (15, 16, 17 en index)
            if df.shape[1] > 17:  # Vérifier qu'on a assez de colonnes
                extrait = df.iloc[1:, [15, 16, 17]].copy()
                extrait.columns = ["P", "Q", "R"]

                # Ajouter le titre de la feuille
                extrait.insert(0, "Titre", f"Aérotherme - {sheet_name}")

                resultats.append(extrait)

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
    print("Colonnes présentes dans df:", df.columns)  # Debugging
    st.write("Colonnes détectées :", df.columns.tolist())

    df.insert(2, "NouvelleCol", df[0])
    df.drop(columns=[0], inplace=True)
    df[8] = df[1].apply(lambda x: "T" if isinstance(x, str) and x.startswith("Aérotherme - ") else "")
    columns = ["Code", "Libellé", "Qté", "", "", "", "", "", "Sous total"]
    df.columns = columns
    df.loc[-1] = columns  # Ajouter en première ligne
    df.index = df.index + 1  # Décalage des index
    df = df.sort_index()
    df.dropna(subset=["Code", "Sous total"], how="all", inplace=True)
    return df

# Interface Streamlit
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xlsm"])

if uploaded_file:
    try:
        st.info("Début du traitement du fichier...")

        # Charger le fichier Excel en tant qu'objet ExcelFile
        xls = pd.ExcelFile(uploaded_file, engine="openpyxl")
        st.write("Feuilles disponibles :", xls.sheet_names)  # Afficher les noms des onglets

        # Étape 1 : Copier les données
        donnees = copier_tableaux_rapid_aero(xls)  # Passer l'objet ExcelFile et non un DataFrame

        if donnees is None or donnees.empty:
            st.error("Aucune donnée extraite. Vérifiez le contenu du fichier.")
        else:
            st.write("Colonnes après extraction :", donnees.columns)

            # Étape 2 : Mise en forme
            resultat = mise_en_forme_rapid_aero(donnees)

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
        st.error(f"Une erreur est survenue : {str(e)}")
        st.text(traceback.format_exc())  # Affiche le détail de l'erreur