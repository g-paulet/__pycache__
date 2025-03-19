"""Module principal qui détecte la nature du fichier et appelle les fonctions en conséquence"""

import io
import traceback
import streamlit as st
import pandas as pd

from panneau import traiter_pulsar, traiter_ds18, modifier_tableau_panneau
from easysel import copier_donnees_easysel, modifier_tableau_easysel
from rapidaero import copier_tableaux_rapid_aero, mise_en_forme_rapid_aero


def identifier_fichier(nom_fichier):

    """Identifie le type de fichier en fonction de son nom et exécute les fonctions appropriées."""
    if "PANNEAU" in nom_fichier.upper():
        st.write("Fichier détecté : PANNEAU")


        # Module issu de panneau.py pour le traitement de données
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

            # Mise en forme du dataframe
            df_export = modifier_tableau_panneau(df_export)

            # Génération du fichier Excel de sortie
            output = io.BytesIO()

            with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
                df_export.to_excel(writer, sheet_name="Données Exportées", index=False)

                # Mise en forme des lignes de titre en Gras + centrer colonne A
                # Récupérer le workbook et le worksheet
                workbook = writer.book
                worksheet = writer.sheets["Données Exportées"]

                # Définir un format en gras
                bold_format = workbook.add_format({"bold": True})

                # Trouver les lignes où "Sous total" contient "T"
                # start=1 car Excel commence à la ligne 2
                for row_num, value in enumerate(df_export["Sous total"], start=1):
                    if value == "T":
                        # Mettre la ligne en gras
                        worksheet.set_row(row_num, cell_format=bold_format)

                # Définir un format centré pour la colonne "Code"
                # Alignement horizontal et vertical
                center_format = workbook.add_format({"align": "center", "valign": "vcenter"})

                # Appliquer l'alignement centré sur toute la colonne "Code"
                #col_index_code = df_export.columns.get_loc("Code")
                # Trouver l'index de la colonne "Code"
                # Centrer le texte dans la colonne A (Code) et ajuster la largeur à 15
                worksheet.set_column(0, 0, 15, center_format)

                # Ajustement automatique des largeurs de colonnes
                for col_num, col_name in enumerate(df_export.columns):
                    # Trouver la longueur maximale
                    max_len = df_export[col_name].astype(str).map(len).max()
                    worksheet.set_column(col_num, col_num, max_len + 2)  # Ajouter un peu d'espace

                # Sauvegarder
                writer._save()

            st.success("Traitement terminé ! Téléchargez votre fichier ci-dessous.")

            # Bouton de téléchargement
            st.download_button(
                label="Télécharger le fichier Excel",
                data=output.getvalue(),
                file_name="export_donnees.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    elif "Offerta" in nom_fichier:
        st.write("Fichier détecté : Easysel")

        # Module issu de Sabiana.py pour le traitement de données
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

    elif "Rapid'Aero" in nom_fichier:
        st.write("Fichier détecté : Rapid'Aero")

        # Module issu de rapidaero.py pour le traitement de données
        if uploaded_file:
            try:
                st.info("Début du traitement du fichier...")

                # Charger le fichier Excel pour récupérer les noms des feuilles
                excel_file = pd.ExcelFile(uploaded_file)

                # Afficher les feuilles disponibles
                # st.write("Feuilles disponibles :", excel_file.sheet_names)

                # Étape 1 : Copier les données
                donnees = copier_tableaux_rapid_aero(excel_file)

                if donnees is None or donnees.empty:
                    st.error("Aucune donnée extraite. Vérifiez le contenu du fichier.")
                else:
                    #st.write("Colonnes après extraction :", donnees.columns) #Debug

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
    else:
        st.error("Le classeur sélectionné ne semble pas correspondre à un fichier Easysel," \
            "Rapid'Aero ou Panneau. Merci de sélectionner un fichier valide.")

#Interface Streamlit
st.image("sabiana-logo.png", use_container_width=True)
st.title("Clic and Selec pour MAGENTA")
st.subheader("Déposez votre fichier Excel ci-dessous \
    (issu des outils de sélection : Easysel, Rapid'Aéro ou Panneau)")

uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    identifier_fichier(uploaded_file.name)
