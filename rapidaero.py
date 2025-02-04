"""Module dev Rapidaero"""

import io
import traceback
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

            df = pd.read_excel(excel_file, sheet_name=sheet_name, header=None)

            # Vérifier que le fichier contient au moins 18 colonnes (jusqu'à la colonne R)
            if df.shape[1] > 17:
                # Trouver la dernière ligne utilisée dans la colonne P (colonne index 15)
                dern_ligne = df.iloc[:, 15].last_valid_index()

                if dern_ligne and dern_ligne >= 1:  # Vérifier qu'il y a bien des données
                    print("Dernière ligne remplie en colonne P :", dern_ligne)

                    # Extraction des colonnes P, Q, R à partir de la ligne 2
                   # extrait = df.iloc[1:dern_ligne + 1, [15, 16, 17]].copy()
                    #extrait.columns = ["P", "Q", "R"]

                    # Ajouter le titre de la feuille
                   # extrait.insert(0, "Titre", f"Aérotherme - {sheet_name}")

                    # Ajouter au résultat final
                    #resultats.append(extrait)
                # Créer une liste vide pour stocker les données formatées
                extrait_liste = []

                #  **Ajouter le titre dans la première colonne (ligne actuelle)**
                extrait_liste.append([f"Aérotherme - {sheet_name}"] + [""] * 2)  # Colonnes P, Q, R vides

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
    print("Colonnes présentes dans df:", df.columns)  # Debugging
    print(df) # Debugging

    #Ajout des colonnes vides pour matcher la mise en forme Dollibar
    #df.drop(df.columns[0], axis=1, inplace=True)
    df.insert(3, "NouvelleCol1", "")
    df.insert(3, "NouvelleCol2", "")
    df.insert(3, "NouvelleCol3", "")
    df.insert(3, "NouvelleCol4", "")
    df.insert(3, "NouvelleCol5", "")
    df.insert(3, "NouvelleCol6", "")

    # Renommer correctement les colonnes + swap A et B
    titres = ["Code", "Libellé", "Qté"] + [f"Col_{i+4}" for i in range(df.shape[1] - 4)] + ["Sous total"]
    df.columns = titres
    df[["Code", "Libellé", "Qté"]] = df[["Libellé","Code", "Qté"]]

    print("Colonnes dans le df :", titres)

    # Ajouter "T" dans "sous-total" si "Aérotherme" est trouvé dans "Libellé"
    df.loc[df["Libellé"].str.contains("Aérotherme", na=False), "Sous total"] = "T"

    # Supprimer les lignes où les colonnes A et I sont vides
    df = df[~(df["Code"].isna() | (df["Code"] == "")) |
            ~(df["Sous total"].isna() | (df["Sous total"] == ""))]

    # Remplir les cellules vides avec des chaînes vides pour éviter les NaN
    # df.fillna("", inplace=True)

    return df

# Interface Streamlit
uploaded_file = st.file_uploader("Choisissez un fichier Excel", type=["xlsx", "xlsm"])

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
