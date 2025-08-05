import streamlit as st
import pandas as pd
import base64

FICHIER_EXCEL = "omp.xlsm"

# Fonction de coloration des pourcentages
def colorer_pourcentages(val, feuille, colonne):
    if isinstance(val, str) and "%" in val:
        try:
            val_num = float(val.replace("%", "").replace(",", ".").strip())
            # Spécifique à SPLITTER CHANGE
            if feuille == "SPLITTER CHANGE":
                return "background-color: red; color: white" if val_num > 10 else "background-color: green; color: white"
            # Spécifique à DECONNEXION CSP
            if colonne == "DÉCONNEXION CSP":
                return "background-color: red; color: white" if val_num > 25 else "background-color: green; color: white"
            # Par défaut
            return "background-color: red; color: white" if val_num < 85 else "background-color: green; color: white"
        except:
            return ""
    return ""

# CSS personnalisé
st.markdown("""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans&family=Roboto&display=swap');

        html, body, [class*="css"] {
            font-family: 'Roboto', 'Open Sans', sans-serif;
            background-color: #1e1e1e;
            color: white;
        }

        .title-container {
            text-align: left;
            margin-bottom: 20px;
        }

        .title-container h1 {
            margin: 0;
            font-weight: 900;
            font-size: 2.4rem;
            color: black;
        }

        .title-container h2 {
            margin: 0;
            font-weight: bold;
            font-size: 1.6rem;
            color: black;
            text-transform: uppercase;
        }

        @media (prefers-color-scheme: dark) {
            .title-container h1,
            .title-container h2 {
                color: white;
            }
        }

        .stDataFrame thead {
            background-color: #333333 !important;
            color: white !important;
        }
    </style>
""", unsafe_allow_html=True)

# Affichage des titres
st.markdown("""
    <div class="title-container">
        <h1>OMP TEAM</h1>
        <h2>COACHING : RECHERCHE PAR TECHNICIEN</h2>
    </div>
""", unsafe_allow_html=True)

@st.cache_data
def charger_toutes_les_feuilles():
    xls = pd.ExcelFile(FICHIER_EXCEL, engine="openpyxl")
    feuilles = {}
    for nom in xls.sheet_names:
        try:
            df = xls.parse(nom, header=4)
            feuilles[nom] = df
        except Exception as e:
            feuilles[nom] = pd.DataFrame({"Erreur": [str(e)]})
    return feuilles

feuilles = charger_toutes_les_feuilles()

# Colonnes à formater en pourcentage
colonnes_pourcent_par_feuille = {
    "TSDC APT & AT (LIGNE ACTIVE)": [
        "USAGE PAIRE ACTIVE", "TEST RÉUSSIS PAIRE ACTIVE",
        "RÉSULTATS APRÈS COACHING APT", "USAGE SYNCH AUTO",
        "TEST RÉUSSI SYNCH AUTO", "RÉSULTAT APRÈS COACHING AT"
    ],
    "SPLITTER CHANGE": [
        "RÉSULTATS INITIALE", "RÉSULTAT APRÈS COACHING"
    ],
    "COACHING OX1": [
        "RÉSULTAT INITIALE (USAGE)", "RÉSULTAT INITIALE (U.R)"
    ],
}

# Interface de recherche
critere = st.radio("Rechercher par :", ["Numéro d'employé", "Nom"])
if critere == "Numéro d'employé":
    recherche = st.text_input("Entrez le numéro d'employé")
    filtre_colonne = "TECH"
else:
    recherche = st.text_input("Entrez le nom")
    filtre_colonne = "NOM DU TECH"

if recherche:
    trouve = False
    for nom_feuille, df in feuilles.items():
        if filtre_colonne in df.columns:
            resultats = df[df[filtre_colonne].astype(str).str.contains(recherche, case=False, na=False)]
            if not resultats.empty:
                # Nettoyage
                resultats = resultats.loc[:, ~resultats.columns.str.contains("Unnamed", na=False)]
                resultats = resultats.dropna(axis=1, how="all")
                if "index" in resultats.columns:
                    resultats = resultats.drop(columns=["index"])
                if resultats.columns[0] == resultats.columns[0] and resultats.columns[0] == 0:
                    resultats = resultats.iloc[:, 1:]

                # Formatage pourcentages
                colonnes_pourcent = colonnes_pourcent_par_feuille.get(nom_feuille, [])
                for col in colonnes_pourcent:
                    if col in resultats.columns:
                        if nom_feuille == "COACHING OX1":
                            resultats[col] = resultats[col].apply(lambda x: f"{x:.0f} %" if pd.notnull(x) else "")
                        else:
                            resultats[col] = resultats[col].apply(lambda x: f"{x:.0%}" if pd.notnull(x) else "")

                # Traitement spécifique sans décimales pour certaines colonnes
                if nom_feuille == "SPLITTER CHANGE":
                    for col in ["NOMBRE DE TÂCHES", "NOMBRE DE CHANGEMENT"]:
                        if col in resultats.columns:
                            resultats[col] = resultats[col].apply(lambda x: f"{int(x)}" if pd.notnull(x) else "")

                # Supprimer décimales inutiles sauf pour colonnes %
                for col in resultats.select_dtypes(include='number').columns:
                    if col not in colonnes_pourcent:
                        resultats[col] = resultats[col].apply(lambda x: int(x) if pd.notnull(x) and x == int(x) else round(x, 2))

                # Format des dates
                for col in resultats.select_dtypes(include='datetime').columns:
                    resultats[col] = resultats[col].dt.strftime('%Y-%m-%d')

                resultats = resultats.reset_index(drop=True)
                trouve = True

                # Supprimer "Feuille :" dans le titre
                nom_affiche = nom_feuille.replace("Feuille :", "").strip()
                st.subheader(nom_affiche)

                # Appliquer les couleurs personnalisées
                def appliquer_couleurs(val, colonne):
                    return colorer_pourcentages(val, nom_feuille, colonne)

                styles = resultats.style
                for col in resultats.columns:
                    styles = styles.applymap(lambda v: appliquer_couleurs(v, col), subset=[col])

                st.dataframe(styles, hide_index=True, use_container_width=True)

    if not trouve:
        st.warning("Aucun résultat trouvé dans aucune feuille.")
