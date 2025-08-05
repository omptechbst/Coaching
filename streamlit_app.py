
import streamlit as st
import pandas as pd

FICHIER_EXCEL = "omp.xlsm"

# Fonction de coloration personnalisée
def colorer_pourcentages(val, feuille, colonne):
    try:
        val_num = float(str(val).replace("%", "").replace(",", ".").strip())
        if feuille == "SPLITTER CHANGE":
            return "background-color: red; color: white" if val_num > 10 else "background-color: green; color: white"
        elif colonne == "DÉCONNEXION CSP":
            return "background-color: red; color: white" if val_num > 25 else "background-color: green; color: white"
        elif val_num < 85:
            return "background-color: red; color: white"
        else:
            return "background-color: green; color: white"
    except:
        return ""

# Titre principal et sous-titre
st.markdown("## OMP TEAM")
st.markdown("# COACHING : RECHERCHE PAR TECHNICIEN")

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

colonnes_pourcent_par_feuille = {
    "TSDC APT & AT (LIGNE ACTIVE)": [
        "USAGE PAIRE ACTIVE",
        "TEST RÉUSSIS PAIRE ACTIVE",
        "RÉSULTATS APRÈS COACHING APT",
        "USAGE SYNCH AUTO",
        "TEST RÉUSSI SYNCH AUTO",
        "RÉSULTAT APRÈS COACHING AT"
    ],
    "SPLITTER CHANGE": [
        "RÉSULTATS INITIALE",
        "RÉSULTAT APRÈS COACHING"
    ],
    "COACHING OX1": [
        "RÉSULTAT INITIALE (USAGE)",
        "RÉSULTAT INITIALE (U.R)"
    ]
}

critere = st.radio("Rechercher par :", ["Numéro d'employé", "Nom"])
recherche = st.text_input("Entrez la valeur à rechercher")
filtre_colonne = "TECH" if critere == "Numéro d'employé" else "NOM DU TECH"

if recherche:
    trouve = False
    for nom_feuille, df in feuilles.items():
        if filtre_colonne in df.columns:
            resultats = df[df[filtre_colonne].astype(str).str.contains(recherche, case=False, na=False)]
            if not resultats.empty:
                resultats = resultats.loc[:, ~resultats.columns.str.contains("Unnamed", na=False)]
                resultats = resultats.dropna(axis=1, how="all")
                if "index" in resultats.columns:
                    resultats = resultats.drop(columns=["index"])
                if "TECH" in resultats.columns:
                    resultats["TECH"] = resultats["TECH"].astype(str)
                if nom_feuille == "SPLITTER CHANGE":
                    for col in ["NOMBRE DE TACHES", "NOMBRE DE CHANGEMENT"]:
                        if col in resultats.columns:
                            resultats[col] = resultats[col].apply(lambda x: f"{int(x)}" if pd.notnull(x) else "")
                for col in resultats.select_dtypes(include="datetime").columns:
                    resultats[col] = resultats[col].dt.strftime("%Y-%m-%d")
                colonnes_pourcent = colonnes_pourcent_par_feuille.get(nom_feuille, [])
                for col in colonnes_pourcent + ["DÉCONNEXION CSP"]:
                    if col in resultats.columns:
                        resultats[col] = resultats[col].apply(lambda x: f"{x:.0f} %" if pd.notnull(x) else "")
                resultats = resultats.reset_index(drop=True)
                st.subheader(nom_feuille.strip())
                st.write(resultats.style.applymap(lambda val: colorer_pourcentages(val, nom_feuille, col), subset=resultats.columns))
                trouve = True
    if not trouve:
        st.warning("Aucun résultat trouvé dans aucune feuille.")
