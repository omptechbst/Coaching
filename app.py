import streamlit as st
import pandas as pd
import base64

FICHIER_EXCEL = "omp.xlsm"
LOGO_PATH = "Bell_White_large_transparent_cropped_to_fit_2524870a3c.png"

# Lire et encoder le logo en base64
def get_base64_image(image_path):
    with open(image_path, "rb") as img_file:
        return base64.b64encode(img_file.read()).decode()

logo_b64 = get_base64_image(LOGO_PATH)

# CSS personnalisé
st.markdown(f"""
    <style>
        @import url('https://fonts.googleapis.com/css2?family=Open+Sans&family=Roboto&display=swap');

        html, body, [class*="css"] {{
            font-family: 'Roboto', 'Open Sans', sans-serif;
            background-color: #1e1e1e;
            color: white;
        }}

        .logo-container {{
            display: flex;
            align-items: center;
            margin-bottom: 20px;
        }}

        .logo-container img {{
            height: 45px;
            margin-right: 15px;
        }}

        .logo-container h1 {{
            margin: 0;
            font-weight: bold;
            font-size: 1.8rem;
            color: white;
            text-transform: uppercase;
        }}

        .stDataFrame thead {{
            background-color: #333333 !important;
            color: white !important;
        }}
    </style>
""", unsafe_allow_html=True)

# Affichage du logo + nouveau titre
st.markdown(f"""
    <div class="logo-container">
        <img src="data:image/png;base64,{logo_b64}" alt="Bell logo">
        <h1>COACHING : RECHERCHE PAR TECHNICIEN</h1>
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

                # Formatage pourcentages
                colonnes_pourcent = colonnes_pourcent_par_feuille.get(nom_feuille, [])
                for col in colonnes_pourcent:
                    if col in resultats.columns:
                        if nom_feuille == "COACHING OX1":
                            resultats[col] = resultats[col].apply(lambda x: f"{x:.0f} %" if pd.notnull(x) else "")
                        else:
                            resultats[col] = resultats[col].apply(lambda x: f"{x:.0%}" if pd.notnull(x) else "")

                # Arrondi des autres colonnes numériques
                for col in resultats.select_dtypes(include='number').columns:
                    if col not in colonnes_pourcent:
                        resultats[col] = resultats[col].round(2)

                # Format des dates
                for col in resultats.select_dtypes(include='datetime').columns:
                    resultats[col] = resultats[col].dt.strftime('%Y-%m-%d')

                resultats = resultats.reset_index(drop=True)
                trouve = True
                st.subheader(f"Feuille : {nom_feuille}")
                st.dataframe(resultats, hide_index=True)

    if not trouve:
        st.warning("Aucun résultat trouvé dans aucune feuille.")
