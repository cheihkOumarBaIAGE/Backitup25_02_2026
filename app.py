import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import unicodedata

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Minatholy - ISM Manager", layout="wide", initial_sidebar_state="expanded")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
DATA_DIR = Path("data")

# -----------------------
# Helper Functions
# -----------------------
def remove_accents_and_cleanup(s: str) -> str:
    """Removes all accents and normalizes text cleanly using standard Unicode decomposition."""
    if not isinstance(s, str): 
        return s
    nfkd_form = unicodedata.normalize('NFKD', s)
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

def sanitize_csv_output(df: pd.DataFrame, header=True) -> BytesIO:
    """Exports DataFrame cleanly into UTF-8-SIG CSV Buffer."""
    b = BytesIO()
    df.to_csv(b, index=False, header=header, encoding="utf-8-sig")
    b.seek(0)
    return b

def read_emails_txt(path: Path) -> list:
    if not path.exists(): 
        return []
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def normalize_and_clean_df(df: pd.DataFrame) -> pd.DataFrame:
    df.columns = [col.strip().capitalize() for col in df.columns]
    df = df[["Classe", "E-mail", "Nom", "Prénom"]].copy()
    df.columns = ["Classroom Name", "Member Email", "Nom", "Prénom"]
    
    df["Classroom Name"] = df["Classroom Name"].fillna("").astype(str).str.replace(r"\s+", "", regex=True).str.upper()
    df["Member Email"] = df["Member Email"].fillna("").astype(str).str.strip().str.replace(r"\s+", "", regex=True)
    df["Nom"] = df["Nom"].fillna("").astype(str).str.strip().apply(remove_accents_and_cleanup)
    df["Prénom"] = df["Prénom"].fillna("").astype(str).str.strip().apply(remove_accents_and_cleanup)
    return df

# -----------------------
# Sidebar & Header
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Urgence IT Settings")
    selected_school = st.selectbox("Sélectionner l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP (Automatique)", value=True)
    st.divider()
    st.caption("v3.2 - Interactive Classroom Course Assignment")

st.title("Urgence IT BASH")

main_tab, manual_tab = st.tabs(["BASH", "MAJ de cours"])

# -----------------------
# TAB 1: AUTOMATIC TREATMENT
# -----------------------
with main_tab:
    with st.container(border=True):
        uploaded_excel = st.file_uploader("📥 Importer le fichier Excel des inscriptions (Étudiants)", type=["xls", "xlsx"])
        run = st.button("🚀 Lancer le traitement", type="primary", use_container_width=True)

    # Initial processing: Parse Excel & load base mapping
    if run:
        if not uploaded_excel:
            st.error("❌ Aucun fichier d'inscriptions détecté.")
        else:
            with st.status("Traitement du fichier Excel...", expanded=True) as status:
                school_
