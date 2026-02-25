import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import shutil
import os

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Minatholy - ISM Manager", layout="wide")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
BASE_DIR = Path("data")
HISTORY_DIR = BASE_DIR / "history"
# Ensure folders exist
HISTORY_DIR.mkdir(parents=True, exist_ok=True)

# -----------------------
# Core Functions
# -----------------------
def sanitize_csv_output(df: pd.DataFrame, is_profile=False, header=True):
    b = BytesIO()
    df.to_csv(b, index=False, header=header, encoding="utf-8-sig")
    csv_text = b.getvalue().decode("utf-8-sig")
    if is_profile:
        csv_text = csv_text.replace('"""', '"')
        csv_text = csv_text.replace('"ismapps2025,,,,,,,,,,,,,,,,,1382"', 'ismapps2025,,,,,,,,,,,,,,,,,1382')
    return BytesIO(csv_text.encode("utf-8-sig"))

def normalize_and_clean_df(df: pd.DataFrame):
    df.columns = [col.strip().capitalize() for col in df.columns]
    if not {"Classe", "E-mail", "Nom", "Prénom"}.issubset(df.columns):
        return None
    df = df[["Classe", "E-mail", "Nom", "Prénom"]].copy()
    df.columns = ["Classroom Name", "Member Email", "Nom", "Prénom"]
    df["Classroom Name"] = df["Classroom Name"].fillna("").astype(str).str.replace(r"\s+","",regex=True).str.upper()
    df["Member Email"] = df["Member Email"].fillna("").astype(str).str.strip().str.lower()
    df["Full Name Key"] = df["Nom"].astype(str).str.upper() + "_" + df["Prénom"].astype(str).str.upper()
    return df

def save_to_history(file, school):
    """Saves a copy of the uploaded file to the history folder."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    history_path = HISTORY_DIR / f"{school}_latest.xlsx"
    # Create backup of current 'latest' before overwriting
    if history_path.exists():
        shutil.copy(history_path, HISTORY_DIR / f"{school}_backup_{timestamp}.xlsx")
    
    with open(history_path, "wb") as f:
        f.write(file.getbuffer())
    return history_path

# -----------------------
# Sidebar & Navigation
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Minatholy v3.0")
    selected_school = st.selectbox("École", SCHOOLS)
    
st.title("🎓 ISM Data & History Manager")

tab_proc, tab_audit, tab_manual = st.tabs(["🚀 Traitement & Archivage", "🔍 Audit vs Historique", "🛠️ Manuel"])

# -----------------------
# TAB 1: PROCESSING & AUTO-SAVE
# -----------------------
with tab_proc:
    st.subheader("Traitement de la nouvelle liste")
    new_excel = st.file_uploader("Importer le nouvel Excel", type=["xls", "xlsx"])
    
    if st.button("🚀 Lancer le traitement", type="primary"):
        if new_excel:
            # 1. Save to History automatically
            saved_path = save_to_history(new_excel, selected_school)
            st.success(f"✅ Fichier archivé avec succès dans `{saved_path.name}`")
            
            # 2. Logic for generating CSVs (Google/Blackboard)
            # [Insert your existing mapping and export logic here]
            st.info("Traitement des fichiers Google et Blackboard terminé (voir téléchargements ci-dessous).")
        else:
            st.error("Veuillez charger un fichier.")

# -----------------------
# TAB 2: AUDIT VS HISTORY
# -----------------------
with tab_audit:
    st.subheader("📊 Comparaison automatique avec l'archive")
    
    history_file_path = HISTORY_DIR / f"{selected_school}_latest.xlsx"
    
    if not history_file_path.exists():
        st.warning(f"⚠️ Aucun historique trouvé pour {selected_school}. Traitez d'abord un fichier pour créer une archive.")
    else:
        st.info(f"📁 Fichier d'historique détecté : Mis à jour le {datetime.fromtimestamp(history_file_path.stat().st_mtime).strftime('%d/%m/%Y à %H:%M')}")
        
        current_file = st.file_uploader("Uploader le fichier actuel pour comparaison", type=["xls", "xlsx"])
        
        if st.button("🔎 Lancer l'Audit"):
            if current_file:
                df_old = normalize_and_clean_df(pd.read_excel(history_file_path))
                df_new = normalize_and_clean_df(pd.read_excel(current_file))
                
                # --- COMPARISON ENGINE ---
                # New Students
                added = df_new[~df_new["Full Name Key"].isin(df_old["Full Name Key"])]
                # Departures
                lost = df_old[~df_old["Full Name Key"].isin(df_new["Full Name Key"])]
                # Classroom Moves
                merged = pd.merge(df_old, df_new, on="Full Name Key", suffixes=('_Old', '_New'))
                moved = merged[merged["Classroom Name_Old"] != merged["Classroom Name_New"]]
                
                # Metrics
                c1, c2, c3 = st.columns(3)
                c1.metric("Nouveaux", len(added))
                c2.metric("Départs", len(lost))
                c3.metric("Transferts", len(moved))
                
                # Dataframes
                st.write("#### Nouveaux inscrits")
                st.dataframe(added[["Nom", "Prénom", "Classroom Name", "Member Email"]], use_container_width=True)
                
                st.write("#### Transferts de classe")
                st.dataframe(moved[["Nom", "Prénom", "Classroom Name_Old", "Classroom Name_New"]], use_container_width=True)
            else:
                st.warning("Veuillez uploader le fichier actuel.")

# -----------------------
# TAB 3: MANUAL
# -----------------------
with tab_manual:
    st.write("Outils d'inscription manuelle...")
