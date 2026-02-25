import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import shutil

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Minatholy - ISM Manager", layout="wide")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
BASE_DIR = Path("data")
HISTORY_DIR = BASE_DIR / "history"
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
    final_buffer = BytesIO(csv_text.encode("utf-8-sig"))
    final_buffer.seek(0)
    return final_buffer

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

# -----------------------
# Sidebar
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Minatholy v3.1")
    selected_school = st.selectbox("École", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP", value=True)

st.title("🎓 ISM Data & History Manager")

tab_proc, tab_audit, tab_manual = st.tabs(["🚀 Traitement & Archivage", "🔍 Audit vs Historique", "🛠️ Manuel"])

# -----------------------
# TAB 1: PROCESSING & DOWNLOADS
# -----------------------
with tab_proc:
    uploaded_excel = st.file_uploader("Importer le fichier Excel Actuel", type=["xls", "xlsx"])
    
    if st.button("🚀 Lancer le traitement", type="primary"):
        if uploaded_excel:
            with st.status("Traitement et Archivage...", expanded=True) as status:
                # 1. Save to History
                history_path = HISTORY_DIR / f"{selected_school}_latest.xlsx"
                with open(history_path, "wb") as f:
                    f.write(uploaded_excel.getbuffer())
                
                # 2. Process Data
                df_raw = pd.read_excel(uploaded_excel)
                df = normalize_and_clean_df(df_raw)
                
                if df is not None:
                    # (Dummy logic for the example - replace with your real mapping logic)
                    google_df = df[["Member Email"]].copy()
                    google_df["Group"] = "test@ism.edu.sn"
                    
                    profiles_df = df.copy()
                    profiles_df["Password"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
                    
                    # 3. Generate Buffers
                    b_google = sanitize_csv_output(google_df)
                    b_profiles = sanitize_csv_output(profiles_df, is_profile=True)
                    
                    status.update(label="✅ Traitement Terminé & Archivé !", state="complete")
                    
                    # 4. Display Download Buttons
                    st.divider()
                    col1, col2 = st.columns(2)
                    col1.download_button("📧 Télécharger Google Groups", b_google, f"google_{selected_school}.csv", "text/csv")
                    col2.download_button("👤 Télécharger Profils Blackboard", b_profiles, f"profiles_{selected_school}.csv", "text/csv")
                else:
                    st.error("Format de colonnes invalide.")
        else:
            st.warning("Veuillez charger un fichier.")

# -----------------------
# TAB 2: AUDIT VS HISTORY
# -----------------------
with tab_audit:
    st.subheader("🔍 Comparaison avec l'archive")
    hist_file = HISTORY_DIR / f"{selected_school}_latest.xlsx"
    
    if hist_file.exists():
        st.info(f"Dernière archive trouvée pour {selected_school}: {datetime.fromtimestamp(hist_file.stat().st_mtime).strftime('%Y-%m-%d %H:%M')}")
        
        comp_file = st.file_uploader("Nouveau fichier pour comparaison", type=["xls", "xlsx"])
        if st.button("Comparer"):
            if comp_file:
                df_old = normalize_and_clean_df(pd.read_excel(hist_file))
                df_new = normalize_and_clean_df(pd.read_excel(comp_file))
                
                # Delta Logic
                added = df_new[~df_new["Full Name Key"].isin(df_old["Full Name Key"])]
                lost = df_old[~df_old["Full Name Key"].isin(df_new["Full Name Key"])]
                
                st.metric("Nouveaux Étudiants", len(added))
                st.dataframe(added, use_container_width=True)
    else:
        st.warning("Aucun historique pour cette école. Traitez un fichier d'abord.")
