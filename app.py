import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Minatholy - ISM Manager", layout="wide", initial_sidebar_state="expanded")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
DATA_DIR = Path("data")

# -----------------------
# Helper Functions
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

def remove_accents_and_cleanup(s: str):
    if not isinstance(s, str): return s
    accent_map = str.maketrans({"à":"a","â":"a","ä":"a","é":"e","è":"e","ê":"e","ë":"e","î":"i","ï":"i","ô":"o","ö":"o","û":"u","ü":"u","ÿ":"y"})
    return s.translate(accent_map).strip().upper()

def normalize_and_clean_df(df: pd.DataFrame):
    df.columns = [col.strip().capitalize() for col in df.columns]
    required = {"Classe", "E-mail", "Nom", "Prénom"}
    if not required.issubset(df.columns):
        return None
    df = df[["Classe", "E-mail", "Nom", "Prénom"]].copy()
    df.columns = ["Classroom Name", "Member Email", "Nom", "Prénom"]
    df["Classroom Name"] = df["Classroom Name"].fillna("").astype(str).str.replace(r"\s+","",regex=True).str.upper()
    df["Member Email"] = df["Member Email"].fillna("").astype(str).str.strip().str.lower()
    df["Nom"] = df["Nom"].apply(remove_accents_and_cleanup)
    df["Prénom"] = df["Prénom"].apply(remove_accents_and_cleanup)
    # Create a unique key for Name+Surname identity
    df["Full Name Key"] = df["Nom"] + "_" + df["Prénom"]
    return df

# -----------------------
# Sidebar & Header
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Minatholy v2.9")
    selected_school = st.selectbox("Sélectionner l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP", value=True)

st.title("🎓 ISM Data Processing")

main_tab, manual_tab, audit_tab = st.tabs(["🚀 Traitement Excel", "🛠️ Inscriptions", "🔍 Audit & Changements"])

# -----------------------
# TAB: AUDIT & CHANGEMENTS
# -----------------------
with audit_tab:
    st.subheader("📊 Analyse approfondie des mouvements")
    st.info("Cette analyse détecte les nouveaux venus, les départs, les changements de classe et les corrections d'emails.")
    
    c_up1, c_up2 = st.columns(2)
    with c_up1:
        old_file = st.file_uploader("📁 Ancien Fichier (Référence)", type=["xls", "xlsx"])
    with c_up2:
        new_file = st.file_uploader("📂 Nouveau Fichier (Actuel)", type=["xls", "xlsx"])

    if st.button("🔎 Lancer l'audit comparatif", use_container_width=True):
        if old_file and new_file:
            df_old = normalize_and_clean_df(pd.read_excel(old_file, dtype=str))
            df_new = normalize_and_clean_df(pd.read_excel(new_file, dtype=str))
            
            if df_old is not None and df_new is not None:
                # 1. NEW STUDENTS (Identity not in old file)
                added_ids = set(df_new["Full Name Key"]) - set(df_old["Full Name Key"])
                df_added = df_new[df_new["Full Name Key"].isin(added_ids)]
                
                # 2. DEPARTURES (Identity not in new file)
                lost_ids = set(df_old["Full Name Key"]) - set(df_new["Full Name Key"])
                df_lost = df_old[df_old["Full Name Key"].isin(lost_ids)]
                
                # 3. INTERNAL MOVEMENTS (Identity present in both)
                common_ids = set(df_old["Full Name Key"]) & set(df_new["Full Name Key"])
                
                # Merge on Identity Key
                merged = pd.merge(
                    df_old[df_old["Full Name Key"].isin(common_ids)][["Full Name Key", "Nom", "Prénom", "Classroom Name", "Member Email"]],
                    df_new[df_new["Full Name Key"].isin(common_ids)][["Full Name Key", "Classroom Name", "Member Email"]],
                    on="Full Name Key",
                    suffixes=('_Old', '_New')
                )
                
                # Flag Class Changes
                df_moved = merged[merged["Classroom Name_Old"] != merged["Classroom Name_New"]]
                
                # Flag Email Changes (Typo corrections)
                df_email_fix = merged[merged["Member Email_Old"] != merged["Member Email_New"]]
                
                # Display Results
                col_m1, col_m2, col_m3, col_m4 = st.columns(4)
                col_m1.metric("Nouveaux", len(df_added))
                col_m2.metric("Départs", len(df_lost))
                col_m3.metric("Transferts", len(df_moved))
                col_m4.metric("Corrections Email", len(df_email_fix))
                
                st.divider()
                
                res1, res2, res3, res4 = st.tabs(["🆕 Nouveaux", "🏃 Départs", "🔄 Transferts de Classe", "📧 Corrections Email"])
                
                with res1:
                    st.dataframe(df_added[["Classroom Name", "Nom", "Prénom", "Member Email"]], use_container_width=True)
                
                with res2:
                    st.dataframe(df_lost[["Classroom Name", "Nom", "Prénom", "Member Email"]], use_container_width=True)

                with res3:
                    st.write("Étudiants ayant changé de classe :")
                    st.dataframe(df_moved[["Nom", "Prénom", "Classroom Name_Old", "Classroom Name_New", "Member Email_New"]], use_container_width=True)
                
                with res4:
                    st.write("Identités dont l'email a été modifié (corrections de typos) :")
                    st.dataframe(df_email_fix[["Nom", "Prénom", "Member Email_Old", "Member Email_New"]], use_container_width=True)
                    st.warning("⚠️ Attention : Si un email a changé, vous devez probablement supprimer l'ancien compte et créer le nouveau sur Blackboard.")

        else:
            st.warning("Chargez deux fichiers pour comparer.")

# --- (Other tabs logic Traitement/Inscriptions remains unchanged) ---
