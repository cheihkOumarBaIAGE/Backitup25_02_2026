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
# Mappings (Example: MADIBA)
# -----------------------
SCHOOL_MAPPINGS = {
    "MADIBA": {
        "JMI1": "jmi1-2526@ism.edu.sn",
        "LCM-1": "lcm1-2526@ism.edu.sn",
        "SPRI1-A": "spri1a-2526@ism.edu.sn",
        "MASTER1SCIENCEPOLITIQUEETRELATIONSINTERNATIONALES": "mba1spri-2526@ism.edu.sn",
    },
    # Add other schools here as needed...
}

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
    buf = BytesIO(csv_text.encode("utf-8-sig"))
    buf.seek(0)
    return buf

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
# UI Logic
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Minatholy v3.2")
    selected_school = st.selectbox("École cible", SCHOOLS)
    zip_opt = st.checkbox("Inclure Archive ZIP", value=True)

tab_proc, tab_audit, tab_manual = st.tabs(["🚀 Traitement & Exports", "🔍 Audit Historique", "🛠️ Outils"])

with tab_proc:
    uploaded_excel = st.file_uploader("📥 Importer l'Excel du mois", type=["xls", "xlsx"])
    
    if st.button("🚀 Lancer le traitement complet", type="primary"):
        if uploaded_excel:
            with st.status("Traitement des données...", expanded=True) as status:
                # 1. Archive current file
                hist_path = HISTORY_DIR / f"{selected_school}_latest.xlsx"
                with open(hist_path, "wb") as f:
                    f.write(uploaded_excel.getbuffer())

                # 2. Process logic
                df_raw = pd.read_excel(uploaded_excel)
                df = normalize_and_clean_df(df_raw)
                
                if df is not None:
                    mapping = SCHOOL_MAPPINGS.get(selected_school, {})
                    
                    # --- OUTPUT 1: Google Diffusion ---
                    valid_df = df[df["Member Email"].str.contains("@", na=False)].copy()
                    valid_df["Group Email"] = valid_df["Classroom Name"].map(mapping)
                    google_list = valid_df.dropna(subset=["Group Email"])[["Group Email", "Member Email"]].copy()
                    google_list.columns = ["Group Email [Required]", "Member Email"]
                    google_list["Member Type"], google_list["Member Role"] = "USER", "MEMBER"
                    
                    # --- OUTPUT 2: Admins ---
                    admin_df = pd.DataFrame([{"Group Email [Required]": g, "Member Email": "admin@ism.edu.sn", "Member Type": "USER", "Member Role": "MANAGER"} for g in set(mapping.values())])

                    # --- OUTPUT 3: Blackboard Profiles ---
                    profiles_df = valid_df.copy()
                    profiles_df["Password"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
                    profiles_export = profiles_df[["Member Email", "Nom", "Prénom", "Member Email", "Password"]]
                    profiles_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Mot de passe"]

                    # --- OUTPUT 4: Course Enrollment ---
                    course_df = valid_df[["Classroom Name", "Member Email"]].copy()
                    # (Simplified for example: logic would use your .txt mapping files here)

                    # --- OUTPUT 5: Report ---
                    report_txt = f"Rapport {selected_school} - {datetime.now()}\nTotal: {len(df)}\nMapped: {len(google_list)}"
                    
                    # Generate Buffers
                    b1 = sanitize_csv_output(google_list)
                    b2 = sanitize_csv_output(admin_df)
                    b3 = sanitize_csv_output(profiles_export, is_profile=True)
                    b4 = sanitize_csv_output(course_df, header=False)
                    b5 = BytesIO(report_txt.encode("utf-8"))

                    status.update(label="✅ Traitement terminé !", state="complete")

                    # --- DISPLAY DOWNLOAD BUTTONS ---
                    st.success("Cliquez sur les boutons ci-dessous pour télécharger vos fichiers :")
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.download_button("📧 1. Liste de Diffusion", b1, f"diffusion_{selected_school}.csv")
                        st.download_button("🔑 2. Ajouter Admins", b2, f"admins_{selected_school}.csv")
                        st.download_button("📄 5. Rapport (TXT)", b5, f"rapport_{selected_school}.txt")
                    
                    with col2:
                        st.download_button("👤 3. Profils Blackboard", b3, f"profils_{selected_school}.csv")
                        st.download_button("📚 4. Inscriptions Cours", b4, f"cours_{selected_school}.csv")
                        
                        if zip_opt:
                            z_buf = BytesIO()
                            with zipfile.ZipFile(z_buf, "w") as zf:
                                zf.writestr("diffusion.csv", b1.getvalue())
                                zf.writestr("admins.csv", b2.getvalue())
                                zf.writestr("profils.csv", b3.getvalue())
                                zf.writestr("cours.csv", b4.getvalue())
                                zf.writestr("rapport.txt", report_txt)
                            st.download_button("📦 Télécharger tout (ZIP)", z_buf.getvalue(), f"export_{selected_school}.zip", type="primary")
                else:
                    st.error("Colonnes manquantes dans le fichier Excel.")
        else:
            st.warning("Veuillez d'abord sélectionner un fichier Excel.")

# (Audit and Manual tabs follow same logic as v3.1...)
