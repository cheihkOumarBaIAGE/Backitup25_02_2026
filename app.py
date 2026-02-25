import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile
import csv

# -----------------------
# Page configuration
# -----------------------
st.set_page_config(page_title="Minatholy - ISM Manager", layout="wide", initial_sidebar_state="expanded")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
DATA_DIR = Path("data")

# -----------------------
# Default School Mappings
# -----------------------
SCHOOL_MAPPINGS = {
    "DROIT": {
        "L1-JURISTED'ENTREPRISE": "lda1c-2025-2026@ism.edu.sn",
        "LDA1-A": "lda1a-2025-2026@ism.edu.sn",
        "LDA2-A": "lda2a-2025-2026@ism.edu.sn",
    },
    "MANAGEMENT": {
        "LG1-ADMINISTRATIONDESAFFAIRES": "lg1-adm-affaires2025-2026@ism.edu.sn",
        "LG1-AGRO": "lg1-agro2025-2026@ism.edu.sn",
    },
    "GRADUATE": {
        "EMBA": "emba-2526@ism.edu.sn",
        "MASTER1-ACG": "mba1acg-2526@ism.edu.sn",
    },
    "MADIBA": {
        "JMI1": "jmi1-2526@ism.edu.sn",
        "LCM-1": "lcm1-2526@ism.edu.sn",
        "SPRI1-A": "spri1a-2526@ism.edu.sn",
        "MASTER1SCIENCEPOLITIQUEETRELATIONSINTERNATIONALES": "mba1spri-2526@ism.edu.sn",
    },
    "INGENIEUR": {
        "L1-CPD": "l1cpd-2526@ism.edu.sn",
        "L1-CYBERSÉCURITÉ": "l1ia-cyber-2526@ism.edu.sn",
        "L3-IAGEA": "licence3iage2025-2026@ism.edu.sn",
    }
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
        bad_pw = '"ismapps2025,,,,,,,,,,,,,,,,,1382"'
        good_pw = 'ismapps2025,,,,,,,,,,,,,,,,,1382'
        csv_text = csv_text.replace(bad_pw, good_pw)

    final_buffer = BytesIO(csv_text.encode("utf-8-sig"))
    final_buffer.seek(0)
    return final_buffer

def remove_accents_and_cleanup(s: str):
    if not isinstance(s, str): return s
    accent_map = str.maketrans({
        "à": "a", "â": "a", "ä": "a", "é": "e", "è": "e", "ê": "e", "ë": "e",
        "î": "i", "ï": "i", "ô": "o", "ö": "o", "û": "u", "ü": "u", "ÿ": "y"
    })
    return s.translate(accent_map)

def read_emails_txt(path: Path):
    if not path.exists(): return []
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def read_cours_mapping(cours_dir: Path):
    mapping = {}
    if not cours_dir.exists(): return mapping
    for txt_file in sorted(cours_dir.glob("*.txt")):
        mapping[txt_file.stem.upper()] = [l.strip() for l in txt_file.open("r", encoding="utf-8") if l.strip()]
    return mapping

def normalize_and_clean_df(df: pd.DataFrame):
    df.columns = [col.strip().capitalize() for col in df.columns]
    df = df[["Classe", "E-mail", "Nom", "Prénom"]].copy()
    df.columns = ["Classroom Name", "Member Email", "Nom", "Prénom"]
    
    df["Classroom Name"] = df["Classroom Name"].fillna("").astype(str).str.replace(r"\s+","",regex=True).str.upper()
    df["Member Email"] = df["Member Email"].fillna("").astype(str).str.strip().str.replace(r"\s+","",regex=True)
    df["Nom"] = df["Nom"].fillna("").astype(str).str.strip().apply(remove_accents_and_cleanup)
    df["Prénom"] = df["Prénom"].fillna("").astype(str).str.strip().apply(remove_accents_and_cleanup)
    return df

# -----------------------
# Sidebar & Header
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Minatholy Settings")
    selected_school = st.selectbox("Sélectionner l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP (Automatique)", value=True)
    st.divider()
    st.caption("v2.6 - Added Manual Inscription Mode")

st.title("🎓 ISM Data Processing")

# --- NAVIGATION TABS ---
main_tab, manual_tab = st.tabs(["🚀 Traitement Automatique (Excel)", "🛠️ Inscription Manuelle"])

# -----------------------
# TAB 1: AUTOMATIC TREATMENT
# -----------------------
with main_tab:
    with st.container(border=True):
        col_up1, col_up2 = st.columns([3, 1])
        with col_up1:
            uploaded_excel = st.file_uploader("Importer le fichier Excel des inscriptions", type=["xls", "xlsx"])
        with col_up2:
            st.write("##")
            run = st.button("🚀 Lancer le traitement", type="primary", use_container_width=True)

    if run:
        if not uploaded_excel:
            st.error("❌ Aucun fichier détecté.")
        else:
            with st.status("Traitement en cours...", expanded=True) as status:
                school_dir = DATA_DIR / selected_school
                admins = read_emails_txt(school_dir / "emails.txt")
                cours_mapping = read_cours_mapping(school_dir / "CoursParClasse")
                
                mapping_csv = school_dir / "mapping.csv"
                if mapping_csv.exists():
                    dfm = pd.read_csv(mapping_csv).fillna("")
                    mapping = dict(zip(dfm.iloc[:,0].str.upper().str.replace(" ",""), dfm.iloc[:,1].str.strip()))
                else:
                    mapping = {k.upper().replace(" ", ""): v for k, v in SCHOOL_MAPPINGS.get(selected_school, {}).items()}

                try:
                    raw_df = pd.read_excel(uploaded_excel, dtype=str)
                    df = normalize_and_clean_df(raw_df)
                    
                    valid_emails_mask = df["Member Email"].str.endswith("@ism.edu.sn", na=False) | df["Member Email"].str.endswith("@groupeism.sn", na=False)
                    invalid_emails_list = df[~valid_emails_mask]["Member Email"].unique().tolist()
                    valid_df = df[valid_emails_mask].copy()
                    
                    valid_df["Group Email"] = valid_df["Classroom Name"].map(mapping)
                    mapped_df = valid_df.dropna(subset=["Group Email"]).copy()
                    unmapped_students_df = valid_df[valid_df["Group Email"].isna()].copy()
                    
                    # 1. Google Export
                    google_list = mapped_df[["Group Email", "Member Email"]].copy()
                    google_list.columns = ["Group Email [Required]", "Member Email"]
                    google_list["Member Type"] = "USER"
                    google_list["Member Role"] = "MEMBER"

                    # 2. Admins
                    admin_rows = [{"Group Email [Required]": g, "Member Email": a, "Member Type": "USER", "Member Role": "MANAGER"} 
                                  for g in set(mapping.values()) for a in admins]
                    admin_df = pd.DataFrame(admin_rows)
                    combined_google = pd.concat([google_list, admin_df], ignore_index=True)

                    # 3. Profiles
                    profiles = valid_df.copy()
                    profiles["Prénom"] = "\"" + profiles["Prénom"] + "\""
                    profiles["Nouveau mot de passe"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
                    profile_export = profiles[["Member Email", "Nom", "Prénom", "Member Email", "Nouveau mot de passe"]]
                    profile_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]

                    # 4. Courses
                    course_rows = []
                    classes_sans_codes = set()
                    for classe in mapped_df["Classroom Name"].unique():
                        if classe not in cours_mapping or not cours_mapping[classe]:
                            classes_sans_codes.add(classe)
                        students_in_class = mapped_df[mapped_df["Classroom Name"] == classe]["Member Email"].unique()
                        codes = cours_mapping.get(classe, [])
                        for email in students_in_class:
                            for code in codes:
                                course_rows.append([code, email, "", ""])
                    course_df = pd.DataFrame(course_rows)

                    # --- REPORT ---
                    now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                    report = f"Report — {selected_school} — {now_str}\n\n"
                    distinct_mapped = mapped_df[["Classroom Name", "Group Email"]].drop_duplicates().sort_values("Classroom Name")
                    report += f"Mapped classes: {len(distinct_mapped)}\n"
                    for _, row in distinct_mapped.iterrows():
                        report += f"- {row['Classroom Name']} -> {row['Group Email']}\n"
                    
                    unmapped_classes = unmapped_students_df["Classroom Name"].unique().tolist()
                    report += f"\nUnmapped classes ({len(unmapped_classes)}):\n"
                    for c in unmapped_classes: report += f"- {c}\n"
                    
                    report += f"\nInvalid emails ({len(invalid_emails_list)}):\n"
                    for m in invalid_emails_list: report += f"- {m if m else '(Vide)'}\n"
                    
                    report += f"\nSummary counts:\n"
                    report += f"- Utilisateurs mappés: {len(mapped_df)}\n- Utilisateurs non mappés: {len(unmapped_students_df)}\n- Emails ignorés: {len(df) - len(valid_df)}\n- Classes sans codes: {len(classes_sans_codes)}\n"

                    bytes_mise = sanitize_csv_output(combined_google)
                    bytes_admin = sanitize_csv_output(admin_df)
                    bytes_profils = sanitize_csv_output(profile_export, is_profile=True)
                    bytes_courses = sanitize_csv_output(course_df, header=False)
                    b_report = BytesIO(report.encode("utf-8"))

                    status.update(label="✅ Traitement réussi !", state="complete", expanded=False)
                except Exception as e:
                    st.error(f"Erreur fatale : {e}")
                    st.stop()

            # --- DOWNLOADS ---
            res_tab1, res_tab2 = st.tabs(["📥 Téléchargements", "📊 Rapport"])
            with res_tab1:
                m1, m2, m3 = st.columns(3)
                m1.download_button("📧 Listes Google", bytes_mise, file_name=f"diffusion_{selected_school}.csv", use_container_width=True)
                m2.download_button("🔑 Admins Google", bytes_admin, file_name=f"admins_{selected_school}.csv", use_container_width=True)
                m3.download_button("👤 Profils Blackboard", bytes_profils, file_name=f"profils_blu_{selected_school}.csv", use_container_width=True)
                
                c1, c2 = st.columns(2)
                c1.download_button("📚 Inscriptions Cours", bytes_courses, file_name=f"cours_{selected_school}.csv", use_container_width=True)
                c2.download_button("📝 Rapport Complet", b_report, file_name=f"rapport_{selected_school}.txt", use_container_width=True)

                if zip_opt:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as zf:
                        zf.writestr("liste_diffusion.csv", bytes_mise.getvalue())
                        zf.writestr("admins.csv", bytes_admin.getvalue())
                        zf.writestr("profils_blackboard.csv", bytes_profils.getvalue())
                        zf.writestr("inscriptions_cours.csv", bytes_courses.getvalue())
                        zf.writestr("rapport.txt", report)
                    z_buf.seek(0)
                    st.download_button("📦 Télécharger ZIP", z_buf, file_name=f"export_{selected_school}.zip", type="primary", use_container_width=True)
            with res_tab2:
                st.text_area("Audit Log", report, height=300)

# -----------------------
# TAB 2: MANUAL INSCRIPTION
# -----------------------
with manual_tab:
    st.subheader("🛠️ Générateur d'inscriptions rapide")
    st.info("Utilisez cet outil pour lier un ou plusieurs codes de cours à une liste d'emails sans passer par l'Excel principal.")
    
    col_man1, col_man2 = st.columns(2)
    with col_man1:
        manual_codes = st.text_area("Codes de cours (un par ligne)", placeholder="Ex: MLI41251\nINFO101", height=200)
    with col_man2:
        manual_emails = st.text_area("Emails étudiants (un par ligne)", placeholder="Ex: jean.dieme@groupeism.sn\npauline.diatta@groupeism.sn", height=200)
    
    if st.button("🛠️ Générer Inscriptions Manuelles", use_container_width=True):
        codes_list = [c.strip() for c in manual_codes.split("\n") if c.strip()]
        emails_list = [e.strip() for e in manual_emails.split("\n") if e.strip()]
        
        if not codes_list or not emails_list:
            st.warning("⚠️ Veuillez remplir les codes et les emails.")
        else:
            manual_rows = []
            for email in emails_list:
                for code in codes_list:
                    manual_rows.append([code, email, "", ""])
            
            manual_df = pd.DataFrame(manual_rows)
            bytes_manual = sanitize_csv_output(manual_df, header=False)
            
            st.success(f"✅ Généré : {len(manual_rows)} inscriptions.")
            st.dataframe(manual_df.rename(columns={0: "Code", 1: "Email"}), use_container_width=True)
            st.download_button("📥 Télécharger inscriptions_manuelles.csv", bytes_manual, file_name="inscriptions_cours_manuel.csv", type="primary")
