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
        "L1-ADM.PUBLIQUE": "lda1c-2025-2026@ism.edu.sn",
        "L1DroitPrivéFondamental": "lda1c-2025-2026@ism.edu.sn",
        "LDA1-A": "lda1a-2025-2026@ism.edu.sn",
        "LDA1-BILINGUE": "lda1bilingue-2025-2026@ism.edu.sn",
        "LDA1-B": "lda1b-2025-2026@ism.edu.sn",
        "LDA1-C": "lda1c-2025-2026@ism.edu.sn",
        "LDA1-D": "lda1d-2025-2026@ism.edu.sn",
        "LDA1-E": "lda1e-2025-2026@ism.edu.sn",
        "LDA1-2R": "lda1-2r-2025-2026@ism.edu.sn",
        "L2-DROITPRIVÉFONDAMENTAL": "l2droitprive-fondamental-2526@ism.edu.sn",
        "L2ALBI-DROITGESTION": "l2albi-droitgestion-2526@ism.edu.sn",
        "L2-JURISTED'ENTREPRISE": "l2juriste-dentreprise-2526@ism.edu.sn",
        "LDA2-A": "lda2a-2025-2026@ism.edu.sn",
        "LDA2-B": "lda2b-2025-2026@ism.edu.sn",
        "LDA2-C": "lda2c-2025-2026@ism.edu.sn",
        "LDA2-D": "lda2d-2025-2026@ism.edu.sn",
        "LDA3-A": "lda3a-2025-2026@ism.edu.sn",
        "LDA3-B": "lda3b-2025-2026@ism.edu.sn",
        "LDA3-C": "lda3c-2025-2026@ism.edu.sn",
        "LDA3-D": "lda3d-2025-2026@ism.edu.sn",
        "LDA3-COURSDUSOIR": "lda3coursdusoir-2025-2026@ism.edu.sn",
        "LICENCE3ALBIEXTERNE": "lda3ae-2025-2026@ism.edu.sn",
        "MBA1-ESGJURISTED'ENTREPRISE": "mba1jusristedentreprise-2526@ism.edu.sn",
        "M1-JURISTED'AFFAIRES,ASSURANCE&COMPLIANCE": "mba1jaac-2526@ism.edu.sn",
        "MASTER1-PASSATIONDESMARCHÉS": "mba1passationdesmarches-2526@ism.edu.sn",
        "MASTER1-PASSATIONDESMARCHÉSSOIR": "mba1passationdesmarchessoir-2526@ism.edu.sn",
        "MASTER1-DROITDESAFFAIRES": "mba1droitdesaffaires-2526@ism.edu.sn",
        "MASTER1-DROITDESAFFAIRESSOIR": "mba1droitdesaffairessoir-2526@ism.edu.sn",
        "MASTER1DROITDESAFFAIRESETBUSINESSPARTNERDEELIJE": "mba1droitdesaffaires.business-2526@ism.edu.sn",
        "MBA1-DROITINTERNATIONALDESAFFAIRES": "mba1dia-2526@ism.edu.sn",
        "MASTER1-FISCALITÉ": "mba1fiscalite-2526@ism.edu.sn",
        "MASTER1-FISCALITESOIR": "mba1fiscalitesoir-2526@ism.edu.sn",
        "MASTER1-FISCALITECOURSDUSOIR": "mba1fiscalitesoir-2526@ism.edu.sn",
        "MBA1-DGEM": "mba1dgem-2526@ism.edu.sn",
        "MBA1-DROITDEL'ENTREPRISE": "mba1delentreprise-2526@ism.edu.sn",
        "MBA1-DROITMARITIMEETACTIVITÉSPORTUAIRES": "mba1droitmaritime.activiteportuaires-2526@ism.edu.sn",
        "MASTER1-DROITNOTARIALETGESTIONDUPATRIMOINE": "mba1droitnotarial.gestionpatrimoine-2526@ism.edu.sn",
        "MBA2-DROITDEL'ENTREPRISE": "mba2delentreprise-2526@ism.edu.sn",
        "MBA1-DROITSINTERNATIONALESDESAFFAIRES": "mba1dia-2526@ism.edu.sn",
        "M1-JURISTEDEBANQUEASSURANCE&COMPLIANCE": "mba1jbac-2526@ism.edu.sn",
        "M1-SOIRJURISTEDEBANQUEASSURANCE&COMPLIANCE": "mba1jbacsoir-2526@ism.edu.sn",
        "MBA2-ESGJURISTED'ENTREPRISE": "mba2jusristedentreprise-2526@ism.edu.sn",
        "MASTER2-PASSATIONDESMARCHÉS": "mba2passationdesmarches-2526@ism.edu.sn",
        "MASTER2-PASSATIONDESMARCHÉSSOIR": "mba2passationdesmarchessoir-2526@ism.edu.sn",
        "MASTER2-DROITDESAFFAIRES": "mba2droitdesaffaires-2526@ism.edu.sn",
        "MASTER2-DROITDESAFFAIRESCOURSDUSOIR": "mba2droitdesaffairessoir-2526@ism.edu.sn",
        "MASTER2-FISCALITÉ": "mba2fiscalite-2526@ism.edu.sn",
        "MASTER2-FISCALITECOURSDUSOIR": "mba2fiscalitesoir-2526@ism.edu.sn",
        "MBA2-DGEM": "mba2dgem-2526@ism.edu.sn",
        "MBA2-DROITMARITIMEETACTIVITÉSPORTUAIRES": "mba2droitmaritime.activiteportuaires-2526@ism.edu.sn",
        "MBA2-DROITNOTARIALETGESTIONDUPATRIMOINE": "mba2droitnotarial.gestionpatrimoine-2526@ism.edu.sn",
        "MBA2-DROITINTERNATIONALDESAFFAIRES": "mba2dia-2526@ism.edu.sn",
        "M2-JURISTEDEBANQUEASSURANCE&COMPLIANCE": "mba2jbac-2526@ism.edu.sn",
    },
    "MANAGEMENT": {
        "LG1-ADMINISTRATIONDESAFFAIRES": "lg1-adm-affaires2025-2026@ism.edu.sn",
        "LG1-ADMINISTRATIONDESAFFAIRESB": "lg1-adm-affairesb2025-2026@ism.edu.sn",
        "LG1-AGRO": "lg1-agro2025-2026@ism.edu.sn",
        "LG2-AGRO": "lg2-agro2025-2026@ism.edu.sn",
        "LG3-AGRO": "lg3-agro2025-2026@ism.edu.sn",
        "LG1-CIM": "lg1-cim2025-2026@ism.edu.sn",
        "LG2-CIM": "lg2-cim2025-2026@ism.edu.sn",
        "LG3-MCI": "lg3-cim2025-2026@ism.edu.sn",
    },
    "GRADUATE": {
        "EMBA": "emba-2526@ism.edu.sn",
        "MASTER1-ACG": "mba1acg-2526@ism.edu.sn",
        "MASTER2-ACG": "mba2acg-2526@ism.edu.sn",
    },
    "MADIBA": {
        "JMI1": "jmi1-2526@ism.edu.sn",
        "LCM-1": "lcm1-2526@ism.edu.sn",
        "LCMFULLENGLISH-1": "lcm1bilingue-2526@ism.edu.sn",
        "SPRI1-A": "spri1a-2526@ism.edu.sn",
        "SPRI1-BILINGUE": "spri1bilingue-2526@ism.edu.sn",
        "SPRI1-B": "spri1b-2526@ism.edu.sn",
        "MLI-R2": "mlir2-2526@ism.edu.sn",
        "JMI-2": "jmi2-2526@ism.edu.sn",
        "LCM-2": "lcm2-2526@ism.edu.sn",
        "SPRI2-A": "spri2a-2526@ism.edu.sn",
        "SPRI2-B": "spri2b-2526@ism.edu.sn",
        "JMI-3": "jmi3-2526@ism.edu.sn",
        "LCM-3": "lcm3-2526@ism.edu.sn",
        "SPRI3-A": "spri3a-2526@ism.edu.sn",
        "SPRI3-B": "spri3b-2526@ism.edu.sn",
        "SPRI3-C": "spri3c-2526@ism.edu.sn",
        "MASTER1SCIENCEPOLITIQUEETRELATIONSINTERNATIONALES": "mba1spri-2526@ism.edu.sn",
        "M1-SPRISOIR": "mba1sprisoir-2526@ism.edu.sn",
        "MBA1-DIPLOMATIEETGÉOSTRATÉGIE": "mba1dg-2526@ism.edu.sn",
        "MBA1-DIPLOMATIEETGÉOSTRATÉGIESOIR": "mba1dgsoir-2526@ism.edu.sn",
        "MBA1-Gestiondeprojetsculturels": "mba1gpc-2526@ism.edu.sn",
        "MBA1-CLRP": "mba1clrp-2526@ism.edu.sn",
        "MBA1-CLRPSOIR": "mba1clrpsoir-2526@ism.edu.sn",
        "MBA1-DGT": "mba1dgt-2526@ism.edu.sn",
        "MBA1-EnvironnementetDéveloppementDurable": "mba1edd-2526@ism.edu.sn",
        "MBA1-SPRIPAIXETSÉCURITÉ": "mba1sps-2526@ism.edu.sn",
        "MASTER2SCIENCEPOLITIQUEETRELATIONSINTERNATIONALES": "mba2spri-2526@ism.edu.sn",
        "MBA2-DGT": "mba2dgt-2526@ism.edu.sn",
        "MBA2-CLRP": "mba2clrp-2526@ism.edu.sn",
        "MBA2-CLRPSOIR": "mba2clrp-2526@ism.edu.sn",
        "MBA2DIPLOMATIEETGÉOSTRATÉGIE": "mba2dg-2526@ism.edu.sn",
        "MBA2-SPRIPAIXETSECURITE": "mba2sps-2526@ism.edu.sn",
        "MBA2-GOUVERNANCEETMANAGEMENTPUBLIC": "mba2gouvernance.management.public-2526@ism.edu.sn",
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
    zip_opt = st.checkbox("Générer Archive ZIP", value=True)

st.title("🎓 ISM Data Processing")

with st.container(border=True):
    col_up1, col_up2 = st.columns([3, 1])
    with col_up1:
        uploaded_excel = st.file_uploader("Importer le fichier Excel des inscriptions", type=["xls", "xlsx"])
    with col_up2:
        st.write("##")
        run = st.button("🚀 Lancer le traitement", type="primary", use_container_width=True)

# -----------------------
# Execution Logic
# -----------------------
if run:
    if not uploaded_excel:
        st.error("❌ Aucun fichier détecté.")
        st.stop()

    with st.status("Traitement en cours...", expanded=True) as status:
        school_dir = DATA_DIR / selected_school
        admins = read_emails_txt(school_dir / "emails.txt")
        cours_mapping = read_cours_mapping(school_dir / "CoursParClasse")
        
        mapping_csv = school_dir / "mapping.csv"
        if mapping_csv.exists():
            dfm = pd.read_csv(mapping_csv).fillna("")
            mapping = dict(zip(dfm.iloc[:,0].str.upper().str.replace(" ",""), dfm.iloc[:,1].str.strip()))
        else:
            mapping = {k.upper().replace(" ", ""): v for k, v in SCHOOL_MAPPINGS[selected_school].items()}

        try:
            raw_df = pd.read_excel(uploaded_excel, dtype=str)
            df = normalize_and_clean_df(raw_df)
            
            # 1. Separate Email types
            valid_emails_mask = df["Member Email"].str.endswith("@ism.edu.sn", na=False)
            invalid_emails_list = df[~valid_emails_mask]["Member Email"].unique().tolist()
            valid_df = df[valid_emails_mask].copy()
            
            # 2. Apply Mapping
            valid_df["Group Email"] = valid_df["Classroom Name"].map(mapping)
            mapped_df = valid_df.dropna(subset=["Group Email"]).copy()
            unmapped_students_df = valid_df[valid_df["Group Email"].isna()].copy()
            
            # 3. Google Export
            google_list = mapped_df[["Group Email", "Member Email"]].copy()
            google_list.columns = ["Group Email [Required]", "Member Email"]
            google_list["Member Type"] = "USER"
            google_list["Member Role"] = "MEMBER"

            # 4. Admins
            admin_rows = [{"Group Email [Required]": g, "Member Email": a, "Member Type": "USER", "Member Role": "MANAGER"} 
                          for g in set(mapping.values()) for a in admins]
            admin_df = pd.DataFrame(admin_rows)
            combined_google = pd.concat([google_list, admin_df], ignore_index=True)

            # 5. Profiles (Blackboard)
            profiles = valid_df.copy()
            profiles["Prénom"] = "\"" + profiles["Prénom"] + "\""
            profiles["Nouveau mot de passe"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
            profile_export = profiles[["Member Email", "Nom", "Prénom", "Member Email", "Nouveau mot de passe"]]
            profile_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]

            # 6. Courses Enrollment
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

            # --- GENERATE REPORT TEXT ---
            now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
            report = f"Report — {selected_school} — {now_str}\n\n"
            
            # Mapped Classes list
            distinct_mapped = mapped_df[["Classroom Name", "Group Email"]].drop_duplicates().sort_values("Classroom Name")
            report += f"Mapped classes: {len(distinct_mapped)}\n"
            for _, row in distinct_mapped.iterrows():
                report += f"- {row['Classroom Name']} -> {row['Group Email']}\n"
            
            # Unmapped Classes
            unmapped_classes = unmapped_students_df["Classroom Name"].unique().tolist()
            report += f"\nUnmapped classes ({len(unmapped_classes)}):\n"
            for c in unmapped_classes:
                report += f"- {c}\n"
            
            # Invalid Emails
            report += f"\nInvalid emails ({len(invalid_emails_list)}):\n"
            for m in invalid_emails_list:
                report += f"- {m if m else '(Vide)'}\n"
            
            # Summary
            report += f"\nSummary counts:\n"
            report += f"- Utilisateurs mappés: {len(mapped_df)}\n"
            report += f"- Utilisateurs non mappés: {len(unmapped_students_df)}\n"
            report += f"- Emails ignorés: {len(df) - len(valid_df)}\n"
            report += f"- Classes sans codes: {len(classes_sans_codes)}\n"
            
            # Classes sans codes list
            report += f"\nClasses sans codes:\n"
            for c in sorted(list(classes_sans_codes)):
                report += f"- {c}\n"

            # Prepare downloads
            bytes_mise = sanitize_csv_output(combined_google)
            bytes_admin = sanitize_csv_output(admin_df)
            bytes_profils = sanitize_csv_output(profile_export, is_profile=True)
            bytes_courses = sanitize_csv_output(course_df, header=False)
            b_report = BytesIO(report.encode("utf-8"))

            status.update(label="✅ Traitement réussi !", state="complete", expanded=False)
        except Exception as e:
            st.error(f"Erreur fatale : {e}")
            st.stop()

    # --- UI RESULTS ---
    tab1, tab2, tab3 = st.tabs(["📥 Téléchargements", "🔍 Aperçu Données", "📊 Rapport d'audit"])

    with tab1:
        st.subheader("Fichiers de sortie")
        m1, m2, m3 = st.columns(3)
        m1.download_button("📧 Listes Google", bytes_mise, file_name=f"diffusion_{selected_school}.csv", use_container_width=True)
        m2.download_button("🔑 Admins Google", bytes_admin, file_name=f"admins_{selected_school}.csv", use_container_width=True)
        m3.download_button("👤 Profils Blackboard", bytes_profils, file_name=f"profils_blu_{selected_school}.csv", use_container_width=True)
        
        c1, c2 = st.columns(2)
        c1.download_button("📚 Inscriptions Cours", bytes_courses, file_name=f"cours_{selected_school}.csv", use_container_width=True)
        c2.download_button("📝 Rapport Complet (TXT)", b_report, file_name=f"rapport_{selected_school}_{now_str}.txt", use_container_width=True)

        if zip_opt:
            st.divider()
            z_buf = BytesIO()
            with zipfile.ZipFile(z_buf, "w") as zf:
                zf.writestr("liste_diffusion.csv", bytes_mise.getvalue())
                zf.writestr("admins.csv", bytes_admin.getvalue())
                zf.writestr("profils_blackboard.csv", bytes_profils.getvalue())
                zf.writestr("inscriptions_cours.csv", bytes_courses.getvalue())
                zf.writestr("rapport.txt", report)
            z_buf.seek(0)
            st.download_button("📦 Télécharger tout (ZIP)", z_buf, file_name=f"export_{selected_school}.zip", type="primary", use_container_width=True)

    with tab2:
        st.dataframe(profile_export.head(50), use_container_width=True)

    with tab3:
        st.text_area("Aperçu du rapport", report, height=400)
