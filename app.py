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

def parse_course_codes_file(uploaded_file) -> dict:
    """
    Parses an uploaded Excel or CSV file with course codes.
    Expects columns like: [Classe, Course Code] or [Classroom Name, Code]
    Returns a dict mapping: CLASSE -> [list of course codes]
    """
    if uploaded_file is None:
        return {}
    
    try:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file, dtype=str)
        else:
            df = pd.read_excel(uploaded_file, dtype=str)
        
        df.columns = [c.strip().lower() for c in df.columns]
        
        # Identify classroom and code columns
        class_col = next((c for c in df.columns if "class" in c or "classe" in c), df.columns[0])
        code_col = next((c for c in df.columns if "code" in c or "cours" in c), df.columns[1])
        
        mapping = {}
        for _, row in df.iterrows():
            c_name = str(row[class_col]).strip().replace(" ", "").upper()
            c_code = str(row[code_col]).strip().upper()
            if c_name and c_code and c_name != "NAN" and c_code != "NAN":
                mapping.setdefault(c_name, []).append(c_code)
                
        return mapping
    except Exception as e:
        st.error(f"Erreur lors de la lecture du fichier de codes de cours : {e}")
        return {}

# -----------------------
# Sidebar & Header
# -----------------------
with st.sidebar:
    st.image("https://upload.wikimedia.org/wikipedia/commons/b/b3/Apollo-kop%2C_objectnr_A_12979.jpg", width=80)
    st.title("Urgence IT Settings")
    selected_school = st.selectbox("Sélectionner l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP (Automatique)", value=True)
    st.divider()
    st.caption("v3.0 - Direct UI Course Code Upload")

st.title("Urgence IT BASH")

main_tab, manual_tab = st.tabs(["BASH", "MAJ de cours"])

# -----------------------
# TAB 1: AUTOMATIC TREATMENT
# -----------------------
with main_tab:
    with st.container(border=True):
        col_up1, col_up2 = st.columns(2)
        with col_up1:
            uploaded_excel = st.file_uploader("1️⃣ Importer le fichier Excel des inscriptions (Étudiants)", type=["xls", "xlsx"])
        with col_up2:
            uploaded_courses = st.file_uploader("2️⃣ Importer le fichier des codes de cours (Classe / Code)", type=["xls", "xlsx", "csv"])
            
        run = st.button("🚀 Lancer le traitement", type="primary", use_container_width=True)

    if run:
        if not uploaded_excel:
            st.error("❌ Aucun fichier d'inscriptions détecté.")
        else:
            with st.status("Traitement en cours...", expanded=True) as status:
                school_dir = DATA_DIR / selected_school
                admins = read_emails_txt(school_dir / "emails.txt")
                
                # Parse Course Mapping dynamically from Uploaded UI File
                cours_mapping = parse_course_codes_file(uploaded_courses)
                
                # Load Group Email Mappings
                mapping_csv = school_dir / "mapping.csv"
                if mapping_csv.exists():
                    dfm = pd.read_csv(mapping_csv).fillna("")
                    mapping = dict(zip(dfm.iloc[:,0].astype(str).str.upper().str.replace(" ", ""), dfm.iloc[:,1].astype(str).str.strip()))
                else:
                    mapping = {}

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
                    profiles["Nouveau mot de passe"] = "ismapps2025"
                    profile_export = profiles[["Member Email", "Nom", "Prénom", "Member Email", "Nouveau mot de passe"]]
                    profile_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]

                    # 4. Courses Mapping Calculation
                    course_rows = []
                    classes_sans_codes = set()
                    for classe in mapped_df["Classroom Name"].unique():
                        codes = cours_mapping.get(classe, [])
                        if not codes:
                            classes_sans_codes.add(classe)
                        students_in_class = mapped_df[mapped_df["Classroom Name"] == classe]["Member Email"].unique()
                        for email in students_in_class:
                            for code in codes:
                                course_rows.append([code, email, "", ""])
                    course_df = pd.DataFrame(course_rows)

                    # 5. Course Mapping Overview File (Classroom -> Codes Summary)
                    summary_course_rows = []
                    for classe, codes in cours_mapping.items():
                        for code in codes:
                            summary_course_rows.append([classe, code])
                    summary_course_df = pd.DataFrame(summary_course_rows, columns=["Classroom", "Course Code"])

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
                    
                    report += f"\nClasses without assigned course codes ({len(classes_sans_codes)}):\n"
                    for c in classes_sans_codes: report += f"- {c}\n"

                    # Save to Session State
                    st.session_state["results"] = {
                        "school": selected_school,
                        "bytes_mise": sanitize_csv_output(combined_google),
                        "bytes_admin": sanitize_csv_output(admin_df),
                        "bytes_profils": sanitize_csv_output(profile_export),
                        "bytes_courses": sanitize_csv_output(course_df, header=False),
                        "bytes_summary_courses": sanitize_csv_output(summary_course_df),
                        "report": report
                    }

                    status.update(label="✅ Traitement réussi !", state="complete", expanded=False)
                except Exception as e:
                    st.error(f"Erreur fatale : {e}")
                    st.stop()

    # Render results dynamically
    if "results" in st.session_state and st.session_state["results"]["school"] == selected_school:
        res = st.session_state["results"]
        b_report = BytesIO(res["report"].encode("utf-8"))

        res_tab1, res_tab2 = st.tabs(["📥 Téléchargements", "📊 Rapport"])
        with res_tab1:
            m1, m2, m3 = st.columns(3)
            m1.download_button("📧 Listes Google", res["bytes_mise"], file_name=f"diffusion_{selected_school}.csv", use_container_width=True)
            m2.download_button("🔑 Admins Google", res["bytes_admin"], file_name=f"admins_{selected_school}.csv", use_container_width=True)
            m3.download_button("👤 Profils Blackboard", res["bytes_profils"], file_name=f"profils_blu_{selected_school}.csv", use_container_width=True)
            
            c1, c2, c3 = st.columns(3)
            c1.download_button("📚 Inscriptions Cours", res["bytes_courses"], file_name=f"cours_{selected_school}.csv", use_container_width=True)
            c2.download_button("🗂️ Table Codes / Classes", res["bytes_summary_courses"], file_name=f"mapping_cours_{selected_school}.csv", use_container_width=True)
            c3.download_button("📝 Rapport Complet", b_report, file_name=f"rapport_{selected_school}.txt", use_container_width=True)

            if zip_opt:
                z_buf = BytesIO()
                with zipfile.ZipFile(z_buf, "w") as zf:
                    zf.writestr("liste_diffusion.csv", res["bytes_mise"].getvalue())
                    zf.writestr("admins.csv", res["bytes_admin"].getvalue())
                    zf.writestr("profils_blackboard.csv", res["bytes_profils"].getvalue())
                    zf.writestr("inscriptions_cours.csv", res["bytes_courses"].getvalue())
                    zf.writestr("mapping_cours_classes.csv", res["bytes_summary_courses"].getvalue())
                    zf.writestr("rapport.txt", res["report"])
                z_buf.seek(0)
                st.download_button("📦 Télécharger ZIP Global", z_buf, file_name=f"export_{selected_school}.zip", type="primary", use_container_width=True)
        with res_tab2:
            st.text_area("Audit Log", res["report"], height=300)

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
            manual_rows = [[code, email, "", ""] for email in emails_list for code in codes_list]
            manual_df = pd.DataFrame(manual_rows)
            bytes_manual = sanitize_csv_output(manual_df, header=False)
            
            st.success(f"✅ Généré : {len(manual_rows)} inscriptions.")
            st.dataframe(manual_df.rename(columns={0: "Code", 1: "Email"}), use_container_width=True)
            st.download_button("📥 Télécharger inscriptions_manuelles.csv", bytes_manual, file_name="inscriptions_cours_manuel.csv", type="primary")
