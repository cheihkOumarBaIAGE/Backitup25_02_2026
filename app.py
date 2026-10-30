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
    st.caption("v3.3 - Fixed Column Typo & Variable Persistence")

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
                school_dir = DATA_DIR / selected_school
                admins = read_emails_txt(school_dir / "emails.txt")
                
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

                    # 3. Profiles (Fixed Typo Here)
                    profiles = valid_df.copy()
                    profiles["Nouveau mot de passe"] = "ismapps2025"
                    profile_export = profiles[["Member Email", "Nom", "Prénom", "Member Email", "Nouveau mot de passe"]]
                    profile_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]

                    # Unique Classrooms detected
                    detected_classrooms = sorted(mapped_df["Classroom Name"].unique().tolist())

                    # Save base state into Session State
                    st.session_state["base_data"] = {
                        "school": selected_school,
                        "mapped_df": mapped_df,
                        "unmapped_students_df": unmapped_students_df,
                        "invalid_emails_list": invalid_emails_list,
                        "combined_google": combined_google,
                        "admin_df": admin_df,
                        "profile_export": profile_export,
                        "detected_classrooms": detected_classrooms,
                        "total_valid": len(valid_df),
                        "total_raw": len(df)
                    }
                    
                    # Reset previous course results if a new file is uploaded
                    if "results" in st.session_state:
                        del st.session_state["results"]

                    status.update(label="✅ Fichier traité ! Saisissez les codes de cours ci-dessous.", state="complete", expanded=False)
                except Exception as e:
                    st.error(f"Erreur fatale : {e}")
                    st.stop()

    # Interactive Course Entry & Download Rendering
    if "base_data" in st.session_state and st.session_state["base_data"]["school"] == selected_school:
        base = st.session_state["base_data"]
        
        tab_config, tab_downloads, tab_report = st.tabs(["📚 Configuration des Codes / Classes", "📥 Téléchargements", "📊 Rapport"])
        
        # TAB CONFIG: Interactive Form for Course Codes
        with tab_config:
            st.subheader("⚙️ Saisie des codes de cours par classe")
            st.info("Collez ou saisissez les codes de cours associés à chaque classe (un code par ligne).")
            
            with st.form("course_code_form"):
                course_input_data = {}
                cols = st.columns(2)
                for idx, classroom in enumerate(base["detected_classrooms"]):
                    col = cols[idx % 2]
                    with col:
                        course_input_data[classroom] = st.text_area(
                            f"🏫 Codes pour **{classroom}**",
                            placeholder="Ex:\nINFO101\nMATH202\nENG101",
                            height=120,
                            key=f"class_code_{classroom}"
                        )
                
                submit_courses = st.form_submit_button("💾 Valider & Générer les Inscriptions", type="primary", use_container_width=True)
            
            if submit_courses:
                course_rows = []
                summary_course_rows = []
                classes_sans_codes = set()
                
                mapped_df = base["mapped_df"]
                
                for classroom, text_input in course_input_data.items():
                    codes = [c.strip().upper() for c in text_input.split("\n") if c.strip()]
                    if not codes:
                        classes_sans_codes.add(classroom)
                    else:
                        for c in codes:
                            summary_course_rows.append([classroom, c])
                            
                    students_in_class = mapped_df[mapped_df["Classroom Name"] == classroom]["Member Email"].unique()
                    for email in students_in_class:
                        for code in codes:
                            course_rows.append([code, email, "", ""])
                            
                course_df = pd.DataFrame(course_rows)
                summary_course_df = pd.DataFrame(summary_course_rows, columns=["Classroom", "Course Code"])
                
                # Build Audit Log Report
                now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
                report = f"Report — {selected_school} — {now_str}\n\n"
                distinct_mapped = mapped_df[["Classroom Name", "Group Email"]].drop_duplicates().sort_values("Classroom Name")
                report += f"Mapped classes: {len(distinct_mapped)}\n"
                for _, row in distinct_mapped.iterrows():
                    report += f"- {row['Classroom Name']} -> {row['Group Email']}\n"
                
                unmapped_classes = base["unmapped_students_df"]["Classroom Name"].unique().tolist()
                report += f"\nUnmapped classes ({len(unmapped_classes)}):\n"
                for c in unmapped_classes: report += f"- {c}\n"
                
                report += f"\nClasses without assigned course codes ({len(classes_sans_codes)}):\n"
                for c in classes_sans_codes: report += f"- {c}\n"

                report += f"\nSummary counts:\n"
                report += f"- Utilisateurs mappés: {len(mapped_df)}\n- Utilisateurs non mappés: {len(base['unmapped_students_df'])}\n- Emails ignorés: {base['total_raw'] - base['total_valid']}\n- Classes sans codes: {len(classes_sans_codes)}\n"

                st.session_state["results"] = {
                    "bytes_mise": sanitize_csv_output(base["combined_google"]),
                    "bytes_admin": sanitize_csv_output(base["admin_df"]),
                    "bytes_profils": sanitize_csv_output(base["profile_export"]),
                    "bytes_courses": sanitize_csv_output(course_df, header=False),
                    "bytes_summary_courses": sanitize_csv_output(summary_course_df),
                    "report": report,
                    "total_inscriptions": len(course_rows)
                }
                
                st.success(f"✅ Inscriptions générées avec succès ({len(course_rows)} inscriptions enregistrées). Rendez-vous dans l'onglet '📥 Téléchargements' !")

        # TAB DOWNLOADS: Output CSVs and ZIP
        with tab_downloads:
            if "results" not in st.session_state:
                st.warning("⚠️ Veuillez remplir et valider la grille des codes de cours dans l'onglet '📚 Configuration des Codes / Classes' d'abord.")
            else:
                res = st.session_state["results"]
                b_report = BytesIO(res["report"].encode("utf-8"))

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

        # TAB REPORT: Audit logs
        with tab_report:
            if "results" in st.session_state:
                st.text_area("Audit Log", st.session_state["results"]["report"], height=300)
            else:
                st.info("Validez la configuration des cours pour voir le rapport d'audit.")

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
