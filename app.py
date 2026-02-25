import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import zipfile

# ---------------------------------------------------------
# 1. MAPPING COMPLET (Directement intégré pour GitHub)
# ---------------------------------------------------------
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
        "LICENCE3ALBIEXTERNE": "lda3ae-2025-2026@ism.edu.sn"
    },
    "MANAGEMENT": {
        "LG1-ADMINISTRATIONDESAFFAIRES": "lg1-adm-affaires2025-2026@ism.edu.sn",
        "LG1-AGRO": "lg1-agro2025-2026@ism.edu.sn",
        "LG1-CIM": "lg1-cim2025-2026@ism.edu.sn",
        "LG1J-ORGA-GRH": "lg1j-orga-grh2025-2026@ism.edu.sn",
        "LG2-ADMINISTRATIONDESAFFAIRES": "lg2-adm-affaires2025-2026@ism.edu.sn",
        "LG3-MARKETINGETCOMMUNICATION": "lg3-marketing-communication2025-2026@ism.edu.sn"
    },
    "INGENIEUR": {
        "L1-CPD": "l1cpd-2526@ism.edu.sn",
        "L1-CDSD": "l1cdsd-2526@ism.edu.sn",
        "L1-CYBERSÉCURITÉ": "l1ia-cyber-2526@ism.edu.sn",
        "L1A-IA(IAGE&TTL)": "l1aia-2526@ism.edu.sn",
        "L2-INTELLIGENCEARTIFICIELLE": "licence2ia2025-2026@ism.edu.sn",
        "M1IDC-BIGDATA&DATASTRATÉGIE": "mba1bigdata-2526@ism.edu.sn"
    }
}

# ---------------------------------------------------------
# 2. FONCTION DE NETTOYAGE ROBUSTE
# ---------------------------------------------------------
def clean_and_prepare(df):
    # Nettoyage des noms de colonnes (espaces invisibles)
    df.columns = [str(c).strip() for c in df.columns]
    
    # Correction automatique si "E-mail" est utilisé
    rename_map = {}
    for col in df.columns:
        if col.lower() in ["e-mail", "email", "mail", "identifiant"]:
            rename_map[col] = "Member Email"
    
    return df.rename(columns=rename_map)

# ---------------------------------------------------------
# 3. INTERFACE STREAMLIT
# ---------------------------------------------------------
st.set_page_config(page_title="Minatholy V3.2", layout="wide")
st.title("🛡️ Minatholy V3.2")

with st.sidebar:
    school = st.selectbox("Ecole", list(SCHOOL_MAPPINGS.keys()))
    file = st.file_uploader("Upload Excel", type=["xlsx", "xls"])
    admins = st.text_area("Managers", "marie-bernadette.ndione@groupeism.sn\nmar-sarr.ndiaye@groupeism.sn")

if file:
    df = clean_and_prepare(pd.read_excel(file))
    
    # Vérification que les colonnes Classe et Member Email (ex E-mail) existent
    if "Member Email" not in df.columns or "Classe" not in df.columns:
        st.error("❌ Colonnes manquantes ! Vérifiez que votre Excel contient bien 'Classe' et 'E-mail'.")
        st.write("Colonnes détectées :", list(df.columns))
        st.stop()

    if st.button("🚀 TRAITER LE FICHIER"):
        mapping = SCHOOL_MAPPINGS.get(school, {})
        
        # Nettoyage des données
        df["Classe_Clean"] = df["Classe"].astype(str).str.strip()
        
        # Filtrage des lignes valides
        valid_df = df[df["Member Email"].str.contains("@", na=False)].copy()
        valid_df["Group Email"] = valid_df["Classe_Clean"].map(mapping)
        
        # --- FICHIER GOOGLE GROUPS ---
        gw = valid_df.dropna(subset=["Group Email"])[["Group Email", "Member Email"]].copy()
        gw.columns = ["Group Email [Required]", "Member Email"]
        gw["Member Type"], gw["Member Role"] = "USER", "MEMBER"
        
        # Ajout des admins
        admin_list = [a.strip() for a in admins.split("\n") if a.strip()]
        admin_rows = []
        for g in gw["Group Email [Required]"].unique():
            for a in admin_list:
                admin_rows.append([g, a, "USER", "MANAGER"])
        final_gw = pd.concat([gw, pd.DataFrame(admin_rows, columns=gw.columns)])

        # --- EXPORT ---
        buf = BytesIO()
        with zipfile.ZipFile(buf, "w") as zf:
            zf.writestr(f"Google_Groups_{school}.csv", final_gw.to_csv(index=False, encoding="utf-8-sig"))
            
        st.success(f"✅ Terminé ! {len(valid_df)} étudiants identifiés.")
        
        # Affichage des classes non trouvées pour correction
        unmapped = valid_df[valid_df["Group Email"].isna()]["Classe_Clean"].unique()
        if len(unmapped) > 0:
            with st.expander("⚠️ Classes non reconnues dans le mapping"):
                st.write(list(unmapped))

        st.download_button("📥 Télécharger ZIP", buf.getvalue(), f"Minatholy_{school}.zip")
