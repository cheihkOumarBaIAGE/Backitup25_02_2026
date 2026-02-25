import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import zipfile

# -----------------------
# 1. CONFIGURATION
# -----------------------
st.set_page_config(page_title="Minatholy V3.2", layout="wide", page_icon="🛡️")

# Intégration de ton dictionnaire de mapping (extrait pour l'exemple)
# Remets ici TOUT ton bloc SCHOOL_MAPPINGS de l'ancien code
SCHOOL_MAPPINGS = {
    "INGENIEUR": {
        "L1-CPD": "l1cpd-2526@ism.edu.sn",
        "L1INGENIEURS-R2A": "l1r2.ingenieura@ism.edu.sn",
        # ... remets tout ici
    },
    "MANAGEMENT": { "LG1-ADMINISTRATIONDESAFFAIRES": "lg1-adm-affaires2025-2026@ism.edu.sn" },
    "DROIT": {}, "GRADUATE": {}, "MADIBA": {}
}

# -----------------------
# 2. FONCTIONS DE NETTOYAGE (FIX "CLASSE")
# -----------------------
def clean_columns(df):
    """Supprime les espaces invisibles et normalise les noms de colonnes."""
    # Nettoyage des noms de colonnes (ex: " Classe " -> "Classe")
    df.columns = [str(c).strip() for c in df.columns]
    
    # Mapping intelligent (cherche 'classe' ou 'email' même avec fautes/espaces)
    for col in df.columns:
        c_low = col.lower()
        if 'classe' in c_low: df = df.rename(columns={col: 'Classe'})
        if 'identifiant' in c_low or 'email' in c_low or 'e-mail' in c_low: 
            df = df.rename(columns={col: 'Member Email'})
        if 'nom' == c_low: df = df.rename(columns={col: 'Nom'})
        if 'prénom' in c_low or 'prenom' in c_low: df = df.rename(columns={col: 'Prénom'})
    return df

def remove_accents(s):
    if not isinstance(s, str): return s
    accents = {'à':'a','â':'a','ä':'a','é':'e','è':'e','ê':'e','ë':'e','î':'i','ï':'i','ô':'o','ö':'o','û':'u','ü':'u'}
    for a, b in accents.items(): s = s.replace(a, b)
    return s

# -----------------------
# 3. INTERFACE UTILISATEUR
# -----------------------
st.title("🛡️ Minatholy V3.2")
st.subheader("Dashboard de Gestion des Effectifs")

with st.sidebar:
    st.header("⚙️ Paramètres")
    selected_school = st.selectbox("Sélectionner l'école", list(SCHOOL_MAPPINGS.keys()))
    uploaded_file = st.file_uploader("Upload Excel (.xlsx ou .xls)", type=["xls", "xlsx"])
    
    st.divider()
    st.header("👥 Administration")
    admin_emails = st.text_area("Emails Admins (Manager)", 
                                "marie-bernadette.ndione@groupeism.sn\nmar-sarr.ndiaye@groupeism.sn")

if uploaded_file:
    # Lecture multi-format
    df_raw = pd.read_excel(uploaded_file)
    df = clean_columns(df_raw)
    
    st.info(f"📁 Fichier détecté : {len(df)} lignes trouvées.")

    if st.button("🚀 LANCER LE TRAITEMENT", type="primary"):
        # --- LOGIQUE MÉTIER ---
        mapping = SCHOOL_MAPPINGS.get(selected_school, {})
        
        # Nettoyage des données
        df['Classe_Clean'] = df['Classe'].fillna('').astype(str).str.strip().str.upper().str.replace(' ', '')
        valid_df = df[df["Member Email"].str.contains("@", na=False)].copy()
        valid_df["Group Email"] = valid_df["Classe_Clean"].map(mapping)
        
        # 1. Mailing Lists (GW)
        gw_df = valid_df.dropna(subset=["Group Email"])[["Group Email", "Member Email"]].copy()
        gw_df.columns = ["Group Email [Required]", "Member Email"]
        gw_df["Member Type"], gw_df["Member Role"] = "USER", "MEMBER"
        
        # Ajout des Admins (MANAGER)
        admins = [a.strip() for a in admin_emails.split('\n') if a.strip()]
        admin_rows = []
        for g_mail in gw_df["Group Email [Required]"].unique():
            for a_mail in admins:
                admin_rows.append([g_mail, a_mail, "USER", "MANAGER"])
        full_gw = pd.concat([gw_df, pd.DataFrame(admin_rows, columns=gw_df.columns)])

        # 2. Profils Blackboard (BLU)
        profils_df = pd.DataFrame({
            "Nom d'utilisateur": valid_df["Member Email"],
            "Nom": valid_df["Nom"].fillna('').str.upper().apply(remove_accents),
            "Prénom": '"' + valid_df["Prénom"].fillna('').apply(remove_accents) + '"',
            "Adresse e-mail": valid_df["Member Email"],
            "Mot de passe": "ismapps2025,,,,,,,,,,,,,,,,,1382"
        })

        # --- DASHBOARD DE RAPPORT (AFFICHAGE DIRECT) ---
        st.divider()
        st.header("📊 Rapport d'Exécution")
        
        m1, m2, m3 = st.columns(3)
        m1.metric("Étudiants validés", len(valid_df))
        m2.metric("Classes traitées", gw_df["Group Email [Required]"].nunique())
        m3.metric("Admins ajoutés", len(admins))

        # Alertes sur les classes non mappées
        unmapped = valid_df[valid_df["Group Email"].isna()]["Classe"].unique()
        if len(unmapped) > 0:
            with st.expander("⚠️ Classes sans correspondance (Mapping manquant)", expanded=True):
                st.write(list(unmapped))
                st.caption("Vérifie si ces noms de classes sont bien dans ton dictionnaire SCHOOL_MAPPINGS.")

        # --- GÉNÉRATION DU ZIP ---
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            zf.writestr(f"1_Mailing_Lists_{selected_school}.csv", full_gw.to_csv(index=False, encoding="utf-8-sig"))
            zf.writestr(f"2_Profils_Blackboard_{selected_school}.csv", profils_df.to_csv(index=False, encoding="utf-8-sig"))
            # Tu peux ajouter ici le CSV d'inscription aux cours si besoin
        
        st.divider()
        st.download_button(
            label="📥 Télécharger les fichiers CSV (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Minatholy_{selected_school}_{datetime.now().strftime('%Hh%M')}.zip",
            mime="application/zip",
            use_container_width=True
        )

st.divider()
st.caption("Minatholy V3.2 | ISM Dakar | Correction automatique des colonnes activée")
