import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO
import zipfile

# 1. Configuration de la page
st.set_page_config(page_title="Minatholy V3.2", layout="wide")

# 2. Chargement du mapping depuis secrets.toml
try:
    SCHOOL_MAPPINGS = st.secrets["SCHOOL_MAPPINGS"]
except Exception:
    st.error("❌ Le fichier secrets.toml est manquant ou mal configuré.")
    st.stop()

def clean_data(df):
    """
    Cette fonction renomme 'E-mail' en 'Member Email' 
    et nettoie les noms de colonnes.
    """
    # Supprime les espaces invisibles autour des noms de colonnes
    df.columns = [str(c).strip() for c in df.columns]
    
    # Dictionnaire de correspondance pour les variantes de noms
    rename_dict = {}
    for col in df.columns:
        c_low = col.lower()
        if c_low in ['e-mail', 'email', 'identifiant', 'mail']:
            rename_dict[col] = "Member Email"
        if c_low in ['classe', 'class', 'groupe']:
            rename_dict[col] = "Classe"
            
    return df.rename(columns=rename_dict)

# --- INTERFACE ---
st.title("🛡️ Minatholy V3.2")

with st.sidebar:
    selected_school = st.selectbox("Sélectionnez l'école", list(SCHOOL_MAPPINGS.keys()))
    uploaded_file = st.file_uploader("Upload du fichier Excel", type=["xlsx", "xls"])
    admin_emails = st.text_area("Admins (Mailing List)", "marie-bernadette.ndione@groupeism.sn\nmar-sarr.ndiaye@groupeism.sn")

if uploaded_file:
    # Lecture du fichier
    df_raw = pd.read_excel(uploaded_file)
    
    # NETTOYAGE (Le fix est ici)
    df = clean_data(df_raw)
    
    # Vérification de sécurité avant de continuer
    if "Member Email" not in df.columns:
        st.error(f"❌ La colonne 'E-mail' n'a pas été détectée. Colonnes présentes : {list(df_raw.columns)}")
        st.stop()

    if st.button("🚀 LANCER LE TRAITEMENT", type="primary"):
        # Récupération du mapping de l'école choisie
        mapping = SCHOOL_MAPPINGS.get(selected_school, {})
        
        # Nettoyage des données
        df["Classe"] = df["Classe"].astype(str).str.strip()
        
        # Ligne 77 (Corrigée par le renommage automatique plus haut)
        valid_df = df[df["Member Email"].str.contains("@", na=False)].copy()
        
        # Mapping des classes vers les emails de groupes
        valid_df["Group Email"] = valid_df["Classe"].map(mapping)
        
        # --- PREPARATION DES FICHIERS ---
        # Google Workspace
        gw_df = valid_df.dropna(subset=["Group Email"])[["Group Email", "Member Email"]].copy()
        gw_df.columns = ["Group Email [Required]", "Member Email"]
        gw_df["Member Type"] = "USER"
        gw_df["Member Role"] = "MEMBER"
        
        # Ajout des admins
        admins = [a.strip() for a in admin_emails.split("\n") if a.strip()]
        admin_rows = []
        for group in gw_df["Group Email [Required]"].unique():
            for admin in admins:
                admin_rows.append([group, admin, "USER", "MANAGER"])
        
        final_gw = pd.concat([gw_df, pd.DataFrame(admin_rows, columns=gw_df.columns)])

        # Blackboard (Profils)
        blu_df = pd.DataFrame({
            "Nom d'utilisateur": valid_df["Member Email"],
            "Nom": valid_df["Nom"].fillna("").str.upper(),
            "Prénom": valid_df["Prénom"].fillna(""),
            "Adresse e-mail": valid_df["Member Email"],
            "Mot de passe": "ismapps2025"
        })

        # --- EXPORT ZIP ---
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zf:
            zf.writestr(f"Google_Groups_{selected_school}.csv", final_gw.to_csv(index=False, encoding="utf-8-sig"))
            zf.writestr(f"Blackboard_Profils_{selected_school}.csv", blu_df.to_csv(index=False, encoding="utf-8-sig"))

        st.success(f"✅ Traitement terminé pour {len(valid_df)} lignes.")
        
        # Affichage des classes non mappées s'il y en a
        unmapped = valid_df[valid_df["Group Email"].isna()]["Classe"].unique()
        if len(unmapped) > 0:
            with st.expander("⚠️ Classes non trouvées dans le mapping"):
                st.write(list(unmapped))

        st.download_button(
            "📥 Télécharger les résultats (ZIP)",
            data=zip_buffer.getvalue(),
            file_name=f"Minatholy_{selected_school}.zip",
            mime="application/zip"
        )
