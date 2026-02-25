import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import os

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="Minatholy V2.0", 
    layout="wide", 
    page_icon="🛡️"
)

# --- CHARGEMENT SÉCURISÉ DES SECRETS ---
try:
    # On récupère les données du fichier .streamlit/secrets.toml
    SCHOOL_MAPPING = st.secrets["mappings"]
    BB_PASSWORD = st.secrets["general"]["password"]
except Exception as e:
    st.error("❌ Erreur de configuration : Le fichier secrets.toml est manquant ou mal rempli.")
    st.info("Vérifie que tu as bien créé le dossier '.streamlit' et le fichier 'secrets.toml'.")
    st.stop()

# --- FONCTION DE NETTOYAGE (L'ARME SECRÈTE) ---
def clean_bb_csv(df):
    """Force le format Blackboard en nettoyant les guillemets après génération."""
    output = io.StringIO()
    # Génération standard sans guillemets automatiques (quoting=0)
    df.to_csv(output, index=False, header=False, sep=',', quoting=0)
    lines = output.getvalue().split('\n')
    
    cleaned_lines = []
    for line in lines:
        if not line.strip(): continue
        parts = line.split(',')
        # On entoure uniquement les 3 premières colonnes (ID, Prénom, Nom)
        # Le reste (Email, Password brut) reste tel quel
        p1, p2, p3 = f'"{parts[0]}"', f'"{parts[1]}"', f'"{parts[2]}"'
        rest = ",".join(parts[3:])
        cleaned_lines.append(f"{p1},{p2},{p3},{rest}")
    
    return "\n".join(cleaned_lines)

# --- INTERFACE PRINCIPALE ---
st.title("🛡️ Minatholy V2.0")
st.caption("Système de gestion des fichiers Blackboard - ISM Dakar")

tab_proc, tab_map, tab_hist = st.tabs([
    "📤 Traitement des fichiers", 
    "🔍 Mappings & Doublons", 
    "📊 Statistiques d'utilisation"
])

# --- ONGLET 1 : TRAITEMENT ---
with tab_proc:
    col_l, col_r = st.columns([1, 2])
    
    with col_l:
        st.subheader("Configuration")
        school_choice = st.selectbox("Sélectionner l'école", list(SCHOOL_MAPPING.keys()))
        uploaded_file = st.file_uploader("Charger le fichier Excel étudiant", type=['xlsx'])
    
    if uploaded_file and school_choice:
        with st.status("Traitement des données en cours...", expanded=True) as status:
            df = pd.read_excel(uploaded_file)
            mapping = SCHOOL_MAPPING[school_choice]
            zip_buffer = io.BytesIO()
            
            success_count = 0
            missing_classes = []

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                classes = df['Classe'].unique()
                
                for cls in classes:
                    if cls in mapping:
                        st.write(f"✅ Traitement : **{cls}**")
                        sub_df = df[df['Classe'] == cls].copy()
                        
                        # Structure Blackboard
                        bb_df = pd.DataFrame({
                            'ID': sub_df['Identifiant de l’utilisateur'],
                            'First': sub_df['Prénom'],
                            'Last': sub_df['Nom'],
                            'Email': mapping[cls],
                            'Pass': BB_PASSWORD
                        })
                        
                        # Application du fix de formatage
                        csv_data = clean_bb_csv(bb_df)
                        zip_file.writestr(f"{cls}.csv", csv_data)
                        success_count += 1
                    else:
                        missing_classes.append(cls)
            
            if missing_classes:
                st.warning(f"⚠️ {len(missing_classes)} classes non trouvées dans le mapping : {', '.join(missing_classes)}")
            
            status.update(label="Fichiers prêts !", state="complete")
            
            st.download_button(
                label=f"📥 Télécharger le ZIP ({success_count} fichiers)",
                data=zip_buffer.getvalue(),
                file_name=f"Blackboard_{school_choice}_{datetime.now().strftime('%d%m_%Hh%M')}.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            # Sauvegarde Log
            log_data = pd.DataFrame([{"Date": datetime.now().strftime("%Y-%m-%d %H:%M"), "Ecole": school_choice, "Classes": success_count, "Etudiants": len(df)}])
            log_data.to_csv("history.csv", mode='a', header=not os.path.exists("history.csv"), index=False)

# --- ONGLET 2 : MAPPINGS & DOUBLONS ---
with tab_map:
    st.subheader("Explorateur de Mappings")
    
    # Préparation des données pour l'affichage
    flat_map = []
    for sch, classes in SCHOOL_MAPPING.items():
        for c, e in classes.items():
            flat_map.append({"Ecole": sch, "Classe": c, "Email": e})
    map_df = pd.DataFrame(flat_map)

    # Recherche
    search = st.text_input("Rechercher une classe ou un email...").strip().upper()
    filtered_df = map_df[map_df.apply(lambda row: search in row.values.astype(str).flat, axis=1)] if search else map_df
    
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    # --- NOUVEAU : VÉRIFICATEUR DE DOUBLONS ---
    st.divider()
    if st.button("🔍 Vérifier les doublons d'emails"):
        duplicates = map_df[map_df.duplicated('Email', keep=False)]
        if not duplicates.empty:
            st.error("🚨 Alerte : Les emails suivants sont assignés à plusieurs classes !")
            st.dataframe(duplicates.sort_values('Email'), use_container_width=True)
        else:
            st.success("✨ Parfait ! Aucun doublon détecté dans tes mappings.")

# --- ONGLET 3 : HISTORIQUE ---
with tab_hist:
    if os.path.exists("history.csv"):
        hist = pd.read_csv("history.csv")
        
        c1, c2, c3
