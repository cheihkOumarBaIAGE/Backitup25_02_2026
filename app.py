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
    SCHOOL_MAPPING = st.secrets["mappings"]
    BB_PASSWORD = st.secrets["general"]["password"]
except Exception as e:
    st.error("❌ Erreur de configuration : Le fichier secrets.toml est manquant ou mal rempli.")
    st.stop()

# --- FONCTION DE NETTOYAGE BLACKBOARD ---
def clean_bb_csv(df):
    """Force le format Blackboard : guillemets sur les textes, brut sur le reste."""
    output = io.StringIO()
    df.to_csv(output, index=False, header=False, sep=',', quoting=0)
    lines = output.getvalue().split('\n')
    
    cleaned_lines = []
    for line in lines:
        if not line.strip(): continue
        parts = line.split(',')
        # On entoure d'un guillemet les colonnes Nom, Prénom, ID
        p1, p2, p3 = f'"{parts[0]}"', f'"{parts[1]}"', f'"{parts[2]}"'
        rest = ",".join(parts[3:])
        cleaned_lines.append(f"{p1},{p2},{p3},{rest}")
    
    return "\n".join(cleaned_lines)

# --- INTERFACE PRINCIPALE ---
st.title("🛡️ Minatholy V2.0")
st.caption("Système intelligent de gestion Blackboard - ISM Dakar")

tab_proc, tab_map, tab_hist = st.tabs([
    "📤 Traitement des fichiers", 
    "🔍 Mappings & Doublons", 
    "📊 Statistiques"
])

# --- ONGLET 1 : TRAITEMENT ---
with tab_proc:
    col_l, col_r = st.columns([1, 2])
    
    with col_l:
        st.subheader("Configuration")
        school_choice = st.selectbox("Sélectionner l'école", list(SCHOOL_MAPPING.keys()))
        # Acceptation des deux formats : XLSX et XLS
        uploaded_file = st.file_uploader("Charger le fichier Excel", type=['xlsx', 'xls'])
    
    if uploaded_file and school_choice:
        with st.status("Analyse et nettoyage du fichier...", expanded=True) as status:
            # Lecture du fichier
            df = pd.read_excel(uploaded_file)
            
            # NETTOYAGE DES ESPACES (Le correctif pour " Classe ")
            df.columns = [str(c).strip() for c in df.columns]
            
            # Vérification de la colonne Classe
            if 'Classe' not in df.columns:
                st.error(f"❌ Colonne 'Classe' introuvable. Colonnes détectées : {list(df.columns)}")
                st.stop()
            
            # Détection flexible de l'Identifiant (gestion des apostrophes)
            id_col = next((c for c in df.columns if 'Identifiant' in c), None)
            if not id_col:
                st.error("❌ Impossible de trouver la colonne 'Identifiant'.")
                st.stop()

            mapping = SCHOOL_MAPPING[school_choice]
            zip_buffer = io.BytesIO()
            success_count = 0
            missing_classes = []

            with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                classes = df['Classe'].unique()
                
                for cls in classes:
                    # On nettoie aussi le nom de la classe au cas où il y a des espaces dans les cellules
                    clean_cls = str(cls).strip()
                    if clean_cls in mapping:
                        st.write(f"✅ Génération : **{clean_cls}**")
                        sub_df = df[df['Classe'] == cls].copy()
                        
                        bb_df = pd.DataFrame({
                            'ID': sub_df[id_col],
                            'First': sub_df['Prénom'],
                            'Last': sub_df['Nom'],
                            'Email': mapping[clean_cls],
                            'Pass': BB_PASSWORD
                        })
                        
                        csv_data = clean_bb_csv(bb_df)
                        zip_file.writestr(f"{clean_cls}.csv", csv_data)
                        success_count += 1
                    else:
                        missing_classes.append(clean_cls)
            
            if missing_classes:
                st.warning(f"⚠️ {len(missing_classes)} classes non mappées : {', '.join(missing_classes)}")
            
            status.update(label="Traitement terminé !", state="complete")
            
            st.download_button(
                label=f"📥 Télécharger le ZIP ({success_count} fichiers)",
                data=zip_buffer.getvalue(),
                file_name=f"Blackboard_{school_choice}_{datetime.now().strftime('%d%m_%Hh%M')}.zip",
                mime="application/zip",
                use_container_width=True
            )
            
            # Historique
            log_data = pd.DataFrame([{"Date": datetime.now().strftime("%Y-%m-%d %H:%M"), "Ecole": school_choice, "Classes": success_count, "Etudiants": len(df)}])
            log_data.to_csv("history.csv", mode='a', header=not os.path.exists("history.csv"), index=False)

# --- ONGLET 2 : MAPPINGS & DOUBLONS ---
with tab_map:
    st.subheader("Explorateur de Mappings")
    flat_map = []
    for sch, classes in SCHOOL_MAPPING.items():
        for c, e in classes.items():
            flat_map.append({"Ecole": sch, "Classe": c, "Email": e})
    map_df = pd.DataFrame(flat_map)

    search = st.text_input("Rechercher une classe ou un email...").strip().upper()
    filtered_df = map_df[map_df.apply(lambda row: search in str(row.values).upper(), axis=1)] if search else map_df
    st.dataframe(filtered_df, use_container_width=True, hide_index=True)

    if st.button("🔍 Vérifier les doublons d'emails"):
        duplicates = map_df[map_df.duplicated('Email', keep=False)]
        if not duplicates.empty:
            st.error("🚨 Doublons détectés !")
            st.dataframe(duplicates.sort_values('Email'), use_container_width=True)
        else:
            st.success("✨ Aucun doublon dans les emails.")

# --- ONGLET 3 : HISTORIQUE ---
with tab_hist:
    if os.path.exists("history.csv"):
        hist = pd.read_csv("history.csv")
        st.metric("Total Étudiants Traités", hist["Etudiants"].sum())
        st.dataframe(hist.sort_values("Date", ascending=False), use_container_width=True)
        if st.button("🗑️ Effacer l'historique"):
            os.remove("history.csv")
            st.rerun()
    else:
        st.info("Aucun historique disponible.")
