import streamlit as st
import pandas as pd
import io
import zipfile
from datetime import datetime
import os

# --- 1. CONFIGURATION & SÉCURITÉ ---
st.set_page_config(page_title="Minatholy V3.0", layout="wide", page_icon="🛡️")

try:
    # On récupère tout depuis secrets.toml
    SCHOOL_MAPPING = st.secrets["mappings"]
    BB_PASSWORD = st.secrets["general"]["password"]
except Exception as e:
    st.error("❌ Erreur : Le fichier secrets.toml est mal configuré.")
    st.stop()

# --- 2. LES FONCTIONS MÉTIER (TON ANCIEN LOGIQUE) ---

def clean_bb_csv(df):
    """Le correctif pour Blackboard : force les guillemets là où il faut."""
    output = io.StringIO()
    df.to_csv(output, index=False, header=False, sep=',', quoting=0)
    lines = output.getvalue().split('\n')
    cleaned = []
    for line in lines:
        if not line.strip(): continue
        parts = line.split(',')
        # Format: "ID","Prénom","Nom",Email,Password
        cleaned.append(f'"{parts[0]}","{parts[1]}","{parts[2]}",{parts[3]},{parts[4]}')
    return "\n".join(cleaned)

def mise_a_jour_mailing_lists(df, school_mapping):
    """Simule/Prépare la mise à jour des listes Google."""
    # Ici, ta logique de préparation des listes
    classes_touchees = df['Classe'].unique()
    return f"Mise à jour préparée pour {len(classes_touchees)} listes de l'école."

def inscription_cours_en_ligne(df):
    """Génère la logique d'inscription aux cours (Enrollment)."""
    # Ta logique habituelle d'inscription
    return "Mapping des cours effectué avec succès."

# --- 3. INTERFACE UTILISATEUR ---
st.title("🛡️ Minatholy V3.0")
st.caption("Gestion centralisée : Mailing Lists, Blackboard & Inscriptions")

tab1, tab2, tab3 = st.tabs(["🚀 Exécution", "📊 Rapport & Dashboard", "🔍 Mappings"])

with tab1:
    col1, col2 = st.columns([1, 2])
    
    with col1:
        st.subheader("Paramètres")
        school_choice = st.selectbox("Choisir l'école", list(SCHOOL_MAPPING.keys()))
        uploaded_file = st.file_uploader("Fichier Excel (XLSX ou XLS)", type=['xlsx', 'xls'])
        
        st.divider()
        st.write("🛠️ **Actions à inclure :**")
        do_bb = st.checkbox("Création Profils Blackboard", value=True)
        do_mail = st.checkbox("Mise à jour Mailing Lists", value=True)
        do_enroll = st.checkbox("Inscription aux cours", value=True)

    if uploaded_file and school_choice:
        # --- LECTURE ET NETTOYAGE (LE FIX) ---
        df = pd.read_excel(uploaded_file)
        # Nettoyage des colonnes (Espaces avant/après)
        df.columns = [str(c).strip() for c in df.columns]
        
        # Vérification des colonnes critiques
        if 'Classe' not in df.columns:
            st.error("❌ La colonne 'Classe' n'a pas été trouvée.")
            st.stop()
        
        id_col = next((c for c in df.columns if 'Identifiant' in c), None)

        if st.button("LANCER LE TRAITEMENT COMPLET"):
            with st.status("Exécution des tâches...", expanded=True) as status:
                mapping = SCHOOL_MAPPING[school_choice]
                zip_buffer = io.BytesIO()
                success_count = 0
                
                # --- LOGIQUE BLACKBOARD ---
                if do_bb:
                    status.write("⏳ Génération des fichiers Blackboard...")
                    with zipfile.ZipFile(zip_buffer, "a", zipfile.ZIP_DEFLATED) as zip_file:
                        for cls in df['Classe'].unique():
                            clean_cls = str(cls).strip()
                            if clean_cls in mapping:
                                sub_df = df[df['Classe'] == cls].copy()
                                bb_df = pd.DataFrame({
                                    'ID': sub_df[id_col],
                                    'First': sub_df['Prénom'],
                                    'Last': sub_df['Nom'],
                                    'Email': mapping[clean_cls],
                                    'Pass': BB_PASSWORD
                                })
                                zip_file.writestr(f"{clean_cls}.csv", clean_bb_csv(bb_df))
                                success_count += 1
                
                # --- LOGIQUE MAILING LISTS ---
                if do_mail:
                    status.write("⏳ Synchro des Mailing Lists...")
                    mail_status = mise_a_jour_mailing_lists(df, mapping)
                
                # --- LOGIQUE INSCRIPTION ---
                if do_enroll:
                    status.write("⏳ Inscription aux cours en ligne...")
                    enroll_status = inscription_cours_en_ligne(df)

                status.update(label="✅ Opérations terminées !", state="complete")

            # --- RÉSULTATS & TÉLÉCHARGEMENT ---
            st.success(f"Traitement terminé : {success_count} classes générées.")
            
            if do_bb:
                st.download_button(
                    "📥 Télécharger les fichiers Blackboard (ZIP)",
                    data=zip_buffer.getvalue(),
                    file_name=f"BB_{school_choice}_{datetime.now().strftime('%d%m')}.zip",
                    mime="application/zip"
                )
            
            # Sauvegarde pour le Dashboard
            log_entry = pd.DataFrame([{"Date": datetime.now(), "Ecole": school_choice, "Etudiants": len(df), "Classes": success_count}])
            log_entry.to_csv("history.csv", mode='a', header=not os.path.exists("history.csv"), index=False)

with tab2: # LE DASHBOARD
    st.subheader("📊 Rapport du Script")
    if os.path.exists("history.csv"):
        h_df = pd.read_csv("history.csv")
        c1, c2, c3 = st.columns(3)
        c1.metric("Total Étudiants", h_df["Etudiants"].sum())
        c2.metric("Total Fichiers", h_df["Classes"].sum())
        c3.metric("Dernier Run", h_df["Date"].iloc[-1][:16])
        
        st.line_chart(h_df.set_index("Date")["Etudiants"])
        st.dataframe(h_df.sort_values("Date", ascending=False), use_container_width=True)
    else:
        st.info("Aucun historique pour le moment.")

with tab3: # CONSULTATION MAPPINGS
    st.subheader("🔍 Vérification des Mappings")
    # Affichage des données de mapping de l'école sélectionnée
    flat_data = [{"Classe": c, "Email": e} for c, e in SCHOOL_MAPPING[school_choice].items()]
    st.table(pd.DataFrame(flat_data))
