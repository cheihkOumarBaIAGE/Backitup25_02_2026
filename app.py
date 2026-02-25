# app.py
import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import datetime
from io import BytesIO
import zipfile

# -----------------------
# Page config & constants
# -----------------------
st.set_page_config(page_title="Hi", layout="wide", initial_sidebar_state="expanded")

SCHOOLS = ["INGENIEUR", "GRADUATE", "MANAGEMENT", "DROIT", "MADIBA"]
DATA_DIR = Path("data")

# -----------------------
# Mappings fournis (par école)
# -----------------------
# NOTE: tu m'as fourni des paires email->email ; je les place ici en majuscules côté clé
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
        "LG1-AGRO-R2": "lg1-agro2025-2026@ism.edu.sn",
        "LG1-ADMINISTRATIONDESAFFAIRES-R2": "lg1-adm-affaires-2r2025-2026@ism.edu.sn",
        "LG1-MI-R2": "lg1-mi-2r2025-2026@ism.edu.sn",
        "LG1-ORGA-GRH-R2": "lg1j-orga-grh-2r2025-2026@ism.edu.sn",
        "LG1-CIM": "lg1-cim2025-2026@ism.edu.sn",
        "LG1-CIMB": "lg1b-cim2025-2026@ism.edu.sn",
        "LG1D-MI": "lg1d-mi2025-2026@ism.edu.sn",
        "LG1E-MI": "lg1e-mi2025-2026@ism.edu.sn",
        "LG1-CF-R2": "lg1-2r2025-2026@ism.edu.sn",
        "LG1-CF-R2B": "lg1-2r2025-2026@ism.edu.sn",
        "LG1-CIM-R2": "lg1-2r2025-2026@ism.edu.sn",
        "LG1-MARKETINGETCOMMUNICATION-R2": "lg1-2r2025-2026@ism.edu.sn",
        "LG1-QHSE-R2": "lg1-2r2025-2026@ism.edu.sn",
        "LG1G-CF": "lg1g-cf2025-2026@ism.edu.sn",
        "LG1H-CF": "lg1h-cf2025-2026@ism.edu.sn",
        "LG1I-CF": "lg1i-cf2025-2026@ism.edu.sn",
        "LG1J-ORGA-GRH": "lg1j-orga-grh2025-2026@ism.edu.sn",
        "LG1J-ORGA-GRHB": "lg1j-orga-grh-b2025-2026@ism.edu.sn",
        "LG1K-CF": "lg1k-cf2025-2026@ism.edu.sn",
        "LG1L-CF": "lg1l-cf2025-2026@ism.edu.sn",
        "LG1-MARKETINGETCOMMUNICATION": "lg1-marketing-communication2025-2026@ism.edu.sn",
        "LG1-MIFULLENGLISH": "lg1-mi-full-english2025-2026@ism.edu.sn",
        "LG1-MIFULLENGLISH-R2": "lg1-mi-full-english-2r2025-2026@ism.edu.sn",
        "LG1-QHSE": "lg1-qhse2025-2026@ism.edu.sn",
        "LG1-R2": "lg1-2r2025-2026@ism.edu.sn",
        "LG2-CIMB": "lg2b-cim2025-2026@ism.edu.sn",
        "LG2-AGRO": "lg2-agro2025-2026@ism.edu.sn",
        "LG2B-ORGA-GRH": "lg2b-orga-grh2025-2026@ism.edu.sn",
        "LG2-CIM": "lg2-cim2025-2026@ism.edu.sn",
        "LG2D-MI": "lg2d-mi2025-2026@ism.edu.sn",
        "LG2E-MI": "lg2e-mi2025-2026@ism.edu.sn",
        "LG2G-CF": "lg2g-cf2025-2026@ism.edu.sn",
        "LG2H-CF": "lg2h-cf2025-2026@ism.edu.sn",
        "LG2I-CF": "lg2i-cf2025-2026@ism.edu.sn",
        "LG2J-ORGA-GRH": "lg2j-orga-grh2025-2026@ism.edu.sn",
        "LG2K-CF": "lg2k-cf2025-2026@ism.edu.sn",
        "LG2L-CF": "lg2l-cf2025-2026@ism.edu.sn",
        "LG2-MARKETINGETCOMMUNICATION": "lg2-marketing-communication2025-2026@ism.edu.sn",
        "LG2-MIFULLENGLISH": "lg2-mi-full-english2025-2026@ism.edu.sn",
        "LG2-QHSE": "lg2-qhse2025-2026@ism.edu.sn",
        "LG2-ADMINISTRATIONDESAFFAIRES": "lg2-adm-affaires2025-2026@ism.edu.sn",
        "LG2-ADMINISTRATIONDESAFFAIRESB": "lg2-adm-affairesb2025-2026@ism.edu.sn",
        "LG3-AGRO": "lg3-agro2025-2026@ism.edu.sn",
        "LG3-MCI": "lg3-cim2025-2026@ism.edu.sn",
        "LG3-MCIB": "lg3-cimb2025-2026@ism.edu.sn",
        "LG3D-MI": "lg3d-mi2025-2026@ism.edu.sn",
        "LG3E-MI": "lg3e-mi2025-2026@ism.edu.sn",
        "LG3G-CF": "lg3g-cf2025-2026@ism.edu.sn",
        "LG3H-CF": "lg3h-cf2025-2026@ism.edu.sn",
        "LG3I-CF": "lg3i-cf2025-2026@ism.edu.sn",
        "LG3K-CF": "lg3k-cf2025-2026@ism.edu.sn",
        "LG3L-CF": "lg3l-cf2025-2026@ism.edu.sn",
        "LG3-MARKETINGETCOMMUNICATION": "lg3-marketing-communication2025-2026@ism.edu.sn",
        "LG3-MIFULLENGLISH": "lg3-mi-full-english2025-2026@ism.edu.sn",
        "LG3-ORGA-GRH": "lg3-orga-grh2025-2026@ism.edu.sn",
        "LG3-ORGA-GRHB": "lg3b-orga-grh2025-2026@ism.edu.sn",
    },
    "GRADUATE": {
        "EMBA": "emba-2526@ism.edu.sn",
        "MASTER1-BUSINESSADMINISTRATION": "mba1business.administration-2526@ism.edu.sn",
        "MASTER1-BUSINESSADMINISTRATIONSOIR": "mba1business.administrationsoir-2526@ism.edu.sn",
        "LG3-ADMINISTRATIONDESAFFAIRESSOIR": "lg3-adm-affairessoir-2526@ism.edu.sn",
        "MASTER1-QHSE": "mba1qhse-2526@ism.edu.sn",
        "MASTER1-ACG": "mba1acg-2526@ism.edu.sn",
        "MASTER1-BANQUE-ASSURANCE": "mba1banque.assurance-2526@ism.edu.sn",
        "MASTER1-CCE": "mba1cce-2526@ism.edu.sn",
        "MASTER1-FINANCEDIGITALE(FINTECH)": "mba1finance-digitale-2526@ism.edu.sn",
        "MASTER1-FINANCEDEMARCHEETTRADING": "mba1finance-de-marche-2526@ism.edu.sn",
        "MASTER1-GCL": "mba1gcl-2526@ism.edu.sn",
        "MASTER1-IF": "mba1if-2526@ism.edu.sn",
        "MASTER1-INTERNATIONALMANAGEMENT": "mba1international.management-2526@ism.edu.sn",
        "MASTER1-MAA": "mba1maa-2526@ism.edu.sn",
        "MASTER1-MANAGEMENTAGROBUSINESS": "mba1management.agrobusiness-2526@ism.edu.sn",
        "MASTER1-MANAGEMENTAGROBUSINESSSOIR": "mba1management.agrobusinesssoir-2526@ism.edu.sn",
        "MASTER1-MANAGEMENT,VENTEETRELATIONCLIENT": "mba1mvente.rc-2526@ism.edu.sn",
        "MASTER1-MRH": "mba1mrh-2526@ism.edu.sn",
        "MASTER1-MARCHEFINANCIERETTRADING": "mba1mft-2526@ism.edu.sn",
        "MASTER1-RSEETDEVELOPPEMENTDURABLE": "mba1rse.dd-2526@ism.edu.sn",
        "MASTER1-QHSESOIR": "mba1qhsesoir-2526@ism.edu.sn",
        "MASTER1-ACGSOIR": "mba1acgsoir-2526@ism.edu.sn",
        "MASTER1-BANQUE-ASSURANCESOIR": "mba1banque.assurancesoir-2526@ism.edu.sn",
        "MASTER1-CCESOIR": "mba1ccesoir-2526@ism.edu.sn",
        "MASTER1-GCLSOIR": "mba1gclsoir-2526@ism.edu.sn",
        "MASTER1-IFSOIR": "mba1ifsoir-2526@ism.edu.sn",
        "MASTER1-MAASOIR": "mba1maasoir-2526@ism.edu.sn",
        "MASTER1-MANAGEMENT,VENTEETRELATIONCLIENTSOIR": "mba1mvente.rcsoir-2526@ism.edu.sn",
        "MASTER1-MRHSOIR": "mba1mrhsoir-2526@ism.edu.sn",
        "BP3-PRO": "bp3-2526@ism.edu.sn",
        "MASTER1-MANAGEMENTDESENERGIESPÉTROLIÈRESETGAZIÈRES": "mba1management.pet.gaz-2526@ism.edu.sn",
        "MASTER2-QHSE": "mba2qhse-2526@ism.edu.sn",
        "MASTER2-ACG": "mba2acg-2526@ism.edu.sn",
        "MASTER2-ACGSOIR": "mba2acgsoir-2526@ism.edu.sn",
        "MASTER2-BANQUE-ASSURANCE": "mba2banque.assurance-2526@ism.edu.sn",
        "MASTER2-BANQUE-ASSURANCESOIR": "mba2banque.assurancesoir-2526@ism.edu.sn",
        "MASTER2-CCE": "mba2cce-2526@ism.edu.sn",
        "MASTER2-CCESOIR": "mba2ccesoir-2526@ism.edu.sn", 
        "MASTER2-GCLSOIR": "mba2gclsoir-2526@ism.edu.sn",
        "MASTER2-GCL": "mba2gcl-2526@ism.edu.sn",
        "MASTER2-IF": "mba2if-2526@ism.edu.sn",
        "MASTER2-INTERNATIONALMANAGEMENT": "mba2international.management-2526@ism.edu.sn",
        "MASTER2-MAA": "mba2maa-2526@ism.edu.sn",
        "MASTER2-MANAGEMENTAGROBUSINESS": "mba2management.agrobusiness-2526@ism.edu.sn",
        "MASTER2-MANAGEMENT,VENTEETRELATIONCLIENT": "mba2mvente.rc-2526@ism.edu.sn",
        "MASTER2-MANAGEMENT,VENTEETRELATIONCLIENTSOIR": "mba2mvente.rc-2526@ism.edu.sn",
        "MASTER2-MRH": "mba2mrh-2526@ism.edu.sn",
        "MASTER2-BUSINESSADMINISTRATION": "mba2business.administration-2526@ism.edu.sn",
        "MASTER2-BUSINESSADMINISTRATIONSOIR": "mba2business.administrationsoir-2526@ism.edu.sn",
        "MASTER2-QHSESOIR": "mba2qhsesoir-2526@ism.edu.sn",
        "MBA2-ACGSoir": "mba2acgsoir-2526@ism.edu.sn",
        "MASTER2-IFSOIR": "mba2ifsoir-2526@ism.edu.sn",
        "MASTER2-MRHSOIR": "mba2mrhsoir-2526@ism.edu.sn",
        "MASTER2-FINANCEDIGITALE(FINTECH)": "mba2finance-digitale-2526@ism.edu.sn",
        "MASTER2-MANAGEMENTDESENERGIESPÉTROLIÈRESETGAZIÈRES": "mba2management.pet.gaz-2526@ism.edu.sn",
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
        "L1INGENIEURS-R2A": "l1r2.ingenieura@ism.edu.sn",
        "L1INGENIEURS-R2B": "l1r2.ingenieurb@ism.edu.sn",
        "L1-CPD": "l1cpd-2526@ism.edu.sn",
        "L1A-IA(IAGE&TTL)": "l1aia-2526@ism.edu.sn",
        "L1B-IA(IAGE&TTL)": "l1bia-2526@ism.edu.sn",
        "L1C-IA(GLRS&ETSE)": "l1cia-2526@ism.edu.sn",
        "L1-CDSD": "l1cdsd-2526@ism.edu.sn",
        "L1-CYBERSÉCURITÉ": "l1ia-cyber-2526@ism.edu.sn",
        "L1D-IA(MAIE&MOSIEF)": "l1dia-2526@ism.edu.sn",
        "L1E-IA(GLRS&ETSE)": "l1eia-2526@ism.edu.sn",
        "L1F-IA(IAGE&TTLC)": "l1fia-2526@ism.edu.sn",
        "L1G-IA(IAGE&TTL)": "l1gia-2526@ism.edu.sn",
        "L1H-IA(IAGE&TTL)": "l1hia-2526@ism.edu.sn",
        "L1-INTELLIGENCEARTIFICIELLE": "l1ia-cyber-2526@ism.edu.sn",
        "L2-INTELLIGENCEARTIFICIELLE": "licence2ia2025-2026@ism.edu.sn",
        "L2-CYBERSÉCURITÉ": "licence2cyber2025-2026@ism.edu.sn",
        "L2-CDSD": "licence2cdsd2025-2026@ism.edu.sn",
        "L2-CPD": "licence2cpd2025-2026@ism.edu.sn",
        "L2-ETSE": "licence2etse2025-2026@ism.edu.sn",
        "L2-GLRSB": "licence2glrsb2025-2026@ism.edu.sn",
        "L2-GLRSA": "licence2glrsa2025-2026@ism.edu.sn",
        "L2-IAGEA": "licence2iagea2025-2026@ism.edu.sn",
        "L2-IAGEB": "licence2iageb2025-2026@ism.edu.sn",
        "L2-MAIE": "licence2maie2025-2026@ism.edu.sn",
        "L2-MOSIEF": "licence2mosief2025-2026@ism.edu.sn",
        "L2-TTLB": "licence2ttlb2025-2026@ism.edu.sn",
        "L2-TTLA": "licence2ttla2025-2026@ism.edu.sn",
        "L3-CDSD": "licence3cdsd2025-2026@ism.edu.sn",
        "L3-CPD": "licence3cpd2025-2026@ism.edu.sn",
        "L3-ETSE": "licence3etse2025-2026@ism.edu.sn",
        "L3-GLRSA": "licence3glrs2025-2026@ism.edu.sn",
        "L3-GLRSB": "licence3glrs2025-2026@ism.edu.sn",
        "L3-IAGEA": "licence3iage2025-2026@ism.edu.sn",
        "L3-IAGEB": "licence3iage2025-2026@ism.edu.sn",
        "L3-MAIE": "licence3maie2025-2026@ism.edu.sn",
        "L3-MOSIEF": "licence3mosief2025-2026@ism.edu.sn",
        "L3-TTLA": "licence3ttl2025-2026@ism.edu.sn",
        "L3-TTLB": "licence3ttl2025-2026@ism.edu.sn",
        "M1IDC-BIGDATA&DATASTRATÉGIE": "mba1bigdata-2526@ism.edu.sn",
        "M1IDC-MARKETINGDIGITAL&BRANDCONTENT": "mba1marketingdigital-2526@ism.edu.sn",
        "M1IDC-UXDESIGN": "mba1ux-2526@ism.edu.sn",
        "M1IRSD": "mba1info-2526@ism.edu.sn",
        "M1PROJETS": "mba1projet-2526@ism.edu.sn",
        "M1PROJETSSOIR": "mba1projetsoir-2526@ism.edu.sn",
        "M1-MSSI": "mba1info-2526@ism.edu.sn",
        "MBA1-FINTECH": "mba1fintech-2526@ism.edu.sn",
        "MBA1-DATA-INTELLIGENCEARTIFICIELLE": "mba1ia-2526@ism.edu.sn",
        "MBA1-ActuariatBigDataetAssuranceQuantitative": "mba1actuariat-2526@ism.edu.sn",
        "MBA1-CDSD": "mba1cdsd-2526@ism.edu.sn",
        "MBA2DATA-INTELLIGENCEARTIFICIELLE": "mba2data_ia-2526@ism.edu.sn",
        "M2BIGDATA&DATASTRATÉGIE": "mba2bigdata-2526@ism.edu.sn",
        "M2MARKETINGDIGITAL&BRANDCONTENT": "mba2marketingdigital-2526@ism.edu.sn",
        "M2MSSI": "mba2mssi-2526@ism.edu.sn",
        "M2UXDESIGN": "mba2uxdesign-2526@ism.edu.sn",
        "MASTER2-MANAGEMENTDEPROJETS": "mba2projetga-2526@ism.edu.sn",
        "MBA2IRSD": "mba2irsd-2526@ism.edu.sn",
        "MBA2-ACTUARIATBIGDATAETASSURANCEQUANTITATIVE": "mba2actuariat-2526@ism.edu.sn",
        "MBA2-CDSD": "mba2cdsd-2526@ism.edu.sn",
        "MBA2-MANAGEMENTDEPROJETSINTERNATIONAUX": "mba2projetinternationaux-2526@ism.edu.sn",
    }
}

# -----------------------
# Helpers
# -----------------------
def read_emails_txt(path: Path):
    if not path.exists():
        return []
    with path.open("r", encoding="utf-8") as f:
        return [line.strip() for line in f if line.strip()]

def read_cours_mapping(cours_dir: Path):
    mapping = {}
    if not cours_dir.exists():
        return mapping
    for txt_file in sorted(cours_dir.glob("*.txt")):
        classe_name = txt_file.stem.upper()
        with txt_file.open("r", encoding="utf-8") as f:
            codes = [line.strip() for line in f if line.strip()]
        mapping[classe_name] = codes
    return mapping

def parse_mapping_textarea(text):
    out = {}
    for line in text.splitlines():
        if not line.strip():
            continue
        if "," in line:
            a, b = line.split(",", 1)
            out[a.strip().upper().replace(" ", "")] = b.strip()
    return out

def read_mapping_csv(mapping_csv: Path):
    if not mapping_csv.exists():
        return {}
    dfm = pd.read_csv(mapping_csv, dtype=str).fillna("")
    cols = [c.strip() for c in dfm.columns]
    c1 = cols[0]
    c2 = cols[1] if len(cols) > 1 else cols[0]
    return dict(zip(dfm[c1].astype(str).str.upper().str.replace(r"\s+","",regex=True), dfm[c2].astype(str).str.strip()))

def remove_accents_and_cleanup(s: str):
    if not isinstance(s, str): return s
    accent_map = str.maketrans({
        "à": "a", "â": "a", "ä": "a", "á": "a", "ã": "a", "ȧ": "a",
        "é": "e", "è": "e", "ê": "e", "ë": "e",
        "ì": "i", "î": "i", "ï": "i", "í": "i",
        "ò": "o", "ô": "o", "ö": "o", "ó": "o", "õ": "o", "ȯ": "o",
        "ù": "u", "û": "u", "ü": "u", "ú": "u",
        "ÿ": "y", "ý": "y"
    })
    return s.translate(accent_map)

def normalize_and_clean_df(df: pd.DataFrame):
    df.columns = [col.strip().capitalize() for col in df.columns]
    required_cols = {"Classe", "E-mail", "Nom", "Prénom"}
    if not required_cols.issubset(df.columns):
        missing = required_cols - set(df.columns)
        raise ValueError(f"Colonnes manquantes : {missing} (attendues: Classe, E-mail, Nom, Prénom)")
    df = df[["Classe", "E-mail", "Nom", "Prénom"]].copy()
    df.columns = ["Classroom Name", "Member Email", "Nom", "Prénom"]

    df["Classroom Name"] = df["Classroom Name"].fillna("").astype(str).str.replace(r"\s+","",regex=True).str.upper()
    df["Member Email"] = df["Member Email"].fillna("").astype(str).str.strip().str.replace(r"\s+","",regex=True)
    df["Nom"] = df["Nom"].fillna("").astype(str).str.strip().str.replace(r"\s+"," ",regex=True).apply(remove_accents_and_cleanup)
    df["Prénom"] = df["Prénom"].fillna("").astype(str).str.strip().str.replace(r"\s+"," ",regex=True).apply(remove_accents_and_cleanup)

    return df

def process_dataframe(df: pd.DataFrame, classroom_email_mapping: dict):
    df = normalize_and_clean_df(df)
    valid_emails_df = df[df["Member Email"].str.endswith("@ism.edu.sn", na=False)].copy()
    invalid_emails_df = df[~df["Member Email"].str.endswith("@ism.edu.sn", na=False)].copy()

    valid_emails_df["Group Email [Required]"] = valid_emails_df["Classroom Name"].map(classroom_email_mapping)
    mapped_df = valid_emails_df.dropna(subset=["Group Email [Required]"]).copy()
    unmapped_df = valid_emails_df[valid_emails_df["Group Email [Required]"].isna()].copy()

    mapped_export_df = mapped_df[["Group Email [Required]", "Member Email"]].copy()
    mapped_export_df["Member Type"] = "USER"
    mapped_export_df["Member Role"] = "MEMBER"

    profile_df = valid_emails_df.copy()
    profile_df["Nom d'utilisateur"] = profile_df["Member Email"]
    profile_df["Adresse e-mail"] = profile_df["Member Email"]
    profile_df["Nom"] = profile_df["Nom"]
    profile_df["Prénom"] = "\"" + profile_df["Prénom"] + "\""
    profile_df["Nouveau mot de passe"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
    profile_export_df = profile_df[["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]]

    return {
        "mapped_export_df": mapped_export_df,
        "mapped_df": mapped_df,
        "unmapped_df": unmapped_df,
        "invalid_emails_df": invalid_emails_df,
        "profile_export_df": profile_export_df
    }

def df_to_bytes(df_obj: pd.DataFrame, index=False, header=True, encoding="utf-8-sig"):
    b = BytesIO()
    df_obj.to_csv(b, index=index, header=header, encoding=encoding)
    b.seek(0)
    return b

# -----------------------
# UI (header / sidebar)
# -----------------------
header_col1, header_col2 = st.columns([1, 4])
with header_col1:
    st.image("https://commons.wikimedia.org/wiki/File:Apollo-kop,_objectnr_A_12979.jpg", width=64)
with header_col2:
    st.title("Minatholy")
    st.markdown("Génère les exports (listes de diffusion, créations et inscriptions de profils sur BLU) à partir d'un export de liste d'éléves. Choisis l'école, upload le fichier, télécharge les fichiers.")

st.markdown("---")

with st.sidebar:
    st.header("Paramètres")
    selected_school = st.selectbox("Choisir l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer un ZIP contenant tous les fichiers", value=True)
    st.markdown("---")
    st.caption("Les fichiers internes doivent être dans data/<ECOLE>/ (emails.txt + CoursParClasse/).")

# -----------------------
# Uploads & options
# -----------------------
# --- UPLOAD EXCEL ---
st.subheader(f" Upload du fichier Excel pour : **{selected_school}**")

uploaded_excel = st.file_uploader(
    "Importer le fichier Excel (.xls/.xlsx)",
    type=["xls", "xlsx"]
)

st.markdown("---")

# --- OPTIONS AVANCÉES ---
show_adv = st.checkbox("Afficher options avancées")
if show_adv:
    st.info("Règle par défaut : un email est valide s'il se termine par '@ism.edu.sn'")
    st.write("Mapping utilisé pour cette école :")
    st.json(SCHOOL_MAPPINGS[selected_school])

# --- BOUTON DE TRAITEMENT ---
run = st.button("🚀 Lancer le traitement", type="primary")
# -----------------------
# Processing
# -----------------------
if run:
    if not uploaded_excel:
        st.error("Veuillez uploader un fichier Excel avant de lancer le traitement.")
        st.stop()

    school_dir = DATA_DIR / selected_school
    emails_path = school_dir / "emails.txt"
    cours_dir = school_dir / "CoursParClasse"
    mapping_csv_path = school_dir / "mapping.csv"

    # mapping priority: mapping.csv in repo > uploaded csv > textarea > built-in
    mapping = {}
    if mapping_csv_path.exists():
        try:
            mapping = read_mapping_csv(mapping_csv_path)
            st.success(f"Mapping chargé depuis {mapping_csv_path} ({len(mapping)} entrées).")
        except Exception as e:
            st.warning(f"Impossible de lire mapping.csv : {e}")
    else:
        mapping = {k.upper().replace(" ", ""): v for k, v in SCHOOL_MAPPINGS[selected_school].items()}
        st.info(f"Mapping par défaut chargé depuis le code ({len(mapping)} entrées).")

    # admins
    admins = read_emails_txt(emails_path)
    if admins:
        st.success(f"{len(admins)} admin(s) lus depuis {emails_path}.")
    else:
        st.warning(f"Aucun emails.txt trouvé dans {school_dir} — le fichier admins sera vide.")

    # cours mapping
    classroom_course_mapping = read_cours_mapping(cours_dir)
    if classroom_course_mapping:
        st.success(f"{len(classroom_course_mapping)} fichier(s) de cours chargés depuis {cours_dir}.")
    else:
        st.warning(f"Aucun fichier de cours trouvé dans {cours_dir}. Les classes seront considérées sans codes.")

    # read excel
    try:
        df = pd.read_excel(uploaded_excel, dtype=str)
    except Exception as e:
        st.error(f"Erreur lecture Excel : {e}")
        st.stop()

    # process
    try:
        with st.spinner("Traitement en cours..."):
            results = process_dataframe(df, mapping)
    except Exception as e:
        st.error(f"Erreur lors du traitement : {e}")
        st.stop()

    mapped_export_df = results["mapped_export_df"]
    mapped_df = results["mapped_df"]
    unmapped_df = results["unmapped_df"]
    invalid_emails_df = results["invalid_emails_df"]
    profile_export_df = results["profile_export_df"]

    # admin rows (one per group email from mapping)
    admin_rows = []
    for group_email in set(mapping.values()):
        for admin_email in admins:
            admin_rows.append({
                "Group Email [Required]": group_email,
                "Member Email": admin_email,
                "Member Type": "USER",
                "Member Role": "MANAGER"
            })
    admin_df = pd.DataFrame(admin_rows)

    # combined
    combined = pd.concat([mapped_export_df, admin_df], ignore_index=True) if not mapped_export_df.empty else admin_df

    # course inscriptions
    course_rows = []
    classes_sans_code = set()
    for classe, group in mapped_df.groupby("Classroom Name"):
        emails = group["Member Email"].dropna().str.strip().unique()
        codes = classroom_course_mapping.get(classe, [])
        if not codes:
            classes_sans_code.add(classe)
        for email in emails:
            for code in codes:
                course_rows.append([code, email, "", ""])
    course_df = pd.DataFrame(course_rows)

    # report
    now_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    report_lines = []
    report_lines.append(f"Report — {selected_school} — {now_str}\n")
    report_lines.append(f"Mapped classes: {mapped_df['Classroom Name'].nunique()}")
    mapped_summary = mapped_df.drop_duplicates(subset=["Classroom Name", "Group Email [Required]"])
    for _, r in mapped_summary.iterrows():
        report_lines.append(f"- {r['Classroom Name']} -> {r['Group Email [Required]']}")
    report_lines.append(f"\nUnmapped classes ({unmapped_df['Classroom Name'].nunique()}):")
    for c in sorted(unmapped_df["Classroom Name"].dropna().unique()):
        report_lines.append(f"- {c}")
    report_lines.append(f"\nInvalid emails ({len(invalid_emails_df)}):")
    for e in invalid_emails_df["Member Email"].dropna():
        report_lines.append(f"- {e}")
    report_lines.append("\nSummary counts:")
    report_lines.append(f"- Utilisateurs mappés: {len(mapped_df)}")
    report_lines.append(f"- Utilisateurs non mappés: {len(unmapped_df)}")
    report_lines.append(f"- Emails ignorés: {len(invalid_emails_df)}")
    report_lines.append(f"- Classes sans codes: {len(classes_sans_code)}")
    if classes_sans_code:
        report_lines.append("\nClasses sans codes:")
        for c in sorted(classes_sans_code):
            report_lines.append(f"- {c}")
    report_text = "\n".join(report_lines)

    # bytes
    def make_bytes(obj):
        b = BytesIO()
        if isinstance(obj, pd.DataFrame):
            obj.to_csv(b, index=False, header=True, encoding="utf-8-sig")
        else:
            b.write(str(obj).encode("utf-8"))
        b.seek(0)
        return b

    fn_mise = f"mise_a_jour_liste_de_diffusion_GW_{selected_school}_{now_str}.csv"
    fn_admin = f"ajouter_membres_admin_GW_{selected_school}_{now_str}.csv"
    fn_profils = f"creation_profils_BLU_{selected_school}_{now_str}.csv"
    fn_courses = f"inscription_au_cours_en_ligne_BLU_{selected_school}_{now_str}.csv"
    fn_report = f"rapport_du_script_{selected_school}_{now_str}.txt"

    bytes_mise = make_bytes(combined) if not combined.empty else make_bytes(pd.DataFrame(columns=["Group Email [Required]","Member Email","Member Type","Member Role"]))
    bytes_admin = make_bytes(admin_df)
    bytes_profils = make_bytes(profile_export_df) if not profile_export_df.empty else make_bytes(pd.DataFrame(columns=["Nom d'utilisateur","Nom","Prénom","Adresse e-mail","Nouveau mot de passe"]))
    # courses: no header
    b_courses = BytesIO()
    if not course_df.empty:
        course_df.to_csv(b_courses, index=False, header=False, encoding="utf-8-sig")
    b_courses.seek(0)
    b_report = BytesIO(report_text.encode("utf-8"))
    b_report.seek(0)

    # downloads
    st.success("✅ Traitement terminé")
    st.markdown("### Fichiers générés — Téléchargements")
    c1, c2 = st.columns(2)
    with c1:
        st.download_button("Télécharger → mise_a_jour_liste_de_diffusion", bytes_mise, file_name=fn_mise, mime="text/csv")
        st.download_button("Télécharger → ajouter_membres_admin", bytes_admin, file_name=fn_admin, mime="text/csv")
        st.download_button("Télécharger → creation_profils_blackboard", bytes_profils, file_name=fn_profils, mime="text/csv")
    with c2:
        st.download_button("Télécharger → inscription_au_cours_en_ligne", b_courses, file_name=fn_courses, mime="text/csv")
        st.download_button("Télécharger → rapport (TXT)", b_report, file_name=fn_report, mime="text/plain")

    if zip_opt:
        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr(fn_mise, bytes_mise.getvalue())
            zf.writestr(fn_admin, bytes_admin.getvalue())
            zf.writestr(fn_profils, bytes_profils.getvalue())
            zf.writestr(fn_courses, b_courses.getvalue())
            zf.writestr(fn_report, report_text)
        zip_buffer.seek(0)
        st.download_button("📦 Télécharger tout en ZIP", zip_buffer, file_name=f"export_{selected_school}_{now_str}.zip", mime="application/zip")

    st.markdown("### Aperçu rapide")
    if not combined.empty:
        st.dataframe(combined.head(200))
    if not course_df.empty:
        st.markdown("Extrait des inscriptions aux cours")
        st.dataframe(course_df.head(200))

    st.info("Tu peux corriger les mappings (upload CSV ou zone texte) puis relancer si besoin.")
