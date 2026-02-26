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
    st.title("Urgence IT Settings")
    selected_school = st.selectbox("Sélectionner l'école", SCHOOLS)
    zip_opt = st.checkbox("Générer Archive ZIP (Automatique)", value=True)
    st.divider()
    st.caption("v2.6 - Added Manual Inscription Mode")

st.title("Urgence IT BASH")

# --- NAVIGATION TABS ---
main_tab, manual_tab = st.tabs(["BASH", "Mise à jour Urgente de cours"])

# -----------------------
# TAB 1: AUTOMATIC TREATMENT
# -----------------------
with main_tab:
    with st.container(border=True):
        col_up1, col_up2 = st.columns([3, 1])
        with col_up1:
            uploaded_excel = st.file_uploader("Importer le fichier Excel des inscriptions", type=["xls", "xlsx"])
        with col_up2:
            st.write("##")
            run = st.button("🚀 Lancer le traitement", type="primary", use_container_width=True)

    if run:
        if not uploaded_excel:
            st.error("❌ Aucun fichier détecté.")
        else:
            with st.status("Traitement en cours...", expanded=True) as status:
                school_dir = DATA_DIR / selected_school
                admins = read_emails_txt(school_dir / "emails.txt")
                cours_mapping = read_cours_mapping(school_dir / "CoursParClasse")
                
                mapping_csv = school_dir / "mapping.csv"
                if mapping_csv.exists():
                    dfm = pd.read_csv(mapping_csv).fillna("")
                    mapping = dict(zip(dfm.iloc[:,0].str.upper().str.replace(" ",""), dfm.iloc[:,1].str.strip()))
                else:
                    mapping = {k.upper().replace(" ", ""): v for k, v in SCHOOL_MAPPINGS.get(selected_school, {}).items()}

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
                    profiles["Prénom"] = "\"" + profiles["Prénom"] + "\""
                    profiles["Nouveau mot de passe"] = "ismapps2025,,,,,,,,,,,,,,,,,1382"
                    profile_export = profiles[["Member Email", "Nom", "Prénom", "Member Email", "Nouveau mot de passe"]]
                    profile_export.columns = ["Nom d'utilisateur", "Nom", "Prénom", "Adresse e-mail", "Nouveau mot de passe"]

                    # 4. Courses
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
                    
                    report += f"\nInvalid emails ({len(invalid_emails_list)}):\n"
                    for m in invalid_emails_list: report += f"- {m if m else '(Vide)'}\n"
                    
                    report += f"\nSummary counts:\n"
                    report += f"- Utilisateurs mappés: {len(mapped_df)}\n- Utilisateurs non mappés: {len(unmapped_students_df)}\n- Emails ignorés: {len(df) - len(valid_df)}\n- Classes sans codes: {len(classes_sans_codes)}\n"

                    bytes_mise = sanitize_csv_output(combined_google)
                    bytes_admin = sanitize_csv_output(admin_df)
                    bytes_profils = sanitize_csv_output(profile_export, is_profile=True)
                    bytes_courses = sanitize_csv_output(course_df, header=False)
                    b_report = BytesIO(report.encode("utf-8"))

                    status.update(label="✅ Traitement réussi !", state="complete", expanded=False)
                except Exception as e:
                    st.error(f"Erreur fatale : {e}")
                    st.stop()

            # --- DOWNLOADS ---
            res_tab1, res_tab2 = st.tabs(["📥 Téléchargements", "📊 Rapport"])
            with res_tab1:
                m1, m2, m3 = st.columns(3)
                m1.download_button("📧 Listes Google", bytes_mise, file_name=f"diffusion_{selected_school}.csv", use_container_width=True)
                m2.download_button("🔑 Admins Google", bytes_admin, file_name=f"admins_{selected_school}.csv", use_container_width=True)
                m3.download_button("👤 Profils Blackboard", bytes_profils, file_name=f"profils_blu_{selected_school}.csv", use_container_width=True)
                
                c1, c2 = st.columns(2)
                c1.download_button("📚 Inscriptions Cours", bytes_courses, file_name=f"cours_{selected_school}.csv", use_container_width=True)
                c2.download_button("📝 Rapport Complet", b_report, file_name=f"rapport_{selected_school}.txt", use_container_width=True)

                if zip_opt:
                    z_buf = BytesIO()
                    with zipfile.ZipFile(z_buf, "w") as zf:
                        zf.writestr("liste_diffusion.csv", bytes_mise.getvalue())
                        zf.writestr("admins.csv", bytes_admin.getvalue())
                        zf.writestr("profils_blackboard.csv", bytes_profils.getvalue())
                        zf.writestr("inscriptions_cours.csv", bytes_courses.getvalue())
                        zf.writestr("rapport.txt", report)
                    z_buf.seek(0)
                    st.download_button("📦 Télécharger ZIP", z_buf, file_name=f"export_{selected_school}.zip", type="primary", use_container_width=True)
            with res_tab2:
                st.text_area("Audit Log", report, height=300)

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
            manual_rows = []
            for email in emails_list:
                for code in codes_list:
                    manual_rows.append([code, email, "", ""])
            
            manual_df = pd.DataFrame(manual_rows)
            bytes_manual = sanitize_csv_output(manual_df, header=False)
            
            st.success(f"✅ Généré : {len(manual_rows)} inscriptions.")
            st.dataframe(manual_df.rename(columns={0: "Code", 1: "Email"}), use_container_width=True)
            st.download_button("📥 Télécharger inscriptions_manuelles.csv", bytes_manual, file_name="inscriptions_cours_manuel.csv", type="primary")
