"""
Dashboard Bancassurance NSIA - Saisie et suivi des objectifs, commissions
et profil partenaire.

Donnees source: TABLEAU DE BORD ET PAC PAR PARTENAIRE 2026 (2).xlsx
FICHIER SUIVI: FICHE DE SUIVI HEBDOMADAIRE DES REALISATIONS 2026 BANCASSURANCE (5) (13).xlsx

Formules:
  - GAP = Objectif(annee N) - Realisation(annee N-1)
  - Taux de realisation = Realisation(annee N) / Objectif(annee N) * 100
"""

import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils.dataframe import dataframe_to_rows
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
import io
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import Wedge
import matplotlib.colors as mcolors

from database_service import (
    BancassuranceService,
    PARTENAIRES,
    MOIS,
    format_fcfa,
    format_pct,
    get_mois_from_fiche,
    result_date_str,
    add_partenaire,
    get_all_partenaires,
)
from historical_service import HistoricalService


# ==========================================================================
# THÈME COULEURS SOMBRES (NSIA)
# ==========================================================================
# Couleurs définies selon les consignes utilisateur:
#   - Fond principal : bleu nuit #002B49
#   - Titres / éléments prioritaires : doré #D4A72C
#   - Bordures : gris argent #A7A9AC
#   - Textes principaux : blanc #FFFFFF
#   - Textes secondaires : gris très clair #F2F4F5
# Le thème Streamlit (config.toml) gère le fond principal, le doré et le blanc.
# Le CSS ci-dessous affine les bordures et textes secondaires.
DARK_THEME_CSS = """
<style>
/* Bordures grises argentées autour des inputs et cartes */
.stTextInput>div>div,
.stTextArea>div>div,
.stSelectbox>div>div,
.stNumberInput>div>div,
.stDateInput>div>div,
.stDataFrame,
.stDataFrame tbody tr td,
.stDataFrame tbody tr th,
.stDataFrame thead tr th,
.stExpander {
    border: 1px solid #A7A9AC !important;
}
/* Textes secondaires (métriques delta, captions, etc.) */
.css-10trbl2,
.stMetric label,
.stMarkdown caption,
.stCaption,
.secondary-text,
small {
    color: #F2F4F5 !important;
}
/* Titres et éléments prioritaires en doré */
h1, h2, h3, h4, h5, h6,
.stMarkdown h1, .stMarkdown h2, .stMarkdown h3,
.stSubheader,
.stHeader {
    color: #D4A72C !important;
    font-weight: 600 !important;
}
/* Liens et éléments interactifs */
.stButton>button,
.stTabs [data-baseweb="tab"] {
    background-color: #003a61 !important;
    color: #D4A72C !important;
}
/* Barres de défilement */
::-webkit-scrollbar-thumb {
    background: #A7A9AC !important;
}
</style>
"""
st.markdown(DARK_THEME_CSS, unsafe_allow_html=True)

# ==========================================================================

# ==========================================================================
# CONFIGURATION STREAMLIT
# ==========================================================================

st.set_page_config(
    page_title="Dashboard Bancassurance NSIA",
    page_icon="NSIA",
    layout="wide",
)

# ==========================================================================
# INITIALISATION
# ==========================================================================

# Charger le service de donnees
svc = BancassuranceService()
hist_svc = HistoricalService()

# Session state
if "partenaire_sel" not in st.session_state:
    st.session_state.partenaire_sel = PARTENAIRES[0]
if "dg_entries" not in st.session_state:
    st.session_state.dg_entries = {p: "" for p in PARTENAIRES}
if "recommandations_entries" not in st.session_state:
    st.session_state.recommandations_entries = {p: "" for p in PARTENAIRES}
if "objectifs_mensuels_override" not in st.session_state:
    st.session_state.objectifs_mensuels_override = {p: {m: None for m in MOIS} for p in PARTENAIRES}
if "realisations_mensuelles_override" not in st.session_state:
    st.session_state.realisations_mensuelles_override = {p: {m: None for m in MOIS} for p in PARTENAIRES}
if "commissions_override" not in st.session_state:
    st.session_state.commissions_override = {
        p: {"commission_reversee": None, "total_a_payer": None,
            "date_traitement": None, "date_paiement": None}
        for p in PARTENAIRES
    }


def get_effective_objectifs_mensuels(partner: str, annee: int = 2026) -> dict:
    """Retourne les objectifs mensuels depuis la base historique SQLite.

    Les objectifs sont saisis par l'utilisateur dans l'onglet
    "Historique & Comparaison" et stockes dans historical_bancassurance.db.
    Si aucune donnée n'existe pour l'annee demandee, retourne des zeros.
    """
    data = hist_svc.get_data_for_partner_annee(partner, annee)
    result = {}
    for m in MOIS:
        entry = data.get(m, {})
        obj = entry.get("objectif")
        result[m] = float(obj) if obj is not None else 0.0
    return result


def get_effective_realisations_mensuelles(partner: str, annee: int = 2026) -> dict:
    """Retourne les realisations mensuelles depuis la base historique SQLite.

    Les realisations sont saisies par l'utilisateur dans l'onglet
    "Historique & Comparaison" et stockées dans historical_bancassurance.db.
    Si aucune donnée n'existe pour l'annee demandee, retourne des zeros.
    """
    data = hist_svc.get_data_for_partner_annee(partner, annee)
    result = {}
    for m in MOIS:
        entry = data.get(m, {})
        real = entry.get("realisation")
        result[m] = float(real) if real is not None else 0.0
    return result


def get_effective_commissions(partner: str, annee: int = 2026) -> dict:
    """Retourne les commissions saisies par l'utilisateur depuis la base historique.

    Les commissions sont stockées comme des valeurs speciales dans la base
    historique (clé "commission_reversee", "total_a_payer", etc.) avec le
    mois = "ANNUEL". Si aucune donnée, retourne des valeurs par defaut.
    """
    data = hist_svc.get_data_for_partner_annee(partner, annee)
    annee_entry = data.get("ANNUEL", {})
    return {
        "commission_reversee": annee_entry.get("objectif") if annee_entry.get("objectif") else None,
        "total_a_payer": annee_entry.get("realisation") if annee_entry.get("realisation") else None,
        "date_traitement": None,
        "date_paiement": None,
    }


# EXPORT POWERPOINT
# ==========================================================================


# EXPORT POWERPOINT
# ==========================================================================


def export_presentation(partner: str) -> str:
    """
    Genere un fichier PowerPoint (.pptx) avec:
    - Page de garde
    - Slide GAP / Taux de realisation
    - Slide objectifs mensuels
    - Slide realisations cumulées & commissions
    - Slide profil partenaire (points forts/faibles, DG, responsables)
    - Slide etat des lieux (difficultes, ameliorations)
    - Slide recommandations
    """
    prs = Presentation()

    # --- Slide 1: Page de garde ---
    slide = prs.slides.add_slide(prs.slide_layouts[6])  # layout vide
    title_box = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(2))
    tf = title_box.text_frame
    p = tf.paragraphs[0]
    p.text = "Dashboard Bancassurance NSIA"
    p.font.size = Pt(32)
    p.font.bold = True
    p.alignment = PP_ALIGN.CENTER

    subtitle_box = slide.shapes.add_textbox(Inches(1), Inches(4), Inches(8), Inches(1))
    tf2 = subtitle_box.text_frame
    p2 = tf2.paragraphs[0]
    p2.text = f"Profil partenaire : {partner}"
    p2.font.size = Pt(18)
    p2.alignment = PP_ALIGN.CENTER

    date_box = slide.shapes.add_textbox(Inches(1), Inches(6), Inches(8), Inches(0.5))
    tf3 = date_box.text_frame
    p3 = tf3.paragraphs[0]
    p3.text = f"Date de genertion : {datetime.now().strftime('%d/%m/%Y')}"
    p3.font.size = Pt(12)
    p3.alignment = PP_ALIGN.CENTER

    # --- Slide 2: GAP et Taux ---
    profil = hist_svc.get_profil_partenaire(partner)
    gap_data = hist_svc.get_gap_taux_historical(partner, 2026)

    slide = prs.slides.add_slide(prs.slide_layouts[1])  # title + content
    slide.shapes.title.text = f"GAP & Taux de realisation - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Objectif 2026 (annee N) : {format_fcfa(gap_data['objectif_N'])}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Realisation 2025 (annee N-1) : {format_fcfa(gap_data['realisation_N_1'])}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Realisation 2026 : {format_fcfa(gap_data['realisation_N'])}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f""
    p = tf.add_paragraph()
    p.text = f"GAP = Objectif(N) - Realisation(N-1) = {format_fcfa(gap_data['gap'])}"
    p.font.size = Pt(14)
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = f"Taux de realisation = Realisation(N) / Objectif(N) = {format_pct(gap_data['taux'])}"
    p.font.size = Pt(14)

    # --- Slide 3: Objectifs mensuels ---
    obj_mensuels = get_effective_objectifs_mensuels(partner)
    real_mensuels = get_effective_realisations_mensuelles(partner)
    slide = prs.slides.add_slide(prs.slide_layouts[5])  # title only
    slide.shapes.title.text = f"Objectifs mensuels 2026 - {partner}"

    # --- Graphique 1: Barres comparatives Objectif vs Réalisation ---
    fig_bar, ax_bar = plt.subplots(figsize=(8, 5), dpi=150)
    x = range(len(MOIS))
    width = 0.35
    obj_vals = [obj_mensuels[m] for m in MOIS]
    real_vals = [real_mensuels[m] for m in MOIS]
    bars1 = ax_bar.bar([i - width/2 for i in x], obj_vals, width, label="Objectif", color="#D4A72C")
    bars2 = ax_bar.bar([i + width/2 for i in x], real_vals, width, label="Réalisation", color="#003a61")
    ax_bar.set_xlabel("Mois")
    ax_bar.set_ylabel("Montant (FCFA)")
    ax_bar.set_title("Objectif vs Réalisation mensuelle")
    ax_bar.set_xticks(list(x))
    ax_bar.set_xticklabels(MOIS, rotation=45, ha="right")
    ax_bar.legend()
    ax_bar.grid(axis="y", alpha=0.3)
    fig_bar.tight_layout()
    buf_bar = io.BytesIO()
    fig_bar.savefig(buf_bar, format="png", facecolor="white")
    buf_bar.seek(0)
    plt.close(fig_bar)
    slide.shapes.add_picture(buf_bar, Inches(0.3), Inches(4.2), width=Inches(9))

    # --- Graphique 2: Camembert parts de marché (Objectif réparti par mois) ---
    fig_pie, ax_pie = plt.subplots(figsize=(6, 6), dpi=150)
    # Part de chaque mois dans l'objectif annuel
    total_obj = sum(obj_vals)
    if total_obj > 0:
        sizes = [v / total_obj * 100 for v in obj_vals]
        explode = [0.05] * len(MOIS)
        colors = plt.cm.Set3(np.linspace(0, 1, len(MOIS)))
        wedges, texts, autotexts = ax_pie.pie(
            sizes, explode=explode, labels=MOIS, autopct="%1.1f%%",
            colors=colors, startangle=90, textprops={"fontsize": 7}
        )
        ax_pie.set_title("Répartition mensuelle des objectifs (%)")
    else:
        ax_pie.text(0.5, 0.5, "Aucune donnée", ha="center", va="center", fontsize=16)
        ax_pie.set_title("Répartition mensuelle des objectifs (%)")
    fig_pie.tight_layout()
    buf_pie = io.BytesIO()
    fig_pie.savefig(buf_pie, format="png", facecolor="white")
    buf_pie.seek(0)
    plt.close(fig_pie)

    # --- Slide 4: Profils & graphiques ---
    slide2 = prs.slides.add_slide(prs.slide_layouts[1])
    slide2.shapes.title.text = f"Répartition & Profil - {partner}"
    body2 = slide2.shapes.placeholders[1]
    tf2 = body2.text_frame

    p = tf2.add_paragraph()
    p.text = f"Objectif annuel 2026: {format_fcfa(total_obj)}"
    p.font.size = Pt(14)

    p = tf2.add_paragraph()
    total_real = sum(real_vals)
    p.text = f"Réalisation annuelle 2026: {format_fcfa(total_real)}"
    p.font.size = Pt(14)

    gap = total_obj - total_real
    p = tf2.add_paragraph()
    p.text = f"GAP: {format_fcfa(gap)}"
    p.font.size = Pt(14)
    p.font.bold = True

    taux = (total_real / total_obj * 100) if total_obj > 0 else 0
    p = tf2.add_paragraph()
    p.text = f"Taux de réalisation: {format_pct(taux)}"
    p.font.size = Pt(14)

    # Insérer le camembert sur la slide 4
    slide2.shapes.add_picture(buf_pie, Inches(6.5), Inches(1.5), width=Inches(3.5))

    # --- Slide suivante: Graphique barres taux par mois ---
    taux_data = []
    for m in MOIS:
        o = obj_mensuels[m]
        r = real_mensuels[m]
        t = (r / o * 100) if o > 0 else 0
        taux_data.append({"Mois": m, "Taux (%)" : t})
    df_taux = pd.DataFrame(taux_data)

    fig_taux, ax_taux = plt.subplots(figsize=(8, 4), dpi=150)
    colors_taux = ["#D4A72C" if t >= 100 else "#003a61" for t in df_taux["Taux (%)"]]
    ax_taux.bar(df_taux["Mois"], df_taux["Taux (%)"], color=colors_taux)
    ax_taux.axhline(y=100, color="red", linestyle="--", linewidth=1, label="Objectif 100%")
    ax_taux.set_xlabel("Mois")
    ax_taux.set_ylabel("Taux de réalisation (%)")
    ax_taux.set_title("Taux de réalisation mensuel 2026")
    ax_taux.legend()
    ax_taux.grid(axis="y", alpha=0.3)
    plt.xticks(rotation=45, ha="right")
    fig_taux.tight_layout()
    buf_taux = io.BytesIO()
    fig_taux.savefig(buf_taux, format="png", facecolor="white")
    buf_taux.seek(0)
    plt.close(fig_taux)

    slide_taux = prs.slides.add_slide(prs.slide_layouts[6])  # layout vide
    title_box_taux = slide_taux.shapes.add_textbox(Inches(1), Inches(0.3), Inches(8), Inches(0.7))
    tf_taux = title_box_taux.text_frame
    p_taux = tf_taux.paragraphs[0]
    p_taux.text = f"Taux de réalisation mensuel - {partner}"
    p_taux.font.size = Pt(24)
    p_taux.font.bold = True
    p_taux.alignment = PP_ALIGN.CENTER
    slide_taux.shapes.add_picture(buf_taux, Inches(1), Inches(1.3), width=Inches(9))

    # Tableau objectifs
    rows = len(MOIS) + 1
    cols = 2
    left = Inches(0.5)
    top = Inches(1.5)
    width = Inches(4)
    height = Inches(0.3 * rows)
    table = slide.shapes.add_table(rows, cols, left, top, width, height).table

    table.cell(0, 0).text = "Mois"
    table.cell(0, 1).text = "Objectif (FCFA)"
    for i, m in enumerate(MOIS):
        table.cell(i + 1, 0).text = m
        table.cell(i + 1, 1).text = f"{obj_mensuels[m]:,.0f}"

    # Graphique barres taux par mois
    real_mensuels = get_effective_realisations_mensuelles(partner)
    txas_data = []
    for m in MOIS:
        o = obj_mensuels[m]
        r = real_mensuels[m]
        taux = (r / o * 100) if o > 0 else 0
        txas_data.append({"Mois": m, "Taux (%)": taux})
    df_taux = pd.DataFrame(txas_data)

    # --- Slide 4: Realisations & Commissions ---
    comm_eff = get_effective_commissions(partner)
    total_real = sum(real_mensuels[m] for m in MOIS)

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"Realisations & Commissions - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Realisation cumulee 2026 : {format_fcfa(total_real)}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Commission reversee : {format_fcfa(comm_eff['commission_reversee'])}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Total a payer : {format_fcfa(comm_eff['total_a_payer'])}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Date de traitement : {comm_eff['date_traitement'] or '-'}"
    p.font.size = Pt(14)

    p = tf.add_paragraph()
    p.text = f"Date de paiement : {comm_eff['date_paiement'] or '-'}"
    p.font.size = Pt(14)

    # --- Slide 5: Profil partenaire ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"Profil partenaire - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Points forts :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["points_forts"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Points faibles :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["points_faibles"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Points focaux :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["points_focaux"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    dg_val = ""
    try:
        dg_val = st.session_state.dg_entries.get(partner, "")
    except Exception:
        dg_val = profil.get("dg", "")
    if not dg_val:
        dg_val = profil.get("dg", "") or "A saisir"
    p.text = f"DG : {dg_val}"
    p.font.size = Pt(12)
    p.font.bold = True

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Personnes ressources :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["personnes_ressources"] or "Non renseigne"
    p.font.size = Pt(11)

    # --- Slide 6: Etat des lieux ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"Etat des lieux - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Difficultes :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["difficultes"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Ameliorations / Solutions :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["ameliorations"] or "Non renseigne"
    p.font.size = Pt(11)

    # --- Slide 7: Recommandations ---
    rec_text = ""
    try:
        rec_text = st.session_state.get(f"rec_{partner}", "")
    except Exception:
        rec_text = ""
    if not rec_text:
        try:
            rec_text = st.session_state.recommandations_entries.get(partner, "")
        except Exception:
            rec_text = ""
    if not rec_text:
        rec_text = profil.get("recommandations", "") or "Non renseigne"

    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"Recommandations - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Recommandations :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = rec_text
    p.font.size = Pt(11)

    # --- Slide 8: Actions PAC ---
    slide = prs.slides.add_slide(prs.slide_layouts[1])
    slide.shapes.title.text = f"Actions PAC planifiees - {partner}"

    body = slide.shapes.placeholders[1]
    tf = body.text_frame

    p = tf.add_paragraph()
    p.text = f"Activites planifiees :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["activites_planifiees"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Cibles :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["cibles"] or "Non renseigne"
    p.font.size = Pt(11)

    p = tf.add_paragraph()
    p.text = ""
    p = tf.add_paragraph()
    p.text = f"Actions sur le dernier trimestre :"
    p.font.size = Pt(12)
    p.font.bold = True
    p = tf.add_paragraph()
    p.text = profil["actions_trimestre"] or "Non renseigne"
    p.font.size = Pt(11)

    # --- Sauvegarde ---
    horodatage = datetime.now().strftime("%Y%m%d_%H%M%S")
    nom_fichier = f"rapport_bancassurance_{partner.replace(' ', '_')}_{horodatage}.pptx"
    chemin = os.path.join(
        r"C:\Users/houno/Desktop/NSIA_assurance_projet/bancassurance_dashboard",
        nom_fichier
    )
    prs.save(chemin)
    return chemin
# ==========================================================================
# SIDEBAR
# ==========================================================================

with st.sidebar:
    st.header("Partenaire")

    # --- Formulaire création de partenaire ---
    with st.expander("➕ Créer un nouveau partenaire", expanded=False):
        new_partenaire = st.text_input("Nom du partenaire", key="new_partenaire_input")
        if st.button("Ajouter le partenaire", key="btn_add_partenaire"):
            if new_partenaire.strip():
                add_partenaire(new_partenaire)
                st.success(f"Partenaire '{new_partenaire}' ajouté !")
                st.rerun()
            else:
                st.warning("Veuillez saisir un nom.")

    # Utiliser get_all_partenaires au lieu de PARTENAIRES pour inclure les dynamiques
    all_partenaires = get_all_partenaires()

    partenaire = st.selectbox(
        "Choisir un partenaire",
        all_partenaires,
        index=all_partenaires.index(st.session_state.partenaire_sel)
            if st.session_state.partenaire_sel in all_partenaires else 0,
        key="partenaire_select_sidebar",
    )
    st.session_state.partenaire_sel = partenaire

    st.divider()
    st.subheader(" Indicateurs cles")

    gap_data = hist_svc.get_gap_taux_historical(partenaire, 2026)
    obj = gap_data["objectif_N"]
    real_nm1 = gap_data["realisation_N_1"]
    real_n = gap_data["realisation_N"]
    gap = gap_data["gap"]
    taux = gap_data["taux"]

    # Taux de realisation: si real_nm1 est disponible, on le calcule; sinon on utilise real_n/obj
    if real_nm1 > 0:
        taux_real = (real_n / obj) * 100 if obj > 0 else 0
    else:
        taux_real = taux
        real_nm1 = real_n  # fallback

    st.metric("Objectif 2026", format_fcfa(obj))
    st.metric("Realisation 2025", format_fcfa(real_nm1))
    st.metric("Realisation 2026 (T1+T2)", format_fcfa(real_n))
    st.metric("GAP", format_fcfa(gap), delta=f"{(gap/real_nm1*100) if real_nm1 > 0 else 0:.1f}% vs 2025")
    st.metric("Taux de realisation 2026", format_pct(taux_real), delta=f"{taux_real - 100:.1f}% vs objectif")

    st.divider()

    if st.button("Sauvegarder saisies (session)", use_container_width=True):
        st.success("Saisies sauvegardees en memoire pour cette session.")

    if st.button("Exporter PowerPoint", use_container_width=True, type="primary"):
        ppt_path = export_presentation(partenaire)
        st.success(f"Presentation exportee: {ppt_path}")

    st.caption("Dashboard Bancassurance NSIA | Donnees: TABLEAU DE BORD ET PAC PAR PARTENAIRE 2026")


# ==========================================================================
# ONGLETS PRINCIPAUX
# ==========================================================================

onglets = st.tabs([
    "Synthèse",
    "Saisie mensuelle & hebdo",
    "Suivi & Commissions",
    "Profil partenaire",
    "Historique multi-années",
])


# --- Onglet 1: Synthese ---

with onglets[0]:

    st.subheader(f"Synthèse GAP et taux de réalisation — {partenaire}")

    # Cartes de performance
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Objectif 2026", format_fcfa(obj),
                  help="Objectif annuel 2026 depuis Feuille1 (colonne Objectif T1 2026)")
    with col2:
        st.metric("Réalisation 2025", format_fcfa(real_nm1),
                  help="Réalisation année N-1 (T1 2025) depuis TB1")
    with col3:
        st.metric("GAP", format_fcfa(gap), delta=format_fcfa(gap),
                  help="GAP = Objectif 2026 - Réalisation 2025")
    with col4:
        st.metric("Taux de réalisation 2026", format_pct(taux),
                  delta=f"{taux - 100:.1f}% vs objectif",
                  help="Taux = Réalisation 2026 / Objectif 2026 * 100")

    st.divider()

    # Tableau de bord detaille
    profil = hist_svc.get_profil_partenaire(partenaire)

    df_synth = pd.DataFrame([
        {
            "Indicateur": "Objectif 2026",
            "Valeur": format_fcfa(obj),
            "Commentaire": "Objectif annuel (Feuille1)",
        },
        {
            "Indicateur": "Réalisation 2025 (N-1)",
            "Valeur": format_fcfa(real_nm1),
            "Commentaire": "Realisation annee precedente",
        },
        {
            "Indicateur": "Réalisation 2026 (T1+T2)",
            "Valeur": format_fcfa(real_n),
            "Commentaire": f"Réalisation N: {format_fcfa(profil['realisation_annuelle'])} — Objectif N: {format_fcfa(profil['objectif_annuel'])}",
        },
        {
            "Indicateur": "GAP",
            "Valeur": format_fcfa(gap),
            "Commentaire": "Objectif(N) - Réalisation(N-1)",
        },
        {
            "Indicateur": "Taux de réalisation 2026",
            "Valeur": format_pct(taux),
            "Commentaire": "Réalisation(N) / Objectif(N)",
        },
    ])
    st.dataframe(df_synth, use_container_width=True, hide_index=True)

    # Graphique comparaison
    st.divider()
    st.markdown("### Comparaison 2025 vs 2026")
    df_comp = pd.DataFrame(
        {"Montant (FCFA)": [real_nm1, obj, real_n]},
        index=["Réalisation 2025", "Objectif 2026", "Réalisation 2026"],
    )
    st.bar_chart(df_comp)

    # Vue d'ensemble tous partenaires
    st.divider()
    st.markdown("### Vue d'ensemble — Tous les partenaires")
    df_tous = hist_svc.get_tous_partenaires_gap_taux()
    if len(df_tous) > 0:
        # Formatter en tant que strings pour le display (on ne garde que le TOTAL ligne)
        df_display = df_tous.copy()
        for col in ["Objectif", "Réalisation", "GAP"]:
            df_display[col] = df_display[col].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else "-")
        df_display["Taux de réalisation (%)"] = df_display["Taux de réalisation (%)"].apply(
            lambda x: f"{x:.1f}%" if pd.notna(x) else "-"
        )

        # --- Ligne TOTALE ---
        total_obj = df_tous["Objectif"].sum()
        total_real = df_tous["Réalisation"].sum()
        total_gap = df_tous["GAP"].sum()
        total_taux = (total_real / total_obj * 100) if total_obj > 0 else 0.0
        total_row = pd.DataFrame([{
            "Partenaire": "TOTAL",
            "Objectif": f"{total_obj:,.0f}",
            "Réalisation": f"{total_real:,.0f}",
            "GAP": f"{total_gap:,.0f}",
            "Taux de réalisation (%)": f"{total_taux:.1f}%",
        }])
        df_tous_total = pd.concat([df_display, total_row], ignore_index=True)
        st.dataframe(df_tous_total, width="stretch", hide_index=True)

        # --- Graphique comparatif Objectif vs Réalisation + Taux par partenaire ---
        st.markdown("#### Comparaison Objectif vs Réalisation par partenaire")
        partners = df_tous["Partenaire"].tolist()
        fig_comp, ax_comp = plt.subplots(figsize=(14, 6), dpi=150)
        x_pos = range(len(partners))
        bar_w = 0.32

        # Barres Objectif (axe gauche)
        obj_vals_all = df_tous["Objectif"].tolist()
        bars_obj_p = ax_comp.bar(
            [i - bar_w for i in x_pos], obj_vals_all, bar_w,
            label="Objectif", color="#D4A72C"
        )

        # Barres Réalisation (axe gauche)
        real_vals_all = df_tous["Réalisation"].tolist()
        bars_real_p = ax_comp.bar(
            x_pos, real_vals_all, bar_w,
            label="Réalisation", color="#003a61"
        )

        # Ligne Taux (%), axe droit
        ax_taux = ax_comp.twinx()
        taux_vals_all = df_tous["Taux de réalisation (%)"].tolist()
        ax_taux.plot(
            list(x_pos), taux_vals_all, "o-", color="#00A86B",
            linewidth=2, markersize=6, label="Taux (%)"
        )
        ax_taux.set_ylabel("Taux de réalisation (%)", color="#00A86B")
        ax_taux.tick_params(axis="y", labelcolor="#00A86B")
        ax_taux.set_ylim(0, 150)

        ax_comp.set_xlabel("Partenaire")
        ax_comp.set_ylabel("Montant (FCFA)")
        ax_comp.set_title("Objectif vs Réalisation vs Taux de réalisation par partenaire")
        ax_comp.set_xticks(list(x_pos))
        ax_comp.set_xticklabels(partners, rotation=45, ha="right", fontsize=9)
        ax_comp.legend(loc="upper left", fontsize=8)
        ax_taux.legend(loc="upper right", fontsize=8)
        ax_comp.grid(axis="y", alpha=0.3)

        fig_comp.tight_layout()
        st.pyplot(fig_comp)
        plt.close(fig_comp)




# --- Onglet 2: Saisie mensuelle & hebdo ---

with onglets[1]:

    st.subheader(f"Objectifs mensuels 2026 — {partenaire}")

    obj_mensuels = get_effective_objectifs_mensuels(partenaire)

    # Inputs en 2 colonnes
    col_obj, col_total = st.columns([3, 1])
    with col_obj:
        st.markdown("### Objectifs par mois (FCFA)")
        obj_modifies = {}
        for m in MOIS:
            val = st.number_input(
                f"Objectif {m}",
                min_value=0.0,
                value=float(obj_mensuels[m]),
                step=1000.0,
                format="%.0f",
                key=f"obj_{partenaire}_{m}",
            )
            obj_modifies[m] = val
            if obj_mensuels[m] != val:
                st.session_state.objectifs_mensuels_override[partenaire][m] = val
            elif st.session_state.objectifs_mensuels_override[partenaire][m] is None:
                pass  # garde la valeur source

    with col_total:
        total_obj = sum(obj_modifies[m] for m in MOIS)
        st.metric("Total annuel", format_fcfa(total_obj))
        st.metric("Objectif annuel source", format_fcfa(obj))

    # Tableau recap
    st.divider()
    recap_data = []
    total_obj = 0
    for m in MOIS:
        o = obj_modifies[m]
        total_obj += o
        recap_data.append({
            "Mois": m,
            "Objectif": f"{o:,.0f}",
        })
    df_obj = pd.DataFrame(recap_data)
    st.dataframe(df_obj, use_container_width=True, hide_index=True)




    with st.expander("📅 Décomposition hebdomadaire", expanded=False):

            st.subheader(f"Décomposition hebdomadaire — {partenaire}")

            mois_sel = st.selectbox(
                "Choisir un mois",
                MOIS,
                index=datetime.now().month - 1,
                key="mois_semaine_select",
            )

            from calendar import monthcalendar
            mois_idx = MOIS.index(mois_sel) + 1
            cal = monthcalendar(2026, mois_idx)
            nb_semaines = len([w for w in cal if w[4] != 0])
            st.caption(f"{mois_sel} 2026 : {nb_semaines} semaines")

            obj_mensuel = get_effective_objectifs_mensuels(partenaire)[mois_sel]
            real_mensuel = get_effective_realisations_mensuelles(partenaire)[mois_sel]

            obj_hebdo = hist_svc.get_objectifs_hebdomadaires(partenaire, mois_sel)

            col_obj_m, col_real_m = st.columns(2)
            with col_obj_m:
                st.metric("Objectif mensuel", format_fcfa(obj_mensuel))
            with col_real_m:
                st.metric("Realisation mensuelle", format_fcfa(real_mensuel))

            st.divider()
            st.markdown(f"### Semaines de {mois_sel}")

            if f"semaines_{partenaire}_{mois_sel}" not in st.session_state:
                st.session_state[f"semaines_{partenaire}_{mois_sel}"] = {
                    i: {"objectif": float(obj_hebdo.get(i, 0)),
                        "realisation": 0.0, "commentaire": ""}
                    for i in range(1, nb_semaines + 1)
                }

            semaines_data = {}
            for s in range(1, nb_semaines + 1):
                with st.expander(f"Semaine {s}", expanded=(s == 1)):
                    col_s1, col_s2, col_s3 = st.columns(3)
                    with col_s1:
                        obj_s = st.number_input(
                            f"Objectif S{s}",
                            min_value=0.0,
                            value=float(st.session_state[f"semaines_{partenaire}_{mois_sel}"][s]["objectif"]),
                            step=1000.0,
                            format="%.0f",
                            key=f"sem_obj_{partenaire}_{mois_sel}_{s}",
                        )
                    with col_s2:
                        real_s = st.number_input(
                            f"Réalisation S{s}",
                            min_value=0.0,
                            value=float(st.session_state[f"semaines_{partenaire}_{mois_sel}"][s]["realisation"]),
                            step=1000.0,
                            format="%.0f",
                            key=f"sem_real_{partenaire}_{mois_sel}_{s}",
                        )
                    with col_s3:
                        comment = st.text_input(
                            f"Commentaire S{s}",
                            value=st.session_state[f"semaines_{partenaire}_{mois_sel}"][s]["commentaire"],
                            key=f"sem_comm_{partenaire}_{mois_sel}_{s}",
                        )
                    semaines_data[s] = {"objectif": obj_s, "realisation": real_s, "commentaire": comment}

            for s in range(1, nb_semaines + 1):
                st.session_state[f"semaines_{partenaire}_{mois_sel}"][s] = semaines_data[s]

            # Recap semaines
            st.divider()
            st.markdown("### Récapitulatif hebdomadaire")
            recap_semaines = []
            total_obj_s = sum(semaines_data[s]["objectif"] for s in semaines_data)
            total_real_s = sum(semaines_data[s]["realisation"] for s in semaines_data)

            for s in range(1, nb_semaines + 1):
                d = semaines_data[s]
                taux_s = (d["realisation"] / d["objectif"] * 100) if d["objectif"] > 0 else 0
                recap_semaines.append({
                    "Semaine": f"S{s}",
                    "Objectif": f"{d['objectif']:,.0f}",
                    "Réalisation": f"{d['realisation']:,.0f}",
                    "Écart": f"{d['realisation'] - d['objectif']:+,.0f}",
                    "Taux (%)": f"{taux_s:.1f}%",
                    "Commentaire": d["commentaire"],
                })
            df_sem = pd.DataFrame(recap_semaines)
            df_sem.loc[len(df_sem)] = {
                "Semaine": "**TOTAL**",
                "Objectif": f"**{total_obj_s:,.0f}**",
                "Réalisation": f"**{total_real_s:,.0f}**",
                "Écart": f"**{total_real_s - total_obj_s:+,.0f}**",
                "Taux (%)": f"**{(total_real_s / total_obj_s * 100) if total_obj_s > 0 else 0:.1f}%**",
                "Commentaire": "",
            }
            st.dataframe(df_sem, use_container_width=True, hide_index=True)

            if nb_semaines > 0:
                st.bar_chart(
                    pd.DataFrame({
                        "Objectif": [semaines_data[s]["objectif"] for s in semaines_data],
                        "Réalisation": [semaines_data[s]["realisation"] for s in semaines_data],
                    }, index=[f"S{s}" for s in semaines_data])
                )






# --- Onglet 3: Suivi & Commissions ---

with onglets[2]:

    st.subheader(f"Réalisations cumulées & Commissions — {partenaire}")

    obj_mensuels_eff = get_effective_objectifs_mensuels(partenaire)
    real_mensuels_eff = get_effective_realisations_mensuelles(partenaire)
    comm_eff = get_effective_commissions(partenaire)

    # Cartes de synthese
    total_obj = sum(obj_mensuels_eff[m] for m in MOIS)
    total_real = sum(real_mensuels_eff[m] for m in MOIS)
    taux_global = (total_real / total_obj * 100) if total_obj > 0 else 0

    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Objectif mensuel total", format_fcfa(total_obj))
    with col2:
        st.metric("Réalisation cumulée", format_fcfa(total_real))
    with col3:
        st.metric("Écart", format_fcfa(total_real - total_obj), delta=f"{taux_global - 100:.1f}%")
    with col4:
        st.metric("Taux global", format_pct(taux_global))

    st.divider()

    # Tableau mensuel objectif vs realisation
    st.markdown("### Tableau Objectif vs Réalisation par mois")
    recap_data = []
    for m in MOIS:
        o = obj_mensuels_eff[m]
        r = real_mensuels_eff[m]
        ecart = r - o
        taux = (r / o * 100) if o > 0 else 0
        recap_data.append({
            "Mois": m,
            "Objectif": f"{o:,.0f}",
            "Réalisation": f"{r:,.0f}",
            "Écart": f"{ecart:+,.0f}",
            "Taux (%)": f"{taux:.1f}%",
        })
    df_recap = pd.DataFrame(recap_data)
    df_recap.loc[len(df_recap)] = {
        "Mois": "**TOTAL**",
        "Objectif": f"**{total_obj:,.0f}**",
        "Réalisation": f"**{total_real:,.0f}**",
        "Écart": f"**{total_real - total_obj:+,.0f}**",
        "Taux (%)": f"**{taux_global:.1f}%**",
    }
    st.dataframe(df_recap, use_container_width=True, hide_index=True)

    # Graphique progression mensuelle (Refait — Janvier à Décembre)
    st.divider()
    st.markdown("### Progression mensuelle — Objectif vs Réalisation")
    df_prog = pd.DataFrame({
        "Objectif": [obj_mensuels_eff[m] for m in MOIS],
        "Réalisation": [real_mensuels_eff[m] for m in MOIS],
    }, index=MOIS)

    # Graphique matplotlib personnalisé
    fig_prog, ax_prog = plt.subplots(figsize=(12, 5), dpi=150)
    x_prog = range(len(MOIS))
    bar_width = 0.35
    bars_obj = ax_prog.bar([i - bar_width/2 for i in x_prog], df_prog["Objectif"].tolist(),
                           bar_width, label="Objectif", color="#D4A72C")
    bars_real = ax_prog.bar([i + bar_width/2 for i in x_prog], df_prog["Réalisation"].tolist(),
                            bar_width, label="Réalisation", color="#003a61")
    ax_prog.set_xlabel("Mois")
    ax_prog.set_ylabel("Montant (FCFA)")
    ax_prog.set_title(f"Progression mensuelle 2026 — {partenaire}")
    ax_prog.set_xticks(list(x_prog))
    ax_prog.set_xticklabels(MOIS, rotation=45, ha="right", fontsize=9)
    ax_prog.legend()
    ax_prog.grid(axis="y", alpha=0.3)

    # Ajouter les valeurs sur les barres
    for bar in bars_obj:
        height = bar.get_height()
        ax_prog.annotate(f"{height:,.0f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=7)
    for bar in bars_real:
        height = bar.get_height()
        ax_prog.annotate(f"{height:,.0f}", xy=(bar.get_x() + bar.get_width() / 2, height),
                        xytext=(0, 3), textcoords="offset points", ha="center", va="bottom", fontsize=7)

    fig_prog.tight_layout()
    st.pyplot(fig_prog)
    plt.close(fig_prog)

    # Barre de progression cumulée
    st.markdown(f"**Taux de réalisation cumulé : {format_pct(taux_global)}**")
    st.progress(min(taux_global / 100, 1.0))

    # Performance par trimestre
    st.divider()
    st.markdown("### Performance par trimestre")
    col_t1, col_t2, col_t3, col_t4 = st.columns(4)
    trimestres = {
        "T1": ["Janvier", "Février", "Mars"],
        "T2": ["Avril", "Mai", "Juin"],
        "T3": ["Juillet", "Août", "Septembre"],
        "T4": ["Octobre", "Novembre", "Décembre"],
    }
    for nom_t, mois_t, col in zip(trimestres.keys(), trimestres.values(), [col_t1, col_t2, col_t3, col_t4]):
        obj_t = sum(obj_mensuels_eff[m] for m in mois_t)
        real_t = sum(real_mensuels_eff[m] for m in mois_t)
        taux_t = (real_t / obj_t * 100) if obj_t > 0 else 0
        with col:
            st.metric(label=f"{nom_t} 2026", value=format_fcfa(real_t), delta=f"{taux_t:.1f}%")

    # Commissions
    st.divider()
    st.markdown("### Commissions cumulées")

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        comm_val = st.number_input(
            "Commission reversée (FCFA)",
            min_value=0.0,
            value=float(comm_eff["commission_reversee"] or 0),
            step=10000.0,
            format="%.0f",
            key=f"comm_rev_{partenaire}",
        )
        total_val = st.number_input(
            "Total à payer (FCFA)",
            min_value=0.0,
            value=float(comm_eff["total_a_payer"] or 0),
            step=10000.0,
            format="%.0f",
            key=f"comm_tot_{partenaire}",
        )
    with col_c2:
        dt_trait = st.date_input(
            "Date de traitement",
            value=datetime.strptime(comm_eff["date_traitement"], "%Y-%m-%d").date()
            if comm_eff["date_traitement"] else date.today(),
            key=f"dt_trait_{partenaire}",
        )
        dt_paie = st.date_input(
            "Date de paiement",
            value=datetime.strptime(comm_eff["date_paiement"], "%Y-%m-%d").date()
            if comm_eff["date_paiement"] else None,
            key=f"dt_paie_{partenaire}",
        )

    st.session_state.commissions_override[partenaire]["commission_reversee"] = comm_val
    st.session_state.commissions_override[partenaire]["total_a_payer"] = total_val
    st.session_state.commissions_override[partenaire]["date_traitement"] = dt_trait.strftime("%Y-%m-%d") if dt_trait else None
    st.session_state.commissions_override[partenaire]["date_paiement"] = dt_paie.strftime("%Y-%m-%d") if dt_paie else None

    # Recap commissions (sans colonnes Date traitement / Date paiement)
    st.markdown("### Récapitulatif commissions")
    df_comm = pd.DataFrame([{
        "Partenaire": partenaire,
        "Commission reversée": f"{comm_val:,.0f} FCFA",
        "Total à payer": f"{total_val:,.0f} FCFA",
    }])
    st.dataframe(df_comm, width="stretch", hide_index=True)

    # Historique commissions tous partenaires (sans colonnes date)
    st.divider()
    st.markdown("### Historique commissions — Tous partenaires")
    df_comm_tous = hist_svc.get_commissions_tous()
    # Supprimer les colonnes date_traitement et date_paiement
    cols_a_garder = ["Partenaire", "Commission reversée", "Total à payer"]
    cols_presentes = [c for c in cols_a_garder if c in df_comm_tous.columns]
    df_comm_tous_aff = df_comm_tous[cols_presentes] if cols_presentes else df_comm_tous
    st.dataframe(df_comm_tous_aff, width="stretch", hide_index=True)




    with st.expander("📊 État des lieux hebdomadaire", expanded=False):

            st.subheader(f"État des lieux — {partenaire}")

            profil = hist_svc.get_profil_partenaire(partenaire)

            # Suivi hebdomadaire (depuis FICHE DE SUIVI)
            mois_fiche = get_mois_from_fiche("Septembre")
            if mois_fiche:
                suivi = hist_svc.get_suivi_hebdomadaire(partenaire, mois_fiche)
                if suivi and "semaines" in suivi:
                    st.markdown("#### Suivi hebdomadaire (Septembre 2026)")
                    col_s1, col_s2, col_s3, col_s4 = st.columns(4)
                    with col_s1:
                        st.metric("Objectif mois", format_fcfa(suivi.get("objectif_mois", 0)))
                    with col_s2:
                        st.metric("Réalisation mois", format_fcfa(suivi.get("realisation_mois", 0)))
                    with col_s3:
                        st.metric("Objectif cumulé", format_fcfa(suivi.get("total_objectif_cumule", 0)))
                    with col_s4:
                        st.metric("Réalisation cumulée", format_fcfa(suivi.get("total_realisation_cumule", 0)))

                    if suivi["semaines"]:
                        st.divider()
                        st.markdown("#### Par semaine")
                        sem_data = []
                        for s_name, s_data in suivi["semaines"].items():
                            taux_s = (s_data["realisation"] / s_data["objectif"] * 100) if s_data["objectif"] > 0 else 0
                            sem_data.append({
                                "Semaine": s_name,
                                "Objectif": f"{s_data['objectif']:,.0f}",
                                "Réalisation": f"{s_data['realisation']:,.0f}",
                                "Taux (%)": f"{taux_s:.1f}%",
                            })
                        st.dataframe(pd.DataFrame(sem_data), use_container_width=True, hide_index=True)

                    if suivi.get("actions_phare"):
                        st.markdown("#### Actions phares du mois")
                        st.write(suivi["actions_phare"])
                    if suivi.get("niveau_rea_commentaires"):
                        st.markdown("#### Niveau de réalisation / Commentaires")
                        st.write(suivi["niveau_rea_commentaires"])

            st.divider()

            # Difficultés et améliorations
            st.markdown("### Difficultés rencontrées")
            st.write(profil["difficultes"] or "Non renseigné")

            st.markdown("### Améliorations / Solutions")
            st.write(profil["ameliorations"] or "Non renseigné")

            # Saisie libre état des lieux
            st.divider()
            st.markdown("### Observations libres")
            obs_key = f"obs_{partenaire}"
            if obs_key not in st.session_state:
                st.session_state[obs_key] = ""
            obs = st.text_area(
                "Observations / état des lieux (saisie libre)",
                value=st.session_state[obs_key],
                height=150,
                key=f"obs_text_{partenaire}",
                help="Saisir les observations, état des lieux, difficultés et améliorations",
            )
            st.session_state[obs_key] = obs






# --- Onglet 4: Profil partenaire ---

with onglets[3]:

    st.subheader(f"Profil partenaire — {partenaire}")

    profil = hist_svc.get_profil_partenaire(partenaire)

    # Cartes indicateurs
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        st.metric("Objectif 2026", format_fcfa(profil["objectif_annuel"]))
    with col2:
        st.metric("Réalisation (annuelle)", format_fcfa(profil["realisation_annuelle"]))
    with col3:
        st.metric("GAP", format_fcfa(profil["gap"]),
                  delta=f"{(profil['gap']/profil['realisation_annuelle']*100) if profil['realisation_annuelle'] > 0 else 0:.1f}%")
    with col4:
        st.metric("Taux 2026", format_pct(profil["taux_realisation"]),
                  delta=f"{profil['taux_realisation'] - 100:.1f}%")

    st.divider()

    # Spécificités (Points Forts / Faibles)
    st.markdown("### Spécificités")
    col_s1, col_s2, col_s3 = st.columns(3)
    with col_s1:
        st.info(f"**Points Forts**\n\n{profil['points_forts'] or 'Non renseigné'}")
    with col_s2:
        st.warning(f"**Points Faibles**\n\n{profil['points_faibles'] or 'Non renseigné'}")
    with col_s3:
        st.success(f"**Points Focaux**\n\n{profil['points_focaux'] or 'Non renseigné'}")

    st.divider()

    # Responsables
    st.markdown("### Responsables")
    col_r1, col_r2 = st.columns(2)
    with col_r1:
        st.markdown("**Personnes Ressources (source Excel):**")
        st.write(profil["personnes_ressources"] or "Non renseigné")
    with col_r2:
        dg_val = st.text_input(
            "DG (à saisir)",
            value=st.session_state.dg_entries.get(partenaire, ""),
            key=f"dg_input_{partenaire}",
            help="Nom du DG / Responsable général du partenaire",
        )
        st.session_state.dg_entries[partenaire] = dg_val

    st.divider()

    # Action planifiée sur le dernier trimestre
    st.markdown("### Actions planifiées sur le dernier trimestre")
    st.write(profil["actions_trimestre"] or "Non renseigné")

    # Activités planifiées
    st.markdown("### Activités planifiées")
    st.write(profil["activites_planifiees"] or "Non renseigné")

    # Cibles
    st.markdown("### Cibles")
    st.write(profil["cibles"] or "Non renseigné")

    # A fin Septembre
    if profil["a_fin_septembre"]:
        st.markdown("### À fin Septembre")
        st.write(profil["a_fin_septembre"])




    with st.expander("💡 Recommandations & Solutions", expanded=False):

            st.subheader(f"Recommandations — {partenaire}")

            profil = hist_svc.get_profil_partenaire(partenaire)

            # Recommandations depuis source Excel
            st.markdown("### Recommandations / Solutions (source Excel)")
            rec_source = profil.get("recommandations", "")
            st.write(rec_source or "Non renseigné")

            st.divider()

            # Saisie recommandations
            st.markdown("### Recommandations (à saisir / modifier)")
            rec_key = f"rec_{partenaire}"
            if rec_key not in st.session_state:
                st.session_state[rec_key] = st.session_state.recommandations_entries.get(partenaire, rec_source or "")

            rec_val = st.text_area(
                "Recommandations pour ce partenaire",
                value=st.session_state[rec_key],
                height=200,
                key=f"rec_text_{partenaire}",
                help="Saisir les recommandations, actions correctrices, priorites",
            )
            st.session_state[rec_key] = rec_val
            st.session_state.recommandations_entries[partenaire] = rec_val

            st.divider()

            # Vue toutes recommandations
            st.markdown("### Toutes les recommandations")
            rec_data = []
            for p in PARTENAIRES:
                prof = hist_svc.get_profil_partenaire(p)
                rec = st.session_state.recommandations_entries.get(p, "")
                if not rec:
                    rec = st.session_state.get(f"rec_{p}", "")
                    if not rec:
                        rec = prof.get("recommandations", "")
                rec_data.append({"Partenaire": p, "Recommandations": rec or "-"})
            st.dataframe(pd.DataFrame(rec_data), use_container_width=True, hide_index=True)









# --- Onglet 5: Historique multi-années ---

with onglets[4]:

    st.subheader(f"Historique & Comparaison inter-années — {partenaire}")

    # --- Barre d'outils: sélection année + actions ---
    existing_annees = hist_svc.get_all_annees(partenaire)
    annee_options = sorted(set(existing_annees + [2021, 2022, 2023, 2024, 2025, 2026]))

    annee_sel = st.selectbox(
        "Année à saisir / consulter",
        annee_options,
        index=len(annee_options) - 1,
        key="hist_annee_select",
    )

    st.divider()
    st.markdown(f"### 📊 Données — Année {annee_sel}")

    # --- Layout en expanders pour une meilleure organisation ---
    # 1. Données numériques (objectifs/realisations)
    with st.expander("Données mensuelles & annuelles (objectifs / réalisations FCFA)", expanded=True):
        data_saisie = {}
        cols_input = st.columns([1, 1, 1, 1, 1, 1])
        for i, m in enumerate(MOIS):
            with cols_input[i % 6]:
                with st.expander(f"{m}", expanded=False):
                    existing = hist_svc.get_data_for_partner_annee(partenaire, annee_sel)
                    mois_data = existing.get(m, {})
                    obj_val = st.number_input(
                        "Objectif (FCFA)",
                        min_value=0.0, value=float(mois_data.get("objectif") or 0),
                        step=1000.0, format="%.0f",
                        key=f"hist_obj_{partenaire}_{annee_sel}_{m}",
                    )
                    real_val = st.number_input(
                        "Réalisation (FCFA)",
                        min_value=0.0, value=float(mois_data.get("realisation") or 0),
                        step=1000.0, format="%.0f",
                        key=f"hist_real_{partenaire}_{annee_sel}_{m}",
                    )
                    data_saisie[m] = {"objectif": obj_val, "realisation": real_val}

        existing_annee = hist_svc.get_data_for_partner_annee(partenaire, annee_sel)
        annee_data = existing_annee.get("ANNUEL", {})
        obj_annuel = st.number_input(
            "Objectif annuel total (FCFA)",
            min_value=0.0, value=float(annee_data.get("objectif") or 0),
            step=10000.0, format="%.0f",
            key=f"hist_obj_annuel_{partenaire}_{annee_sel}",
        )
        real_annuel = st.number_input(
            "Réalisation annuelle totale (FCFA)",
            min_value=0.0, value=float(annee_data.get("realisation") or 0),
            step=10000.0, format="%.0f",
            key=f"hist_real_annuel_{partenaire}_{annee_sel}",
        )
        data_saisie["ANNUEL"] = {"objectif": obj_annuel, "realisation": real_annuel}

        # Boutons d'action pour les données numériques
        col_save, col_calc, col_del = st.columns(3)
        with col_save:
            save_clicked = st.button("💾 Sauvegarder", key=f"save_hist_{partenaire}_{annee_sel}")
        with col_calc:
            calc_clicked = st.button("🧮 Calculer automatiquement", key=f"calc_hist_{partenaire}_{annee_sel}")
        with col_del:
            del_clicked = st.button("🗑️ Supprimer cette année", key=f"del_hist_{partenaire}_{annee_sel}", type="secondary")

        if save_clicked:
            hist_svc.save_values_batch(partenaire, annee_sel, data_saisie)
            for champ, val in st.session_state.get(f"textual_entries_{partenaire}_{annee_sel}", {}).items():
                hist_svc.save_text(partenaire, annee_sel, champ, val)
            st.success(f"Données sauvegardées pour {partenaire} — {annee_sel}")

        if calc_clicked:
            total_obj, total_real = hist_svc.calculate_and_save_annee_totale(partenaire, annee_sel)
            st.success(f"Année calculée: Objectif={format_fcfa(total_obj)}, Réalisation={format_fcfa(total_real)}")

        if del_clicked:
            hist_svc.remove_annee(partenaire, annee_sel)
            remaining = hist_svc.get_all_annees(partenaire)
            st.warning(f"Année {annee_sel} supprimée. Années restantes: {remaining if remaining else 'aucune'}")

    # 2. Informations textuelles (profil, difficultés, recommandations)
    with st.expander("Profil & Informations textuelles (spécificités, DG, points forts/faibles, etc.)", expanded=True):
        st.markdown(f"### Profil & Informations textuelles — Année {annee_sel}")

        if f"textual_entries_{partenaire}_{annee_sel}" not in st.session_state:
            existing_text = hist_svc.get_text_for_partner_annee(partenaire, annee_sel)
            st.session_state[f"textual_entries_{partenaire}_{annee_sel}"] = {
                "specifications": existing_text.get("specifications", ""),
                "responsables_bancaires": existing_text.get("responsables_bancaires", ""),
                "dg": existing_text.get("dg", ""),
                "points_forts": existing_text.get("points_forts", ""),
                "points_faibles": existing_text.get("points_faibles", ""),
                "points_focaux": existing_text.get("points_focaux", ""),
                "activites_planifiees": existing_text.get("activites_planifiees", ""),
                "cibles": existing_text.get("cibles", ""),
                "actions_trimestre": existing_text.get("actions_trimestre", ""),
                "personnes_ressources": existing_text.get("personnes_ressources", ""),
                "difficultes": existing_text.get("difficultes", ""),
                "ameliorations": existing_text.get("ameliorations", ""),
                "recommandations": existing_text.get("recommandations", ""),
            }

        text_entries = st.session_state[f"textual_entries_{partenaire}_{annee_sel}"]

        # SPECIFICITE + RESPONSABLES + DG (colonnes)
        col1, col2, col3 = st.columns(3)
        with col1:
            st.markdown("**Spécificité**")
            text_entries["specifications"] = st.text_area(
                "Spécificité du partenaire",
                value=text_entries["specifications"],
                key=f"txt_spec_{partenaire}_{annee_sel}",
                height=100,
            )
        with col2:
            st.markdown("**Responsables bancaires**")
            text_entries["responsables_bancaires"] = st.text_area(
                "Responsables bancaires (DG, points focaux, etc.)",
                value=text_entries["responsables_bancaires"],
                key=f"txt_resp_{partenaire}_{annee_sel}",
                height=100,
            )
        with col3:
            st.markdown("**DG**")
            text_entries["dg"] = st.text_area(
                "DG du partenaire",
                value=text_entries["dg"],
                key=f"txt_dg_{partenaire}_{annee_sel}",
                height=100,
            )

        # Points Forts / Faibles / Focaux
        st.markdown("**Points Forts / Faibles / Focaux**")
        col_pf, col_pfa, col_pfo = st.columns(3)
        with col_pf:
            text_entries["points_forts"] = st.text_area(
                "Points Forts",
                value=text_entries["points_forts"],
                key=f"txt_pf_{partenaire}_{annee_sel}",
                height=120,
            )
        with col_pfa:
            text_entries["points_faibles"] = st.text_area(
                "Points Faibles",
                value=text_entries["points_faibles"],
                key=f"txt_pfa_{partenaire}_{annee_sel}",
                height=120,
            )
        with col_pfo:
            text_entries["points_focaux"] = st.text_area(
                "Points Focaux",
                value=text_entries["points_focaux"],
                key=f"txt_pfo_{partenaire}_{annee_sel}",
                height=120,
            )

        # Activités, Cibles, Actions, Personnes Ressources
        st.markdown("**Activités, Cibles & Actions**")
        col_act, col_cib = st.columns(2)
        with col_act:
            text_entries["activites_planifiees"] = st.text_area(
                "Activités planifiées",
                value=text_entries["activites_planifiees"],
                key=f"txt_act_{partenaire}_{annee_sel}",
                height=100,
            )
        with col_cib:
            text_entries["cibles"] = st.text_area(
                "Cibles",
                value=text_entries["cibles"],
                key=f"txt_cib_{partenaire}_{annee_sel}",
                height=100,
            )

        st.markdown("**Actions & Personnes Ressources**")
        col_act_p, col_pers = st.columns(2)
        with col_act_p:
            text_entries["actions_trimestre"] = st.text_area(
                "Actions (planifiées)",
                value=text_entries["actions_trimestre"],
                key=f"txt_actplan_{partenaire}_{annee_sel}",
                height=100,
            )
        with col_pers:
            text_entries["personnes_ressources"] = st.text_area(
                "Personnes Ressources",
                value=text_entries["personnes_ressources"],
                key=f"txt_pers_{partenaire}_{annee_sel}",
                height=100,
            )

        st.session_state[f"textual_entries_{partenaire}_{annee_sel}"] = text_entries

        if st.button("💾 Sauvegarder les informations textuelles", key=f"save_text_{partenaire}_{annee_sel}"):
            for champ, val in text_entries.items():
                hist_svc.save_text(partenaire, annee_sel, champ, val)
            st.success(f"Informations textuelles sauvegardées pour {partenaire} — {annee_sel}")

        # Suppression de champs textuels spécifiques
        st.divider()
        st.markdown("**🗑️ Supprimer un champ textuel**")
        champ_to_delete = st.selectbox(
            "Champ à supprimer",
            ["", "specifications", "responsables_bancaires", "dg", "points_forts",
             "points_faibles", "points_focaux", "activites_planifiees", "cibles",
             "actions_trimestre", "personnes_ressources", "difficultes", "ameliorations",
             "recommandations"],
            key=f"del_field_select_{partenaire}_{annee_sel}",
            help="Sélectionner un champ pour le supprimer de la base",
        )
        if champ_to_delete:
            if st.button(f"🗑️ Supprimer « {champ_to_delete} »", key=f"del_field_{partenaire}_{annee_sel}_{champ_to_delete}"):
                hist_svc.delete_text_field(partenaire, annee_sel, champ_to_delete)
                text_entries[champ_to_delete] = ""
                if f"textual_entries_{partenaire}_{annee_sel}" in st.session_state:
                    st.session_state[f"textual_entries_{partenaire}_{annee_sel}"][champ_to_delete] = ""
                st.warning(f"Champ « {champ_to_delete} » supprimé pour {partenaire} — {annee_sel}")

    # 3. État des lieux et recommandations
    with st.expander("État des lieux & Recommandations", expanded=True):
        st.markdown("**Difficultés**")
        text_entries["difficultes"] = st.text_area(
            "Difficultés rencontrées",
            value=text_entries["difficultes"],
            key=f"txt_diff_{partenaire}_{annee_sel}",
            height=100,
        )
        st.markdown("**Améliorations**")
        text_entries["ameliorations"] = st.text_area(
            "Améliorations / solutions",
            value=text_entries["ameliorations"],
            key=f"txt_amp_{partenaire}_{annee_sel}",
            height=100,
        )
        st.markdown("**Recommandations**")
        text_entries["recommandations"] = st.text_area(
            "Recommandations / Solutions",
            value=text_entries["recommandations"],
            key=f"txt_rec_{partenaire}_{annee_sel}",
            height=100,
        )

        st.session_state[f"textual_entries_{partenaire}_{annee_sel}"] = text_entries

        if st.button("💾 Sauvegarder l'état des lieux & recommandations", key=f"save_lieux_{partenaire}_{annee_sel}"):
            for champ, val in text_entries.items():
                hist_svc.save_text(partenaire, annee_sel, champ, val)
            st.success(f"État des lieux et recommandations sauvegardés pour {partenaire} — {annee_sel}")

    # --- Comparaison inter-années ---
    st.divider()
    st.markdown("### Comparaison inter-années")

    annees_db = hist_svc.get_all_annees(partenaire)
    if not annees_db:
        st.info("Aucune donnée historique enregistrée. Saisissez des données ci-dessus pour commencer.")
    else:
        annees_a_compar = st.multiselect(
            "Sélectionner les années à comparer",
            annees_db,
            default=annees_db,
            key=f"compare_annees_{partenaire}",
        )

        if annees_a_compar:
            annees_compare = sorted(set(annees_a_compar))

            # Comparaison par mois
            st.markdown("#### Comparaison par mois")
            df_compare = hist_svc.get_comparison_dataframe(partenaire, annees_compare)
            st.dataframe(df_compare, use_container_width=True, hide_index=True)

            chart_data = {}
            for a in annees_compare:
                if f"Réalisation {a}" in df_compare.columns:
                    chart_data[f"Réalisation {a}"] = df_compare[f"Réalisation {a}"]
                if f"Objectif {a}" in df_compare.columns:
                    chart_data[f"Objectif {a}"] = df_compare[f"Objectif {a}"]
            if chart_data:
                df_chart = pd.DataFrame(chart_data, index=df_compare["Mois"])
                st.bar_chart(df_chart)

            # Comparaison par semaine
            st.divider()
            st.markdown("#### Comparaison hebdomadaire par mois")

            mois_comparaison = st.selectbox(
                "Choisir un mois pour la comparaison hebdomadaire",
                MOIS,
                index=8,
                key=f"weekly_compare_mois_{partenaire}",
            )

            df_weekly = hist_svc.get_weekly_comparison_dataframe(partenaire, annees_compare, mois_comparaison)
            if not df_weekly.empty:
                st.dataframe(df_weekly, use_container_width=True, hide_index=True)

                chart_data_weekly = {}
                for a in annees_compare:
                    if f"Réalisation {a}" in df_weekly.columns:
                        chart_data_weekly[f"Réalisation {a}"] = df_weekly[f"Réalisation {a}"]
                    if f"Objectif {a}" in df_weekly.columns:
                        chart_data_weekly[f"Objectif {a}"] = df_weekly[f"Objectif {a}"]
                if chart_data_weekly:
                    df_chart_w = pd.DataFrame(chart_data_weekly, index=df_weekly["Semaine"])
                    st.bar_chart(df_chart_w)

                # Saisie hebdomadaire
                st.markdown("##### Saisie / modification des données hebdomadaires")
                existing_weekly = hist_svc.get_weekly_for_partner_mois(partenaire, annee_sel, mois_comparaison)
                for s in range(1, len(df_weekly) + 1):
                    s_data = existing_weekly.get(s, {"objectif": 0, "realisation": 0, "commentaire": ""})
                    col_ws, col_wr, col_wc = st.columns(3)
                    with col_ws:
                        obj_s = st.number_input(
                            f"Obj. S{s}",
                            min_value=0.0, value=float(s_data.get("objectif", 0)),
                            step=1000.0, format="%.0f",
                            key=f"week_obj_{partenaire}_{annee_sel}_{mois_comparaison}_{s}",
                        )
                    with col_wr:
                        real_s = st.number_input(
                            f"Réal. S{s}",
                            min_value=0.0, value=float(s_data.get("realisation", 0)),
                            step=1000.0, format="%.0f",
                            key=f"week_real_{partenaire}_{annee_sel}_{mois_comparaison}_{s}",
                        )
                    with col_wc:
                        comm_s = st.text_input(
                            f"Commentaire S{s}",
                            value=s_data.get("commentaire", ""),
                            key=f"week_comm_{partenaire}_{annee_sel}_{mois_comparaison}_{s}",
                        )
                    if st.button(f"Sauvegarder S{s}", key=f"save_week_{partenaire}_{annee_sel}_{mois_comparaison}_{s}"):
                        hist_svc.save_weekly(partenaire, annee_sel, mois_comparaison, s, obj_s, real_s, comm_s)
                        st.success(f"S{s} sauvegardé pour {mois_comparaison} {annee_sel}")
            else:
                st.info("Aucune donnée hebdomadaire pour ce mois. Saisissez des données hebdomadaires via l'onglet 3.")

            # Export
            st.divider()
            if st.button("📥 Exporter l'historique comparé (Excel)", key=f"export_excel_{partenaire}"):
                export_path = os.path.join(
                    os.path.dirname(os.path.abspath(__file__)),
                    f"historique_comparaison_{partenaire.replace(' ', '_')}.xlsx",
                )
                hist_svc.export_to_excel(partenaire, annees_compare, export_path)
                st.success(f"Exporté: {export_path}")
