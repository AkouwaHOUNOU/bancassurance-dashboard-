"""Bancassurance NSIA : consulter, saisir durablement, comparer.

SQLite est l'unique source des valeurs affichées ; aucun brouillon de session
n'est présenté comme une sauvegarde. GAP = objectif - réalisation de l'année.
"""
import os
from datetime import date
from pathlib import Path
import altair as alt

import pandas as pd
import streamlit as st

from database_service import MOIS, format_fcfa, format_pct
from historical_service import DB_PATH, HistoricalService

st.set_page_config(page_title="Bancassurance NSIA", layout="wide")
st.markdown("<style>" + Path(__file__).with_name("styles.css").read_text(encoding="utf-8") + "</style>", unsafe_allow_html=True)

# Aucune suppression automatique au lancement.

# L'injection permet des tests sur copie sans jamais ouvrir la base de production.
hist_svc = HistoricalService(os.environ.get("BANCASSURANCE_DB_PATH", DB_PATH))
partners = hist_svc.get_all_partenaires()
with st.sidebar:
    st.markdown('<div class="nsia-brand"><strong>NSIA</strong><span>Bancassurance · Pilotage</span></div><div class="nsia-sidebar-label">PÉRIMÈTRE DE TRAVAIL</div>', unsafe_allow_html=True)
    partenaire = st.selectbox("Partenaire", partners, key="global_partner")
    annee = int(st.number_input("Année", min_value=1900, max_value=2200,
                               value=date.today().year, step=1, key="global_year"))
    with st.expander("Ajouter un partenaire", expanded=False):
        with st.form("new_partner"):
            name = st.text_input("Nom du partenaire", key="new_partner_name")
            if st.form_submit_button("Ajouter le partenaire", key="add_partner"):
                if name.strip():
                    hist_svc.add_partenaire(name)
                    st.session_state["saved_notice"] = "Partenaire enregistré en SQLite."
                    st.rerun()
                else:
                    st.warning("Veuillez saisir un nom.")
    page = st.radio("Navigation", ["Tableau de bord", "Production", "Suivi hebdomadaire",
                                  "Commissions", "Profil", "Comparaison"], key="navigation")

st.title(f"{partenaire} — {annee}")
if "saved_notice" in st.session_state:
    # Un export préparé avant une modification ne doit plus être téléchargé.
    for key in list(st.session_state):
        if key.startswith("exports_"):
            del st.session_state[key]
    st.success(st.session_state.pop("saved_notice"))
raw = hist_svc.get_data_for_partner_annee(partenaire, annee)
text = hist_svc.get_text_for_partner_annee(partenaire, annee)
context = f"{partenaire}_{annee}"
def amount_deletion(scope, values, labels, widget_keys, delete, help_text=""):
    """Choix, aperçu et confirmation hors formulaire ; aucun effacement implicite."""
    with st.expander("Supprimer des montants", expanded=False):
        st.caption(scope.replace("_", " · "))
        options = {"Choisir un montant": (), **{label: (field,) for field, label in labels.items()},
                   "Les deux montants": tuple(labels)}
        choice = st.selectbox("Montant à supprimer", list(options), key=f"delete_metric_{scope}")
        fields = options[choice]
        for field in fields:
            value = values.get(field)
            rendered = "— (absent)" if value is None or value == "" else format_fcfa(value) if isinstance(value, (int, float)) else str(value)
            st.write(f"{labels[field]} : {rendered}")
        if help_text:
            st.caption(help_text)
        confirmation_key = f"confirm_amount_{scope}_{choice}"
        confirmed = st.checkbox("Je confirme la suppression de ces montants uniquement.", key=confirmation_key)
        present = any(values.get(field) is not None and values.get(field) != "" for field in fields)
        def remove_selected():
            # Callback exécuté avant le rendu : relecture SQLite et widget
            # désactivé utilisent immédiatement les nouvelles valeurs.
            delete(fields)
            for field in fields:
                st.session_state[widget_keys[field]] = None if field in ("objectif", "realisation") else ""
            st.session_state[confirmation_key] = False
            st.session_state["saved_notice"] = "Montants sélectionnés supprimés. Les autres données sont conservées."
        st.button("Supprimer les montants sélectionnés", key=f"delete_amount_{scope}",
                  disabled=not (fields and present and confirmed), on_click=remove_selected)

monthly = pd.DataFrame([
    {"Mois": month, "Objectif": raw.get(month, {}).get("objectif"),
     "Réalisation": raw.get(month, {}).get("realisation")}
    for month in MOIS
])
if page == "Tableau de bord":
    kpi = hist_svc.get_gap_taux_historical(partenaire, annee)
    for column, label, value in zip(st.columns(4),
            ["Objectif", "Réalisation", "GAP", "Taux de réalisation"],
            [format_fcfa(kpi["objectif_N"]), format_fcfa(kpi["realisation_N"]),
             format_fcfa(kpi["gap"]), format_pct(kpi["taux"])] if raw else ["—"] * 4):
        column.metric(label, value)
    total_source = "somme des mois" if text.get("annual_mode") == "monthly" or "ANNUEL" not in raw else "totaux annuels saisis"
    st.caption(f"GAP = objectif − réalisation de l'année. Total annuel : {total_source}. Une donnée absente est indiquée par —.")
    with st.container(key="evolution_panel"):
        st.subheader("Évolution mensuelle")
        st.caption("Objectifs et réalisations enregistrés · FCFA · Janvier à décembre")
        if monthly[["Objectif", "Réalisation"]].notna().any().any():
            points = monthly.melt(id_vars="Mois", var_name="Indicateur", value_name="Montant").dropna(subset=["Montant"])
            # Ordre calendaire explicite et infobulles : ni tri alphabétique,
            # ni valeurs annotées qui se chevauchent au-dessus des barres.
            evolution = alt.Chart(points).mark_bar(cornerRadiusTopLeft=2, cornerRadiusTopRight=2).encode(
                x=alt.X("Mois:N", sort=MOIS, axis=alt.Axis(title=None, labelAngle=-25, labelPadding=12)),
                xOffset="Indicateur:N",
                y=alt.Y("Montant:Q", axis=alt.Axis(title="Montant (FCFA)", format="~s", tickCount=5)),
                color=alt.Color("Indicateur:N", scale=alt.Scale(domain=["Objectif", "Réalisation"], range=["#D4A72C", "#002B49"]), legend=alt.Legend(title=None, orient="top", symbolType="square")),
                tooltip=["Mois:N", "Indicateur:N", alt.Tooltip("Montant:Q", format=",.0f")],
            ).properties(height=300, background="transparent").configure_view(stroke=None).configure_axis(
                labelColor="#002B49", titleColor="#002B49", domain=False, gridColor="#E2E8F0", gridOpacity=.8, labelFontSize=16, titleFontSize=16,
            ).configure_legend(labelColor="#002B49", labelFontSize=16, padding=10)
            # Streamlit 1.49 accepte use_container_width sur Altair, pas width.
            st.altair_chart(evolution, use_container_width=True)
        else:
            st.info("Aucune donnée mensuelle enregistrée pour cette année.")
    with st.expander("Vue de tous les partenaires", expanded=False):
        rows = []
        for partner in partners:
            values = hist_svc.get_gap_taux_historical(partner, annee)
            known = bool(hist_svc.get_data_for_partner_annee(partner, annee))
            rows.append({"Partenaire": partner, "Objectif": format_fcfa(values["objectif_N"]) if known else "—",
                         "Réalisation": format_fcfa(values["realisation_N"]) if known else "—", "GAP": format_fcfa(values["gap"]) if known else "—",
                         "Taux de réalisation (%)": format_pct(values["taux"]) if known else "—"})
        st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
elif page == "Production":
    st.subheader("Production mensuelle et annuelle")
    mois = st.selectbox("Mois à saisir", MOIS, key=f"month_{context}")
    entry = raw.get(mois, {})
    with st.form(f"monthly_{context}_{mois}"):
        left, right = st.columns(2)
        obj = left.number_input("Objectif du mois (FCFA)", value=entry.get("objectif"),
                                step=1000.0, key=f"monthly_obj_{context}_{mois}")
        real = right.number_input("Réalisation du mois (FCFA)", value=entry.get("realisation"),
                                  step=1000.0, key=f"monthly_real_{context}_{mois}")
        if st.form_submit_button("Enregistrer le mois", key=f"save_month_{context}_{mois}"):
            # Un seul mois est modifié : les autres mois et totaux manuels restent intacts.
            hist_svc.save_values_batch(partenaire, annee, {mois: {field: value for field, value in {"objectif": obj, "realisation": real}.items() if value is not None}})
            st.session_state["saved_notice"] = f"{mois} {annee} enregistré en SQLite."
            st.rerun()

    amount_deletion(f"monthly_{context}_{mois}", entry,
                    {"objectif": "Objectif", "realisation": "Réalisation"},
                    {"objectif": f"monthly_obj_{context}_{mois}", "realisation": f"monthly_real_{context}_{mois}"},
                    lambda fields: hist_svc.delete_amounts(partenaire, annee, mois, fields))

    with st.expander("Total annuel", expanded=False):
        # Un total manuel existant n'est jamais remplacé lors d'une saisie mensuelle.
        automatic = text.get("annual_mode") == "monthly" or ("ANNUEL" not in raw and text.get("annual_mode") != "manual")
        mode = st.radio("Mode du total annuel", ["Somme des mois", "Totaux annuels manuels"],
                        index=0 if automatic else 1, key=f"annual_mode_{context}")
        with st.form(f"annual_{context}"):
            if mode == "Totaux annuels manuels":
                annual = raw.get("ANNUEL", {})
                a_left, a_right = st.columns(2)
                a_obj = a_left.number_input("Objectif annuel (FCFA)", value=annual.get("objectif"),
                                        step=1000.0, key=f"annual_obj_{context}")
                a_real = a_right.number_input("Réalisation annuelle (FCFA)", value=annual.get("realisation"),
                                         step=1000.0, key=f"annual_real_{context}")
            else:
                st.caption("Le total suivra la somme des mois enregistrés. Les anciens totaux manuels restent conservés en base.")
            if st.form_submit_button("Enregistrer le mode et le total annuel", key=f"save_annual_{context}"):
                if mode == "Totaux annuels manuels":
                    hist_svc.save_values_batch(partenaire, annee, {"ANNUEL": {field: value for field, value in {"objectif": a_obj, "realisation": a_real}.items() if value is not None}})
                hist_svc.save_text(partenaire, annee, "annual_mode", "monthly" if mode == "Somme des mois" else "manual")
                st.session_state["saved_notice"] = "Mode annuel enregistré en SQLite."
                st.rerun()
        amount_deletion(f"annual_{context}", raw.get("ANNUEL", {}),
                        {"objectif": "Objectif", "realisation": "Réalisation"},
                        {"objectif": f"annual_obj_{context}", "realisation": f"annual_real_{context}"},
                        lambda fields: hist_svc.delete_amounts(partenaire, annee, "ANNUEL", fields),
                        "Seuls les totaux manuels archivés sont supprimés ; le mode annuel est conservé. Un montant annuel absent est recalculé depuis les mois (0 si aucun mois). L'autre total manuel reste prioritaire en mode manuel.")
elif page == "Suivi hebdomadaire":
    st.subheader("Suivi hebdomadaire")
    w_month = st.selectbox("Mois hebdomadaire", MOIS, key=f"weekly_month_{context}")
    # S1–S5 sont les périodes métier de la fiche, pas des semaines ISO :
    # conserver S5 même en février évite de masquer les anciennes saisies.
    week = st.selectbox("Semaine", list(range(1, 6)), format_func=lambda value: f"S{value}",
                        key=f"weekly_week_{context}_{w_month}")
    weekly = hist_svc.get_weekly_for_partner_mois(partenaire, annee, w_month)
    w_data = weekly.get(week, {})
    w_key = f"{context}_{w_month}_{week}"
    with st.form(f"weekly_{w_key}"):
        left, right = st.columns(2)
        w_obj = left.number_input("Objectif de la semaine (FCFA)", value=w_data.get("objectif"),
                                 step=1000.0, key=f"weekly_obj_{w_key}")
        w_real = right.number_input("Réalisation de la semaine (FCFA)", value=w_data.get("realisation"),
                                   step=1000.0, key=f"weekly_real_{w_key}")
        comment = st.text_input("Commentaire", value=w_data.get("commentaire", ""), key=f"weekly_comment_{w_key}")
        if st.form_submit_button("Enregistrer la semaine", key=f"save_week_{w_key}"):
            hist_svc.save_weekly(partenaire, annee, w_month, week, w_obj, w_real, comment)
            st.session_state["saved_notice"] = f"S{week} — {w_month} {annee} enregistrée en SQLite."
            st.rerun()
    amount_deletion(f"weekly_{w_key}", w_data,
                    {"objectif": "Objectif", "realisation": "Réalisation"},
                    {"objectif": f"weekly_obj_{w_key}", "realisation": f"weekly_real_{w_key}"},
                    lambda fields: hist_svc.delete_weekly_amounts(partenaire, annee, w_month, week, fields),
                    "Le commentaire et l'autre montant sont conservés. Semaine sélectionnée : S" + str(week))
    if weekly:
        st.dataframe(pd.DataFrame([{"Semaine": f"S{s}", "Objectif": value["objectif"],
                                   "Réalisation": value["realisation"], "Commentaire": value["commentaire"]}
                                  for s, value in weekly.items()]), hide_index=True, width="stretch")
    st.caption("Le suivi hebdomadaire reste indépendant du mensuel : pas de cumul automatique ni de double comptage.")
elif page == "Commissions":
    st.subheader("Commissions")
    with st.form(f"commissions_{context}"):
        commissions = {}
        # Les champs textuels dédiés préservent les anciennes écritures FCFA
        # (espaces, virgules, annotations) au lieu de les confondre avec ANNUEL.
        cols = st.columns(2)
        for index, (field, label) in enumerate([("commission_reversee", "Commission reversée (FCFA)"),
                             ("total_a_payer", "Total à payer (FCFA)"),
                             ("date_traitement", "Date de traitement"),
                             ("date_paiement", "Date de paiement")]):
            commissions[field] = cols[index % 2].text_input(label, value=text.get(field, "") or "", key=f"{field}_{context}")
        if st.form_submit_button("Enregistrer les commissions", key=f"save_commissions_{context}"):
            hist_svc.save_text_batch(partenaire, annee, {field: value for field, value in commissions.items() if field in text or value != ""})
            st.session_state["saved_notice"] = "Commissions enregistrées en SQLite."
            st.rerun()
    amount_deletion(f"commissions_{context}", text,
                    {"commission_reversee": "Commission reversée", "total_a_payer": "Total à payer"},
                    {field: f"{field}_{context}" for field in ("commission_reversee", "total_a_payer")},
                    lambda fields: hist_svc.delete_commission_amounts(partenaire, annee, fields),
                    "Les dates de traitement et de paiement sont conservées.")
elif page == "Profil":
    st.subheader("Profil du partenaire")
    with st.form(f"profile_{context}"):
        profile_entries = {}
        sections = {
            "Profil et responsables": [("specifications", "Spécificités du partenaire"),
                ("responsables_bancaires", "Responsables bancaires"), ("dg", "DG du partenaire"),
                ("points_forts", "Points forts"), ("points_faibles", "Points faibles"),
                ("points_focaux", "Points focaux"), ("personnes_ressources", "Personnes ressources")],
            "Actions et activités": [("activites_planifiees", "Activités planifiées"),
                ("cibles", "Cibles"), ("actions_trimestre", "Actions du dernier trimestre"),
                ("a_fin_septembre", "Situation à fin septembre")],
            "Difficultés et recommandations": [("difficultes", "Difficultés rencontrées"),
                ("ameliorations", "Améliorations / solutions"), ("recommandations", "Recommandations"),
                ("observations", "Observations / état des lieux")],
        }
        for section, fields in sections.items():
            with st.expander(section, expanded=False):
                cols = st.columns(2)
                for index, (field, label) in enumerate(fields):
                    profile_entries[field] = cols[index % 2].text_area(label, value=text.get(field, "") or "",
                                                                      key=f"profile_{field}_{context}", height=100)
        if st.form_submit_button("Enregistrer le profil, les actions et les recommandations", key=f"save_profile_{context}"):
            hist_svc.save_text_batch(partenaire, annee, profile_entries)
            st.session_state["saved_notice"] = "Profil, actions et recommandations enregistrés en SQLite."
            st.rerun()
    with st.expander("Options avancées — suppression", expanded=False):
        expected = f"{partenaire} / {annee}"
        st.warning(f"Supprime uniquement {partenaire}, année {annee} : mois, hebdomadaire, profil et commissions. Cette opération est irréversible.")
        confirmation = st.text_input("Confirmation de suppression", help=f"Recopiez exactement : {expected}",
                                     key=f"delete_confirmation_{context}")
        if st.button("Supprimer cette année", key=f"delete_year_{context}", disabled=confirmation != expected):
            # Le service supprime les trois tables dans une transaction ciblée.
            hist_svc.remove_annee(partenaire, annee)
            for key in list(st.session_state):
                if context in key:
                    del st.session_state[key]
            st.session_state["saved_notice"] = f"Année {annee} supprimée uniquement pour {partenaire}."
            st.rerun()
elif page == "Comparaison":
    st.subheader("Comparaison inter-années")
    years = sorted(set(hist_svc.get_all_annees(partenaire) + [annee]))
    selected_years = st.multiselect("Années à comparer", years, default=[annee], key=f"compare_years_{context}")
    if selected_years:
        comparison = hist_svc.get_comparison_dataframe(partenaire, sorted(selected_years))
        has_numeric = any(hist_svc.get_data_for_partner_annee(partenaire, year) for year in selected_years)
        if has_numeric:
            # L'affichage est distinct des exports numériques : aucun « None »
            # ni zéro supposé pour les mois non saisis.
            display = comparison.copy()
            for col in display.columns[1:]:
                display[col] = display[col].map(lambda value: "—" if pd.isna(value) else f"{value:,.1f}".replace(",", " ") if col.startswith("Taux") else f"{value:,.0f}".replace(",", " "))
            st.dataframe(display, hide_index=True, width="stretch", height=360)
            chart = comparison[comparison["Mois"] != "ANNUEL"].set_index("Mois")
            realization_cols = [f"Réalisation {year}" for year in sorted(selected_years)]
            if chart[realization_cols].notna().any().any():
                # Un axe nominal Streamlit trie les mois alphabétiquement :
                # imposer l'ordre calendaire pour une comparaison fidèle.
                series = chart[realization_cols].reset_index().melt(
                    id_vars="Mois", var_name="Exercice", value_name="Réalisation"
                ).dropna(subset=["Réalisation"])
                comparison_chart = alt.Chart(series).mark_line(point=True).encode(
                    x=alt.X("Mois:N", sort=MOIS, axis=alt.Axis(title=None, labelAngle=-25)),
                    y=alt.Y("Réalisation:Q", axis=alt.Axis(title="Réalisation (FCFA)", format="~s")),
                    color=alt.Color("Exercice:N", scale=alt.Scale(range=["#002B49", "#D4A72C", "#547084", "#8F6B16"]), legend=alt.Legend(title=None, orient="top")),
                    tooltip=["Mois:N", "Exercice:N", alt.Tooltip("Réalisation:Q", format=",.0f")],
                ).properties(height=280, background="transparent").configure_view(stroke=None).configure_axis(
                    labelColor="#002B49", titleColor="#002B49", gridColor="#E2E8F0", domain=False, labelFontSize=16, titleFontSize=16,
                ).configure_legend(labelColor="#002B49", labelFontSize=16)
                st.altair_chart(comparison_chart, use_container_width=True)
        else:
            st.info("Aucune donnée chiffrée pour les années sélectionnées. Choisissez une année renseignée ou ouvrez Production pour enregistrer vos premiers chiffres.")
        with st.expander("Comparaison hebdomadaire", expanded=False):
            compare_month = st.selectbox("Mois à comparer", MOIS, key=f"compare_month_{context}")
            weekly_compare = hist_svc.get_weekly_comparison_dataframe(partenaire, sorted(selected_years), compare_month)
            if weekly_compare.iloc[:, 1:].notna().any().any():
                weekly_display = weekly_compare.copy()
                for col in weekly_display.columns[1:]:
                    weekly_display[col] = weekly_display[col].map(lambda value: "—" if pd.isna(value) else f"{value:,.1f}".replace(",", " "))
                st.dataframe(weekly_display, hide_index=True, width="stretch")
            else:
                st.caption("Aucune saisie hebdomadaire pour ce mois et ces années.")
        st.subheader("Rapports à télécharger")
        st.caption("Excel : exercices sélectionnés. PowerPoint : exercice choisi dans le panneau latéral.")
        from reporting import export_excel, export_presentation
        export_key = f"exports_{context}_{sorted(selected_years)}"
        exports = st.session_state.setdefault(export_key, {})
        left, right = st.columns(2)
        with left:
            if st.button("Préparer l'export Excel", key=f"export_excel_{context}"):
                exports["excel"] = export_excel(hist_svc, partenaire, sorted(selected_years))
            if "excel" in exports:
                st.download_button("Télécharger Excel", exports["excel"],
                                   file_name=f"historique_{partenaire}_{annee}.xlsx",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key=f"download_excel_{context}")
        with right:
            if st.button(f"Préparer le PowerPoint {annee}", key=f"export_pptx_{context}"):
                exports["pptx"] = export_presentation(hist_svc, partenaire, annee)
            if "pptx" in exports:
                st.download_button("Télécharger PowerPoint", exports["pptx"],
                                   file_name=f"rapport_{partenaire}_{annee}.pptx",
                                   mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                                   key=f"download_pptx_{context}")
    else:
        st.info("Sélectionnez au moins une année pour comparer et exporter.")
