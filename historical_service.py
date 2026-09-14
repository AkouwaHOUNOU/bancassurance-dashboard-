"""
Service SQLite pour le stockage et la comparaison d'historiques bancassurance.

Permet de saisir des objectifs et réalisations par partenaire, par annee et par mois,
ainsi que des champs textuels (profil, difficultés, recommandations, DG, etc.) et
des données de suivi hebdomadaire. Toutes ces données sont persistées dans une base
SQLite locale afin de garder un historique exploitable pour les comparaisons
inter-annees (ex: 2021 vs 2026).

Tables:
  - historical_data: (partenaire, annee, mois, type, valeur)
    type = 'objectif' ou 'realisation' (montant en FCFA)
  - textual_data: (partenaire, annee, champ, valeur)
    champ = 'points_forts', 'points_faibles', 'points_focaux', 'dg',
            'activites_planifiees', 'cibles', 'personnes_ressources',
            'difficultes', 'ameliorations', 'recommandations', etc.
  - weekly_data: (partenaire, annee, mois, semaine, objectif, realisation, commentaire)

Usage:
  svc = HistoricalService()
  svc.save_value("AFRICAN LEASE", 2021, "Janvier", "objectif", 24307056)
  svc.save_value("AFRICAN LEASE", 2021, "Janvier", "realisation", 11022530)
  svc.save_text("AFRICAN LEASE", 2021, "dg", "YOMA MANIWA SIKA")
  svc.get_data_for_partner("AFRICAN LEASE", 2021)
  svc.get_all_annee("AFRICAN LEASE")
"""

import sqlite3
import os
from typing import Dict, List, Optional
import pandas as pd

DB_PATH = os.path.join(
    os.path.dirname(os.path.abspath(__file__)),
    "historical_bancassurance.db",
)


class HistoricalService:
    """Gère la persistance SQLite des donnees historiques saisies par l'utilisateur."""

    def __init__(self, db_path: str = DB_PATH):
        self.db_path = db_path
        self._init_db()

    def _get_conn(self) -> sqlite3.Connection:
        """Retourne une connexion SQLite avec row_factory active."""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        return conn

    def _init_db(self):
        """Cree les tables principales si elles n'existent pas encore.

        Schemas:
          historical_data:
            - partenaire: TEXT
            - annee: INTEGER
            - mois: TEXT ("Janvier".."Décembre" ou "ANNUEL")
            - type: TEXT ("objectif" | "realisation")
            - valeur: REAL (montant en FCFA)
            PK: (partenaire, annee, mois, type)

          textual_data:
            - partenaire: TEXT
            - annee: INTEGER
            - champ: TEXT (nom du champ textuel)
            - valeur: TEXT
            PK: (partenaire, annee, champ)

          weekly_data:
            - partenaire: TEXT
            - annee: INTEGER
            - mois: TEXT
            - semaine: INTEGER (1..5)
            - objectif: REAL
            - realisation: REAL
            - commentaire: TEXT
            PK: (partenaire, annee, mois, semaine)
        """
        conn = self._get_conn()
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS historical_data (
                partenaire  TEXT,
                annee       INTEGER,
                mois        TEXT,
                type        TEXT,
                valeur      REAL,
                PRIMARY KEY (partenaire, annee, mois, type)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS textual_data (
                partenaire  TEXT,
                annee       INTEGER,
                champ       TEXT,
                valeur      TEXT,
                PRIMARY KEY (partenaire, annee, champ)
            )
            """
        )
        conn.execute(
            """
            CREATE TABLE IF NOT EXISTS weekly_data (
                partenaire  TEXT,
                annee       INTEGER,
                mois        TEXT,
                semaine     INTEGER,
                objectif    REAL,
                realisation REAL,
                commentaire TEXT,
                PRIMARY KEY (partenaire, annee, mois, semaine)
            )
            """
        )
        conn.commit()
        conn.close()

    # ------------------------------------------------------------------
    # Sauvegarde
    # ------------------------------------------------------------------
    def save_value(
        self,
        partenaire: str,
        annee: int,
        mois: str,
        type_valeur: str,
        valeur: float,
    ) -> None:
        """Insere ou met a jour une valeur (objectif ou realisation).

        Args:
            partenaire: nom du partenaire (ex: "AFRICAN LEASE")
            annee: annee concernee (ex: 2021)
            mois: nom du mois ou "ANNUEL" pour un total annuel
            type_valeur: "objectif" ou "realisation"
            valeur: montant en FCFA
        """
        conn = self._get_conn()
        conn.execute(
            """
            INSERT OR REPLACE INTO historical_data (partenaire, annee, mois, type, valeur)
            VALUES (?, ?, ?, ?, ?)
            """,
            (partenaire, annee, mois, type_valeur, float(valeur)),
        )
        conn.commit()
        conn.close()

    def save_values_batch(
        self,
        partenaire: str,
        annee: int,
        data: Dict[str, Dict[str, float]],
    ) -> None:
        """Sauvegarde un lot de valeurs pour un partenaire et une annee.

        Args:
            data: {mois: {"objectif": x, "realisation": y}}
                  peut contenir la clé "ANNUEL" pour un total annuel.
        """
        conn = self._get_conn()
        for mois, types in data.items():
            for type_valeur, valeur in types.items():
                conn.execute(
                    """
                    INSERT OR REPLACE INTO historical_data
                    (partenaire, annee, mois, type, valeur)
                    VALUES (?, ?, ?, ?, ?)
                    """,
                    (partenaire, annee, mois, type_valeur, float(valeur)),
                )
        conn.commit()
        conn.close()

    # ------------------------------------------------------------------
    # Lecture
    # ------------------------------------------------------------------
    def get_all_annees(self, partenaire: str) -> List[int]:
        """Retourne la liste des annees pour lesquelles des donnees existent."""
        conn = self._get_conn()
        rows = conn.execute(
            "SELECT DISTINCT annee FROM historical_data WHERE partenaire = ? ORDER BY annee",
            (partenaire,),
        ).fetchall()
        conn.close()
        return [r["annee"] for r in rows]

    def get_data_for_partner_annee(
        self, partenaire: str, annee: int
    ) -> Dict[str, Dict[str, Optional[float]]]:
        """Retourne un dictionnaire {mois: {"objectif": x, "realisation": y}} pour un partenaire et une annee."""
        conn = self._get_conn()
        rows = conn.execute(
            """
            SELECT mois, type, valeur FROM historical_data
            WHERE partenaire = ? AND annee = ?
            ORDER BY mois
            """,
            (partenaire, annee),
        ).fetchall()
        conn.close()

        MOIS_ORDER = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
            "ANNUEL",
        ]
        result: Dict[str, Dict[str, Optional[float]]] = {}
        for r in rows:
            mois = r["mois"]
            type_v = r["type"]
            if mois not in result:
                result[mois] = {"objectif": None, "realisation": None}
            result[mois][type_v] = r["valeur"]

        # Reordonner selon MOIS_ORDER
        ordered = {k: result[k] for k in MOIS_ORDER if k in result}
        return ordered

    def get_comparison_dataframe(
        self, partenaire: str, annees: List[int]
    ) -> pd.DataFrame:
        """Construit un DataFrame de comparaison inter-annees.

        Colonnes: Mois, Objectif <annee>, Réalisation <annee>, Taux <annee>, ...
        Inclut une ligne "ANNUEL" si presente.
        """
        MOIS_ORDER = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
        ]

        data = {m: {} for m in MOIS_ORDER + ["ANNUEL"]}
        for annee in annees:
            raw = self.get_data_for_partner_annee(partenaire, annee)
            for mois in MOIS_ORDER + ["ANNUEL"]:
                entry = raw.get(mois, {"objectif": None, "realisation": None})
                obj = entry.get("objectif")
                real = entry.get("realisation")
                taux = None
                if obj is not None and obj > 0 and real is not None:
                    taux = (real / obj) * 100
                data[mois][f"Objectif {annee}"] = obj if obj is not None else None
                data[mois][f"Réalisation {annee}"] = real if real is not None else None
                data[mois][f"Taux {annee}"] = round(taux, 1) if taux is not None else None

        rows = []
        for mois in MOIS_ORDER + ["ANNUEL"]:
            row = {"Mois": mois}
            row.update(data[mois])
            rows.append(row)

        return pd.DataFrame(rows)

    # ------------------------------------------------------------------
    # Données textuelles (profil, difficultés, recommandations, DG, etc.)
    # ------------------------------------------------------------------
    def save_text(
        self,
        partenaire: str,
        annee: int,
        champ: str,
        valeur: str,
    ) -> None:
        """Insere ou met a jour une valeur textuelle pour un partenaire et une annee.

        Args:
            partenaire: nom du partenaire
            annee: année concernée
            champ: nom du champ (points_forts, points_faibles, dg,
                   activites_planifiees, cibles, difficultes, ameliorations,
                   recommandations, personne_ressources, etc.)
            valeur: texte saisi par l'utilisateur
        """
        conn = self._get_conn()
        conn.execute(
            """
            INSERT OR REPLACE INTO textual_data (partenaire, annee, champ, valeur)
            VALUES (?, ?, ?, ?)
            """,
            (partenaire, annee, champ, valeur),
        )
        conn.commit()
        conn.close()

    def get_text_for_partner_annee(
        self, partenaire: str, annee: int
    ) -> Dict[str, str]:
        """Retourne tous les champs textuels pour un partenaire et une annee."""
        conn = self._get_conn()
        rows = conn.execute(
            """
            SELECT champ, valeur FROM textual_data
            WHERE partenaire = ? AND annee = ?
            ORDER BY champ
            """,
            (partenaire, annee),
        ).fetchall()
        conn.close()
        return {r["champ"]: r["valeur"] for r in rows}

    # ------------------------------------------------------------------
    # Données hebdomadaires
    # ------------------------------------------------------------------
    def save_weekly(
        self,
        partenaire: str,
        annee: int,
        mois: str,
        semaine: int,
        objectif: float,
        realisation: float,
        commentaire: str = "",
    ) -> None:
        """Insere ou met a jour une entree hebdomadaire."""
        conn = self._get_conn()
        conn.execute(
            """
            INSERT OR REPLACE INTO weekly_data
            (partenaire, annee, mois, semaine, objectif, realisation, commentaire)
            VALUES (?, ?, ?, ?, ?, ?, ?)
            """,
            (partenaire, annee, mois, semaine, float(objectif), float(realisation), commentaire),
        )
        conn.commit()
        conn.close()

    def get_weekly_for_partner_mois(
        self, partenaire: str, annee: int, mois: str
    ) -> Dict[int, Dict[str, any]]:
        """Retourne les donnees hebdomadaires d'un partenaire pour un mois et annee."""
        conn = self._get_conn()
        rows = conn.execute(
            """
            SELECT semaine, objectif, realisation, commentaire
            FROM weekly_data
            WHERE partenaire = ? AND annee = ? AND mois = ?
            ORDER BY semaine
            """,
            (partenaire, annee, mois),
        ).fetchall()
        conn.close()
        return {
            r["semaine"]: {
                "objectif": r["objectif"],
                "realisation": r["realisation"],
                "commentaire": r["commentaire"] or "",
            }
            for r in rows
        }

    # ------------------------------------------------------------------
    # Profil partenaire complet (numérique + textuel)
    # ------------------------------------------------------------------
    def get_profil_partenaire(self, partner: str, annee: int = 2026) -> Dict[str, any]:
        """Construit un profil complet d'un partenaire pour une annee donnee.

        Combine les donnees numeriques (objectifs, realisations, commissions)
        et les champs textuels (profil, difficultés, recommandations, etc.)
        saisis par l'utilisateur via l'interface Streamlit.
        """
        # Données numériques
        raw = self.get_data_for_partner_annee(partner, annee)
        text = self.get_text_for_partner_annee(partner, annee)

        # Calculs GAP/Taux
        annee_entry = raw.get("ANNUEL", {})
        obj_annuel = annee_entry.get("objectif") or 0.0
        real_annuel = annee_entry.get("realisation") or 0.0
        gap_data = self.get_gap_taux_historical(partner, annee)

        profil = {
            "partenaire": partner,
            "annee": annee,
            "objectif_annuel": float(obj_annuel),
            "realisation_annuelle": float(real_annuel),
            "gap": gap_data["gap"],
            "taux_realisation": gap_data["taux"],
            "points_forts": text.get("points_forts", ""),
            "points_faibles": text.get("points_faibles", ""),
            "points_focaux": text.get("points_focaux", ""),
            "dg": text.get("dg", ""),
            "responsables_bancaires": text.get("responsables_bancaires", ""),
            "specifications": text.get("specifications", ""),
            "activites_planifiees": text.get("activites_planifiees", ""),
            "cibles": text.get("cibles", ""),
            "actions_trimestre": text.get("actions_trimestre", ""),
            "personnes_ressources": text.get("personnes_ressources", ""),
            "a_fin_septembre": text.get("a_fin_septembre", ""),
            "difficultes": text.get("difficultes", ""),
            "ameliorations": text.get("ameliorations", ""),
            "recommandations": text.get("recommandations", ""),
            "commission_reversee": text.get("commission_reversee"),
            "total_a_payer": text.get("total_a_payer"),
            "date_traitement": text.get("date_traitement"),
            "date_paiement": text.get("date_paiement"),
        }
        return profil

    def get_gap_taux_historical(self, partner: str, annee: int) -> Dict[str, float]:
        """Calcule le GAP et le taux de realisation depuis la base historique.

        - Objectif = total des objectifs mensuels + objectif annuel si present
        - Realisation = total des realisations mensuelles + realisation annuelle
        """
        raw = self.get_data_for_partner_annee(partner, annee)
        annee_entry = raw.get("ANNUEL", {})
        obj_annuel = float(annee_entry.get("objectif") or 0)
        real_annuel = float(annee_entry.get("realisation") or 0)

        # Si l'objectif annuel n'est pas saisi, calculer a partir des mensuels
        if obj_annuel == 0:
            obj_annuel = sum(
                raw.get(m, {}).get("objectif") or 0
                for m in [
                    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
                ]
            )

        # Si la realisation annuelle n'est pas saisie, calculer a partir des mensuels
        if real_annuel == 0:
            real_annuel = sum(
                raw.get(m, {}).get("realisation") or 0
                for m in [
                    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
                    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
                ]
            )

        gap = obj_annuel - real_annuel
        taux = (real_annuel / obj_annuel * 100) if obj_annuel > 0 else 0.0

        return {
            "objectif_N": obj_annuel,
            "realisation_N_1": 0.0,   # pas de données N-1 dans l'historique libre
            "realisation_N": real_annuel,
            "gap": gap,
            "taux": taux,
        }

    def get_tous_partenaires_gap_taux(self, annee: int = 2026) -> pd.DataFrame:
        """Retourne un DataFrame avec GAP et Taux pour tous les partenaires d'une annee."""
        from database_service import PARTENAIRES  # éviter import circulaire
        rows = []
        for p in PARTENAIRES:
            data = self.get_gap_taux_historical(p, annee)
            rows.append({
                "Partenaire": p,
                "Objectif": data["objectif_N"],
                "Réalisation": data["realisation_N"],
                "GAP": data["gap"],
                "Taux de réalisation (%)": round(data["taux"], 1),
            })
        return pd.DataFrame(rows)

    def remove_annee(self, partenaire: str, annee: int) -> None:
        """Supprime toutes les données (numeriques, textuelles, hebdomadaires)
        pour un partenaire et une annee donnee.
        """
        conn = self._get_conn()
        conn.execute(
            "DELETE FROM historical_data WHERE partenaire = ? AND annee = ?",
            (partenaire, annee),
        )
        conn.execute(
            "DELETE FROM textual_data WHERE partenaire = ? AND annee = ?",
            (partenaire, annee),
        )
        conn.execute(
            "DELETE FROM weekly_data WHERE partenaire = ? AND annee = ?",
            (partenaire, annee),
        )
        conn.commit()
        conn.close()

    def delete_text_field(
        self,
        partenaire: str,
        annee: int,
        champ: str,
    ) -> None:
        """Supprime un champ textuel specifique pour un partenaire et une annee.

        Args:
            partenaire: nom du partenaire
            annee: annee concernee
            champ: nom du champ a supprimer (ex: "dg", "points_forts", etc.)
        """
        conn = self._get_conn()
        conn.execute(
            "DELETE FROM textual_data WHERE partenaire = ? AND annee = ? AND champ = ?",
            (partenaire, annee, champ),
        )
        conn.commit()
        conn.close()

    def calculate_and_save_annee_totale(self, partenaire: str, annee: int) -> tuple:
        """Calcule automatiquement l'objectif et realisation annuels a partir
        des donnees mensuelles, puis les sauvegarde dans la table historical_data
        avec mois="ANNUEL".

        Args:
            partenaire: nom du partenaire
            annee: annee concernée

        Returns:
            Tuple (objectif_annuel, realisation_annuelle) sauvegardés en base.
        """
        raw = self.get_data_for_partner_annee(partenaire, annee)
        MOIS_LIST = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
        ]

        total_obj = sum(raw.get(m, {}).get("objectif") or 0 for m in MOIS_LIST)
        total_real = sum(raw.get(m, {}).get("realisation") or 0 for m in MOIS_LIST)

        self.save_value(partenaire, annee, "ANNUEL", "objectif", total_obj)
        self.save_value(partenaire, annee, "ANNUEL", "realisation", total_real)

        return total_obj, total_real

    def get_weekly_comparison_dataframe(self, partner: str, annees: List[int], mois: str) -> pd.DataFrame:
        """Construit un DataFrame de comparaison hebdomadaire inter-annees pour un mois donne.

        Colonnes: Semaine, Objectif <annee>, Réalisation <annee>, Taux <annee>, ...
        """
        MOIS_LIST = [
            "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
            "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
        ]

        if mois not in MOIS_LIST:
            return pd.DataFrame()

        from calendar import monthcalendar
        from database_service import MOIS as MOIS_ORDER

        mois_idx = MOIS_ORDER.index(mois) + 1
        cal = monthcalendar(annees[0] if annees else 2026, mois_idx)
        nb_semaines = len([w for w in cal if w[4] != 0])

        data = {s: {} for s in range(1, nb_semaines + 1)}
        for annee in annees:
            weekly = self.get_weekly_for_partner_mois(partner, annee, mois)
            for s in range(1, nb_semaines + 1):
                entry = weekly.get(s, {})
                obj = entry.get("objectif")
                real = entry.get("realisation")
                taux = None
                if obj is not None and obj > 0 and real is not None:
                    taux = (real / obj) * 100
                data[s][f"Objectif {annee}"] = obj
                data[s][f"Réalisation {annee}"] = real
                data[s][f"Taux {annee}"] = round(taux, 1) if taux is not None else None

        rows = []
        for s in range(1, nb_semaines + 1):
            row = {"Semaine": f"S{s}"}
            row.update(data[s])
            rows.append(row)

        return pd.DataFrame(rows)

    def get_objectifs_hebdomadaires(self, partner: str, mois: str, annee: int = 2026) -> Dict[int, float]:
        """Retourne les objectifs hebdomadaires pour un partenaire, un mois et une annee.

        Les objectifs hebdomadaires sont calculés à partir de l'objectif
        mensuel divisé par le nombre de semaines du mois.
        """
        result = {}
        raw = self.get_data_for_partner_annee(partner, annee)
        mois_data = raw.get(mois, {})
        obj_mensuel = float(mois_data.get("objectif") or 0)

        from calendar import monthcalendar
        from database_service import MOIS as MOIS_LIST
        mois_idx = MOIS_LIST.index(mois) + 1
        cal = monthcalendar(annee, mois_idx)
        nb_semaines = len([w for w in cal if w[4] != 0])

        if obj_mensuel > 0 and nb_semaines > 0:
            obj_hebdo = round(obj_mensuel / nb_semaines, 2)
            for s in range(1, nb_semaines + 1):
                result[s] = obj_hebdo

        return result

    def get_suivi_hebdomadaire(self, partner: str, mois: str, annee: int = 2026) -> Dict[str, any]:
        """Retourne un dictionnaire de suivi hebdomadaire formaté.

        Structure utilisee par l'ancien onglet 'État des lieux':
            {
                "semaines": {"S1": {"objectif": x, "realisation": y}, ...},
                "objectif_mois": x,
                "realisation_mois": y,
                "total_objectif_cumule": x,
                "total_realisation_cumule": y,
                "actions_phare": "...",
                "niveau_rea_commentaires": "...",
            }
        """
        weekly = self.get_weekly_for_partner_mois(partner, annee, mois)
        raw = self.get_data_for_partner_annee(partner, annee)
        mois_data = raw.get(mois, {})

        semaines = {}
        for s, d in weekly.items():
            semaines[f"S{s}"] = {
                "objectif": d["objectif"],
                "realisation": d["realisation"],
            }

        return {
            "semaines": semaines,
            "objectif_mois": float(mois_data.get("objectif") or 0),
            "realisation_mois": float(mois_data.get("realisation") or 0),
            "total_objectif_cumule": float(mois_data.get("objectif") or 0),
            "total_realisation_cumule": float(mois_data.get("realisation") or 0),
            "actions_phare": "",
            "niveau_rea_commentaires": "",
        }

    def get_commissions_tous(self) -> pd.DataFrame:
        """Retourne un DataFrame avec les commissions pour tous les partenaires.

        Les commissions sont sauvegardées via save_text avec les clés
        'commission_reversee' et 'total_a_payer'.
        """
        from database_service import PARTENAIRES
        rows = []
        for p in PARTENAIRES:
            text = self.get_text_for_partner_annee(p, 2026)
            rows.append({
                "Partenaire": p,
                "Commission reversée": text.get("commission_reversee"),
                "Total à payer": text.get("total_a_payer"),
                "Date traitement": text.get("date_traitement"),
                "Date paiement": text.get("date_paiement"),
            })
        return pd.DataFrame(rows)

    # ------------------------------------------------------------------
    # ------------------------------------------------------------------
    def export_to_excel(
        self, partenaire: str, annees: List[int], path: str
    ) -> str:
        """Exporte les donnees historiques d'un partenaire vers un fichier Excel."""
        df = self.get_comparison_dataframe(partenaire, annees)
        df.to_excel(path, index=False, sheet_name="Historique")
        return path
