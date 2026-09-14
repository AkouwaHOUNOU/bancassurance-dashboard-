"""
Service de base de donnees pour le dashboard bancassurance NSIA.

APPROCHE: Le fichier Excel source (TABLEAU DE BORD ET PAC PAR PARTENAIRE 2026)
serve uniquement a fournir la liste des partenaires. Toutes les donnees
(objectifs, realisations, specificites, responsables, commissions, etc.)
sont saisies manuellement par l'utilisateur dans l'onglet "Historique &
Comparaison" et persistees dans une base SQLite locale
(historical_bancassurance.db) via le module historical_service.

Cela permet de:
  - Saisir des donnees pour n'importe quelle annee (2021, 2022, ..., 2026)
  - Comparer facilement plusieurs exercices
  - Garder un controle total sur les informations saisies

La liste PARTENAIRES est definie statiquement (identique au fichier Excel source)
afin de ne pas dependre d'une ouverture de fichier pour afficher les noms.
"""

import pandas as pd
from typing import Optional, List, Dict, Any
from datetime import datetime


# ==========================================================================
# Constantes
# ==========================================================================

# Liste des partenaires bancassurance (identique au fichier Excel source).
# Ces noms servent uniquement pour la liste deroulante du dashboard.
PARTENAIRES = [
    "AFRICAN LEASE",
    "BIA TOGO",
    "NSIA BANQUE",
    "ECOBANK",
    "ORABANK",
    "LA POSTE",
    "UTB",
    "IB BANK",
]

# Noms des mois en francais
MOIS = [
    "Janvier", "Février", "Mars", "Avril", "Mai", "Juin",
    "Juillet", "Août", "Septembre", "Octobre", "Novembre", "Décembre",
]

# Mapping mois -> trimestre
MOIS_TRIM = {
    "Janvier": "T1", "Février": "T1", "Mars": "T1",
    "Avril": "T2", "Mai": "T2", "Juin": "T2",
    "Juillet": "T3", "Août": "T3", "Septembre": "T3",
    "Octobre": "T4", "Novembre": "T4", "Décembre": "T4",
}


# ==========================================================================
# Fonctions utilitaires (conservées pour compatibilite descendante)
# ==========================================================================

def result_date_str(val) -> Optional[str]:
    """Convertit une valeur date/datetime en chaine YYYY-MM-DD."""
    if pd.isna(val):
        return None
    if hasattr(val, "strftime"):
        return val.strftime("%Y-%m-%d")
    if isinstance(val, str) and val.strip():
        return val.strip()
    return None


def format_fcfa(valeur: Optional[float]) -> str:
    """Formate un montant en FCFA avec separateurs."""
    if valeur is None or (isinstance(valeur, float) and pd.isna(valeur)):
        return "-"
    return f"{valeur:,.0f} FCFA"


def format_pct(valeur: float) -> str:
    """Formate un pourcentage."""
    return f"{valeur:.1f}%"


def get_mois_from_fiche(mois: str) -> str:
    """Retourne le nom de la feuille FICHE DE SUIVI pour un mois donne.

    NOTE: Cette fonction est conservee pour compatibilite mais n'est plus
    utilisee car les donnees sont saisies manuellement.
    """
    mois_map = {
        "Janvier": "REAL JANVIER 2026",
        "Février": "REAL FEVRIER 2026",
        "Mars": "REAL MARS 2026",
        "Avril": "REAL AVRIL 2026",
        "Mai": "REAL MAI 2026",
        "Juin": "REAL JUIN 2026",
        "Juillet": "REAL JUILLET 2026",
        "Août": "REAL AOUT 2026",
        "Septembre": "REAL SEPTEMBRE 2026",
    }
    return mois_map.get(mois, "")


# ==========================================================================
# Fonctions dynamiques (ajout/suppression de partenaires en session)
# ==========================================================================

# Stockage en mémoire des partenaires supplémentaires créés par l'utilisateur
# lors de la session courante. Ces partenaires sont ajoutés à PARTENAIRES
# et persistés dans la base SQLite via HistoricalService.
_PARTENAIRES_DYNAMIQUES: List[str] = []


def add_partenaire(nom: str) -> List[str]:
    """Ajoute un nouveau partenaire à la liste dynamique.

    Args:
        nom: Nom du partenaire à ajouter.

    Returns:
        Nouvelle liste complète des partenaires (statique + dynamique).
    """
    nom = nom.strip().upper()
    if nom and nom not in PARTENAIRES and nom not in _PARTENAIRES_DYNAMIQUES:
        _PARTENAIRES_DYNAMIQUES.append(nom)
    return PARTENAIRES + _PARTENAIRES_DYNAMIQUES


def get_all_partenaires() -> List[str]:
    """Retourne la liste complète des partenaires (statique + dynamique)."""
    return PARTENAIRES + _PARTENAIRES_DYNAMIQUES


def reset_partenaires_dynamiques() -> None:
    """Vide la liste des partenaires dynamiques (pour réinitialiser la session)."""
    _PARTENAIRES_DYNAMIQUES.clear()


# ==========================================================================
# CLASSE PRINCIPALE - Version legacy (ne charge plus le fichier Excel)
# ==========================================================================
# La classe BancassuranceService est conservee pour compatibilite avec
# l'ancien code, mais toutes les methodes retournent des valeurs par defaut
# (zero) car les donnees sont maintenant saisies via HistoricalService.
# Le dashboard utilise historique_service.HistoricalService pour tout.

class BancassuranceService:
    """
    Service herite de l'ancienne version. NE CHARGE PLUS le fichier Excel.

    Toutes les donnees sont saisies manuellement par l'utilisateur et
    gerées par HistoricalService (base SQLite).

    Cette classe est conservee pour permettre a l'application de demarrer
    sans dependre du fichier Excel source. Les methodes retournent des
    valeurs par défaut (0.0, "", None) car les donnees proviennent désormais
    de la base SQLite via HistoricalService.
    """

    def __init__(self):
        # Ne charge aucun fichier Excel; les donnees proviennent de la base SQLite
        self._fichiers: Dict[str, Any] = {}
        self._loaders: List[str] = []

    # ------------------------------------------------------------------
    # Méthodes de compatibilité (retournent des valeurs par défaut)
    # ------------------------------------------------------------------
    # Toutes les donnees sont saisies via l'onglet "Historique & Comparaison"
    # et stockées dans historical_bancassurance.db (SQLite)

    def get_gap_taux(self, partner: str) -> Dict[str, float]:
        """Retourne des valeurs par defaut (donnees saisies par l'utilisateur)."""
        return {
            "objectif_N": 0.0,
            "realisation_N_1": 0.0,
            "realisation_N": 0.0,
            "gap": 0.0,
            "taux": 0.0,
        }

    def get_objectifs_mensuels(self, partner: str) -> Dict[str, Dict[str, float]]:
        """Retourne des objectifs par defaut (saisis via l'onglet historique)."""
        return {"objectifs": {m: 0.0 for m in MOIS}}

    def get_realisations_mensuelles(self, partner: str) -> Dict[str, float]:
        """Retourne des realisations par defaut (saisis via l'onglet historique)."""
        return {m: 0.0 for m in MOIS}

    def get_commissions_cumulees(self, partner: str) -> Dict[str, Any]:
        """Retourne des commissions par defaut."""
        return {
            "commission_reversee": None,
            "total_a_payer": None,
            "date_traitement": None,
            "date_paiement": None,
        }

    def get_profil_partenaire(self, partner: str) -> Dict[str, Any]:
        """Retourne un profil partenaire vide (saisie manuelle via historique)."""
        return {
            "partenaire": partner,
            "objectif_annuel_2026": 0.0,
            "realisation_2025": 0.0,
            "realisation_2026_t1": 0.0,
            "realisation_2026_t2": 0.0,
            "gap": 0.0,
            "taux_realisation": 0.0,
            "points_forts": "",
            "points_faibles": "",
            "points_focaux": "",
            "dg": "",
            "responsables_bancaires": "",
            "specifications": "",
            "activites_planifiees": "",
            "cibles": "",
            "actions_trimestre": "",
            "personnes_ressources": "",
            "a_fin_septembre": "",
            "difficultes": "",
            "ameliorations": "",
            "recommandations": "",
            "commission_reversee": None,
            "total_a_payer": None,
            "date_traitement": None,
            "date_paiement": None,
        }

    def get_objectifs_hebdomadaires(self, partner: str, mois: str) -> Dict[int, float]:
        """Retourne des objectifs hebdomadaires par defaut."""
        return {}

    def get_tous_partenaires_gap_taux(self) -> pd.DataFrame:
        """Retourne un DataFrame vide pour tous les partenaires."""
        return pd.DataFrame(columns=[
            "Partenaire", "Objectif 2026", "Réalisation 2025",
            "Réalisation 2026", "GAP", "Taux de réalisation (%)",
        ])

    def get_objectifs_mensuels_tous(self) -> pd.DataFrame:
        """Retourne un DataFrame vide."""
        return pd.DataFrame({"Partenaire": PARTENAIRES})

    def get_realisations_tous(self) -> pd.DataFrame:
        """Retourne un DataFrame vide."""
        return pd.DataFrame({"Partenaire": PARTENAIRES})

    def get_commissions_tous(self) -> pd.DataFrame:
        """Retourne un DataFrame vide."""
        return pd.DataFrame({"Partenaire": PARTENAIRES})

    def get_suivi_hebdomadaire(self, partner: str, mois_sheet: str) -> Dict[str, Any]:
        """Retourne un dictionnaire vide."""
        return {}
