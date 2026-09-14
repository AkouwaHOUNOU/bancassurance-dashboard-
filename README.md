# NSIA Bancassurance - Tableau de Bord

Application Streamlit de suivi et d'analyse du chiffre d'affaires bancassurance par partenaire.

## 🚀 Installation rapide

### 1. Prérequis
- Python 3.11+ installé
- pip ou uv disponible

### 2. Installation des dépendances

```bash
# Avec pip
pip install -r requirements.txt

# Ou avec uv (recommandé)
uv pip install -r requirements.txt
```

### 3. Lancement de l'application

```bash
streamlit run app.py
```

L'application s'ouvre automatiquement dans votre navigateur sur `http://localhost:8501`

## 📋 Fonctionnalités

### 1. 📥 Saisie des CA
- Enregistrement du chiffre d'affaires par partenaire, semaine et mois
- Ajout de nouveaux partenaires
- Saisie du responsable bancaire associé
- Historique des dernières saisies

### 2. 📊 Tableau de Bord Général
- CA total annuel
- Nombre de partenaires actifs
- Part de marché par partenaire
- Graphique en barres de répartition du CA

### 3. 📅 Comparaison N vs N-1
- Comparaison mensuelle entre année N et N-1
- Évolution en valeur et en pourcentage
- Tendance par partenaire (croissance/baisse/stable)
- Graphique comparatif groupé

### 4. 🏆 Classement des Partenaires
- Classement par période personnalisable
- Top 3 avec médailles (🥇🥈🥉)
- Part de marché par partenaire
- Graphique en barres et camembert

### 5. 📈 Évolution Annuelle
- Évolution mensuelle du CA par partenaire
- Heatmap du CA par partenaire et par mois
- Courbes d'évolution interactives

### 6. 💾 Export
- Export CSV des CA hebdomadaires
- Export CSV des comparaisons N vs N-1
- Export CSV du classement des partenaires

## 🗄️ Structure de la base de données

L'application utilise SQLite avec 4 tables:
- **partenaires**: Liste des partenaires bancaires
- **ca_hebdomadaire**: CA par semaine, mois et partenaire
- **objectifs_mensuels**: Objectifs mensuels par partenaire
- **notes_partenaire**: Notes et commentaires par partenaire

## 📂 Structure du projet

```
bancassurance_dashboard/
├── app.py                      # Application Streamlit principale
├── database_service.py         # Service de base de données SQLite
├── requirements.txt            # Dépendances Python
├── README.md                   # Ce fichier
├── .streamlit/
│   ├── config.toml            # Configuration Streamlit
│   └── secrets.example.toml   # Exemple de secrets (optionnel)
└── data/
    ├── bancassurance.db       # Base de données SQLite (générée automatiquement)
    └── export_*.csv           # Fichiers d'export CSV
```

## 🎯 Usage

### Première utilisation
1. Lancez l'application: `streamlit run app.py`
2. Dans la sidebar, cliquez sur "🔄 Charger données exemples" pour peupler la base avec les données du fichier Excel source
3. Ou saisissez manuellement les CA dans l'onglet "Saisie CA"

### Saisie de nouvelles données
1. Allez dans l'onglet "📥 Saisie CA"
2. Sélectionnez le partenaire
3. Indiquez l'année, le mois, la semaine
4. Entrez le chiffre d'affaires
5. Cliquez sur "✅ Enregistrer le CA"

### Analyse et comparaison
1. Onglet "📊 Tableau de Bord" pour une vue globale
2. Onglet "📅 Comparaison N vs N-1" pour comparer deux années
3. Onglet "🏆 Classement Partenaires" pour identifier les meilleurs performeurs
4. Onglet "📈 Évolution Annuelle" pour suivre l'évolution mensuelle

### Export
1. Allez dans l'onglet "💾 Export"
2. Sélectionnez l'année et/ou la période
3. Générez le CSV et téléchargez-le

## 🔧 Configuration

### Streamlit Community Cloud (déploiement)
Pour déployer sur Streamlit Community Cloud:
1. Poussez ce dossier sur GitHub
2. Connectez votre repo sur https://streamlit.io/cloud
3. Déployez en sélectionnant `app.py` comme fichier principal
4. Ajoutez les dépendances depuis `requirements.txt`

### Fichier de secrets (optionnel)
Copiez `.streamlit/secrets.example.toml` vers `.streamlit/secrets.toml` si vous avez besoin de secrets (clés API, tokens, etc.)

## 📊 Source des données

Les données exemples sont basées sur le fichier Excel:
`TABLEAU DE BORD ET PAC PAR PARTENAIRE 2026 (2).xlsx`

Partenaires inclus:
- AFRICAN LEASE
- BIA TOGO
- NSIA BANQUE
- ECOBANK
- ORABANK
- LA POSTE
- UTB
- IB BANK

## 🛠️ Développement

### Ajout d'un nouveau partenaire
- Via l'interface: onglet "Saisie CA" → formulaire "Nouveau partenaire"
- Via le code Python:
  ```python
  db.ajouter_partenaire("NOM_PARTENAIRE", "Responsable")
  ```

### Ajout d'une nouvelle métrique
Éditez `database_service.py` pour ajouter une nouvelle table et requête SQL, puis ajoutez l'onglet correspondant dans `app.py`.

## 📝 Notes

- Les valeurs sont stockées en FCFA (Franc CFA)
- La base de données est locale (SQLite) et stockée dans `data/bancassurance.db`
- Les exports CSV sont encodés en UTF-8 avec BOM pour compatibilité Excel
- Les données exemples sont des approximations basées sur les totaux trimestriels du fichier Excel source

## 👨‍💻 Auteur

Développé pour le mémoire NSIA NON-VIE - Licence Pro IA & Big Data
HOUNOU Akouwa Lumière Elisée Éliane
