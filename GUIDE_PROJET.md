# Guide Complet – Collateral Management Tool (Natixis SOU)

> Ce guide couvre l'intégralité du projet de A à Z : contexte financier, architecture technique,
> logique métier, résultats, et préparation aux questions d'entretien.

---

## 1. Le collateral management : comprendre le contexte

### 1.1 Qu'est-ce que le collateral ?

Le **collateral** (ou garantie financière) est un actif financier — obligation, action, ETF — qu'une contrepartie remet à une autre pour sécuriser une exposition de crédit. Dans les opérations de marché (repos, prêts/emprunts de titres, swaps de taux ou de change), le collateral fonctionne comme un dépôt de garantie : si la contrepartie fait défaut, le créancier peut liquider les actifs reçus pour se rembourser.

**Exemple concret :** Natixis prête des liquidités à un hedge fund. Pour s'assurer que le hedge fund remboursera, Natixis exige qu'il poste des obligations d'État en garantie. Si le hedge fund fait défaut, Natixis conserve et vend ces obligations.

### 1.2 Pourquoi gérer le collateral ?

Avant 2008, une grande partie des opérations de gré à gré (OTC) se faisait sans garantie ou avec des garanties très légères. La crise a montré le danger de cette pratique (ex. : faillite de Lehman Brothers). En réponse, les régulateurs ont imposé :

- **EMIR** (European Market Infrastructure Regulation, 2012) : compensation centrale obligatoire des dérivés standardisés via des contreparties centrales (CCP comme LCH ou Eurex), et collatéralisation bilatérale des dérivés OTC non compensés centralement.
- **Basel III** : exigences de capital et de liquidité accrues, notamment via les ratios LCR (Liquidity Coverage Ratio) et NSFR, qui valorisent les actifs liquides de haute qualité (HQLA).
- **SFDR** / **SFTR** : transparence sur les opérations de financement sur titres (repos, prêts de titres).

**Conséquence directe** : les besoins en collateral ont explosé post-2008. Les banques doivent gérer des centaines de milliers d'opérations simultanément, en s'assurant que chaque garantie est :
- **Éligible** : acceptée contractuellement par la contrepartie
- **Suffisante** : d'une valeur éligible couvrant l'exposition
- **Disponible** : non déjà utilisée ailleurs

C'est dans ce contexte qu'opère la **Securities Optimisation Unit (SOU)** de Natixis CIB, dont la mission est de minimiser le coût du collateral posté tout en satisfaisant toutes les contraintes.

### 1.3 Le problème de la maturité

Un titre posté en collateral a une date de maturité. Quand cette date approche, le titre n'a plus de valeur dans le temps et ne peut plus servir de garantie. Il faut donc le **substituer** : trouver un autre titre du portefeuille disponible qui joue le même rôle.

**Seuil d'alerte dans ce projet : 30 jours avant maturité.**

Si on attend le dernier moment, on risque de ne pas trouver de substitut à temps, ce qui entraîne un **margin call non satisfait** (défaut de collateral) → pénalités contractuelles, dégradation de la relation avec la contrepartie, voire problème de compliance.

---

## 2. La logique CTD (Cheapest-to-Deliver)

### 2.1 Principe

CTD signifie **Cheapest-to-Deliver** : lorsqu'on a le choix entre plusieurs actifs à poster, on choisit en priorité le **moins coûteux à immobiliser**. Poster un actif a un coût d'opportunité : cet actif est bloqué, on ne peut pas l'utiliser ailleurs (le repo, le vendre, etc.). L'objectif est de minimiser ce coût en gardant les actifs les plus précieux disponibles le plus longtemps possible.

### 2.2 Hiérarchie CTD dans cet outil

| Rang | Type d'actif | Score CTD | Raisonnement |
|------|-------------|-----------|-------------|
| 1 (poster en premier) | Corp HY | 1 | Obligations High Yield : les plus risquées, avec les plus gros haircuts, peu demandées. Les "perdre" coûte le moins. |
| 2 | ETF Actions | 2 | Volatiles, décote élevée (haircut 20%). Moins précieux que des obligations. |
| 3 | Corp IG | 3 | Investment Grade : bonne qualité mais moins liquides que les souverains. |
| 4 (garder le plus longtemps) | Souverain EU | 4 | OAT, Bunds : actifs les plus liquides, les plus demandés sur le marché des repos. Coût d'opportunité maximal à immobiliser. |

**Score 1 = poster en premier** (cheap to deliver, on s'en débarrasse en priorité)
**Score 4 = conserver** (les souverains sont précieux, on les réserve)

### 2.3 Le haircut

Le **haircut** est une décote prudentielle appliquée à la valeur de marché pour obtenir la **valeur éligible** (la valeur reconnue comme garantie) :

```
Valeur de Marché (VM)  = Valeur_Nominale × (Prix_Pct / 100)
Valeur Éligible (VE)   = VM × (1 - Haircut_Pct / 100)
```

**Exemples :**
- OAT (souverain, haircut 3%) : 1 000 000 € nominale × 100% prix × 0.97 = **970 000 € de VE**
- Corp HY (haircut 15%) : 1 000 000 € × 91.5% × 0.85 = **777 750 € de VE**

Le haircut reflète le risque de marché : si la contrepartie fait défaut et qu'on doit liquider le collateral en urgence, le marché peut avoir bougé défavorablement. Un actif volatile nécessite une décote plus élevée.

**Subtilité CTD** : à score égal (même type d'actif), l'outil préfère le substitut avec le **plus fort haircut** (il est moins efficace comme collateral → moins précieux → moins cher à poster).

### 2.4 L'éligibilité

Chaque titre du portefeuille disponible n'est pas forcément accepté par toutes les contreparties. Le champ `Eligibilite` du CSV liste les contreparties qui acceptent ce titre (au titre d'accords cadres ISDA/GMRA). Avant de proposer un substitut, l'outil vérifie que la contrepartie du titre sortant figure dans cette liste.

---

## 3. Architecture du projet

### 3.1 Structure des fichiers

```
collateral_tool/
├── main.py                       ← Point d'entrée, orchestration
├── data/
│   ├── collateral_poste.csv      ← 12 titres actuellement postés en collateral
│   └── portefeuille_dispo.csv    ← 14 titres disponibles pour substitution
├── engine/
│   ├── detector.py               ← Détection des maturités < 30 jours
│   ├── substitutor.py            ← Logique CTD et sélection des substituts
│   └── reporter.py               ← Génération du rapport Excel (4 onglets)
└── output/
    └── collateral_report.xlsx    ← Rapport généré automatiquement
```

### 3.2 Flux de données de bout en bout

```
collateral_poste.csv  +  portefeuille_dispo.csv
         │                         │
         ▼                         ▼
   [detector.py]            [substitutor.py]
   charger_portefeuille      charger_disponible()
   _poste()                  → calcule VE, Score CTD
   → calcule VM, VE          → initialise VE_Dispo
   → calcule Jours_Restants  │
         │                   │
         ▼                   │
   detecter_maturites        │
   _proches()                │
   → filtre < 30 jours       │
         │                   │
         └────────┬──────────┘
                  ▼
          construire_alertes()
          → trie par urgence (jours croissants)
          → appelle trouver_substitut() pour chaque alerte
          → gère l'allocation (anti-double comptage)
                  │
                  ▼
            [reporter.py]
            generer_rapport()
            → 4 onglets Excel
                  │
                  ▼
          collateral_report.xlsx
```

---

## 4. Module par module : le code expliqué

### 4.1 detector.py — Détection des alertes

```python
def charger_portefeuille_poste(chemin_csv: str) -> pd.DataFrame:
    df = pd.read_csv(chemin_csv, parse_dates=["Maturite"])
    aujourd_hui = date.today()

    df["Valeur_Marche"]   = df["Valeur_Nominale"] * df["Prix_Pct"] / 100
    df["Valeur_Eligible"] = df["Valeur_Marche"] * (1 - df["Haircut_Pct"] / 100)
    df["Jours_Restants"]  = (df["Maturite"].dt.date - aujourd_hui).apply(lambda d: d.days)

    return df

def detecter_maturites_proches(df: pd.DataFrame) -> pd.DataFrame:
    return df[df["Jours_Restants"] < SEUIL_ALERTE_JOURS].copy()
```

**Ligne par ligne :**
1. `pd.read_csv(..., parse_dates=["Maturite"])` : charge le CSV et convertit automatiquement la colonne Maturite en datetime pandas.
2. `Valeur_Marche = Nominale × Prix/100` : prix coté en % du nominal, donc on divise par 100.
3. `Valeur_Eligible = VM × (1 - Haircut/100)` : applique la décote prudentielle.
4. `Jours_Restants = date maturité - aujourd'hui` : calcul en jours calendaires. `dt.date` extrait la date pure (sans heure) du timestamp pandas, nécessaire pour comparer avec `date.today()`.
5. Filtre final : on ne garde que les titres avec moins de 30 jours (`SEUIL_ALERTE_JOURS = 30`).

---

### 4.2 substitutor.py — Moteur de substitution CTD

**Constante de score :**
```python
SCORE_CTD: dict[str, int] = {
    "Corp HY":      1,
    "ETF Actions":  2,
    "Corp IG":      3,
    "Souverain EU": 4,
}
```

**Chargement du disponible :**
```python
def charger_disponible(chemin_csv: str) -> pd.DataFrame:
    df = pd.read_csv(chemin_csv, parse_dates=["Maturite"])
    df["Valeur_Marche"]         = df["Valeur_Nominale"] * df["Prix_Pct"] / 100
    df["Valeur_Eligible"]       = df["Valeur_Marche"] * (1 - df["Haircut_Pct"] / 100)
    df["Score_CTD"]             = df["Type"].map(SCORE_CTD)
    df["Valeur_Eligible_Dispo"] = df["Valeur_Eligible"].copy()
    return df
```

La colonne `Valeur_Eligible_Dispo` est fondamentale : elle représente la **capacité résiduelle** de chaque substitut. Elle est initialisée à la VE totale et décrémentée à chaque allocation.

**Vérification d'éligibilité :**
```python
def _est_eligible(eligibilite_str: str, contrepartie: str) -> bool:
    noms = [x.strip() for x in str(eligibilite_str).split(";")]
    return contrepartie in noms
```

Le champ Eligibilite du CSV est de la forme `"BNP Paribas;Goldman Sachs;HSBC London"`. On split sur `;`, on strip les espaces, et on vérifie que la contrepartie est dans la liste.

**Algorithme principal — `trouver_substitut()` :**
```python
def trouver_substitut(sortant: pd.Series, dispo: pd.DataFrame) -> dict:
    contrepartie = sortant["Contrepartie"]
    valeur_cible = sortant["Valeur_Eligible"]

    # Étape 1 : filtrer par éligibilité ET capacité résiduelle > 0
    masque = (
        dispo["Eligibilite"].apply(lambda e: _est_eligible(e, contrepartie))
        & (dispo["Valeur_Eligible_Dispo"] > 0)
    )
    eligibles = dispo[masque].copy()

    if eligibles.empty:
        return _resultat_vide(sortant)  # AUCUN SUBSTITUT

    # Étape 2 : trier CTD (score croissant = moins cher d'abord)
    # À score égal : haircut décroissant (moins efficace = moins précieux)
    eligibles = eligibles.sort_values(
        ["Score_CTD", "Haircut_Pct"], ascending=[True, False]
    )

    # Étape 3 : chercher une couverture complète
    for idx, sub in eligibles.iterrows():
        if sub["Valeur_Eligible_Dispo"] >= valeur_cible:
            dispo.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible
            return _construire_resultat(sortant, sub, "SUBSTITUT TROUVE", valeur_cible)

    # Étape 4 : aucune couverture complète → prendre le meilleur (partiel)
    meilleur_idx = eligibles.index[0]
    meilleur     = eligibles.iloc[0]
    ve_allouee   = meilleur["Valeur_Eligible_Dispo"]
    dispo.at[meilleur_idx, "Valeur_Eligible_Dispo"] = 0
    return _construire_resultat(sortant, meilleur, "COUVERTURE PARTIELLE", ve_allouee)
```

**Points clés :**
- `dispo.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible` : **mutation en place** du DataFrame. Comme `dispo` est passé par référence en Python, chaque allocation est immédiatement visible pour les appels suivants. C'est ce qui garantit l'absence de double comptage.
- La boucle `for idx, sub in eligibles.iterrows()` cherche **séquentiellement** dans l'ordre CTD optimal. Dès qu'un candidat peut couvrir entièrement, on l'alloue et on s'arrête.

**Priorité aux cas urgents — `construire_alertes()` :**
```python
def construire_alertes(df_alerte: pd.DataFrame, dispo: pd.DataFrame) -> pd.DataFrame:
    df_tri = df_alerte.sort_values("Jours_Restants").reset_index(drop=True)
    lignes = [trouver_substitut(row, dispo) for _, row in df_tri.iterrows()]
    return pd.DataFrame(lignes)
```

On trie les alertes par `Jours_Restants` **croissants** : le titre qui expire dans 6 jours passe avant celui qui expire dans 18 jours. Si un substitut est rare (capacité limitée), il va au cas le plus urgent.

---

### 4.3 reporter.py — Rapport Excel professionnel

Le rapport est généré via **openpyxl**, la librairie Python de référence pour manipuler des fichiers Excel sans avoir besoin de Microsoft Office installé.

**Structure :**
```python
def generer_rapport(df_poste, df_dispo, df_substitutions, chemin_sortie):
    wb = Workbook()
    wb.remove(wb.active)  # Supprime la feuille vide par défaut

    ws1 = wb.create_sheet("Tableau de Bord")       # KPIs globaux
    ws2 = wb.create_sheet("Alertes & Substitutions") # Résultats CTD
    ws3 = wb.create_sheet("Portefeuille Poste")     # Détail posté
    ws4 = wb.create_sheet("Disponible")             # Détail disponible

    wb.save(chemin_sortie)
```

**4 onglets :**
1. **Tableau de Bord** : KPIs (exposition totale, nb titres, alertes, taux de couverture), répartition du portefeuille par type d'actif.
2. **Alertes & Substitutions** : résultats CTD avec code couleur (vert = substitut trouvé, orange = partiel, rouge = aucun).
3. **Portefeuille Posté** : tous les titres postés avec leur statut d'alerte (URGENT < 7j, ATTENTION < 30j, OK).
4. **Disponible** : portefeuille disponible trié par Score CTD (du moins cher au plus cher).

**Technique de formatage :**
- `PatternFill` pour les couleurs de fond des cellules
- `Font` pour la police, le gras, la couleur du texte
- `Alignment` pour centrage, indentation
- `number_format` pour les formats monétaires (`#,##0`) et pourcentages
- `freeze_panes` pour geler la ligne de header lors du scroll
- `showGridLines = False` pour un rendu propre sans quadrillage

---

## 5. Les données CSV : décryptage

### 5.1 collateral_poste.csv (12 titres)

| Champ | Signification |
|-------|-------------|
| ISIN | Identifiant international du titre (format : 2 lettres pays + 10 chiffres) |
| Nom | Description du titre (émetteur + coupon + maturité) |
| Type | Catégorie : Souverain EU, Corp IG, Corp HY |
| Contrepartie | La banque ou institution à qui ce collateral est posté |
| Valeur_Nominale | Montant nominal (face value) en euros |
| Prix_Pct | Prix de marché en % du nominal (ex. : 99.92 = légèrement sous le pair) |
| Haircut_Pct | Décote appliquée (en %) |
| Maturite | Date d'expiration du titre |
| Rating | Note de crédit (AAA = meilleur, B = spéculatif) |

**Contreparties présentes :** BNP Paribas, Goldman Sachs, HSBC London, Nomura International plc, Société Générale, Hedge Fund Alpha.

### 5.2 portefeuille_dispo.csv (14 titres)

Mêmes colonnes, plus un champ crucial :

| Champ | Signification |
|-------|-------------|
| Eligibilite | Liste séparée par `;` des contreparties qui acceptent ce titre comme collateral |

**Types présents dans le dispo :** Souverain EU (5), Corp IG (5), Corp HY (2), ETF Actions (1).

**Absence notable :** aucun titre du disponible n'est éligible pour Nomura. C'est une contrainte réelle (pas d'accord cadre GMRA/CSA permettant de poster ces titres à Nomura).

---

## 6. Résultats du 7 mai 2026 — Analyse détaillée

### 6.1 Titres en alerte (maturité < 30 jours)

| # | ISIN | Titre | Contrepartie | Maturité | Jours | Statut |
|---|------|-------|-------------|---------|-------|--------|
| 1 | DE000BU0E295 | Schatz 0% 13 mai 2026 | Goldman Sachs | 13/05/26 | **6j** | URGENT |
| 2 | XS2597973812 | Vestas Wind 4.125% 15 mai 2026 | Nomura | 15/05/26 | **8j** | URGENT |
| 3 | XS2176783319 | ENI SpA 1.25% 18 mai 2026 | HSBC London | 18/05/26 | **11j** | URGENT |
| 4 | FR0013131877 | OAT 0.50% 25 mai 2026 | BNP Paribas | 25/05/26 | **18j** | ATTENTION |

### 6.2 Résultats de substitution (traitement par urgence croissante)

**Traitement 1 : Schatz (Goldman Sachs, 6j)**
- VE requise : ~19.89M€
- Candidats éligibles GS dans le dispo : uniquement les Souverains EU (OAT et Bunds)
- Aucun souverain n'atteint seul 19.89M€
- Meilleur candidat CTD (haircut le plus élevé à score égal) : OAT 3.00% 2054 (haircut 5%, VE ~9.50M€)
- **→ COUVERTURE PARTIELLE** : ~9.50M€ alloués sur ~19.89M€ requis
- `Valeur_Eligible_Dispo` de l'OAT 2054 est mise à 0

**Traitement 2 : Vestas (Nomura, 8j)**
- Aucun titre du disponible n'est éligible pour Nomura
- **→ AUCUN SUBSTITUT**

**Traitement 3 : ENI SpA (HSBC London, 11j)**
- VE requise : ~5.94M€
- Candidats éligibles HSBC : Corp IG (score 3) + Souverain EU (score 4)
- Meilleur CTD : Corp IG avant Souverain EU
- Premier candidat Corp IG à fort haircut qui couvre : Aéroports de Paris 2.125% 2026 (score 3, haircut 1% mais VE ~6.91M€)
- **→ SUBSTITUT TROUVÉ** via Aéroports de Paris, VE allouée ~5.94M€

**Traitement 4 : OAT 0.50% (BNP Paribas, 18j)**
- VE requise : ~14.91M€
- Candidats éligibles BNP : ETF Actions (score 2) + Corp IG (score 3) + Souverain EU (score 4)
- Corp HY non éligible BNP dans notre dispo
- Meilleur CTD : ETF Actions (score 2, haircut 20%) = Lyxor CAC40 ETF, VE ~16.46M€ → couvre
- **→ SUBSTITUT TROUVÉ** via Lyxor CAC40, VE allouée ~14.91M€

### 6.3 Synthèse des KPIs

| Indicateur | Valeur |
|-----------|--------|
| Titres en alerte | 4 |
| Substituts complets trouvés | 2 |
| Couvertures partielles | 1 |
| Sans substitut | 1 |
| Taux de couverture complète | 50% |

---

## 7. Points techniques avancés

### 7.1 Mutation en place vs. copie

```python
# Mutation en place : modifie le DataFrame original
dispo.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible

# Ceci serait une erreur : modifie une copie locale, pas l'original
eligibles.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible
```

On passe `dispo` (et non `eligibles` qui est un subset copié) pour que la déduction soit visible pour tous les appels suivants à `trouver_substitut`.

### 7.2 Algorithme greedy vs. optimal

L'algorithme est **greedy** (glouton) : il alloue chaque titre sortant au meilleur substitut disponible au moment où il est traité, sans vision globale de l'avenir. C'est une simplification réaliste pour un MVP.

L'optimum global nécessiterait de la **programmation linéaire** (LP) : minimiser le coût total d'allocation en respectant toutes les contraintes simultanément. Sur Python : `scipy.optimize.linprog` ou `PuLP`.

### 7.3 Complexité algorithmique

- Tri des candidats : O(M log M) avec M = nombre de titres disponibles
- Recherche du premier couvrant : O(M) dans le pire cas
- Pour N titres en alerte : **O(N × M log M)**
- Avec N=100 alertes et M=5000 disponibles : ~100 × 62 000 opérations → exécution en millisecondes

### 7.4 Séparation des responsabilités (architecture propre)

- `detector.py` : logique de détection uniquement (aucun import de substitutor ou reporter)
- `substitutor.py` : logique métier CTD uniquement (aucun I/O fichier)
- `reporter.py` : présentation uniquement (aucune logique métier)
- `main.py` : orchestre, aucune logique métier propre

Cette séparation (pattern MVC-like) facilite les tests unitaires, la maintenance et l'évolution.

### 7.5 Pourquoi openpyxl plutôt que XlsxWriter ou pandas.to_excel ?

- **openpyxl** : lecture et écriture, contrôle fin du formatage cellule par cellule, idéal pour les rapports élaborés.
- **XlsxWriter** : écriture seule mais API plus simple pour les graphiques. 
- **pandas.to_excel** : encapsule openpyxl ou XlsxWriter, bon pour un export rapide sans formatage poussé.
On a choisi openpyxl car le rapport nécessite un formatage conditionnel fin (couleurs par statut, headers stylisés, KPI blocs).

---

## 8. Questions-réponses : préparation à l'entretien

### BLOC A — Finance / Métier

**Q1. Qu'est-ce que le collateral management ?**
C'est l'ensemble des processus qui permettent de gérer les actifs postés comme garanties dans les opérations de marché. L'objectif est de minimiser le coût du collateral (logique CTD) tout en satisfaisant les obligations contractuelles et réglementaires vis-à-vis des contreparties. Post-EMIR/Basel III, cette fonction est devenue critique dans toutes les grandes banques d'investissement.

**Q2. Qu'est-ce que la logique CTD et pourquoi l'utiliser ?**
La logique Cheapest-to-Deliver consiste à poster en priorité les actifs les moins précieux du portefeuille pour conserver les actifs de haute qualité (obligations souveraines) disponibles pour d'autres usages. Concrètement, on sacrifie en premier le High Yield (risqué, faible liquidité), puis les ETF actions, puis le Corp IG, et on conserve au maximum les OAT et Bunds. L'enjeu économique est réel : les souverains sont massivement utilisés dans les repos de trésorerie, et les immobiliser en collateral a un coût d'opportunité significatif.

**Q3. Qu'est-ce qu'un haircut et pourquoi l'appliquer ?**
Le haircut est une décote prudentielle appliquée à la valeur de marché d'un actif pour calculer sa valeur éligible comme collateral. Il protège le créancier contre le risque de marché : si la contrepartie fait défaut et que le marché a bougé défavorablement entre le moment du défaut et la liquidation du collateral, le créancier est quand même couvert. Plus un actif est volatile (actions, HY), plus son haircut est élevé.

**Q4. Pourquoi le seuil de 30 jours ?**
En pratique opérationnelle, une substitution de collateral prend 2 à 5 jours ouvrés (identification du substitut, validation par la contrepartie, règlement/livraison via Euroclear ou Clearstream). 30 jours laisse une marge confortable pour anticiper, tout en restant suffisamment proche pour que l'information soit pertinente. En dessous de 7 jours, c'est une urgence qui requiert un traitement prioritaire.

**Q5. Qu'est-ce qu'EMIR et en quoi impacte-t-il le collateral management ?**
EMIR (European Market Infrastructure Regulation, entré en vigueur progressivement depuis 2012) impose : (1) la compensation centrale des dérivés standardisés via des CCPs (LCH, Eurex), créant des besoins de collateral initial et de variation margin importants ; (2) la collatéralisation bilatérale des dérivés OTC non compensés, avec des appels de marge quotidiens. Cela a multiplié par 5 à 10 les besoins en collateral de haute qualité dans le système financier européen.

**Q6. Pourquoi Nomura n'a aucun substitut dans notre portefeuille ?**
Aucun actif du portefeuille disponible ne liste Nomura International plc dans son champ d'éligibilité. Cela peut signifier : (a) absence d'accord cadre GMRA (Global Master Repurchase Agreement) ou CSA (Credit Support Annex) autorisant ces actifs avec Nomura, (b) Nomura n'accepte pas ces catégories d'actifs comme collateral selon les termes de notre accord. C'est une situation réelle dans la gestion du collateral : l'éligibilité est contractuelle, pas universelle.

**Q7. Comment calculez-vous la valeur éligible ?**
```
VE = Valeur_Nominale × (Prix_Pct / 100) × (1 - Haircut_Pct / 100)
```
On multiplie le nominal par le prix coté (en % du pair) pour obtenir la valeur de marché, puis on applique le haircut. Par exemple : 20M€ nominaux × 99.96% prix × (1 - 0.5%) = environ 19.89M€ de valeur éligible.

**Q8. Qu'est-ce qu'une couverture partielle ?**
Quand aucun substitut seul ne peut couvrir l'intégralité de la VE du titre sortant, on retourne une couverture partielle avec le meilleur candidat CTD disponible. Dans notre outil, on n'agrège pas plusieurs substituts pour une même alerte (simplification MVP). En production, on pourrait allouer plusieurs substituts complémentaires pour couvrir entièrement.

---

### BLOC B — Technique / Python

**Q9. Pourquoi Python plutôt qu'Excel/VBA ?**
Python offre : reproductibilité (code versionnable sur Git, résultats identiques à chaque exécution), maintenabilité (fonctions testables unitairement, lisibles), performance (traitement vectorisé avec pandas sur de gros volumes), et extensibilité (connecteurs Bloomberg, bases de données SQL, APIs REST). Excel/VBA atteint ses limites sur des volumes réels (dizaines de milliers de titres) et est difficile à débugger et maintenir.

**Q10. Comment as-tu évité le double comptage des substituts ?**
En maintenant une colonne `Valeur_Eligible_Dispo` sur le DataFrame du disponible, initialisée à la VE totale et décrémentée à chaque allocation via `dispo.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible`. Puisque pandas DataFrames sont mutables et passés par référence en Python, chaque allocation est immédiatement visible pour les traitements suivants. Le filtre `dispo["Valeur_Eligible_Dispo"] > 0` exclut les titres épuisés.

**Q11. Pourquoi trier par Jours_Restants avant d'allouer ?**
Pour que les titres les plus urgents obtiennent les substituts en priorité. Si deux titres sont en compétition pour le même substitut (car c'est le seul éligible pour leurs contreparties respectives), celui qui expire dans 6 jours doit être servi avant celui qui expire dans 18 jours. Sinon on risque de "gaspiller" la capacité sur un cas moins urgent et de laisser le cas urgent sans solution.

**Q12. Quelle est la complexité algorithmique du moteur de substitution ?**
Pour N titres en alerte et M titres disponibles : O(N × M log M). Le tri des candidats est O(M log M) et la recherche linéaire est O(M) dans le pire cas. Sur les volumes réels des desks (~100-500 alertes, ~1000-5000 disponibles), c'est extrêmement rapide (quelques millisecondes).

**Q13. Qu'est-ce que `iterrows()` et y a-t-il une alternative plus performante ?**
`iterrows()` itère ligne par ligne sur un DataFrame, retournant un tuple (index, Series) par ligne. C'est lisible mais pas le plus rapide. Pour de très gros volumes, on pourrait vectoriser avec `apply()` ou `numpy.where()`. Dans notre cas (N < 1000 alertes), la performance est largement suffisante.

**Q14. Pourquoi as-tu utilisé openpyxl plutôt que pandas.to_excel directement ?**
`pandas.to_excel()` génère un fichier Excel mais avec un formatage très limité. openpyxl permet de contrôler finement chaque cellule : couleurs conditionnelles (vert/orange/rouge selon le statut), polices, tailles, formats numériques personnalisés, gel de volets, masquage du quadrillage, cellules fusionnées pour les titres de section. Le résultat est un rapport professionnel, pas une simple export de données.

**Q15. Comment gérer les données manquantes ou mal formées dans les CSV ?**
Dans le contexte de ce prototype, les données sont supposées propres (données simulées). En production, on ajouterait : validation des types avec `dtype` dans `read_csv`, gestion des NaN avec `dropna()`/`fillna()`, vérification des ISIN via regex, et logging des anomalies détectées.

---

### BLOC C — Critique / Approfondissement

**Q16. Quelles sont les limites de cet outil ?**
- **Données statiques** : le portefeuille est chargé depuis un CSV figé, pas mis à jour intraday (pas de flux Bloomberg/MarkIT).
- **Pas de combinaison de substituts** : une couverture partielle ne cherche pas à combiner plusieurs titres.
- **Algorithme greedy** : non optimal en théorie (une LP donnerait la solution globalement optimale).
- **Pas de gestion du variation margin** : uniquement la substitution à maturité, pas les appels de marge quotidiens.
- **Pas de settlement** : aucune interface avec Euroclear, Clearstream ou les systèmes internes de Natixis.
- **Données simulées** : les prix, haircuts et eligibilités sont fictifs ; en production, ils viendraient de systèmes référentiels (Murex, Calypso, Bloomberg).

**Q17. Comment amélioreriez-vous cet outil si vous aviez plus de temps ?**
1. **Optimisation LP** : remplacer l'algorithme greedy par un solveur de programmation linéaire (`scipy.linprog` ou `PuLP`) pour minimiser le coût total d'allocation.
2. **Flux temps réel** : connexion à l'API Bloomberg (`blpapi`) pour des prix et positions live.
3. **Couverture multi-substituts** : combiner plusieurs titres du disponible pour couvrir une alerte en couverture partielle.
4. **Interface web** : tableau de bord interactif avec Dash (Plotly) pour les opérateurs du desk.
5. **Alertes automatiques** : envoi d'emails ou messages Bloomberg MSG aux traders concernés.
6. **Tests unitaires** : couverture pytest de chaque fonction métier.
7. **Base de données** : persistance des alertes et décisions dans PostgreSQL pour historique et audit.

**Q18. Comment tester unitairement le moteur de substitution ?**
```python
# Exemple de test pytest
def test_substitut_trouve():
    sortant = pd.Series({
        "Contrepartie": "BNP Paribas", "ISIN": "XX0000000001",
        "Nom": "Test", "Maturite": pd.Timestamp("2026-05-20"),
        "Valeur_Eligible": 5_000_000
    })
    dispo = charger_disponible("data/portefeuille_dispo.csv")
    resultat = trouver_substitut(sortant, dispo)
    assert resultat["Statut"] == "SUBSTITUT TROUVE"
    assert resultat["Valeur_Eligible_Substitut"] >= 5_000_000
```

**Q19. Le score CTD est-il configurable ? Comment le rendre paramétrable ?**
Oui. Le dictionnaire `SCORE_CTD` est défini en constante en tête de `substitutor.py`. Pour le rendre paramétrable, on pourrait le charger depuis un fichier de configuration YAML ou JSON, ou le passer en argument à `charger_disponible()`. En production, ces règles CTD peuvent être définies par le desk et changer selon les conditions de marché.

**Q20. Comment intégrer des données en temps réel depuis Bloomberg ?**
Via la librairie officielle Bloomberg `blpapi` (Bloomberg API). On ferait des appels `BDH` (Historical) ou `BLPAPI_SUBSCRIBE` (streaming) pour récupérer les prix mid, les haircuts réglementaires (via le champ `DUR_ADJ_MID`), et les positions. Alternativement, Natixis utilise des systèmes comme Murex ou Calypso qui exposent ces données via des APIs REST ou des extractions batch en CSV/XML.

---

## 9. Lexique

| Terme | Définition |
|-------|-----------|
| Collateral | Actif financier remis en garantie d'une opération |
| Haircut | Décote prudentielle sur la valeur de marché |
| Valeur Éligible (VE) | Valeur marché × (1 − haircut) |
| CTD | Cheapest-to-Deliver : poster en priorité les actifs les moins précieux |
| HQLA | High Quality Liquid Assets : actifs liquides de haute qualité (OAT, Bunds) |
| Repo | Pension livrée : cession temporaire d'un titre avec engagement de rachat |
| SOU | Securities Optimisation Unit : desk de gestion et optimisation du collateral |
| EMIR | European Market Infrastructure Regulation (2012) |
| CCP | Central Counterparty Clearing : chambre de compensation centrale (LCH, Eurex) |
| ISDA | International Swaps and Derivatives Association : accord cadre pour les dérivés |
| GMRA | Global Master Repurchase Agreement : accord cadre pour les repos |
| CSA | Credit Support Annex : annexe de l'ISDA régissant le collatéral des dérivés |
| Margin Call | Appel de marge : demande de complément de garantie |
| Variation Margin | Appel de marge quotidien basé sur la valeur mark-to-market des positions |
| Initial Margin | Marge initiale déposée à l'entrée d'une opération |
| ISIN | International Securities Identification Number : identifiant unique d'un titre |
| VM | Valeur de Marché |
| VE | Valeur Éligible |
| OAT | Obligation Assimilable du Trésor (obligation d'État française) |
| Bund | Obligation d'État allemande |
| Corp IG | Corporate Investment Grade (rating BBB- et au-dessus) |
| Corp HY | Corporate High Yield (rating BB+ et en dessous) |
| ETF | Exchange-Traded Fund : fonds indiciel coté en bourse |
| openpyxl | Librairie Python pour lire/écrire des fichiers Excel (.xlsx) |
| pandas | Librairie Python de manipulation de données tabulaires |
