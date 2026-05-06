# CLAUDE.md – Contexte du projet pour l'IA

## Qui est l'utilisateur

Louis JEAATE, étudiant en M1 Finance et Ingénierie Quantitative à ECE Paris.
- Email : jeaate.louis@gmail.com
- Téléphone : +33 7 83 28 78 58
- Contexte : stage / projet réalisé pour Natixis CIB, Securities Optimisation Unit (SOU)

## Ce que ce projet est

Un outil Python de **gestion et d'optimisation du collatéral** développé pour la SOU de Natixis CIB. Il automatise trois choses :
1. Détecter les titres postés en collatéral arrivant à maturité dans moins de 30 jours
2. Trouver le meilleur substitut pour chaque alerte via la logique CTD (Cheapest-to-Deliver)
3. Générer un rapport Excel professionnel multi-onglets

## Ce qui a été fait dans cette session

### Code Python (collateral_tool/)
- `engine/detector.py` : charge le portefeuille posté, calcule VM/VE/jours restants, filtre les alertes < 30j
- `engine/substitutor.py` : moteur CTD — filtre par éligibilité, tri par score CTD (HY=1, ETF=2, IG=3, Souverain=4), gestion de la capacité résiduelle (colonne `Valeur_Eligible_Dispo` mutuée en place pour éviter le double comptage), priorité aux cas les plus urgents
- `engine/reporter.py` : génération Excel via openpyxl — 4 onglets (Tableau de Bord, Alertes & Substitutions, Portefeuille Posté, Disponible), formatage professionnel avec couleurs conditionnelles
- `main.py` : orchestration du pipeline complet

### Documents produits
- `GUIDE_PROJET.md` : guide complet de A à Z pour que Louis maîtrise le projet (contexte financier, logique CTD, code commenté, résultats analysés, 20+ Q&A d'entretien, lexique)
- `rapport_natixis.tex` : rapport LaTeX professionnel A4 à remettre à Natixis — page de garde avec logo, coordonnées de Louis, ECE Paris ; 6 sections ; tableaux formatés ; code couleur ; rédigé à la première personne comme si Louis l'avait écrit

### GitHub
- Dépôt créé : https://github.com/ye7bi/collateral-management-tool
- gh CLI installé via winget, auth faite par Louis
- Tout pushé sur main

## Structure des fichiers

```
PROJET NATIXIS/
├── CLAUDE.md                          ← ce fichier
├── GUIDE_PROJET.md                    ← guide complet pour Louis
├── README.md                          ← présentation du projet
├── rapport_natixis.tex                ← rapport LaTeX pro (compiler avec pdflatex x2)
├── Natixis-CIB-312x351px.png          ← logo Natixis (utilisé dans le .tex)
├── .gitignore
└── collateral_tool/
    ├── main.py
    ├── data/
    │   ├── collateral_poste.csv       ← 12 titres postés (données simulées)
    │   └── portefeuille_dispo.csv     ← 14 substituts potentiels (données simulées)
    ├── engine/
    │   ├── detector.py
    │   ├── substitutor.py
    │   └── reporter.py
    └── output/
        └── collateral_report.xlsx     ← rapport généré (ne pas modifier à la main)
```

## Logique métier clé à retenir

**Score CTD :** Corp HY (1) → ETF Actions (2) → Corp IG (3) → Souverain EU (4)
- Score 1 = poster en premier (moins précieux)
- Score 4 = conserver (obligations souveraines = actifs HQLA, coût d'opportunité max)

**Anti double-comptage :** la colonne `Valeur_Eligible_Dispo` du DataFrame `dispo` est mutée en place à chaque allocation. Ne jamais travailler sur une copie (`eligibles`) pour les déductions.

**Priorité aux urgences :** `construire_alertes()` trie par `Jours_Restants` croissants avant d'allouer. Le titre qui expire dans 6 jours est servi avant celui qui expire dans 18 jours.

**Éligibilité :** le champ `Eligibilite` du CSV disponible est une liste séparée par `;`. Un substitut n'est proposé que si la contrepartie du titre sortant y figure.

## Résultats du 7 mai 2026 (données de test)

4 titres en alerte :
- Schatz 0% mai 2026 (Goldman Sachs, 6j) → COUVERTURE PARTIELLE via OAT 3.00% 2054
- Vestas Wind 4.125% mai 2026 (Nomura, 8j) → AUCUN SUBSTITUT (Nomura absent de toutes les listes d'éligibilité)
- ENI SpA 1.25% mai 2026 (HSBC London, 11j) → SUBSTITUT TROUVÉ via Aéroports de Paris 2.125% 2026
- OAT 0.50% mai 2026 (BNP Paribas, 18j) → SUBSTITUT TROUVÉ via Lyxor CAC40 ETF

## Comment lancer l'outil

```bash
cd collateral_tool
python main.py
# Rapport généré dans output/collateral_report.xlsx
```

## Comment compiler le rapport LaTeX

```bash
# Depuis PROJET NATIXIS/
pdflatex rapport_natixis.tex
pdflatex rapport_natixis.tex   # 2e passe pour la table des matières
```

## Préférences de Louis

- Répond en français
- Veut du code propre sans commentaires inutiles
- Les documents doivent être rédigés comme s'il les avait écrits lui-même (pas de style IA)
- Projet destiné à être présenté en entretien chez Natixis
