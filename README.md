# Collateral Management Tool – Natixis CIB

Outil Python de gestion et d'optimisation du collatéral développé pour la **Securities Optimisation Unit (SOU)** de Natixis CIB.

## Fonctionnalités

- **Détection automatique** des titres postés en collatéral arrivant à maturité dans moins de 30 jours
- **Sélection CTD** (Cheapest-to-Deliver) du meilleur substitut pour chaque alerte, avec gestion de l'allocation et de la capacité résiduelle
- **Rapport Excel** multi-onglets généré automatiquement (tableau de bord, alertes, substitutions, portefeuilles)

## Logique CTD

| Score | Type d'actif | Priorité |
|-------|-------------|----------|
| 1 | Corp HY | Poster en premier |
| 2 | ETF Actions | En deuxième |
| 3 | Corp IG | En troisième |
| 4 | Souverain EU | Conserver |

## Structure

```
collateral_tool/
├── main.py                        # Point d'entrée
├── data/
│   ├── collateral_poste.csv       # Portefeuille posté (12 titres)
│   └── portefeuille_dispo.csv     # Portefeuille disponible (14 titres)
├── engine/
│   ├── detector.py                # Détection des maturités proches
│   ├── substitutor.py             # Moteur de substitution CTD
│   └── reporter.py                # Génération du rapport Excel
└── output/
    └── collateral_report.xlsx     # Rapport généré
```

## Utilisation

```bash
cd collateral_tool
python main.py
```

Le rapport Excel est généré dans `collateral_tool/output/collateral_report.xlsx`.

## Stack technique

- Python 3.13
- pandas — traitement des données
- openpyxl — génération du rapport Excel

## Auteur

**Louis Jeaate** – ECE Paris, M1 Finance et Ingénierie Quantitative  
Stage – Natixis CIB, Securities Optimisation Unit – Mai 2026
