"""
Solveur LP (programmation linéaire) pour l'allocation optimale du collatéral.
Alternative à l'algorithme greedy de substitutor.py.

Formulation :
  Variables  : x[i,j] = VE allouée du substitut j à l'alerte i  (continu, >= 0)
  Objectif   : min  sum(score_ctd[j] * x[i,j])  +  P * sum(s[i])
  où s[i]    = portion non couverte de l'alerte i  (variable de relâchement)
  P = 5 > max(score_ctd) = 4 : couvrir est toujours préférable, quel que soit l'actif.

  Contrainte capacité : sum_i x[i,j] <= VE_disponible[j]
  Contrainte couverture: sum_j x[i,j] + s[i] = VE_requise[i]
  Inéligibilité : x[i,j] = 0 si j non éligible pour la contrepartie de i

Avantage vs greedy : peut combiner plusieurs substituts pour couvrir une même alerte.
Solver : CBC (inclus dans PuLP, open-source, aucune licence).
"""

import pandas as pd
import pulp

from .substitutor import STATUT_TROUVE, STATUT_PARTIEL, STATUT_AUCUN, _est_eligible

_PENALITE = 5  # pénalité de non-couverture, strictement > max(Score_CTD) = 4


def optimiser_substituts(df_alerte: pd.DataFrame, df_dispo: pd.DataFrame) -> pd.DataFrame:
    """
    Résout le problème d'allocation du collatéral par LP (solver CBC).
    Retourne un DataFrame au même format que construire_alertes() de substitutor.py.
    Plusieurs substituts peuvent être combinés pour couvrir une même alerte.
    """
    if df_alerte.empty:
        return pd.DataFrame()

    alertes = df_alerte.reset_index(drop=True)
    dispos  = df_dispo.reset_index(drop=True)
    n_i, n_j = len(alertes), len(dispos)

    # Paires (i, j) éligibles : substitut j accepté par la contrepartie de l'alerte i
    paires = {
        (i, j)
        for i in range(n_i)
        for j in range(n_j)
        if _est_eligible(dispos.at[j, "Eligibilite"], alertes.at[i, "Contrepartie"])
    }

    # ── Construction du problème LP ───────────────────────────────────────────
    prob = pulp.LpProblem("collateral_lp", pulp.LpMinimize)

    # Variables d'allocation
    x = {
        (i, j): pulp.LpVariable(
            f"x_{i}_{j}",
            lowBound=0,
            upBound=min(
                alertes.at[i, "Valeur_Eligible"],
                dispos.at[j, "Valeur_Eligible"],
            ),
        )
        for (i, j) in paires
    }

    # Variables de relâchement (non-couverture pénalisée)
    s = {
        i: pulp.LpVariable(f"s_{i}", lowBound=0, upBound=alertes.at[i, "Valeur_Eligible"])
        for i in range(n_i)
    }

    # Objectif : coût CTD + pénalité de non-couverture
    prob += (
        pulp.lpSum(dispos.at[j, "Score_CTD"] * x[(i, j)] for (i, j) in paires)
        + _PENALITE * pulp.lpSum(s.values())
    )

    # Contrainte capacité : chaque substitut ne peut excéder sa VE disponible
    for j in range(n_j):
        termes = [x[(i, j)] for i in range(n_i) if (i, j) in paires]
        if termes:
            prob += pulp.lpSum(termes) <= dispos.at[j, "Valeur_Eligible"]

    # Contrainte couverture : allocation + non-couverture = besoin de chaque alerte
    for i in range(n_i):
        termes = [x[(i, j)] for j in range(n_j) if (i, j) in paires]
        ve = alertes.at[i, "Valeur_Eligible"]
        if termes:
            prob += pulp.lpSum(termes) + s[i] == ve
        else:
            prob += s[i] == ve  # aucun substitut éligible → non-couverture totale

    # ── Résolution ────────────────────────────────────────────────────────────
    prob.solve(pulp.PULP_CBC_CMD(msg=0))

    # ── Construction du DataFrame résultat ───────────────────────────────────
    lignes = []
    for i, alerte in alertes.iterrows():
        ve_requise = alerte["Valeur_Eligible"]

        allocs = [
            (j, round(pulp.value(x[(i, j)]) or 0))
            for j in range(n_j)
            if (i, j) in paires and (pulp.value(x[(i, j)]) or 0) > 1e-4
        ]
        ve_allouee = sum(v for _, v in allocs)

        if ve_allouee < 1:
            lignes.append({
                "Contrepartie":              alerte["Contrepartie"],
                "ISIN_Sortant":              alerte["ISIN"],
                "Titre_Sortant":             alerte["Nom"],
                "Maturite":                  alerte["Maturite"],
                "Valeur_Eligible_Sortant":   round(ve_requise),
                "ISIN_Substitut":            "",
                "Titre_Substitut":           "",
                "Valeur_Eligible_Substitut": 0,
                "Score_CTD":                 None,
                "Statut":                    STATUT_AUCUN,
            })
        else:
            statut = STATUT_TROUVE if ve_allouee >= round(ve_requise) - 1 else STATUT_PARTIEL
            score_pond = sum(dispos.at[j, "Score_CTD"] * v for j, v in allocs) / ve_allouee
            lignes.append({
                "Contrepartie":              alerte["Contrepartie"],
                "ISIN_Sortant":              alerte["ISIN"],
                "Titre_Sortant":             alerte["Nom"],
                "Maturite":                  alerte["Maturite"],
                "Valeur_Eligible_Sortant":   round(ve_requise),
                "ISIN_Substitut":            " + ".join(dispos.at[j, "ISIN"] for j, _ in allocs),
                "Titre_Substitut":           " + ".join(dispos.at[j, "Nom"] for j, _ in allocs),
                "Valeur_Eligible_Substitut": ve_allouee,
                "Score_CTD":                 round(score_pond, 2),
                "Statut":                    statut,
            })

    return pd.DataFrame(lignes)
