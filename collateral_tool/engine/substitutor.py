"""
Logique cheapest-to-deliver (CTD) : scoring et selection des meilleurs substituts.

Regle CTD (score croissant = moins cher, a poster en priorite) :
  Corp HY      -> 1
  ETF Actions  -> 2
  Corp IG      -> 3
  Souverain EU -> 4

Gestion de l'allocation : un substitut ne peut pas etre assigne deux fois.
Le nominal eligible restant est deduit apres chaque affectation.
"""

import pandas as pd

SCORE_CTD: dict[str, int] = {
    "Corp HY":      1,
    "ETF Actions":  2,
    "Corp IG":      3,
    "Souverain EU": 4,
}

STATUT_TROUVE  = "SUBSTITUT TROUVE"
STATUT_PARTIEL = "COUVERTURE PARTIELLE"
STATUT_AUCUN   = "AUCUN SUBSTITUT"


def charger_disponible(chemin_csv: str) -> pd.DataFrame:
    """Charge le portefeuille disponible et calcule valeurs + score CTD."""
    df = pd.read_csv(chemin_csv, parse_dates=["Maturite"])

    df["Valeur_Marche"]        = df["Valeur_Nominale"] * df["Prix_Pct"] / 100
    df["Valeur_Eligible"]      = df["Valeur_Marche"] * (1 - df["Haircut_Pct"] / 100)
    df["Score_CTD"]            = df["Type"].map(SCORE_CTD)
    # Capacite residuelle : diminuee au fil des allocations
    df["Valeur_Eligible_Dispo"] = df["Valeur_Eligible"].copy()

    return df


def _est_eligible(eligibilite_str: str, contrepartie: str) -> bool:
    noms = [x.strip() for x in str(eligibilite_str).split(";")]
    return contrepartie in noms


def _construire_resultat(
    sortant: pd.Series,
    substitut: pd.Series,
    statut: str,
    ve_allouee: float,
) -> dict:
    return {
        "Contrepartie":              sortant["Contrepartie"],
        "ISIN_Sortant":              sortant["ISIN"],
        "Titre_Sortant":             sortant["Nom"],
        "Maturite":                  sortant["Maturite"],
        "Valeur_Eligible_Sortant":   round(sortant["Valeur_Eligible"]),
        "ISIN_Substitut":            substitut["ISIN"],
        "Titre_Substitut":           substitut["Nom"],
        "Valeur_Eligible_Substitut": round(ve_allouee),
        "Score_CTD":                 int(substitut["Score_CTD"]),
        "Statut":                    statut,
    }


def _resultat_vide(sortant: pd.Series) -> dict:
    return {
        "Contrepartie":              sortant["Contrepartie"],
        "ISIN_Sortant":              sortant["ISIN"],
        "Titre_Sortant":             sortant["Nom"],
        "Maturite":                  sortant["Maturite"],
        "Valeur_Eligible_Sortant":   round(sortant["Valeur_Eligible"]),
        "ISIN_Substitut":            "",
        "Titre_Substitut":           "",
        "Valeur_Eligible_Substitut": 0,
        "Score_CTD":                 None,
        "Statut":                    STATUT_AUCUN,
    }


def trouver_substitut(sortant: pd.Series, dispo: pd.DataFrame) -> dict:
    """
    Selectionne le meilleur substitut CTD pour un titre sortant,
    en tenant compte de la capacite residuelle du portefeuille dispo.
    """
    titre_sortant = sortant
    contrepartie  = sortant["Contrepartie"]
    valeur_cible  = sortant["Valeur_Eligible"]

    # Filtrer par eligibilite et capacite restante > 0
    masque = (
        dispo["Eligibilite"].apply(lambda e: _est_eligible(e, contrepartie))
        & (dispo["Valeur_Eligible_Dispo"] > 0)
    )
    eligibles = dispo[masque].copy()

    if eligibles.empty:
        return _resultat_vide(titre_sortant)

    # Tri CTD : score croissant, haircut decroissant a egalite
    eligibles = eligibles.sort_values(
        ["Score_CTD", "Haircut_Pct"], ascending=[True, False]
    )

    # Chercher un substitut qui couvre entierement avec la capacite residuelle
    for idx, sub in eligibles.iterrows():
        if sub["Valeur_Eligible_Dispo"] >= valeur_cible:
            # Deduire la valeur allouee de la capacite residuelle
            dispo.at[idx, "Valeur_Eligible_Dispo"] -= valeur_cible
            return _construire_resultat(sortant, sub, STATUT_TROUVE, valeur_cible)

    # Aucune couverture complete : prendre le meilleur disponible (partiel)
    meilleur_idx = eligibles.index[0]
    meilleur     = eligibles.iloc[0]
    ve_allouee   = meilleur["Valeur_Eligible_Dispo"]
    dispo.at[meilleur_idx, "Valeur_Eligible_Dispo"] = 0
    return _construire_resultat(sortant, meilleur, STATUT_PARTIEL, ve_allouee)


def construire_alertes(df_alerte: pd.DataFrame, dispo: pd.DataFrame) -> pd.DataFrame:
    """
    Construit le tableau des substitutions pour tous les titres en alerte.
    Trie les alertes par urgence (jours restants croissants) pour allouer
    en priorite les substituts aux titres les plus proches de maturite.
    """
    if df_alerte.empty:
        return pd.DataFrame()

    df_tri = df_alerte.sort_values("Jours_Restants").reset_index(drop=True)
    lignes = [trouver_substitut(row, dispo) for _, row in df_tri.iterrows()]
    return pd.DataFrame(lignes)
