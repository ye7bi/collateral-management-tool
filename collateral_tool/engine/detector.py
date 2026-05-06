"""
Détection des titres postés en collateral arrivant à maturité dans moins de 30 jours.
"""

import pandas as pd
from datetime import date

SEUIL_ALERTE_JOURS = 30


def charger_portefeuille_poste(chemin_csv: str) -> pd.DataFrame:
    """Charge le portefeuille posté et calcule les colonnes dérivées."""
    df = pd.read_csv(chemin_csv, parse_dates=["Maturite"])
    aujourd_hui = date.today()

    df["Valeur_Marche"]   = df["Valeur_Nominale"] * df["Prix_Pct"] / 100
    df["Valeur_Eligible"] = df["Valeur_Marche"] * (1 - df["Haircut_Pct"] / 100)
    df["Jours_Restants"]  = (df["Maturite"].dt.date - aujourd_hui).apply(lambda d: d.days)

    return df


def detecter_maturites_proches(df: pd.DataFrame) -> pd.DataFrame:
    """Filtre les titres dont la maturité est inférieure au seuil d'alerte."""
    return df[df["Jours_Restants"] < SEUIL_ALERTE_JOURS].copy()
