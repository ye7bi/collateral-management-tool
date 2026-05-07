"""
Collateral Management Tool -- Natixis Securities Optimisation Unit
Lancement : python main.py (depuis collateral_tool/)
"""

import sys
from pathlib import Path

if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

import pandas as pd
from engine.detector    import charger_portefeuille_poste, detecter_maturites_proches
from engine.substitutor import charger_disponible, construire_alertes
from engine.optimizer   import optimiser_substituts
from engine.reporter    import generer_rapport


def _cout_total(df: pd.DataFrame) -> float:
    """Coût total : sum(score_ctd * VE_allouée) sur toutes les allocations effectives."""
    masque = df["Score_CTD"].notna()
    return (df.loc[masque, "Score_CTD"] * df.loc[masque, "Valeur_Eligible_Substitut"]).sum()


def _afficher_stats(label: str, df: pd.DataFrame) -> None:
    n_trouves  = (df["Statut"] == "SUBSTITUT TROUVE").sum()
    n_partiels = (df["Statut"] == "COUVERTURE PARTIELLE").sum()
    n_aucun    = (df["Statut"] == "AUCUN SUBSTITUT").sum()
    cout       = _cout_total(df)
    print(
        f"  {label:<10} ->  {n_trouves} substitut(s) complet(s), "
        f"{n_partiels} partiel(s), {n_aucun} aucun  |  "
        f"cout total : {cout:>15,.0f}"
    )


def main() -> None:
    print("=" * 65)
    print("  Collateral Management Tool - Natixis SOU")
    print("=" * 65)

    print("\nChargement des portefeuilles...")
    df_poste        = charger_portefeuille_poste(DATA_DIR / "collateral_poste.csv")
    df_dispo_greedy = charger_disponible(DATA_DIR / "portefeuille_dispo.csv")
    df_dispo_lp     = charger_disponible(DATA_DIR / "portefeuille_dispo.csv")

    print("Detection des maturites proches (< 30 jours)...")
    df_alerte = detecter_maturites_proches(df_poste)
    n = len(df_alerte)
    print(f"  -> {n} titre{'s' if n > 1 else ''} en alerte\n")

    # ── Approche 1 : Greedy (CTD séquentiel) ─────────────────────────────────
    print("[ 1/2 ] Approche Greedy (CTD sequentiel)...")
    df_greedy     = construire_alertes(df_alerte, df_dispo_greedy)
    chemin_greedy = OUTPUT_DIR / "collateral_report_greedy.xlsx"
    generer_rapport(df_poste, df_dispo_greedy, df_greedy, chemin_greedy)
    print(f"        -> {chemin_greedy.name}")

    # ── Approche 2 : LP (programmation linéaire) ──────────────────────────────
    print("[ 2/2 ] Approche LP (programmation lineaire, solver CBC)...")
    df_lp     = optimiser_substituts(df_alerte, df_dispo_lp)
    chemin_lp = OUTPUT_DIR / "collateral_report_lp.xlsx"
    generer_rapport(df_poste, df_dispo_lp, df_lp, chemin_lp)
    print(f"        -> {chemin_lp.name}")

    # ── Comparaison synthétique ────────────────────────────────────────────────
    print()
    print("─" * 65)
    print("  COMPARAISON DES DEUX APPROCHES")
    print("─" * 65)
    _afficher_stats("Greedy", df_greedy)
    _afficher_stats("LP", df_lp)
    print("─" * 65)
    cout_g = _cout_total(df_greedy)
    cout_l = _cout_total(df_lp)
    if cout_l < cout_g:
        print(f"  LP economise {cout_g - cout_l:,.0f} unites de cout d'opportunite.")
    elif cout_l > cout_g:
        print(
            f"  LP utilise {cout_l - cout_g:,.0f} unites de cout supplementaires "
            f"pour une meilleure couverture."
        )
    else:
        print("  Les deux approches donnent le meme cout total.")

    print()
    print(f"[OK] Rapports generes dans : {OUTPUT_DIR.resolve()}")
    print("=" * 65)


if __name__ == "__main__":
    main()
