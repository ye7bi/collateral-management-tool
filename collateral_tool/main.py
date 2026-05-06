"""
Collateral Management Tool -- Natixis Securities Optimisation Unit
Lancement : python main.py (depuis collateral_tool/)
"""

import sys
from pathlib import Path

# Forcer UTF-8 sur la sortie standard (Windows)
if sys.stdout.encoding != "utf-8":
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

BASE_DIR   = Path(__file__).parent
DATA_DIR   = BASE_DIR / "data"
OUTPUT_DIR = BASE_DIR / "output"
OUTPUT_DIR.mkdir(exist_ok=True)

from engine.detector    import charger_portefeuille_poste, detecter_maturites_proches
from engine.substitutor import charger_disponible, construire_alertes
from engine.reporter    import generer_rapport


def main() -> None:
    print("=" * 60)
    print("  Collateral Management Tool - Natixis SOU")
    print("=" * 60)

    print("Chargement du portefeuille poste...")
    df_poste = charger_portefeuille_poste(DATA_DIR / "collateral_poste.csv")

    print("Chargement du portefeuille disponible...")
    df_dispo = charger_disponible(DATA_DIR / "portefeuille_dispo.csv")

    print("Detection des maturites proches (< 30 jours)...")
    df_alerte = detecter_maturites_proches(df_poste)
    n = len(df_alerte)
    print(f"  -> {n} titre{'s' if n > 1 else ''} en alerte")

    print("Selection des substituts (logique cheapest-to-deliver)...")
    df_substitutions = construire_alertes(df_alerte, df_dispo)

    if not df_substitutions.empty:
        for statut in ["SUBSTITUT TROUVE", "COUVERTURE PARTIELLE", "AUCUN SUBSTITUT"]:
            # Comparaison insensible aux accents manquants dans le terminal
            nb = df_substitutions["Statut"].str.startswith(statut[:10]).sum()
            if nb:
                print(f"  {statut} : {nb}")

    chemin_rapport = OUTPUT_DIR / "collateral_report.xlsx"
    print("Generation du rapport Excel...")
    generer_rapport(df_poste, df_dispo, df_substitutions, chemin_rapport)

    print(f"\n[OK] Rapport genere : {chemin_rapport.resolve()}")
    print("=" * 60)


if __name__ == "__main__":
    main()
