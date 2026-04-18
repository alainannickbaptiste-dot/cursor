#!/usr/bin/env python3
"""
Fusionne tous les fichiers CSV d'un dossier en un seul fichier CSV.

Usage:
    python merge_csv.py                              # csv/ -> output/merged.csv
    python merge_csv.py -i data/ -o resultat.csv     # data/ -> resultat.csv
"""

import argparse
import csv
import os
import sys
from pathlib import Path


def find_csv_files(input_dir: Path) -> list[Path]:
    """Retourne la liste des fichiers CSV triés par nom dans le dossier donné."""
    csv_files = sorted(
        p for p in input_dir.iterdir()
        if p.is_file() and p.suffix.lower() == ".csv"
    )
    return csv_files


def merge_csv_files(csv_files: list[Path], output_path: Path) -> int:
    """
    Fusionne les fichiers CSV en un seul.
    Tous les fichiers doivent partager le même en-tête (colonnes).
    Retourne le nombre total de lignes de données écrites.
    """
    if not csv_files:
        print("Aucun fichier CSV trouvé.")
        return 0

    header = None
    total_rows = 0

    output_path.parent.mkdir(parents=True, exist_ok=True)

    with open(output_path, "w", newline="", encoding="utf-8") as out_file:
        writer = None

        for csv_file in csv_files:
            with open(csv_file, "r", newline="", encoding="utf-8") as in_file:
                reader = csv.reader(in_file)

                try:
                    file_header = next(reader)
                except StopIteration:
                    print(f"  ⚠ {csv_file.name} est vide, ignoré.")
                    continue

                if header is None:
                    header = file_header
                    writer = csv.writer(out_file)
                    writer.writerow(header)
                elif file_header != header:
                    print(
                        f"  ⚠ {csv_file.name} a un en-tête différent "
                        f"({file_header} vs {header}), ignoré."
                    )
                    continue

                rows = list(reader)
                rows = [r for r in rows if any(cell.strip() for cell in r)]
                writer.writerows(rows)
                total_rows += len(rows)
                print(f"  ✓ {csv_file.name} — {len(rows)} lignes ajoutées")

    return total_rows


def main():
    parser = argparse.ArgumentParser(
        description="Fusionne tous les fichiers CSV d'un dossier en un seul fichier."
    )
    parser.add_argument(
        "-i", "--input-dir",
        default="csv",
        help="Dossier contenant les fichiers CSV (défaut: csv/)",
    )
    parser.add_argument(
        "-o", "--output",
        default="output/merged.csv",
        help="Chemin du fichier CSV de sortie (défaut: output/merged.csv)",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    output_path = Path(args.output)

    if not input_dir.is_dir():
        print(f"Erreur : le dossier '{input_dir}' n'existe pas.")
        sys.exit(1)

    csv_files = find_csv_files(input_dir)
    print(f"Dossier source : {input_dir}/")
    print(f"Fichiers CSV trouvés : {len(csv_files)}")
    print()

    total = merge_csv_files(csv_files, output_path)

    print()
    print(f"Fusion terminée → {output_path}")
    print(f"Total : {total} lignes de données dans un seul fichier.")


if __name__ == "__main__":
    main()
