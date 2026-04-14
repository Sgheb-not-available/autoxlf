"""
Estrai Risorse Umane da file Blumatica DVR (.blps) -> Excel (.xls)
Uso: python autoxlf.py
Se non si specifica la cartella, usa la directory corrente.
"""

import os
import glob
from estrattore import Estrattore
from creatore_excel import Excel

# Entry point
def main():
    cartella_IN = input('Inserisci la cartella di input per DVR: ').strip().strip('"')
    cartella_IN = os.path.abspath(cartella_IN)

    cartella_OUT = input('Inserisci il percorso di destinazione file: ').strip().strip('"')
    cartella_OUT = os.path.abspath(cartella_OUT)

    print(f"\n{'='*60}")
    print(f"  Blumatica DVR -> Excel  |  Estrazione Risorse Umane")
    print(f"{'='*60}")
    print(f"  Input:  {cartella_IN}")
    print(f"  Output: {cartella_OUT}\n")

    file_blps = glob.glob(os.path.join(cartella_IN, "**", "*.blps"), recursive=True)
    if not file_blps:
        print("  Nessun file .blps trovato.")
        return

    print(f"  Trovati {len(file_blps)} file .blps\n")

    tutte_risorse = []
    for i, percorso in enumerate(sorted(file_blps), 1):
        nome_file = os.path.basename(percorso)
        print(f"  [{i}/{len(file_blps)}] {nome_file}")
        nome_az, risorse = Estrattore().estrai_da_blps(percorso)
        lavoratori = sum(1 for r in risorse if r.get('_tipo') == 1)
        altri      = len(risorse) - lavoratori
        print(f"        Azienda: {nome_az}")
        print(f"        Lavoratori (Tipo=1): {lavoratori}  |  Altri ruoli: {altri}")
        tutte_risorse.extend(risorse)

    if not tutte_risorse:
        print("\n  Nessuna risorsa trovata.")
        return

    os.makedirs(cartella_OUT, exist_ok=True)
    output = os.path.join(cartella_OUT, "risorse_umane_export.xlf")
    Excel().crea_excel(tutte_risorse, output)

    lav_tot = sum(1 for r in tutte_risorse if r.get('_tipo') == 1)
    print(f"\n{'='*60}")
    print(f"  Completato!")
    print(f"  Totale righe:       {len(tutte_risorse)}")
    print(f"  di cui lavoratori:  {lav_tot}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()