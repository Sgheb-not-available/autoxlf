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
    cartella_IN = input('Inserisci la cartella di input per dvr: ')
    cartella_IN = os.path.abspath(cartella_IN)

    cartella_OUT = input('Inserisci il percorso di destinazione file: ')
    cartella_OUT = os.path.abspath(cartella_OUT)

    print(f"\n{'='*60}")
    print(f"  Blumatica DVR -> Excel  |  Estrazione Risorse Umane")
    print(f"{'='*60}")
    print(f"  Cartella Input: {cartella_IN}\n")
    print(f"  Cartella Output: {cartella_OUT}\n")

    # Trova tutti i .blps nella cartella (ricorsivo)
    pattern = os.path.join(cartella_IN, "**", "*.blps")
    file_blps = glob.glob(pattern, recursive=True)

    if not file_blps:
        print("  Nessun file .blps trovato nella cartella specificata.")
        return

    print(f"  Trovati {len(file_blps)} file .blps\n")

    tutte_risorse = []
    for i, percorso in enumerate(sorted(file_blps), 1):
        nome_file = os.path.basename(percorso)
        print(f"  [{i}/{len(file_blps)}] Elaborazione: {nome_file}")
        nome_az, risorse = Estrattore().estrai_da_blps(percorso, Excel().CAMPI)
        print(f"        -> Azienda: {nome_az} | Risorse trovate: {len(risorse)}")
        tutte_risorse.extend(risorse)

    if not tutte_risorse:
        print("\n Nessuna risorsa umana trovata nei file elaborati.")
        print("  Possibile causa: struttura XML diversa dal previsto.")
        print("  Contatta il supporto allegando un file .blps di esempio.")
        return

    output = os.path.join(cartella_OUT, "risorse_umane_export.xls")
    Excel().crea_excel(tutte_risorse, output)

    print(f"\n{'='*60}")
    print(f"  Completato!")
    print(f"  Totale risorse estratte: {len(tutte_risorse)}")
    print(f"  File salvato: {output}")
    print(f"{'='*60}\n")

if __name__ == "__main__":
    main()
