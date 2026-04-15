"""
Estrai Risorse Umane da file Blumatica DVR (.blps) -> Excel (.xls)
Uso: python autoxlf.py
Se non si specifica la cartella, usa la directory corrente.
"""

import os
from ricerca_blps import RicercaBLPS
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

    file_blps = sorted(RicercaBLPS().trova_dvr(cartella_IN))
    if not file_blps:
        print("  Nessun file .blps trovato.")
        return

    print(f"  Trovati {len(file_blps)} file .blps\n")

    os.makedirs(cartella_OUT, exist_ok=True)
    
    totale_generale = 0
    lav_totale_generale = 0
    file_processati = 0
    
    for i, percorso in enumerate(sorted(file_blps), 1):
        nome_file = os.path.basename(percorso)
        print(f"  [{i}/{len(file_blps)}] {nome_file}")
        nome_az, risorse = Estrattore().estrai_da_blps(percorso)
        
        if not risorse:
            print(f"        Nessuna risorsa trovata.")
            continue
        
        lavoratori = sum(1 for r in risorse if r.get('_tipo') == 1)
        altri      = len(risorse) - lavoratori
        print(f"        Azienda: {nome_az}")
        print(f"        Lavoratori (Tipo=1): {lavoratori}  |  Altri ruoli: {altri}")
        
        # Crea un file Excel separato per ogni .blps
        output = os.path.join(cartella_OUT, f"{nome_az}.xls")
        Excel().crea_excel(risorse, output)
        
        totale_generale += len(risorse)
        lav_totale_generale += lavoratori
        file_processati += 1

    if file_processati == 0:
        print("\n  Nessuna risorsa trovata.")
        return

    print(f"\n{'='*60}")
    print(f"  Completato!")
    print(f"  File creati:       {file_processati}")
    print(f"  Totale righe:      {totale_generale}")
    print(f"  di cui lavoratori: {lav_totale_generale}")
    print(f"{'='*60}\n")


if __name__ == "__main__":
    main()