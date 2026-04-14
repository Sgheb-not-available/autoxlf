"""
Esplora struttura database SQLite di un file .blps Blumatica
Uso: python esplora_sqlite.py "percorso\al\file.blps"
"""

import sqlite3
import sys
import os

def esplora(percorso):
    print(f"\n{'='*60}")
    print(f"  ESPLORAZIONE SQLite: {os.path.basename(percorso)}")
    print(f"{'='*60}\n")

    con = sqlite3.connect(percorso)
    cur = con.cursor()

    # Lista tutte le tabelle
    cur.execute("SELECT name FROM sqlite_master WHERE type='table' ORDER BY name;")
    tabelle = [r[0] for r in cur.fetchall()]
    print(f"  Tabelle trovate ({len(tabelle)}):")
    for t in tabelle:
        cur.execute(f"SELECT COUNT(*) FROM [{t}]")
        n = cur.fetchone()[0]
        print(f"    - {t}  ({n} righe)")

    print()

    # Per ogni tabella, mostra colonne e prime 2 righe
    KEYWORDS = ['lavorat', 'dipend', 'person', 'risorsa', 'umana',
                'anagrafi', 'addett', 'operai', 'staff', 'worker',
                'cognome', 'nome', 'cf', 'fiscal']

    tabelle_interessanti = [
        t for t in tabelle
        if any(k in t.lower() for k in KEYWORDS)
    ]

    # Se nessuna trovata con keyword, mostra tutte
    da_mostrare = tabelle_interessanti if tabelle_interessanti else tabelle

    print(f"  Tabelle potenzialmente rilevanti: {da_mostrare or '(nessuna — mostro tutte)'}\n")

    for t in da_mostrare:
        print(f"  {'─'*55}")
        print(f"  TABELLA: {t}")
        cur.execute(f"PRAGMA table_info([{t}])")
        colonne = cur.fetchall()
        print(f"  Colonne: {[c[1] for c in colonne]}")

        cur.execute(f"SELECT * FROM [{t}] LIMIT 3")
        righe = cur.fetchall()
        for i, riga in enumerate(righe):
            # Tronca valori lunghi
            riga_troncata = tuple(
                (str(v)[:80] + '…' if isinstance(v, str) and len(str(v)) > 80 else v)
                for v in riga
            )
            print(f"  Riga {i+1}: {riga_troncata}")
        print()

    # Cerca anche tra TUTTE le tabelle colonne con nomi sospetti
    print(f"\n  {'='*55}")
    print(f"  RICERCA COLONNE CON NOMI SOSPETTI IN TUTTE LE TABELLE:")
    for t in tabelle:
        cur.execute(f"PRAGMA table_info([{t}])")
        colonne = [c[1] for c in cur.fetchall()]
        sospette = [c for c in colonne if any(k in c.lower() for k in KEYWORDS)]
        if sospette:
            print(f"    Tabella [{t}] -> colonne: {sospette}")

    con.close()

"""
Ispezione dettagliata tabelle Utente, UtenteMansioni, MansPost
"""

def ispezione(percorso):
    con = sqlite3.connect(percorso)
    con.row_factory = sqlite3.Row
    cur = con.cursor()

    tabelle_da_vedere = ['Utente', 'UtenteMansioni', 'MansPost', 'UtenteLogistica', 'Nomina', 'Lavoro', 'SchedaGenerale']

    for t in tabelle_da_vedere:
        print(f"\n{'='*60}")
        print(f"  TABELLA: {t}")
        cur.execute(f"PRAGMA table_info([{t}])")
        colonne = cur.fetchall()
        print(f"  Colonne: {[c['name'] for c in colonne]}")
        cur.execute(f"SELECT * FROM [{t}] LIMIT 5")
        righe = cur.fetchall()
        for i, riga in enumerate(righe):
            print(f"  Riga {i+1}: {dict(riga)}")

    con.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        percorso = input("Percorso del file .blps: ").strip().strip('"')
    else:
        percorso = sys.argv[1]

    if not os.path.exists(percorso):
        print(f"File non trovato: {percorso}")
    else:
        ispezione(percorso)