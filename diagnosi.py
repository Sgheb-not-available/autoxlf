"""
Diagnostica formato file .blps Blumatica
Uso: python diagnosi_blps.py "percorso\al\file.blps"
"""

import os
import sys
import struct

def diagnosi(percorso):
    print(f"\n{'='*60}")
    print(f"  DIAGNOSI FILE: {os.path.basename(percorso)}")
    print(f"  Dimensione: {os.path.getsize(percorso):,} byte")
    print(f"{'='*60}\n")

    with open(percorso, 'rb') as f:
        header = f.read(32)

    # --- Magic bytes noti ---
    magic_map = {
        b'PK\x03\x04': 'ZIP / PKZIP',
        b'PK\x05\x06': 'ZIP vuoto',
        b'\x1f\x8b':   'GZIP',
        b'BZh':         'BZIP2',
        b'\xfd7zXZ':   'XZ / LZMA',
        b'7z\xbc\xaf': '7-ZIP',
        b'Rar!':        'RAR',
        b'<?xml':       'XML puro (UTF-8)',
        b'\xef\xbb\xbf<?': 'XML puro (UTF-8 con BOM)',
        b'\xff\xfe<':   'XML puro (UTF-16 LE)',
        b'\xfe\xff<':   'XML puro (UTF-16 BE)',
        b'SQLite':      'Database SQLite',
        b'\xd0\xcf\x11\xe0': 'OLE2 / MS Compound (vecchio formato Office)',
        b'BLPS':        'Blumatica proprietario con magic BLPS',
        b'BLU':         'Blumatica proprietario con magic BLU',
    }

    formato_rilevato = "Sconosciuto"
    for magic, nome in magic_map.items():
        if header.startswith(magic):
            formato_rilevato = nome
            break

    print(f"  Formato rilevato: {formato_rilevato}")
    print(f"  Primi 32 byte (hex): {header.hex(' ')}")
    print(f"  Primi 32 byte (ascii): {header.decode('latin-1', errors='replace')!r}\n")

    # --- Prova a leggere come testo ---
    for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'utf-16']:
        try:
            with open(percorso, 'r', encoding=enc) as f:
                testo = f.read(2000)
            print(f"  Leggibile come testo ({enc}):")
            print(f"  {'-'*50}")
            print(testo[:500])
            print(f"  {'-'*50}")
            break
        except Exception:
            continue
    else:
        print("  Non leggibile come testo in nessuna codifica comune.\n")
        # Dump esadecimale dei primi 256 byte
        with open(percorso, 'rb') as f:
            raw = f.read(256)
        print("  Dump esadecimale primi 256 byte:")
        for i in range(0, len(raw), 16):
            chunk = raw[i:i+16]
            hex_part = ' '.join(f'{b:02x}' for b in chunk)
            asc_part = ''.join(chr(b) if 32 <= b < 127 else '.' for b in chunk)
            print(f"  {i:04x}  {hex_part:<48}  {asc_part}")

    # --- Prova come ZIP con password o ZIP64 ---
    import zipfile
    try:
        with zipfile.ZipFile(percorso) as z:
            print(f"\n  È un ZIP valido! File interni:")
            for info in z.infolist():
                print(f"    - {info.filename}  ({info.file_size:,} byte)")
    except zipfile.BadZipFile:
        print("\n  Confermato: NON è un file ZIP standard.")
    except Exception as e:
        print(f"\n  Errore ZIP: {e}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        percorso = input("Percorso del file .blps: ").strip().strip('"')
    else:
        percorso = sys.argv[1]
    
    if not os.path.exists(percorso):
        print(f"File non trovato: {percorso}")
    else:
        diagnosi(percorso)
