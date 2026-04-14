import os
import sqlite3

class Estrattore:

    def __init__(self):
        pass

    def formatta_data(self, valore):
        """Converte '1986-09-11 00:00:00' -> '11/09/1986'"""
        if not valore:
            return ""
        try:
            return valore[:10].split('-')[2] + '/' + valore[:10].split('-')[1] + '/' + valore[:10].split('-')[0]
        except Exception:
            return str(valore)

    def estrai_da_blps(self, percorso_blps):
        """
        Legge il database SQLite da un file .blps e restituisce:
        - nome_azienda (str)
        - lista di dict con i dati di ogni risorsa umana
        """
        nome_azienda = os.path.splitext(os.path.basename(percorso_blps))[0]
        risorse = []

        try:
            con = sqlite3.connect(percorso_blps)
            con.row_factory = sqlite3.Row
            cur = con.cursor()

            # Nome azienda da tabella Lavoro
            try:
                cur.execute("SELECT RagioneSociale FROM Lavoro LIMIT 1")
                row = cur.fetchone()
                if row and row['RagioneSociale']:
                    nome_azienda = row['RagioneSociale'].strip()
            except Exception:
                pass

            # Mappa IdUtente -> mansioni (lista nomi)
            mansioni_map = {}
            try:
                cur.execute("""
                    SELECT um.IdUtente, mp.Nome
                    FROM UtenteMansioni um
                    JOIN MansPost mp ON um.IdMansPost = mp.IdMansPost
                    ORDER BY um.IdUtente, mp.Ordine
                """)
                for row in cur.fetchall():
                    uid = row['IdUtente']
                    mansioni_map.setdefault(uid, []).append(row['Nome'])
            except Exception as e:
                print(f"        Attenzione: impossibile leggere mansioni ({e})")

            # Estrai utenti — Tipo=1 sono i lavoratori
            # Tipo=2 tipicamente è il consulente/RSPP esterno → incluso con flag
            cur.execute("""
                SELECT *
                FROM Utente
                ORDER BY Cognome, Nome
            """)
            utenti = cur.fetchall()

            for u in utenti:
                uid = u['IdUtente']
                tipo = u['Tipo']

                # Salta utenti senza cognome né nome
                cognome = (u['Cognome'] or '').strip()
                nome    = (u['Nome']    or '').strip()
                if not cognome and not nome:
                    continue

                mansioni_list = mansioni_map.get(uid, [])
                mansioni_str  = '; '.join(mansioni_list)

                risorsa = {
                    'Azienda':            nome_azienda,
                    'File sorgente':      os.path.basename(percorso_blps),
                    'Cognome':            cognome,
                    'Nome':               nome,
                    'ComuneNascita':      (u['ComuneNascita']  or '').strip(),
                    'DataNascita':        self.formatta_data(u['DataNascita']),
                    'Sesso':              (u['Sesso']           or '').strip(),
                    'CF':                 (u['CF']              or '').strip(),
                    'ComuneResidenza':    (u['Comune']          or '').strip(),
                    'PR':                 (u['PR']              or '').strip(),
                    'CAP':                (u['CAP']             or '').strip(),
                    'Indirizzo':          (u['Indirizzo']       or '').strip(),
                    'DataAssunzione':     self.formatta_data(u['DataAssunzione']),
                    'Note':               (u['Note']            or '').strip(),
                    'Mansioni':           mansioni_str,
                    'Email':              (u['Email']           or '').strip(),
                    'Matricola':          (u['Matricola']       or '').strip(),
                    'Titolo':             (u['Titolo']          or '').strip(),
                    'Nazionalita':        (u['Nazionalita']     or '').strip(),
                    'ProvinciaNascita':   (u['ProvinciaNascita'] or '').strip(),
                    'FormazioneCogente':  'Sì' if u['FormazioneCogente'] else 'No',
                    'SorveglianzaSanitaria': 'Sì' if u['SorveglianzaSanitaria'] else 'No',
                    '_tipo': tipo,  # campo interno, non va in Excel
                }
                risorse.append(risorsa)

            con.close()

        except Exception as e:
            print(f"  ERRORE su '{percorso_blps}': {e}")

        return nome_azienda, risorse