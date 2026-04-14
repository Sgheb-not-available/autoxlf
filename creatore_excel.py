from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

class Excel:

    # Campi da estrarre per ogni risorsa umana
    # Mappa: (etichetta colonna Excel) -> (possibili tag XML nel .blps)
    # Blumatica usa nomi come "Nome", "Cognome", "CodiceFiscale", ecc.
    CAMPI = {
        'Cognome': ['Cognome'],
        'Nome': ['Nome'],
        'ComuneNascita': ['ComuneNascita'],
        'DataNascita':  ['DataNascita'],
        'Sesso': ['Sesso'],
        'CF': ['CF'],
        'ComuneResidenza': ['ComuneResidenza'],
        'PR': ['PR'],
        'CAP': ['CAP'],
        'Indirizzo': ['Indirizzo'],
        'DataAssunzione': ['DataAssunzione'],
        'Note': ['Note'],
        'Immagine': ['Immagine'],
        'Mansioni': ['Mansioni'],
        'Email': ['Email']
    }

    HEADER_FISSI = list(CAMPI.keys())
    

    def __init__(self):
        pass

    def crea_excel(self, tutte_risorse, percorso_output):
        """Crea il file Excel con un foglio riepilogo + un foglio per azienda."""
        wb = Workbook()

        # Stili
        colore_header = "1F4E79"
        colore_altriga = "DEEAF1"

        header_font = Font(name="Arial", bold=True, color="FFFFFF", size=10)
        cell_font = Font(name="Arial", size=10)
        header_fill = PatternFill("solid", start_color=colore_header)
        alt_fill = PatternFill("solid", start_color=colore_altriga)
        border_thin = Border(
            left=Side(style='thin'),
            right=Side(style='thin'),
            top=Side(style='thin'),
            bottom=Side(style='thin')
        )
        center = Alignment(horizontal="center", vertical="center")
        left = Alignment(horizontal="left", vertical="center")

        def scrivi_foglio(ws, righe, titolo):
            ws.title = titolo[:31]  # Excel max 31 caratteri per nome foglio
            ws.freeze_panes = "A2"

            # Intestazioni
            for col_idx, header in enumerate(self.HEADER_FISSI, 1):
                cell = ws.cell(row=1, column=col_idx, value=header)
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = center
                cell.border = border_thin

            # Dati
            for row_idx, risorsa in enumerate(righe, 2):
                fill = alt_fill if row_idx % 2 == 0 else PatternFill()
                for col_idx, header in enumerate(self.HEADER_FISSI, 1):
                    valore = risorsa.get(header, "")
                    cell = ws.cell(row=row_idx, column=col_idx, value=valore)
                    cell.font = cell_font
                    cell.fill = fill
                    cell.alignment = left
                    cell.border = border_thin

            # Larghezze colonne
            larghezze = {
                "Azienda": 30, "File sorgente": 25, "Cognome": 18, "Nome": 16,
                "Codice Fiscale": 18, "Data di Nascita": 15, "Sesso": 8,
                "Mansione": 22, "Reparto / Sede": 20, "Data Assunzione": 15,
                "Tipo Contratto": 16, "Formazione": 12, "Sorveglianza San.": 16,
                "Nomina": 18, "Note": 25,
            }
            for col_idx, header in enumerate(self.HEADER_FISSI, 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = larghezze.get(header, 15)
            ws.row_dimensions[1].height = 20

        # Foglio RIEPILOGO con tutte le aziende
        ws_riepilogo = wb.active
        scrivi_foglio(ws_riepilogo, tutte_risorse, "RIEPILOGO")

        # Un foglio per ogni azienda
        aziende = {}
        for r in tutte_risorse:
            az = r.get("Azienda", "Sconosciuta")
            aziende.setdefault(az, []).append(r)

        for nome_az, righe in sorted(aziende.items()):
            # Pulisci il nome per usarlo come nome foglio
            nome_pulito = "".join(c for c in nome_az if c not in r'\/*?:[]\x00-\x1f')[:31] or "Azienda"
            ws = wb.create_sheet(title=nome_pulito)
            scrivi_foglio(ws, righe, nome_pulito)

        wb.save(percorso_output)