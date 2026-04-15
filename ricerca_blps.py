import os

class RicercaBLPS:
    def __init__(self):
        pass

    def trova_dvr_per_cartella(self, cartella):
        """
        Ricerca ricorsiva depth-first per file .blps.
        Restituisce un dict {cartella: file_più_recente} per ogni cartella che contiene .blps.
        """
        dvr_cartella = {}
        try:
            items = os.listdir(cartella)
        except PermissionError:
            return dvr_cartella

        # Controlla i file nella cartella corrente
        blps_in_folder = []
        for item in items:
            path = os.path.join(cartella, item)
            if os.path.isfile(path) and item.lower().endswith('.blps'):
                blps_in_folder.append(path)

        # Se ci sono file .blps in questa cartella, salva il più recente
        if blps_in_folder:
            most_recent = max(blps_in_folder, key=os.path.getmtime)
            dvr_cartella[cartella] = most_recent

        # Esplora le sottocartelle
        for item in items:
            path = os.path.join(cartella, item)
            if os.path.isdir(path):
                sub_files = self.trova_dvr_per_cartella(path)
                dvr_cartella.update(sub_files)

        return dvr_cartella

    def trova_dvr(self, cartella_radice):
        """
        Trova il file .blps più recente in ogni cartella che ne contiene.
        Ritorna una lista ordinata dei file più recenti (uno per cartella).
        """
        dvr_cartella = self.trova_dvr_per_cartella(cartella_radice)
        # Ritorna la lista dei file ordinati per data di modifica (più recenti prima)
        lista_dvr = list(dvr_cartella.values())
        return sorted(lista_dvr, key=os.path.getmtime, reverse=True)