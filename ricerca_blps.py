import os

class RicercaBLPS:
    def __init__(self):
        pass

    def _trova_blps_ricorsiva(self, cartella):
        """
        Ricerca ricorsiva depth-first per file .blps.
        """
        blps = []
        try:
            items = os.listdir(cartella)
        except PermissionError:
            return blps

        # Controlla i file nella cartella corrente
        for item in items:
            path = os.path.join(cartella, item)
            if os.path.isfile(path) and item.lower().endswith('.blps'):
                blps.append(path)

        # Esplora le sottocartelle
        for item in items:
            path = os.path.join(cartella, item)
            if os.path.isdir(path):
                sub_blps = self._trova_blps_ricorsiva(path)
                blps.extend(sub_blps)

        return blps

    def trova_blps(self, cartella_radice):
        """
        Trova tutti i file .blps nella cartella radice e sottocartelle,
        quindi restituisce solo il più recente come lista.
        Ritorna una lista con un singolo elemento (il file .blps più recente) o lista vuota.
        """
        blps = []
        try:
            items = os.listdir(cartella_radice)
        except PermissionError:
            return []

        # Controlla file nella cartella radice
        for item in items:
            path = os.path.join(cartella_radice, item)
            if os.path.isfile(path) and item.lower().endswith('.blps'):
                blps.append(path)

        # Ricerca ricorsiva nelle sottocartelle
        for item in items:
            path = os.path.join(cartella_radice, item)
            if os.path.isdir(path):
                sub_blps = self._trova_blps_ricorsiva(path)
                blps.extend(sub_blps)

        # Restituisce solo il più recente
        if blps:
            most_recent = max(blps, key=os.path.getmtime)
            return [most_recent]
        return []