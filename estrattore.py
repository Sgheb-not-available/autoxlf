import os
import zipfile
import xml.etree.ElementTree as ET

class Estrattore:

    def __init__(self):
        pass

    def estrai_da_blps(self, percorso_blps, CAMPI):
        """
        Apre un .blps (ZIP), trova l'XML principale e restituisce:
        - nome_azienda (str)
        - lista di dict con i dati di ogni risorsa umana
        """
        nome_azienda = os.path.splitext(os.path.basename(percorso_blps))[0]
        risorse = []

        try:
            with zipfile.ZipFile(percorso_blps, 'r') as z:
                # Lista tutti i file nell'archivio
                nomi_file = z.namelist()
                print(nomi_file)

                # Cerca il file XML principale (spesso è il primo, o si chiama
                # "data.xml", "project.xml", o ha estensione .xml/.blpx)
                xml_candidates = [f for f in nomi_file if f.endswith(('.xml', '.blpx', '.blp'))]
                if not xml_candidates:
                    xml_candidates = nomi_file  # prova tutti

                xml_principale = None
                for candidato in xml_candidates:
                    try:
                        with z.open(candidato) as f:
                            contenuto = f.read()
                            # Il file principale contiene tag risorse umane
                            if any(kw in contenuto for kw in
                                [b'RisorseUmane', b'RisorsaUmana', b'Lavoratore',
                                    b'Dipendente', b'Cognome', b'CodiceFiscale']):
                                xml_principale = candidato
                                xml_contenuto = contenuto
                                break
                    except Exception:
                        continue

                if xml_principale is None:
                    # Fallback: prende il file più grande
                    sizes = {f: z.getinfo(f).file_size for f in nomi_file}
                    xml_principale = max(sizes, key=sizes.get)
                    with z.open(xml_principale) as f:
                        xml_contenuto = f.read()

                # Parse XML
                try:
                    root = ET.fromstring(xml_contenuto)
                except ET.ParseError:
                    # Prova a rimuovere BOM se presente
                    xml_contenuto = xml_contenuto.lstrip(b'\xef\xbb\xbf')
                    root = ET.fromstring(xml_contenuto)

                # Estrai nome azienda dal XML se disponibile
                for tag_az in ['NomeAzienda', 'RagioneSociale', 'Azienda', 'nomeAzienda']:
                    el = root.find(f'.//{tag_az}')
                    if el is not None and el.text:
                        nome_azienda = el.text.strip()
                        break

                # Trova il contenitore delle risorse umane
                container_tags = [
                    'RisorseUmane', 'Lavoratori', 'Dipendenti', 'Personale',
                    'risorseUmane', 'lavoratori'
                ]
                container = None
                for tag in container_tags:
                    container = root.find(f'.//{tag}')
                    if container is not None:
                        break

                # Tag per ogni singola risorsa
                item_tags = [
                    'RisorsaUmana', 'Lavoratore', 'Dipendente', 'Persona',
                    'risorsaUmana', 'lavoratore', 'Item', 'Record'
                ]

                elementi = []
                if container is not None:
                    for tag in item_tags:
                        elementi = container.findall(tag)
                        if elementi:
                            break
                    if not elementi:
                        elementi = list(container)  # tutti i figli diretti

                # Fallback: cerca nel documento intero
                if not elementi:
                    for tag in item_tags:
                        elementi = root.findall(f'.//{tag}')
                        if elementi:
                            break

                for el in elementi:
                    risorsa = {"Azienda": nome_azienda,
                            "File sorgente": os.path.basename(percorso_blps)}
                    for colonna, tag_list in CAMPI.items():
                        risorsa[colonna] = self.trova_testo(el, tag_list)
                    # Includi solo se ha almeno cognome o nome
                    if risorsa.get("Cognome") or risorsa.get("Nome"):
                        risorse.append(risorsa)

        except zipfile.BadZipFile:
            print(f"'{percorso_blps}' non è un file ZIP valido — saltato.")
        except Exception as e:
            print(f"Errore su '{percorso_blps}': {e}")

        return nome_azienda, risorse
    
    def trova_testo(self, element, tag_list):
        """Cerca il testo di un tag tra diverse varianti possibili."""
        for tag in tag_list:
            found = element.find(f".//{tag}")
            if found is not None and found.text:
                return found.text.strip()
        return ""