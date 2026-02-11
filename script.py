import pandas as pd
from lxml import etree
import os
import sys

# Optional: Für die Summe in Worten
try:
    from num2words import num2words
except ImportError:
    num2words = None

# --- KONFIGURATION ---
XML_VORLAGE = "vorlage.xml"
COL_DATUM = "Belegdatum"
COL_BETRAG = "Betrag (Brutto)"

def generate_xml(excel_pfad, vorname, nachname, strasse, hausnr, plz, ort):
    # 1. Prüfen, ob Dateien existieren
    if not os.path.exists(excel_pfad):
        print(f"Fehler: Tabelle '{excel_pfad}' nicht gefunden.")
        return
    if not os.path.exists(XML_VORLAGE):
        print(f"Fehler: Vorlage '{XML_VORLAGE}' nicht gefunden.")
        return

    # 2. Dynamischen Output-Namen generieren
    basis_name = os.path.splitext(os.path.basename(excel_pfad))[0]
    output_datei = f"output_{basis_name}_{nachname}.xml"

    # 3. Tabelle laden
    try:
        df = pd.read_csv(excel_pfad) if excel_pfad.lower().endswith('.csv') else pd.read_excel(excel_pfad)
        df.columns = [str(c).strip() for c in df.columns]
    except Exception as e:
        print(f"Fehler beim Lesen der Tabelle: {e}")
        return

    # 4. XML-Vorlage laden
    parser = etree.XMLParser(remove_blank_text=True)
    tree = etree.parse(XML_VORLAGE, parser)
    root = tree.getroot()
    ns = {"ns": "http://www.lucom.com/ffw/xml-data-1.0.xsd"}

    # 5. Personen-Informationen im XML setzen (ID "name")
    # Format: Vorname Nachname | Strasse Hausnummer | PLZ Ort
    name_string = f"{vorname} {nachname} | {strasse} {hausnr} | {plz} {ort}"
    try:
        name_element = root.xpath("//ns:element[@id='name']", namespaces=ns)[0]
        name_element.text = name_string
    except IndexError:
        print("Hinweis: Element mit ID 'name' nicht in der Vorlage gefunden.")

    # 6. Den Tabellen-Bereich (betraege) leeren und neu befüllen
    dataset = root.xpath("//ns:dataset[@id='betraege']", namespaces=ns)[0]
    for old_row in dataset.xpath("ns:datarow", namespaces=ns):
        dataset.remove(old_row)

    for index, row in df.iterrows():
        datarow = etree.SubElement(dataset, "{http://www.lucom.com/ffw/xml-data-1.0.xsd}datarow")
        
        def add_el(parent, id_name, val):
            el = etree.SubElement(parent, "{http://www.lucom.com/ffw/xml-data-1.0.xsd}element")
            el.set("id", id_name)
            el.text = str(val)

        d_raw = row[COL_DATUM]
        d_str = d_raw.strftime('%d.%m.%Y 00:00:00') if hasattr(d_raw, 'strftime') else f"{d_raw} 00:00:00"

        add_el(datarow, "ID_LINE", index + 1)
        add_el(datarow, "dat1", d_str)
        add_el(datarow, "art", "Geldzuwendung")
        add_el(datarow, "ja_nein", "nein")
        add_el(datarow, "betrag1", f"{float(row[COL_BET_RAG] if 'COL_BET_RAG' in locals() else row[COL_BETRAG]):.2f}")

    # 7. Summen aktualisieren
    gesamtsumme = df[COL_BETRAG].sum()
    root.xpath("//ns:element[@id='gesamtsumme']", namespaces=ns)[0].text = f"{gesamtsumme:.2f}"
    
    if num2words:
        try:
            wort_el = root.xpath("//ns:element[@id='wert2']", namespaces=ns)[0]
            wort_el.text = num2words(int(gesamtsumme), lang='de').upper()
        except: pass

    # 8. Speichern
    tree.write(output_datei, encoding="UTF-8", xml_declaration=True, pretty_print=True)
    print(f"--- ERFOLG ---")
    print(f"Datei erstellt: {output_datei}")
    print(f"Empfänger:      {name_string}")
    print(f"Gesamtsumme:    {gesamtsumme:.2f}")

if __name__ == "__main__":
    # Erwartet: script.py + 7 Argumente = 8 insgesamt
    if len(sys.argv) < 8:
        print("\nFehler: Zu wenige Argumente!")
        print("Nutzung:")
        print("python3 script.py DATEI VORNAME NACHNAME STRASSE HAUSNR PLZ ORT")
        print("\nBeispiel:")
        print("python3 script.py daten.xlsx Max Mustermann Straße 1 12345 Musterort")
    else:
        generate_xml(
            sys.argv[1], # datei
            sys.argv[2], # vorname
            sys.argv[3], # nachname
            sys.argv[4], # strasse
            sys.argv[5], # hausnr
            sys.argv[6], # plz
            sys.argv[7]  # ort
        )