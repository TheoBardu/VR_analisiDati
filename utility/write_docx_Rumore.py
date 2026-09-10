#! /opt/anaconda3/bin/python3

# Scrittura automatica della relazione di valutazione del rischio rumore (.docx).
#
# Il documento viene compilato a partire dal template corretto da fix_template_RUM.py
# e da due file excel gia' prodotti dalla pipeline:
#   - scheda_gruppi_dpi.xlsx  -> mansioni, gruppi omogenei, DPI, esposizione a vibrazioni
#   - VR8h_riepilogo.xlsx     -> Lex8h, incertezza, Lex max, picco e classe di rischio
#
# Il file e' diviso in sezioni: prima quello che l'utente deve compilare a mano,
# poi le parti che il codice recupera da solo.

import sys
import tempfile
from os import path

from docxtpl import DocxTemplate, InlineImage, RichText
from docx.shared import Mm

# insert e non append: esiste un pacchetto 'config' in site-packages che altrimenti vince
sys.path.insert(0, path.dirname(path.dirname(path.abspath(__file__))))

from config import (SCHEDA_MANSIONI, SCHEDA_DPI, SHEET_RIEPILOGO, NOME_VR8h_riepilogo,
                    Nome_colonna_IDgrom, Nome_colonna_Descrizione_GrOm,
                    COLORI_CLASSE_RISCHIO)


# ==========================================================================
# SEZIONE 0 - PERCORSI DEI FILE
# ==========================================================================
# Cartella dell'azienda: contiene scheda_gruppi_dpi.xlsx e la sottocartella output.
MAIN_DIRECTORY = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/AL7/rev/rev2/Rumore'
OUTPUT_DIRECTORY = MAIN_DIRECTORY + '/output'

# File excel di input
FILE_SCHEDA_GRUPPI_DPI = MAIN_DIRECTORY + '/scheda_gruppi_dpi.xlsx'
FILE_VR8H_RIEPILOGO    = OUTPUT_DIRECTORY + '/' + NOME_VR8h_riepilogo

# Template word gia' corretto (prodotto da utility/fix_template_RUM.py)
DIR_MODELLI       = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Modelli/docx/strutture'
DOCUMENTO_WORD_TEMPLATE = DIR_MODELLI + '/Modello_RUM.docx'

# Cartella con i frontespizi selezionabili e nome di quello da usare
DIR_FRONTESPIZI = DIR_MODELLI + '/frontespizi'
FRONTESPIZIO    = 'frontespizio_relyon.docx'

# Logo aziendale inserito nell'intestazione e nel frontespizio
LOGO_AZIENDA = '/Users/theo/Desktop/1631305251718.jpeg'
LARGHEZZA_LOGO_MM = 50

# Documento prodotto
OUTPUT_DOCUMENT = OUTPUT_DIRECTORY + '/Relazione_RUM.docx'


# ==========================================================================
# SEZIONE 1 - PARTI DA COMPILARE A MANO
# ==========================================================================

# Flag orari: se True tutte le mansioni condividono lo stesso orario e la tabella
# ha una sola riga; se False il codice carica le mansioni da scheda_gruppi_dpi.xlsx e
# assegna a tutte l'orario di default, che poi l'utente modifica una per una nel docx.
ORARI_UGUALI_PER_TUTTI = True
ORARIO_LAVORO_DEFAULT  = "Lunedì – Venerdì \n 8:00÷12:00 13:00÷17:00"

# Colonne "Esposizione a ototossici" e "Rumori impulsivi" del quadro sinottico HEG:
# non sono presenti negli excel, restano al valore di default e vanno riviste a mano.
OTOTOSSICI_DEFAULT = "NO"
IMPULSIVI_DEFAULT  = "NO"

context = {
    # Compila i campi vuoti
    "nome_azienda": "PARESA S.R.L.",
    "indirizzo_azienda": "vicolo malvasia 980 (FC)",
    "data_emissione": "16/03/2026",
    "revisione": "rev.01",
    "data_scadenza": "16/03/2030",
    "datore_di_lavoro": "",
    "RSPP": "",
    "medico_competente": "",
    "RLS": "",
    "delegato_sicurezza": "",
    # Info generali azienda
    "attivita_azienda": "",
    "processo_produttivo": "",
    "sede_legale": "",
    "sede_operativa": "",
    "date_misurazione": ["05 dicembre 2025 dalle ore 08:00 alle ore 16:00",
                         "",
                         ""],   # le voci vuote vengono scartate
    "sostanze_ototossiche": "Si",   # presenza di sostanze ototossiche o meno
    "misure_attuative_ototossiche": "Si faccia riferimento al documento di valutazione del rischio chimico.",
    "interazione_vib_rum": "Si",    # presenza di interazione tra rumore e vibrazione
    "misure_attuative_vib_rum": "Certamente si considerato l’utilizzo di attrezzature elettriche portatili. Vi è dunque trasmissione ossea delle vibrazioni e del rumore all’orecchio medio. Si faccia riferimento alla valutazione del rischio chimico.",
    "effetti_indesiderati": "Si",
    "misure_attuative_effetti_indesiderati": "Nelle zone/postazioni di lavoro è possibile che gli addetti possano incorrere in tali situazioni. Si consiglia pertanto di utilizzare D.P.I. con grado di protezione SNR come prescritto dalla presente relazione e l’adozione di sistemi alternativi quali segnali oto-acustici.",
}

# Motivazioni del metodo di valutazione adottato (tabella "Metodo adottato")
valutazione_context = {
    "base_giornaliera": "",
    "base_settimanale": "",
    "esposizioni_variabili": "",
}


# ==========================================================================
# SEZIONE 2 - FUNZIONI DI CARICAMENTO AUTOMATICO DAGLI EXCEL
# ==========================================================================

def carica_dpi(file_scheda_gruppi_dpi):
    '''
    Carica la tabella dei DPI dal foglio Scheda_DPI.

    INPUT:
        file_scheda_gruppi_dpi = <str>, percorso del file scheda_gruppi_dpi.xlsx

    OUTPUT:
        <list> di dizionari con le chiavi attese dal template word
        (codice_DPI, descrizione, marca, modello, snr, H, L, M, note).
    '''
    from analisi_datiVR import files

    df_dpi = files.get_scheda_DPI(path.dirname(file_scheda_gruppi_dpi),
                                  path.basename(file_scheda_gruppi_dpi),
                                  sheet_name=SCHEDA_DPI)
    if df_dpi is None:
        raise RuntimeError(f'Lettura del foglio {SCHEDA_DPI} non riuscita: {file_scheda_gruppi_dpi}')

    # Mappatura per nome e non per posizione: nell'excel l'ordine e' H, L, M mentre
    # nel word l'intestazione e' H, M, L.
    return [
        {"codice_DPI": riga['codice_dpi'],
         "descrizione": riga['descrizione'],
         "marca": riga['marca'],
         "modello": riga['modello'],
         "snr": formatta_numero(riga['SNR'], decimali=0),
         "H": formatta_numero(riga['H'], decimali=0),
         "L": formatta_numero(riga['L'], decimali=0),
         "M": formatta_numero(riga['M'], decimali=0),
         "note": ""}   # colonna non presente nell'excel
        for _, riga in df_dpi.iterrows()
    ]


def leggi_scheda_mansioni(file_scheda_gruppi_dpi):
    '''
    Legge il foglio Scheda_mansioni saltando la riga di titolo.

    OUTPUT:
        <pd.DataFrame> con le colonne originali del foglio.
    '''
    import pandas as pd

    # ID_GrOm letto come testo: gli ID sono etichette (es. '8.2', 'A3'), non numeri
    return pd.read_excel(file_scheda_gruppi_dpi, sheet_name=SCHEDA_MANSIONI, skiprows=1,
                         dtype={Nome_colonna_IDgrom: str})


def carica_mansioni(df_mansioni):
    '''
    Estrae l'elenco dei gruppi omogenei (numero scheda + descrizione), una voce per
    ID_GrOm, mantenendo l'ordine del foglio.

    OUTPUT:
        <list> di dizionari {"ID": <ID scheda, str>, "Mansione": <descrizione>}.
    '''
    df_gruppi = df_mansioni.dropna(subset=[Nome_colonna_IDgrom])
    df_gruppi = df_gruppi.drop_duplicates(subset=[Nome_colonna_IDgrom], keep='first')

    return [
        {"ID": formatta_id(riga[Nome_colonna_IDgrom]),
         "Mansione": str(riga[Nome_colonna_Descrizione_GrOm]).strip()}
        for _, riga in df_gruppi.iterrows()
    ]


def carica_orari(mansioni):
    '''
    Costruisce la tabella degli orari di lavoro.
    Con ORARI_UGUALI_PER_TUTTI a True restituisce una riga sola; altrimenti una riga
    per mansione, tutte con l'orario di default.

    L'orario e' passato come RichText perche' docxtpl non converte il \\n di una
    stringa normale in un a capo del documento.
    '''
    orario = RichText(ORARIO_LAVORO_DEFAULT)

    if ORARI_UGUALI_PER_TUTTI:
        return [{"mansione": "Tutte le mansioni", "orario_lavoro": orario}]

    return [{"mansione": voce["Mansione"], "orario_lavoro": RichText(ORARIO_LAVORO_DEFAULT)}
            for voce in mansioni]


def carica_vibrazioni(df_mansioni):
    '''
    Ricava per ogni gruppo omogeneo l'esposizione a vibrazioni dalle colonne WBV e HAV
    del foglio Scheda_mansioni.

    OUTPUT:
        <dict> {ID_GrOm (str): "HAV" / "WBV" / "HAV + WBV" / "NO"}
    '''
    import pandas as pd

    esposizione = {}

    for id_grom, gruppo in df_mansioni.groupby(Nome_colonna_IDgrom, sort=False):
        presenti = [colonna for colonna in ('HAV', 'WBV')
                    if colonna in gruppo.columns and gruppo[colonna].notna().any()]
        chiave = formatta_id(id_grom)
        esposizione[chiave] = ' + '.join(presenti) if presenti else 'NO'

    return esposizione


def carica_riepilogo_heg(file_riepilogo, esposizione_vibrazioni):
    '''
    Carica il quadro sinottico dal foglio Riepilogo di VR8h_riepilogo.xlsx.
    La lettura e' fatta con openpyxl e non con pandas perche' serve anche il colore di
    riempimento della cella classe_rischio, da riportare identico nel documento word.

    OUTPUT:
        <list> di dizionari, uno per gruppo omogeneo, con i campi attesi dal template
        (numero_scheda, gruppo_HEG, lex8h, U, lexmax, peakmax, classe_rischio, colore,
        vib, oto, imp). classe_rischio e' un RichText per poter forzare il colore del
        testo, colore e' il riempimento della cella usato da {% cellbg %}; la chiave
        classe conserva la classe come stringa, usata per filtrare le tabelle.
    '''
    from openpyxl import load_workbook

    foglio = load_workbook(file_riepilogo, data_only=True)[SHEET_RIEPILOGO]

    intestazioni = {str(cella.value).strip(): cella.column
                    for cella in foglio[1] if cella.value is not None}

    mancanti = [nome for nome in ('ID_GrOm', 'Mansione', 'Lex8h', 'U', 'Lex_max',
                                  'L_picco_C', 'classe_rischio')
                if nome not in intestazioni]
    if mancanti:
        raise RuntimeError(f'Colonne mancanti nel foglio {SHEET_RIEPILOGO}: {mancanti}')

    def valore(riga, nome):
        return foglio.cell(row=riga, column=intestazioni[nome]).value

    righe = []
    for numero_riga in range(2, foglio.max_row + 1):
        if valore(numero_riga, 'ID_GrOm') is None:
            continue

        id_grom = formatta_id(valore(numero_riga, 'ID_GrOm'))
        classe = str(valore(numero_riga, 'classe_rischio') or '').strip().upper()
        cella_classe = foglio.cell(row=numero_riga, column=intestazioni['classe_rischio'])

        righe.append({
            "numero_scheda": id_grom,
            "gruppo_HEG": str(valore(numero_riga, 'Mansione') or '').strip(),
            "lex8h": formatta_numero(valore(numero_riga, 'Lex8h')),
            "U": formatta_numero(valore(numero_riga, 'U')),
            "lexmax": formatta_numero(valore(numero_riga, 'Lex_max')),
            "peakmax": formatta_numero(valore(numero_riga, 'L_picco_C')),
            "classe": classe,
            "classe_rischio": RichText(classe, color=colore_font_classe(classe), bold=True),
            "colore": colore_sfondo_classe(cella_classe, classe),
            "vib": esposizione_vibrazioni.get(id_grom, 'NO'),
            "oto": OTOTOSSICI_DEFAULT,
            "imp": IMPULSIVI_DEFAULT,
        })

    return righe


def colore_sfondo_classe(cella, classe):
    '''
    Colore di riempimento della cella excel, normalizzato a 6 cifre esadecimali
    (openpyxl lo restituisce come AARRGGBB). Se la cella non e' colorata si ricade
    sulla palette condivisa in config.py.
    '''
    rgb = getattr(cella.fill.fgColor, 'rgb', None)
    if isinstance(rgb, str) and len(rgb) in (6, 8):
        rgb = rgb[-6:]
        if rgb != '000000':
            return rgb

    return COLORI_CLASSE_RISCHIO.get(classe, {}).get('sfondo', 'FFFFFF')


def colore_font_classe(classe):
    '''Colore del testo della classe di rischio (bianco sul rosso scuro del rischio ALTA).'''
    return COLORI_CLASSE_RISCHIO.get(classe, {}).get('font', '000000')


def filtra_per_classe(righe_heg, classe):
    '''Sottoinsieme del quadro sinottico per una data classe di rischio.'''
    return [riga for riga in righe_heg if riga['classe'] == classe]


def formatta_id(valore):
    '''
    Restituisce un ID (gruppo omogeneo, scheda, ...) come testo, senza mai passare
    per float: '8.2' resta '8.2', 'A3' resta 'A3'. Solo se la cella excel e' un
    numero intero salvato come float (8.0) viene riportato senza il '.0'.
    '''
    import pandas as pd

    if valore is None or (isinstance(valore, float) and pd.isna(valore)):
        return ''
    if isinstance(valore, float) and valore.is_integer():
        return str(int(valore))
    return str(valore).strip()


def formatta_numero(valore, decimali=1):
    '''
    Formatta un valore numerico per la scrittura nel documento. I valori non numerici
    (o vuoti) vengono restituiti come stringa senza modifiche.
    '''
    import pandas as pd

    if valore is None or (isinstance(valore, float) and pd.isna(valore)):
        return ''
    try:
        return f'{float(valore):.{decimali}f}'
    except (TypeError, ValueError):
        return str(valore).strip()


def logo_inline(template):
    '''
    Logo aziendale come immagine da inserire nel documento. InlineImage e' legata al
    template che la ospita, quindi ne va creata una per ogni documento renderizzato.
    Se il file non esiste il logo viene lasciato vuoto invece di bloccare la scrittura.
    '''
    if not path.isfile(LOGO_AZIENDA):
        print(f'ATTENZIONE: logo non trovato, campo lasciato vuoto -> {LOGO_AZIENDA}')
        return ''

    return InlineImage(template, LOGO_AZIENDA, width=Mm(LARGHEZZA_LOGO_MM))


def costruisci_frontespizio(documento, file_frontespizio, contesto):
    '''
    Renderizza il frontespizio scelto come documento a se' stante e lo restituisce come
    sottodocumento da inserire nel template principale.

    Il render va fatto in due passate: docxtpl inserisce il sottodocumento cosi' com'e',
    senza rielaborarne i tag Jinja, quindi le variabili del frontespizio vanno risolte
    prima. Il file intermedio viene scritto in una cartella temporanea, cancellata
    subito dopo la lettura, per non lasciare residui nella cartella di output.
    '''
    if not path.isfile(file_frontespizio):
        from glob import glob
        disponibili = sorted(path.basename(f) for f in glob(path.join(DIR_FRONTESPIZI, '*.docx')))
        raise FileNotFoundError(f'Frontespizio non trovato: {file_frontespizio}\n'
                                f'Frontespizi disponibili in {DIR_FRONTESPIZI}: {disponibili}')

    template_frontespizio = DocxTemplate(file_frontespizio)
    contesto_frontespizio = dict(contesto)
    # L'immagine e' legata al documento che la ospita: ne serve una per ogni template.
    contesto_frontespizio['img_logo_azienda'] = logo_inline(template_frontespizio)
    template_frontespizio.render(contesto_frontespizio)

    with tempfile.TemporaryDirectory() as cartella_temporanea:
        frontespizio_reso = path.join(cartella_temporanea, '_frontespizio_reso.docx')
        template_frontespizio.save(frontespizio_reso)
        # new_subdoc legge subito il file: la cartella puo' essere rimossa qui.
        return documento.new_subdoc(frontespizio_reso)


# ==========================================================================
# SEZIONE 3 - COSTRUZIONE DEI CONTEXT AUTOMATICI
# ==========================================================================
doc = DocxTemplate(DOCUMENTO_WORD_TEMPLATE)

df_mansioni = leggi_scheda_mansioni(FILE_SCHEDA_GRUPPI_DPI)
mansioni = carica_mansioni(df_mansioni)
righe_heg = carica_riepilogo_heg(FILE_VR8H_RIEPILOGO, carica_vibrazioni(df_mansioni))

context_tabella_dpi = {
    "tabella_dpi": carica_dpi(FILE_SCHEDA_GRUPPI_DPI),
}

context_tabella_orari = {
    "tabella_orario_lavoro_mansione": carica_orari(mansioni),
}

context_tabella_mansioni = {
    "tabella_mansioni": mansioni,
}

context_tabella_heg = {
    "tabella_HEG": righe_heg,
}

context_tabella_heg_medio = {
    "HEG_med": filtra_per_classe(righe_heg, 'MEDIA'),
}

context_tabella_heg_alto = {
    "HEG_alto": filtra_per_classe(righe_heg, 'ALTA'),
}

# Il template inserisce le date dei rilievi a meta' frase: serve una stringa, non una lista.
context["date_misurazione"] = ', '.join(data for data in context["date_misurazione"] if data.strip())

# Unione di tutti i dizionari
context_completo = (context | valutazione_context | context_tabella_dpi |
                    context_tabella_orari | context_tabella_mansioni |
                    context_tabella_heg | context_tabella_heg_medio |
                    context_tabella_heg_alto)

context_completo["frontespizio"] = costruisci_frontespizio(
    doc, path.join(DIR_FRONTESPIZI, FRONTESPIZIO), context_completo)
context_completo["img_logo_azienda"] = logo_inline(doc)


# ==========================================================================
# SEZIONE 4 - SCRITTURA DEL DOCUMENTO
# ==========================================================================
doc.render(context_completo)
doc.save(OUTPUT_DOCUMENT)

print(f'Docx scritto: {OUTPUT_DOCUMENT}')
print(f'  frontespizio      : {FRONTESPIZIO}')
print(f'  mansioni / schede : {len(mansioni)}')
print(f'  DPI               : {len(context_tabella_dpi["tabella_dpi"])}')
print(f'  gruppi omogenei   : {len(righe_heg)} '
      f'(medio: {len(context_tabella_heg_medio["HEG_med"])}, '
      f'alto: {len(context_tabella_heg_alto["HEG_alto"])})')
