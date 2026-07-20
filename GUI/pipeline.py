#! /opt/anaconda3/bin/python3

"""
Copia refactorata di main.py: la stessa sequenza di analisi, ma spezzata in
step parametrici richiamabili singolarmente, senza input() bloccanti e senza
os.system.

main.py resta la versione di riferimento per l'esecuzione da terminale: questo
modulo esiste per poter essere pilotato dalla GUI.
"""

import os
import sys
from dataclasses import dataclass
from os import chdir, getcwd

# La root del progetto deve essere importabile anche se il modulo viene
# caricato da una cwd arbitraria.
_ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if _ROOT not in sys.path:
    sys.path.insert(0, _ROOT)

from analisi_datiVR import manager, analisi
from config import (
    NAME_RILIEVI_FONOMETRICI,
    NOME_VR8h_aggiornato,
    NOME_VR8h_riepilogo,
    NOME_VR8h_totale,
)

# Default presi da main.py
DEFAULT_MISURE = '/misure'
DEFAULT_RISULTATI = '/output'
DEFAULT_DPI = '/DPI_check'

NOME_SCHEDA_GRUPPI_DPI = 'scheda_gruppi_dpi.xlsx'

# Testo mostrato all'utente nel punto in cui main.py:49 chiama input()
MSG_CONTROLLO_TI_GROM = (
    "Controlla il file averaged_data.csv e inserisci a mano i Ti e i GrOm, "
    "poi premi Continua."
)


class PipelineAbort(Exception):
    """Sollevata da un callback di step per fermare la pipeline in modo pulito."""


@dataclass
class Paths:
    """Percorsi di lavoro. Solo main_directory e' obbligatoria."""

    main_directory: str
    misure: str = DEFAULT_MISURE
    risultati: str = DEFAULT_RISULTATI
    dpi: str = DEFAULT_DPI

    @property
    def misure_dir(self):
        return self.main_directory + self.misure

    @property
    def data_folder(self):
        return self.misure_dir + '/data'

    @property
    def risultati_dir(self):
        return self.main_directory + self.risultati

    @property
    def dpi_dir(self):
        return self.main_directory + self.dpi

    @property
    def scheda_dpi_xlsx(self):
        return os.path.join(self.main_directory, NOME_SCHEDA_GRUPPI_DPI)

    @property
    def vr8h_totale(self):
        return self.risultati_dir + '/' + NOME_VR8h_totale

    @property
    def vr8h_riepilogo(self):
        return self.risultati_dir + '/' + NOME_VR8h_riepilogo

    @property
    def vr8h_aggiornato(self):
        return self.risultati_dir + '/' + NOME_VR8h_aggiornato

    @property
    def rilievi_fonometrici(self):
        return os.path.join(self.main_directory, NAME_RILIEVI_FONOMETRICI)


def _richiedi(path, cosa):
    if not os.path.exists(path):
        raise FileNotFoundError(f"{cosa} non trovato: {path}")
    return path


# ── Step ───────────────────────────────────────────────────────────────────────

def step_lettura_misure(paths, file_name='dati.txt', formato='csv', versione='1'):
    """Legge i file di misura ed esporta csv/xlsx in <misure>/data. (main.py:20-26)

    manager() fa getcwd() nel costruttore e iterate_directory() fa chdir() per
    ogni cartella senza mai ripristinarla: la cwd va salvata e rimessa a posto.
    """
    _richiedi(paths.misure_dir, "Cartella delle misure")
    cwd_iniziale = getcwd()
    try:
        chdir(paths.misure_dir)
        m = manager()
        m.iterate_directory(file_name=file_name, format=formato,
                            versione_lettura=versione)
    finally:
        chdir(cwd_iniziale)
    print(f"Lettura misure completata (cwd ripristinata: {getcwd()})")


def step_medie(paths):
    """Calcola le medie delle misure -> averaged_data.csv. (main.py:32-34)"""
    _richiedi(paths.data_folder, "Cartella dei dati elaborati")
    a = analisi(paths.data_folder)
    return a.average_values()


def step_scheda_info(paths, df_avg=None):
    """Copia Ti e GrOm dalla scheda gruppi DPI sulle medie -> df_HEG. (main.py:39)"""
    _richiedi(paths.scheda_dpi_xlsx, NOME_SCHEDA_GRUPPI_DPI)
    if df_avg is None:
        df_avg = step_medie(paths)
    a = analisi(paths.data_folder)
    return a.get_scheda_info(df_avg, excel_info_dir=paths.main_directory)


def step_analisi_8h(paths, df_HEG=None):
    """Valutazione del rischio su 8 h. (main.py:57)"""
    if df_HEG is None:
        df_HEG = step_scheda_info(paths)
    a = analisi(paths.data_folder)
    a.analisi_8h(paths.risultati_dir, df_HEG)
    print(f"Analisi 8h completata: {paths.risultati_dir}")


def step_dpi(paths):
    """Applica il metodo HML dei DPI. (main.py:64-67)

    applica_DPI_HML cattura FileNotFoundError limitandosi a stamparlo e poi
    prosegue su variabili non inizializzate: i controlli vanno fatti prima.
    """
    _richiedi(paths.scheda_dpi_xlsx, NOME_SCHEDA_GRUPPI_DPI)
    _richiedi(paths.vr8h_totale, NOME_VR8h_totale)
    _richiedi(paths.vr8h_riepilogo, NOME_VR8h_riepilogo)

    a = analisi(paths.data_folder)
    a.applica_DPI_HML(
        excel_info_scheda_dpi=paths.scheda_dpi_xlsx,
        excel_total=paths.vr8h_totale,
        excel_output=paths.vr8h_riepilogo,
        excel_aggiornato=paths.vr8h_aggiornato,
    )
    print("Applicazione DPI completata")


def step_rilievi_fonometrici(paths):
    """Crea il file dei rilievi fonometrici. (main.py:70, senza os.system)"""
    from utility.crea_excel_dati import (
        load_averaged_data, load_mis_files, load_scheda, write_excel,
    )

    df_avg = load_averaged_data(paths.data_folder)
    df_mis = load_mis_files(paths.data_folder)
    df_scheda = load_scheda(paths.main_directory)
    write_excel(df_avg, df_mis, df_scheda, paths.rilievi_fonometrici)
    print(f"{NAME_RILIEVI_FONOMETRICI} creato correttamente")
    return paths.rilievi_fonometrici


def step_export_pdf(paths):
    """Esporta in PDF tutti i fogli del file aggiornato. (main.py:74-75)"""
    from utility.export_excel2pdf import esporta_pdf

    _richiedi(paths.vr8h_aggiornato, NOME_VR8h_aggiornato)
    try:
        return esporta_pdf(paths.vr8h_aggiornato)
    except FileNotFoundError as e:
        # _export_libreoffice fallisce cosi' se il binario soffice manca
        if 'soffice' in str(e) or 'LibreOffice' in str(e):
            raise RuntimeError(
                "LibreOffice (soffice) non trovato: e' necessario per l'export "
                "in PDF. Installa LibreOffice e riprova."
            ) from e
        raise


# ── Orchestratore ──────────────────────────────────────────────────────────────

# (chiave, etichetta) nell'ordine di main.py
STEPS = [
    ('lettura_misure',      'Leggi misure (iterate_directory)'),
    ('medie',               'Calcola medie'),
    ('scheda_info',         'Leggi scheda gruppi DPI'),
    ('controllo_ti_grom',   'Controllo manuale Ti / GrOm'),
    ('analisi_8h',          'Analisi 8h'),
    ('dpi',                 'Applica DPI'),
    ('rilievi_fonometrici', 'Crea Rilievi Fonometrici'),
    ('export_pdf',          'Esporta XLSX -> PDF'),
]


def run_pipeline(paths, on_step=None, opts=None):
    """Esegue tutti gli step nell'ordine di main.py.

    on_step(chiave, etichetta, indice, totale) viene chiamato PRIMA di ogni
    step; puo' sollevare PipelineAbort per fermare la pipeline in modo pulito.
    Lo step 'controllo_ti_grom' non esegue nulla: e' il punto in cui main.py:49
    chiama input(), qui diventa uno step-boundary come gli altri.
    """
    opts = opts or {}
    totale = len(STEPS)
    stato = {}

    for i, (chiave, etichetta) in enumerate(STEPS):
        if on_step is not None:
            on_step(chiave, etichetta, i, totale)

        if chiave == 'lettura_misure':
            step_lettura_misure(
                paths,
                file_name=opts.get('file_name', 'dati.txt'),
                formato=opts.get('formato', 'csv'),
                versione=opts.get('versione', '1'),
            )
        elif chiave == 'medie':
            stato['df_avg'] = step_medie(paths)
        elif chiave == 'scheda_info':
            stato['df_HEG'] = step_scheda_info(paths, stato.get('df_avg'))
        elif chiave == 'controllo_ti_grom':
            print(MSG_CONTROLLO_TI_GROM)
        elif chiave == 'analisi_8h':
            step_analisi_8h(paths, stato.get('df_HEG'))
        elif chiave == 'dpi':
            step_dpi(paths)
        elif chiave == 'rilievi_fonometrici':
            step_rilievi_fonometrici(paths)
        elif chiave == 'export_pdf':
            step_export_pdf(paths)

    print("Pipeline completata.")
    return stato


if __name__ == '__main__':
    try:
        run_pipeline(Paths(main_directory=sys.argv[1]))
    except (KeyboardInterrupt, PipelineAbort):
        print('End Program')
