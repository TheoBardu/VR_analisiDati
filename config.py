READING_CSV_DATA_FOLDER_NAME = ["misD","misE","misF"]
READING_TXT_DATA_FOLDER_NAME = ["misW"]
READING_EXEL_DATA_FOLDER_NAME = ["misA", "misB","misC","misG","misH"]

SHEET_NAME_XLSX = 'Profilo storico'

#Variabile per la scrittura
NAME_RILIEVI_FONOMETRICI = "Rilievi_Fonometrici.xlsx"

# Variabili per la lettura del file modello_scheda_gruppi_dpi
NOME_VR8h_totale = 'VR8h_totale.xlsx'
NOME_VR8h_riepilogo = 'VR8h_riepilogo.xlsx'
NOME_VR8h_aggiornato = 'VR8h_totale_aggiornato.xlsx'

# Nomi schede file scheda_gruppi_dpi.xlsx
SCHEDA_MANSIONI = 'Scheda_mansioni' 
SCHEDA_DPI = 'Scheda_DPI'
# Nomi colonne Scheda_HEG_dpi
Nome_colonna_IDgrom = 'ID_GrOm'
Nome_colonna_Descrizione_GrOm = 'Descrizione_GrOm'

# Layout blocco intestazione GrOm/reparto (righe 1-4, colonne A-B) nei fogli "Scheda N"
RIGA_INIZIO_INTESTAZIONE_GROM = 1
COL_ETICHETTA_INTESTAZIONE_GROM = 1   # colonna A: etichette
COL_VALORE_INTESTAZIONE_GROM = 2      # colonna B: valori

# Layout blocco "VALUTAZIONE SU BASE GIORNALIERA" nei fogli "Scheda N"
RIGA_INIZIO_VALUTAZIONE = 7           # era riga 2
COL_INIZIO_VALUTAZIONE = 1            # era ultima_colonna_popolata + 3
OFFSET_ETICHETTA_VALUTAZIONE = 6      # colonna (relativa a COL_INIZIO_VALUTAZIONE) dell'etichetta
                                       # "LEX MAX ="/"Massimo dei Lpicco,C misurati =". Usata sia da
                                       # inserisci_valutazione_schede (per scriverla) sia da
                                       # formatta_dimensioni_celle (per restringerne la larghezza):
                                       # non duplicare questo offset altrove.

# VARIABILI VALUAZIONE DPI in excel_aggiornato
COL_INIZIO_DPI = COL_INIZIO_VALUTAZIONE   # deve coincidere: applica_DPI_HML cerca "CLASSE RISCHIO"
                                           # solo in questa colonna, che è quella in cui
                                           # inserisci_valutazione_schede scrive quel blocco.
FIND_TESTO_FINE_TABELLA_VALUTAZIONE = "CLASSE RISCHIO"
SEPARAZIONE_RIGHE_DA_VALUTAZIONE = 1  # era 4
TESTO_TITOLO_DPI = "Analisi DPI in dotazione"

# Layout tabella misure nei fogli "Scheda N" (dopo il blocco DPI)
SEPARAZIONE_RIGHE_DPI_MISURE = 2
TESTO_INTESTAZIONE_MISURE = "ID_misura"
COLONNE_TABELLA_MISURE = ['ID_misura', 'Descrizione_compito', 'Ti', 'WBV', 'HAV', 'U', 'LeqA', 'LeqC', 'Ppeak']
COLORE_FILL_U = '87CEEB'      # azzurro, stessa palette di formatting_excel_VR8h_totale
COLORE_FILL_LEQA = 'FFA07A'   # arancione
COLORE_FILL_PPEAK = 'FF4500'  # rosso scuro