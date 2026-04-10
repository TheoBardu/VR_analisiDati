import pandas as pd
import openpyxl as ex
from openpyxl.styles import PatternFill
from numpy import zeros, arange, mean, std, max, round, ones, log10, sum, dot, sqrt
from os import chdir, getcwd, error


# file = '/Users/theo/Desktop/Ermes/Misure/misF/misF/LSOURCES - Copia.txt'

        

class exel_file:
    '''
    Classe che serve per la manipolazione dei file excel
    '''
    
    def color_column(file_path, column_letters, colors):
        '''
        Funzione che colora una colonna specifica di un file Excel.
        
        INPUT:
            file_path = <str>, percorso del file Excel
            column_letter = <list str>, lettera della colonna da colorare (es. 'A', 'B', ecc.)
            color = < list str>, colore in formato esadecimale (es. 'FFFF00' per il giallo)
        '''
        wb = ex.load_workbook(file_path) #carica il file exel
        ws = wb.active #seleziona il fogio attivo 


        
        idx_color = 0
        for letter in column_letters:
            fill = PatternFill(start_color=colors[idx_color], end_color=colors[idx_color], fill_type="solid")
            for cell in ws[letter]:
                cell.fill = fill
            idx_color += 1

        wb.save(file_path)
        print(f'Colonna/e {column_letters} colorata/e con successo')
    

    def adjust_column_lenght(file_path, column_letters):
        '''
        Funzione che adatta la lunghezza delle colonne in un file Excel.
        '''
        wb = ex.load_workbook(file_path) #cario il file excel
        ws = wb.active #seleziona il foglio attivo

        for letter in column_letters:
            max_lenght = 0 #ausiliaria, per salvare la lunghezza massima del nome della colonna
            for cell in ws[letter]:
                if len(str(cell.value)) > max_lenght:
                    max_lenght = len(str(cell.value))

            ws.column_dimensions[letter].width = max_lenght
            wb.save(file_path)


    def color_cell_VR8h(file_path, column_names = ['Peak', 'Leq_max','GrOm'], colors = ['e8643c','e8bd3c', '18AB49']):
        '''
        Funzione che colora la cella del file in maniera opportuna in base al rischio
        'e8643c','e8bd3c', '18AB49' : rosso, giallo, verde
        '''
        wb = ex.load_workbook(file_path) #carica il file exel
        ws = wb.active #seleziona il fogio attivo 
        
        #trovo l'indice delle colonne
        col = {}
        for cell in ws[1]:
            if cell.value in column_names:
                col[cell.value] = cell.column
        



        # itero sul numero di righe per colorare le celle
        for i in range(2 , ws.max_row + 1):
        
            #rischio basso
            if ws.cell(row=i, column = col[column_names[1]]).value <= 80.0 and ws.cell(row=i, column = col[column_names[0]]).value <= 135.0:
                #creo il settaggio del riempimento verde
                fill = PatternFill(start_color=colors[2], end_color=colors[2], fill_type="solid")
                ws.cell(row=i,column= col[column_names[2]]).fill = fill
            
            #rischio medio
            if ws.cell(row=i, column = col[column_names[1]]).value > 80.0 and ws.cell(row=i, column = col[column_names[1]]).value <= 85.0 and ws.cell(row=i, column = col[column_names[0]]).value > 135.0 and ws.cell(row=i, column = col[column_names[0]]).value <= 137.0:
                #creo il settaggio del riempimento giallo
                fill = PatternFill(start_color=colors[1], end_color=colors[1], fill_type="solid")
                ws.cell(row=i,column= col[column_names[2]]).fill = fill
            
            #rischio alto
            if ws.cell(row=i, column = col[column_names[1]]).value > 85.0 and ws.cell(row=i, column = col[column_names[1]]).value <= 87.0 and ws.cell(row=i, column = col[column_names[0]]).value > 137.0 and ws.cell(row=i, column = col[column_names[0]]).value <= 140.0:
                #creo il settaggio del riempimento giallo
                fill = PatternFill(start_color=colors[0], end_color=colors[0], fill_type="solid")
                ws.cell(row=i,column= col[column_names[2]]).fill = fill
            
            #rischio fuori scala (Warning)
            if ws.cell(row=i, column = col[column_names[1]]).value > 87 or ws.cell(row=i, column = col[column_names[0]]).value > 140.0:
                fill = PatternFill(start_color=colors[0], end_color=colors[0], fill_type="solid")
                ws.cell(row=i,column= col[column_names[2]]).fill = fill
                print('########## Attenzione ##########\n####################\nPossibile errore nel calcolo oppure valore rumore troppo alto!\nControlla i dati.\n####################')


        
        
        
        wb.save(file_path)



class files:
    '''
    Classe che serve per leggere e scrivere i file
    '''

    def read_measure_file(file, letter_ID, format='txt',ntrack = 6, decimals = 1):
        '''
        Funzione che legge il file di misura in txt o csv e restituisce un dataframe pandas.

        INPUT:
            file = <str>, directory del file da leggere
            letter_ID = <str>, lettera della cartella dati (tipicamente D,E,F o W)
            format = <str>, "txt" o "csv" a seconda se leggere in txt o csv
            ntrack = <int>, default = 3. Solo per format = 'csv'. Numero di tracce di divisione del file intero.
        OUTPUT:
            df = <pd.DataFrame>, dataframe contenente i dati della misura
        '''
        if format == 'txt':
            df = pd.DataFrame(columns=['fileID','letter_ID','nTrack','LeqA_min','LeqA_max','LeqA_eq','LeqC_min','LeqC_max','LeqC_eq','PeakC_max','PeakC_eq','durata','inizio','fine'])
            nFiles = 0 #numero di file nella misura
            

            #apertura e lettura del file
            with open(file, 'r', encoding='utf-16') as f:
                lines = f.readlines() #lettura delle linee

                
            #iterazione per leggere le singola linee
            for line in lines:
                    if 'File' in line:
                        nFiles += 1 #tengo conto del numero di misure presenti nel file
            
            # inizializzazione vettori
            # fileIDs = zeros(nFiles)
            # inizio = zeros(nFiles)
            # fine = zeros(nFiles)
            # nSorgenti = zeros(nFiles)
            fileIDs = []
            letter_IDs = []
            inizio=[]
            fine = []
            nSorgenti = []
            
            LeqA_min = [] # livello equalizzato A minimo 
            LeqA_max = [] # livello equalizzato A massimo
            LeqA_eq = [] # livello equalizzato A equivalente
            
            LeqC_min = [] # livello equalizzato C minimo
            LeqC_max = [] # livello equalizzato C massimo
            LeqC_eq = [] # livello equalizzato C equivalente

            PeakC_max = [] # picco C minimo
            PeakC_eq = [] # picco C massimo

            
            durata = [] # durata della traccia
            

            inxd_letter_IDs = 0 # tengo conto del numero di misura (es: D1, D2, D3, ...)
            indx_inizio = 0
            indx_fine = 0
            indx = 0
            for line in lines:
                #trovo dove sta la voce File
                if 'File' in line:
                    l0 = line.split() #questa è la riga del nome del file
                    # print(l)
                    # print(lines[indx])
                    # fileIDs[inxd_fileIDs] = l[1]
                    inxd_letter_IDs += 1
                    
                    l1 = lines[indx+1].split() # questa è la riga dell'inizio
                    

                    l2 = lines[indx+2].split() #questa è la riga della fine
                    
                    l3 = lines[indx+3].split() #questa è la riga del numero di sorgenti
                    
                    ntracks = arange(len(l3)-1) #creo un array con il numero di sorgenti in ordine crescente

                    #itero sul numero di sorgenti
                    for i in range(len(ntracks)):
                        fileIDs.append(l0[1]) #creo tante copie del nome sorgente quante sono le sorgenti
                        letter_IDs.append(letter_ID+str(inxd_letter_IDs)) #creo tante copie del letter ID  quante sono le sorgenti
                        inizio.append(l1[2]) # creo tante copie dell'inizio quante sono le sorgenti
                        fine.append(l2[2]) # creo tante copie della fine quante sono le sorgenti
                        nSorgenti.append(ntracks[i]+1) # creo un numero di sorgenti in ordine crescente

                    

                    #salvataggio livelli euqalizzati A e durata tracciati
                    l7 = lines[indx+7].split() #questa è la riga dei livelli LeqA
                    # print(l7)
                    for n in range(len(ntracks)):
                        LeqA_max.append(float(l7[5 + (n * 4)])); durata.append(l7[8 + (n*4)])
                        LeqA_min.append(float(l7[6 + (n * 4)]))
                        LeqA_eq.append(float(l7[7 + (n * 4)]))
            
                    # salvataggio livelli equalizzati C
                    l8 = lines[indx+8].split() #questa è la riga dei livelli C
                    # print(l8)
                    for n in range(len(ntracks)):
                        LeqC_max.append(float(l8[5 + (n * 4)]))
                        LeqC_min.append(float(l8[6 + (n * 4)]))
                        LeqC_eq.append(float(l8[7 + (n * 4)]))

                    #salvataggio picchi C
                    l9 = lines[indx+9].split()
                    # print(l9)
                    for n in range(len(ntracks)):
                        PeakC_eq.append(float(l9[5 + (n * 3)]))
                        PeakC_max.append(float(l9[6 + (n * 3)]))
                
                indx += 1

                

            df['fileID'] = fileIDs
            df['letter_ID'] = letter_IDs
            df['inizio'] = inizio
            df['fine'] = fine
            df['nTrack'] = nSorgenti
            df['LeqA_min'] = LeqA_min
            df['LeqA_max'] = LeqA_max
            df['LeqA_eq'] = LeqA_eq 
            df['LeqC_min'] = LeqC_min
            df['LeqC_max'] = LeqC_max
            df['LeqC_eq'] = LeqC_eq
            df['PeakC_max'] = PeakC_max
            df['PeakC_eq'] = PeakC_eq
            df['durata'] = durata
            # df['inizio'] = inizio
            # print(nFiles) #for debug
            print('Lettura e creazione dataFrame completata')
            return df, ntracks[-1], nFiles
        
        elif format == 'csv':
            import glob 
            from numpy import average, min, max

            #inizializzo il dataFrame pandas con il riassunto delle misure
            df_tot = pd.DataFrame(columns=['fileID','letter_ID','nTrack','LeqA_min','LeqA_max','LeqA_eq','LeqC_min','LeqC_max','LeqC_eq','PeakC_max','PeakC_eq','LASeq_T','LAIeq_T','durata','inizio','fine'])
            
            
            #Inizializzazione delle variabili di df_tot
            fileIDs = []
            letter_IDs = []
            inizio = []
            fine = []
            durata = []
            ntrack_id = []
            LeqA_min = []
            LeqA_max = []
            LeqA_eq = []
            LeqC_min = []
            LeqC_max = []
            LeqC_eq = []
            PeakC_max = []
            PeakC_eq = []
            LASeq_T = []
            LAIeq_T = []
            


            print('Reading csv files only')
            csv_files = glob.glob('*.csv') # salvo la lista di tutti i file csv che ci sono
            csv_files.sort() # riordino per nome la lista dei file

            letter_id_number = 1 #inizializzo il numero della misura ad 1 per ogni nuovo file misure ( in modo che d1,d2,d3...f1,f2,f3...)
            #itero sulla lista del numero di files csv
            for file in csv_files:
                
                df = pd.read_csv(file,encoding='latin', skiprows=1, sep=';', engine = 'python') #leggo il file csv delle misure
                df.iloc[:,0] = pd.to_datetime(df.iloc[:,0],format='%H:%M:%S', errors='coerce') # converto i dati della colonna in datetime
                df.dropna(subset='Time',inplace=True) #tolgo tutte le righe che sono NaN nella colonna time
                
                df[df.columns[1:7]] = df[df.columns[1:7]].apply(pd.to_numeric, errors='coerce')

                #identifico se la lunghezza della traccia può essere divisa perfettamente per il numero di tracce desiderate
                n = (len(df)-1)%ntrack
                if n == 0: #se può essere divisa perfettamente
                    sep = int((len(df)-1)/ntrack)
                else: #se non può essere divisa perfettamente
                    df.drop(arange(n),inplace=True) #tolgo gli ultimi valori utili per avere una divisibilità buona
                    df = df.reset_index(drop=True)
                    sep = int((len(df)-1)/ntrack)
                
                # print(sep) #for debug

                # Aggiunta valori alle liste <=============================
                

                ntrack_id_letter = 1 #variabile che mi tiene conto del numero di traccia di una misura (es: D1 track 1, D1 trak 2, D1 track 3, ...)
                for i in range(ntrack):
                    fileIDs.append(file) #salvo il nome del file 
                    letter_IDs.append(letter_ID + str(letter_id_number)) #salvo la lettera del file
                    ntrack_id.append(ntrack_id_letter)
                    ntrack_id_letter += 1
                    
                    # print(letter_ID) # for debug
                    # print('##',i, i+i*sep) #for debug

                    # Inizio e Fine
                    inizio.append(df['Time'][i + i * sep].time()) #salvo la fine
                    fine.append(df['Time'][(i+1)*sep].time()) #salvo l'inizio
                    
                    # Durata
                    dT = df['Time'][(i+1)*sep] - df['Time'][i + i * sep]
                    dT_str = f"{dT.components[1]:02d}:{dT.components[2]:02d}:{dT.components[3]:02d}"
                    durata.append(dT_str)

                    
                    #LeqA
                    LeqA_min.append(round(min(df['LAeq'][0+i*int(sep):(i+1)*int(sep)]),decimals))
                    LeqA_max.append(round(max(df['LAeq'][0+i*int(sep):(i+1)*int(sep)]),decimals))
                    LeqA_eq.append(round(10*log10(sum(10**(df['LAeq'][0+i*int(sep):(i+1)*int(sep)]/10))/(len(df['LAeq'][0+i*int(sep):(i+1)*int(sep)])) ),decimals))
                    # print(len(df['LAeq'][0+i*int(sep):(i+1)*int(sep)]))
                    # input()

                    #LeqC
                    LeqC_min.append(round(min(df['LCeq'][0+i*int(sep):(i+1)*int(sep)]),decimals))
                    LeqC_max.append(round(max(df['LCeq'][0+i*int(sep):(i+1)*int(sep)]),decimals))
                    LeqC_eq.append(round( 10*log10(sum(10**(df['LCeq'][0+i*int(sep):(i+1)*int(sep)]/10))/(len(df['LCeq'][0+i*int(sep):(i+1)*int(sep)])) ),decimals))

                    PeakC_max.append(round(max(df['LCpeak'][0+i*int(sep):(i+1)*int(sep)]),decimals))
                    PeakC_eq.append(round(average(df['LCpeak'][0+i*int(sep):(i+1)*int(sep)]),decimals))

                    #LAeqS and LAeqI
                    LASeq_T.append(round(10*log10(sum(10**(df['LASeqT'][0+i*int(sep):(i+1)*int(sep)]/10))/(len(df['LASeqT'][0+i*int(sep):(i+1)*int(sep)]))),decimals))
                    LAIeq_T.append(round(10*log10(sum(10**(df['LAIeqT'][0+i*int(sep):(i+1)*int(sep)]/10))/(len(df['LAIeqT'][0+i*int(sep):(i+1)*int(sep)]))),decimals))

                letter_id_number += 1
            
            df_tot['fileID'] = fileIDs
            df_tot['letter_ID'] = letter_IDs
            df_tot['nTrack'] = ntrack_id
            df_tot['inizio'] = inizio
            df_tot['fine'] = fine
            df_tot['durata'] = durata
            df_tot['LeqA_min'] = LeqA_min
            df_tot['LeqA_max'] = LeqA_max
            df_tot['LeqA_eq'] = LeqA_eq 
            df_tot['LASeq_T'] = LASeq_T
            df_tot['LAIeq_T'] = LAIeq_T
            df_tot['LeqC_min'] = LeqC_min
            df_tot['LeqC_max'] = LeqC_max
            df_tot['LeqC_eq'] = LeqC_eq
            df_tot['PeakC_max'] = PeakC_max
            df_tot['PeakC_eq'] = PeakC_eq
            
            

            return df_tot




        else:
            error("Sono supportati solo formati 'txt' o 'csv' nella voce format")


    def write_csv(df,file):
        '''
        Funzione che scrive il dataframe df in un file csv
        '''
        df.to_csv(file)
        print('Salvataggio in csv del file completato')


    def write_exel(df, file):
        '''
        Funzione che scrive il dataframe df df in un file excel
        '''
        df.to_excel(file, index=False)
        print('Salvataggio del file exel completato')


    def get_scheda_DPI(excel_info_dir, name_excel_info, sheet_name='Scheda_DPI'):
        '''
        Funzione che legge il foglio "Scheda_DPI" del file excel informativo e restituisce
        un dataframe pandas con i dati dei dispositivi di protezione individuale (DPI).

        INPUT:
            excel_info_dir   = <str>, directory in cui è salvato il file excel informativo
            name_excel_info  = <str>, nome del file excel informativo (es. 'scheda_gruppi_dpi.xlsx')
            sheet_name       = <str>, nome del foglio da leggere (default: 'Scheda_DPI')

        OUTPUT:
            df_dpi = <pd.DataFrame>, dataframe con le seguenti colonne:
                - codice_dpi  : codice identificativo del DPI (str)
                - descrizione : descrizione del DPI (str)
                - marca       : marca del DPI (str)
                - modello     : modello del DPI (str)
                - SNR         : valore SNR del DPI (float)
                - H           : valore di attenuazione alle alte frequenze (float)
                - L           : valore di attenuazione alle basse frequenze (float)
                - M           : valore di attenuazione alle medie frequenze (float)
                - beta        : coefficiente correttivo per l'attenuazione reale (float)
        '''
        import os
        import warnings

        file_path = os.path.join(excel_info_dir, name_excel_info)

        try:
            # Leggo il foglio saltando la riga del titolo (riga 1),
            # usando la riga 2 come intestazione e i dati dalla riga 3 in avanti
            df_raw = pd.read_excel(
                file_path,
                sheet_name=sheet_name,
                skiprows=1,   # salta la riga del titolo "Tabella DPI"
                header=0      # usa la riga successiva come intestazione
            )

            # Rinomino le colonne secondo la convenzione richiesta
            col_map = {
                df_raw.columns[0]: 'codice_dpi',
                df_raw.columns[1]: 'descrizione',
                df_raw.columns[2]: 'marca',
                df_raw.columns[3]: 'modello',
                df_raw.columns[4]: 'SNR',
                df_raw.columns[5]: 'H',
                df_raw.columns[6]: 'L',
                df_raw.columns[7]: 'M',
                df_raw.columns[8]: 'beta',
            }
            df_dpi = df_raw.rename(columns=col_map)

            # Forzo i tipi di dato
            str_cols   = ['codice_dpi', 'descrizione', 'marca', 'modello']
            float_cols = ['SNR', 'H', 'L', 'M', 'beta']

            for col in str_cols:
                df_dpi[col] = df_dpi[col].astype(str).replace('nan', None)

            for col in float_cols:
                df_dpi[col] = pd.to_numeric(df_dpi[col], errors='coerce')

            # Messaggio di conferma con elenco marca-modello
            print('Lettura scheda DPI avvenuta con successo. Verifica la correttezza dei dati:')
            for _, row in df_dpi.iterrows():
                print(f'  [{row["codice_dpi"]}]  Marca: {row["marca"]}  -  Modello: {row["modello"]}')

            return df_dpi

        except Exception as e:
            warnings.warn(f'Lettura file excel non avvenuta con successo. Errore: {e}')
            return None



    def read_csv(file):
        '''
        Funzione che legge il file csv grezzo e lo trasforma in un dataFrame Pandas
        '''
        df = pd.read_csv(file) 

        return df


class manager:
    '''
    Classe che gestisce la lettura e la creazione dei file.
    '''
    
    def __init__(self):
        '''
        Salvataggio delle cartelle di misura
        '''

        print(''''
        #############################
            Analisi VR Rumore
        #############################
        ''')

        from os import getcwd, listdir
        self.main_dir = getcwd() #salva la directory principale di lavoro
        print(f'Directory di lavoro: {self.main_dir}') #stampa la directory di lavoro

        self.file_list = listdir(self.main_dir) #salva la lista dei file nella directory principale
        # print(self.file_list)
        if ".DS_Store" in self.file_list:
            self.file_list.remove(".DS_Store")
        
        if "data" in self.file_list:
            self.file_list.remove("data")
        
        if "VR_8h.csv" in self.file_list and "VR_8h.xlsx" in self.file_list:
            self.file_list.remove("VR_8h.csv")
            self.file_list.remove("VR_8h.xlsx")
        
        if "VR_rumore.out" in self.file_list:
            self.file_list.remove("VR_rumore.out")

        
        print(f'Lista dei file nella directory: {self.file_list}') # stampa la lista dei file nella directory principale


    def iterate_directory(self,file_name = 'LSOURCES.txt', format = 'txt'):
        '''
        Funzione che itera su tutte le directory per salvare i file.
        Crea i file csv e xlsx e colora le colonne opportune del file xlsx mediando tutti i dati.
        INPUT: 
            file_name = <str>, default: LSOURCES.txt. Nome dei file di input.
            format = <str>, formato file. Disponibili: 'txt' o 'csv'
        '''
        from os import path, mkdir
        def read_W_file_txt(file: str, letter_ID: str, decimals: int = 1) -> pd.DataFrame:
            """
            Legge un file in formato 'dati.txt' (output strumento WED con 6 sorgenti,
            separatore decimale virgola, etichette 'Fast A' / 'Fast C' / 'Picco C')
            e restituisce un DataFrame pandas compatibile con l'output di read_measure_file()
            nel formato txt classico (LSOURCES.txt).

            INPUT:
                file       : <str>  percorso del file da leggere
                letter_ID  : <str>  lettera identificativa della sessione (es. 'D', 'F', ...)
                decimals   : <int>  cifre decimali per l'arrotondamento (default 1)

            OUTPUT:
                df : pd.DataFrame con colonne:
                    fileID, letter_ID, nTrack,
                    LeqA_min, LeqA_max, LeqA_eq,
                    LeqC_min, LeqC_max, LeqC_eq,
                    PeakC_max, PeakC_eq,
                    durata, inizio, fine
            """

            # ── lettura file ─────────────────────────────────────────────────────────
            with open(file, 'r', encoding='utf-16') as f:
                lines = f.readlines()

            # ── inizializzazione colonne output ──────────────────────────────────────
            fileIDs    = []
            letter_IDs = []
            nTrack_col = []
            LeqA_min   = []
            LeqA_max   = []
            LeqA_eq    = []
            LeqC_min   = []
            LeqC_max   = []
            LeqC_eq    = []
            PeakC_max  = []
            PeakC_eq   = []
            durata_col = []
            inizio_col = []
            fine_col   = []

            # ── helper: converte stringa con virgola decimale → float ────────────────
            def to_float(s: str) -> float:
                return float(s.strip().replace(',', '.'))

            # ── helper: estrae i valori di una sorgente da una riga tabulare ─────────
            # Ogni sorgente occupa 4 campi: [Leq_eq, Lmin, Lmax, durata]
            # (per Picco C: [vuoto, Lmin, Lmax, durata] → Lmin=PeakC_min skip, Lmax=PeakC_max)
            # La riga inizia con l'etichetta di ubicazione, poi i campi per tutte le sorgenti.
            def parse_row(tokens: list, n_sorgenti: int):
                """
                Restituisce lista di dict con chiavi eq, min, max, dur per ogni sorgente.
                Per 'Picco C' il campo 'eq' sarà vuoto-stringa (lo gestiamo dopo).
                """
                result = []
                # tokens[0] = etichetta ubicazione, poi gruppi di 4 per ogni sorgente
                for i in range(n_sorgenti):
                    base = 1 + i * 4
                    result.append({
                        'eq':  tokens[base].strip(),
                        'min': tokens[base + 1].strip(),
                        'max': tokens[base + 2].strip(),
                        'dur': tokens[base + 3].strip(),
                    })
                return result

            # ── scansione linee ───────────────────────────────────────────────────────
            file_counter = 0  # contatore file nella sessione (→ letter_ID + numero)

            i = 0
            while i < len(lines):
                line = lines[i]

                if line.startswith('File\t'):
                    # ── nuova misurazione ─────────────────────────────────────────────
                    file_counter += 1
                    fileID = line.split('\t')[1].strip()

                    # Inizio e Fine (righe +1, +2)
                    inizio_str = lines[i + 1].split('\t')[1].strip()   # es. "24/09/2025 09:43:24"
                    fine_str   = lines[i + 2].split('\t')[1].strip()

                    # Numero di sorgenti dalla riga +3  →  "Sorgente\t1\t2\t3\t4\t5\t6"
                    sorgente_tokens = lines[i + 3].strip().split('\t')
                    n_sorgenti = len(sorgente_tokens) - 1  # esclude la parola 'Sorgente'

                    # Righe dati: +7 = Fast A, +8 = Fast C, +9 = Picco C
                    row_fastA  = lines[i + 7].strip().split('\t')
                    row_fastC  = lines[i + 8].strip().split('\t')
                    row_piccoC = lines[i + 9].strip().split('\t')

                    vals_A = parse_row(row_fastA,  n_sorgenti)
                    vals_C = parse_row(row_fastC,  n_sorgenti)
                    vals_P = parse_row(row_piccoC, n_sorgenti)

                    # ── una riga per ogni sorgente ────────────────────────────────────
                    for s in range(n_sorgenti):
                        fileIDs.append(fileID)
                        letter_IDs.append(letter_ID + str(file_counter))
                        nTrack_col.append(s + 1)

                        # LeqA
                        LeqA_eq .append(round(to_float(vals_A[s]['eq']),  decimals))
                        LeqA_min.append(round(to_float(vals_A[s]['min']), decimals))
                        LeqA_max.append(round(to_float(vals_A[s]['max']), decimals))

                        # LeqC
                        LeqC_eq .append(round(to_float(vals_C[s]['eq']),  decimals))
                        LeqC_min.append(round(to_float(vals_C[s]['min']), decimals))
                        LeqC_max.append(round(to_float(vals_C[s]['max']), decimals))

                        # Picco C  →  'eq' è vuoto nel file; usiamo min=PeakC_min (non salvato),
                        #             max=PeakC_max; PeakC_eq calcolato come media aritmetica
                        #             (per coerenza con il codice originale che usa average())
                        PeakC_max.append(round(to_float(vals_P[s]['max']), decimals))
                        # PeakC_eq: nel formato LSOURCES era calcolato come media dei picchi
                        # istantanei; qui abbiamo solo il valore complessivo → lo usiamo come eq
                        PeakC_eq .append(round(to_float(vals_P[s]['min']), decimals))  # 'min' = Lmin del picco

                        # durata, inizio, fine
                        durata_col.append(vals_A[s]['dur'])
                        inizio_col.append(inizio_str)
                        fine_col  .append(fine_str)

                    i += 10  # salta l'intero blocco (10 righe per blocco)
                    continue

                i += 1

            # ── costruzione DataFrame ─────────────────────────────────────────────────
            df = pd.DataFrame({
                'fileID'    : fileIDs,
                'letter_ID' : letter_IDs,
                'nTrack'    : nTrack_col,
                'LeqA_min'  : LeqA_min,
                'LeqA_max'  : LeqA_max,
                'LeqA_eq'   : LeqA_eq,
                'LeqC_min'  : LeqC_min,
                'LeqC_max'  : LeqC_max,
                'LeqC_eq'   : LeqC_eq,
                'PeakC_max' : PeakC_max,
                'PeakC_eq'  : PeakC_eq,
                'durata'    : durata_col,
                'inizio'    : inizio_col,
                'fine'      : fine_col,
            })

            print(f'Lettura completata: {file_counter} file, {len(df)} righe totali ({n_sorgenti} sorgenti per file)')
            return df


        #Verifica se esiste la cartella '/data' ed in caso la crea
        self.out_file_dir = self.main_dir + '/' + 'data'
        if path.isdir(self.out_file_dir) == False:
                    mkdir(self.out_file_dir)


        # Caso TXT ==============================
        if format=='txt':
            print('! TXT mode !')
            

            for dir in self.file_list:
                print(f'Iterazione sulla directory: {dir}') 
                chdir(self.main_dir + '/' + dir + '/' + dir) #entro nella cartella della misura i-esima
            
                if dir == "misW":
                    df = read_W_file_txt('misW.txt','W')
                else:
                    df, num_of_track, n_files = files.read_measure_file(file_name,list(dir)[-1], format = format)
                
                files.write_csv(df, f'{self.out_file_dir}/{dir}.csv')
                files.write_exel(df, f'{self.out_file_dir}/{dir}.xlsx')
                exel_file.adjust_column_lenght(f'{self.out_file_dir}/{dir}.xlsx', ['A'])
                exel_file.color_column(f'{self.out_file_dir}/{dir}.xlsx', ['F','I','K'], ['FFFF00','FFFF00','FFFF00'])
       
        # Caso CSV ==============================
        elif format=='csv':
            print('! CSV mode !')
            
            
            for dir in self.file_list:
                print(f'Iterazione sulla directory: {dir}')    
                chdir(self.main_dir + '/' + dir) # entro nella cartella dei dati (misE, misF, ecc...)
                
                if dir == "misW":
                    df = read_W_file_txt('misW.txt','W')
                else:
                    df = files.read_measure_file(file_name,letter_ID=list(dir)[-1], format='csv') # salvo il DF totale delle misure  (D,E,F,ecc)
                
                #salvo i file nella cartella opportuna e nelle versioni csv ed exel
                files.write_csv(df, f'{self.out_file_dir}/{dir}.csv')
                files.write_exel(df, f'{self.out_file_dir}/{dir}.xlsx')
                exel_file.adjust_column_lenght(f'{self.out_file_dir}/{dir}.xlsx', ['A'])
                exel_file.color_column(f'{self.out_file_dir}/{dir}.xlsx', ['F','I','J'], ['FFFF00','FFFF00','FFFF00'])






        
        
        else:
            error("Disponibili solo i formati 'csv' o 'txt'. Controlla di aver inserito il testo correttamente.")


    
    def VR8h_exel(self, name_averaged_data, out_VR8h_name):
        '''
        Funzione che gestisce il file finale della valutazione del rischio in 8h dei lavori. 
        Importa il file xlsx e ne colora le colonne del colore specifico in base al grado di rischio.
        '''
        a = analisi(self.main_dir)
        a.VR_8h(name_averaged_data)
        exel_file.color_cell_VR8h(out_VR8h_name)

    


class analisi:
    '''
    Classe con i metodi per l'analisi delle misure
    '''

    def __init__(self, csv_data_directory):
        '''
        INPUT:
            df = <pd.DataFrame>, dataframe con i dati della misura
        '''
        self.main_dir = csv_data_directory #salvataggio della directory principale
        


    # prende i valori raw e crea un 
    def average_values(self):
        '''
        Funzione che concatena i dataframe di tutte le misure e scrive un file completo con le medie
        e le std sulla media (incertezze)
        '''
        import glob 
        from os.path import exists

        if not exists(self.main_dir + '/averaged_data.csv'):
            

            files_csv = glob.glob(self.main_dir + '/*.csv') #leggo solo i file con estensione csv
            # print(files)

            df_list = [] # lista dei dataframe che ci sono
            
            #inizializzo il dataFrame pandas con i valori medi di tutte le misure
            df_avg = pd.DataFrame(columns=['jobName', 'ID', 'U' ,'LeqA','LeqC','Ppeak'])


            # salvo in un dataframe tutti i file csv
            for file in files_csv:
                # print(file)
                df_list.append(pd.read_csv(file))
            
            # iterazione su tutti i dataframe
            for df in df_list:
                
                fileIDs = df[df.columns[1]].unique() #lista dei fileID
                letter_IDs = df[df.columns[2]].unique() #lista dei letterID
                
                LeqA_mean = zeros(len(fileIDs)) # inizializzo l'array di LeqA_mean
                LeqC_mean = zeros(len(fileIDs)) # inizializzo l'array di LeqC_mean
                Ppeak_mean = zeros(len(fileIDs)) # inizializzo l'array di LeqC_mean
                U_sdom = zeros(len(fileIDs)) #standard deviation of mean (incertezza misure)
                

                for i in range(len(fileIDs)):
                    idx = df[df.columns[1]] == fileIDs[i] # prendo solo i valori opportuni
                    
                    #calcolo i valori medi
                    LeqA_mean[i] = round(mean(df['LeqA_eq'][idx]),1)
                    LeqC_mean[i] = round(mean(df['LeqC_eq'][idx]),1) # + std(df['LeqC_max'][idx],ddof=1) (in caso da aggiungere se serve)
                    Ppeak_mean[i] = round(max(df['PeakC_max'][idx]),1) # + 1.56 in caso da aggiungere 
                    # print(LeqA_mean)
                    
                    #calcolo l'incertezza sulla misura LeqA (SDOM)
                    U_sdom[i] = round(std(df['LeqA_max'][idx], ddof=1) * sqrt(1/sum(idx)),1)

                
                new_df = pd.DataFrame({'jobName': fileIDs ,
                                       'ID':letter_IDs ,
                                       'U':U_sdom,
                                       'LeqA' : LeqA_mean, 
                                       'LeqC' : LeqC_mean ,
                                       'Ppeak' : Ppeak_mean})
                df_avg = pd.concat([df_avg,new_df], ignore_index=True)
            
            
            df_avg['Ti'] = [[]] * len(df_avg) #creo la colonna con i valori di exposure time
            df_avg['GrOm'] = [[]] * len(df_avg) #creo la colonna con gli ID del gruppo omogeneo
            df_avg['DPI'] = [[]] * len(df_avg) #creo la colonna contenente il riferimento all'uso del DPI


            files.write_csv(df_avg, self.main_dir + '/averaged_data.csv')
            files.write_exel(df_avg, self.main_dir + '/averaged_data.xlsx')
            exel_file.adjust_column_lenght(self.main_dir + '/averaged_data.xlsx', 'A')
            print('Averaged data files created')
            return df_avg
        else:
            print('File averaged.csv already exists!')
            df_avg = pd.read_csv(self.main_dir + '/averaged_data.csv')
            return df_avg

    
    def get_scheda_info(self, df_avg, excel_info_dir=None, name_exel_info="scheda_gruppi_dpi.xlsx"):
        '''
        Funzione che legge il file excel "scheda_gruppi_dpi.xlsx" e popola le colonne
        GrOm e Ti del dataframe df_avg con i valori dei gruppi omogenei e dei tempi
        di esposizione per ciascuna misura.

        La funzione va chiamata DOPO average_values() e PRIMA di VR_8h().
        Sostituisce il passaggio manuale di inserimento di GrOm e Ti in averaged_data.csv.

        INPUT:
            df_avg     = <pd.DataFrame>, dataframe con i valori medi delle misure,
                         output di average_values(). Deve contenere almeno la colonna 'ID'
                         (es. 'F4', 'D3', 'E1', ...) e le colonne 'GrOm' e 'Ti'.
            excel_info = <str>, percorso completo del file excel con la scheda dei gruppi
                         omogenei. Default: main_directory + "/scheda_gruppi_dpi.xlsx"
                         (una directory sopra self.main_dir, che corrisponde a data_folder).

        OUTPUT:
            df_avg = <pd.DataFrame>, stesso dataframe in input con le colonne aggiornate:
                     - 'GrOm': lista di stringhe con gli ID_GrOm di tutti i gruppi
                                omogenei in cui compare la misura (es. ['M_01', 'M_03'])
                     - 'Ti'  : lista di interi con i tempi di esposizione [min] per
                                ciascun gruppo omogeneo (es. [90, 120])
                     Il dataframe aggiornato viene salvato sovrascrivendo
                     averaged_data.csv e averaged_data.xlsx nella cartella self.main_dir.

        RAISES:
            ValueError: se una mansione presente nel foglio 'Gruppi_omogenei' non ha
                        corrispondenza nel foglio 'Scheda_mansioni'.
            ValueError: se un ID presente in df_avg non compare in nessun gruppo
                        omogeneo nel foglio 'Gruppi_omogenei'.
            FileNotFoundError: se il file excel_info non esiste al percorso specificato.
        '''
        import openpyxl
        from os.path import dirname, exists

        # ------------------------------------------------------------------
        # Variabili interne (modificabili facilmente)
        # ------------------------------------------------------------------
        sheet_gruppi   = 'Gruppi_omogenei'   # nome foglio gruppi omogenei
        sheet_mansioni = 'Scheda_mansioni'    # nome foglio scheda mansioni

        # Percorsi di output (sovrascrittura dei file averaged_data)
        out_csv  = self.main_dir + '/averaged_data.csv'
        out_xlsx = self.main_dir + '/averaged_data.xlsx'

        # Percorso default del file excel: una directory sopra self.main_dir
        # (corrisponde a main_directory, dato che self.main_dir = main_directory/data)
        if excel_info_dir is None:
            excel_info_dir = dirname(self.main_dir) + ''
        else:
            excel_info_dir = excel_info_dir + '/' + name_exel_info  # usa il percorso specificato in input

        # ------------------------------------------------------------------
        # Verifica esistenza file excel
        # ------------------------------------------------------------------
        if not exists(excel_info_dir):
            raise FileNotFoundError(
                f"File excel non trovato: '{excel_info_dir}'\n"
                f"Controlla che il file {name_exel_info} sia presente in: "
                f"'{dirname(excel_info_dir)}'"
            )

        print(f'Lettura file scheda gruppi: {excel_info_dir}')
        wb = openpyxl.load_workbook(excel_info_dir, data_only=True)

        # ==================================================================
        # STEP 1 — Lettura foglio 'Scheda_mansioni'
        # Costruisce dizionario lookup: {Descrizione_GrOm -> ID_GrOm}
        # Header in riga 2, dati da riga 3 in avanti.
        # ==================================================================
        ws_mansioni = wb[sheet_mansioni]

        lookup_mansioni = {}  # { 'Addetto lavaggio': 'M_01', ... }
        for row in ws_mansioni.iter_rows(min_row=3, values_only=True):
            id_grom    = row[0]   # colonna A: ID_GrOm   (es. 'M_01')
            descrizione = row[1]  # colonna B: Descrizione_GrOm (es. 'Addetto lavaggio')
            if id_grom is not None and descrizione is not None:
                lookup_mansioni[str(descrizione).strip()] = str(id_grom).strip()

        print(f'Lookup mansioni caricato ({len(lookup_mansioni)} voci): {lookup_mansioni}')

        # ==================================================================
        # STEP 2 — Parsing foglio 'Gruppi_omogenei'
        # Costruisce df_grom con colonne: [ID_misura, ID_GrOm, Descrizione, Ti]
        #
        # Struttura del foglio:
        #   Riga 1-2 : titolo/vuote
        #   Riga 3   : nomi mansioni nelle colonne dispari (A, C, E, ...)
        #              es. ('Addetto lavaggio', None, 'Adetto sollevamento', None, ...)
        #   Riga 4   : header 'ID', 'Ti', 'ID', 'Ti', ... (ripetuto per ogni gruppo)
        #   Righe 5+ : dati (ID_misura, Ti) per ogni gruppo
        # ==================================================================
        ws_gruppi = wb[sheet_gruppi]
        all_rows  = list(ws_gruppi.iter_rows(values_only=True))

        row_mansioni = all_rows[2]  # riga 3 (0-indexed -> indice 2)
        data_rows    = all_rows[4:] # righe dati da riga 5 in avanti (indice 4+)

        n_cols  = len(row_mansioni)
        n_groups = n_cols // 2  # ogni gruppo occupa 2 colonne: ID e Ti

        df_grom_rows = []

        for g in range(n_groups):
            col_id = g * 2        # indice colonna ID
            col_ti = g * 2 + 1   # indice colonna Ti

            mansione_nome = row_mansioni[col_id]
            if mansione_nome is None:
                continue  # colonna vuota: nessun gruppo definito, si passa al prossimo

            mansione_nome = str(mansione_nome).strip()

            # Ricerca ID_GrOm nel dizionario di lookup
            if mansione_nome not in lookup_mansioni:
                raise ValueError(
                    f"Mansione '{mansione_nome}' (foglio '{sheet_gruppi}', gruppo {g+1}) "
                    f"non trovata nel foglio '{sheet_mansioni}'.\n"
                    f"Verifica che il nome sia identico nella colonna 'Descrizione_GrOm'.\n"
                    f"Voci disponibili: {list(lookup_mansioni.keys())}"
                )

            id_grom = lookup_mansioni[mansione_nome]

            # Lettura righe dati per questo gruppo
            for data_row in data_rows:
                id_misura = data_row[col_id]
                ti_val    = data_row[col_ti]

                if id_misura is None:
                    continue  # riga vuota per questo gruppo

                df_grom_rows.append({
                    'ID_misura':   str(id_misura).strip(),
                    'ID_GrOm':     id_grom,
                    'Descrizione': mansione_nome,
                    'Ti':          int(ti_val)
                })

        df_grom = pd.DataFrame(df_grom_rows, columns=['ID_misura', 'ID_GrOm', 'Descrizione', 'Ti'])
        print(f'df_grom costruito con {len(df_grom)} righe:')
        print(df_grom.to_string(index=False))

        # ==================================================================
        # STEP 3 — Popolamento colonne GrOm e Ti in df_avg
        #
        # NOTA: le colonne GrOm e Ti in df_avg sono inizializzate da average_values()
        # come interi (dtype int64). Prima di assegnare liste occorre forzare
        # il tipo a object, altrimenti pandas lancia TypeError.
        # ==================================================================
        df_avg['GrOm'] = df_avg['GrOm'].astype(object)
        df_avg['Ti']   = df_avg['Ti'].astype(object)
        #
        # Per ogni riga di df_avg (identificata da colonna 'ID'):
        #   - cerca tutte le occorrenze di quell'ID in df_grom
        #   - scrive la lista degli ID_GrOm in df_avg['GrOm']
        #   - scrive la lista dei Ti in df_avg['Ti']
        # Se un ID non è trovato in nessun gruppo -> Warning

        missing_ids = []

        for idx, row in df_avg.iterrows():
            id_misura = str(row['ID']).strip()

            # Filtra df_grom per trovare tutti i gruppi che contengono questa misura
            matches = df_grom[df_grom['ID_misura'] == id_misura]

            if matches.empty:
                missing_ids.append(id_misura)
                continue  # lascia GrOm e Ti invariati per questa riga, prosegue


            # Scrivi le liste nelle colonne GrOm e Ti
            df_avg.at[idx, 'GrOm'] = matches['ID_GrOm'].tolist()
            df_avg.at[idx, 'Ti']   = matches['Ti'].tolist()

        print('\nPopolamento GrOm e Ti completato.')
        print(df_avg[['ID', 'GrOm', 'Ti']].to_string(index=False))

            # Warning finale non bloccante per le misure non trovate
        if missing_ids:
            print(
                f'\n*** WARNING: Le seguenti misure di df_avg non sono state trovate '
                f'in nessun gruppo omogeneo nel file:\n'
                f'    {excel_info_dir}\n'
                f'    Misure mancanti: {missing_ids}\n'
                f'    Le colonne GrOm e Ti di queste righe sono rimaste invariate.\n'
                f'    Verifica che gli ID siano presenti nel foglio "{sheet_gruppi}".\n'
                f'    ID disponibili in df_grom: {sorted(df_grom["ID_misura"].unique().tolist())}\n'
                f'***'
            )

        # ==================================================================
        # STEP 4 — Salvataggio (sovrascrittura averaged_data.csv e .xlsx)
        # ==================================================================
        files.write_csv(df_avg, out_csv)
        files.write_exel(df_avg, out_xlsx)
        print(f'\nFile aggiornati salvati:\n  CSV  -> {out_csv}\n  XLSX -> {out_xlsx}')

        return df_avg

    #DEPRECATED, sostituita da analisi_8h()
    def calcolo_Leq8h(self, df_GrOm, T0 = 480):
        '''
        funzione che calcola il livello equivalente nelle 8h
        INPUT:
            df_GrOm = <dataFrame>, contenente i valori medi delle misure di uno specifico gruppo omogeneo
            T0 = <int>, numero di minuti di esposizione di una giornata lavorativa
        OUTPUT:
            df_avg = <dataFrame> specifico della mansione
        
        
        Theory:
        livello sonoro equivalente 8h
        Lex,8h = 10*log( 1/T0 * sum( T_i * 10^(0.1 * L_i)  )    (dBA)
            
            La i indica la sorgente sonora i_esima
            T0 è il tempo totale di lavoro in ore (8 h lavorative in genere)
            T_i è il tempo di esposizione quotidiana, in ore, di un lavoratore alla fonte i-esima
            L_i è il livello equivalente continuo ponderato A della fonte i-esima
        '''

        # Calcolo il Leq_8h
        self.LeqA_8h = 10 * log10( 1/T0 * dot( df_GrOm['Ti'], 10**(0.1 * df_GrOm['LeqA']) ) )
        return self.LeqA_8h
    

    def calcolo_U_estesa(self, df_GrOm, T0 = 480, u2m = 0.7, u_pos = 1):
        '''
        Funzione che calcola l'incertezza combinata standard e quella estesa per il LeqA_8h
        INPUT:
            df_GrOm = <dataFrame>, contenente i valori medi delle misure di uno specifico gruppo omogeneo
            T0 = <int>, minuti di esposizione di una giornata lavorativa, in genere 8 ore ossia 480 minuti
            u2m = <float>, errore secondo normativa in base allo strumento (0.7 o 1.5)
            u_pos = <float>, errore nel posizionamento dello strumento (in metri)
        OUTPUT:
            U_estesa = <float>, incertezza estesa
            U_comb_std = <float>, iincertezza combinata standard

        THEORY:

        '''
        LeqA8H = self.calcolo_Leq8h(df_GrOm, T0=T0)
        self.U_comb_std = sum(  (df_GrOm['Ti']/T0  * 10**(0.1*( df_GrOm['LeqA'] - LeqA8H ) ) )**2   * ( df_GrOm['U']**2 + u2m**2 + u_pos **2 ) +
                    ( 4.34 * (1/T0  * 10**(0.1*( df_GrOm['LeqA'] - LeqA8H ) )) * (std(df_GrOm['Ti'], ddof=1) * sqrt(1/len(df_GrOm)) ) )**2 )
        
        self.U_ext = self.U_comb_std * 1.65

        return self.U_ext, self.U_comb_std


    def analisi_8h(self, output_dir, df_avg, T0=480, u2m=0.7, u_pos=1.0):
        '''
        Funzione che calcola la valutazione del rischio rumore su base giornaliera (8h)
        per ogni gruppo omogeneo presente in df_avg.

        Per ogni gruppo omogeneo vengono calcolati:
            - Lex8h    : livello sonoro equivalente ponderato A su 8h  [cella I44 / G44 del foglio Excel]
            - U        : incertezza estesa                              [cella K44 / G46 del foglio Excel]
            - Lex_max  : Lex8h + U                                     [cella O44 del foglio Excel]
            - L_picco_C: massimo dei Ppeak nel gruppo omogeneo         [cella O45 del foglio Excel]

        TEORIA (D.Lgs. 81/08, norma ISO 9612):

            AG_i    = Ti/T0 * 10^(LeqA_i / 10)              [col. AG foglio Excel]
            Lex8h   = 10 * log10( SUM(AG_i) )                [G44 = I44]

            Z_i     = Ti/T0 * 10^((LeqA_i - Lex8h) / 10)   [col. Z foglio Excel]
            W_i     = max(0, Z_i^2 * (u_i^2 + u2m^2 + u_pos^2))  [col. W foglio Excel]
                      (il II termine X_i, legato alla variabilita` di Tm, e` posto = 0)
            U_comb  = SUM(W_i)                               [G45]
            U       = 1.65 * sqrt(U_comb)                    [G46 = K44]

            Lex_max   = Lex8h + U                            [O44]
            L_picco_C = max(Ppeak nel gruppo)                [O45]

        INPUT:
            output_dir = <str>, directory in cui salvare i file di output
            df_avg     = <pd.DataFrame>, dataframe con le medie delle misure.
                         Le colonne Ti e GrOm contengono liste (una voce per ogni
                         gruppo omogeneo di appartenenza della misura).
                         Esempio riga: ID='F1', Ti=[230,90,180], GrOm=['M1','M2','M3']
            T0         = <int>, tempo di riferimento in minuti (default 480 = 8h)
            u2m        = <float>, incertezza strumentale U2 (default 0.7 dB)
            u_pos      = <float>, incertezza di posizione U3 (default 1.0 dB)

        OUTPUT (file scritti in output_dir):
            VR8h.csv  e  VR8h.xlsx  : file unico riepilogativo con i risultati di tutti
                                       i gruppi omogenei (GrOm, Lex8h, U, Lex_max, L_picco_C)
            VR8h_<GrOm>.csv  e  VR8h_<GrOm>.xlsx  : un file per ogni gruppo omogeneo con
                                       il dettaglio delle misure (ID, LeqA, LeqC, U, Ti)
        '''
        import ast
        import os

        # ----------------------------------------------------------------
        # STEP 1 — Parsing di Ti e GrOm: lista Python o stringa da CSV
        #          Esplosione del dataframe: una riga per ogni (ID, GrOm, Ti)
        # ----------------------------------------------------------------
        rows = []
        for _, row in df_avg.iterrows():
            ti_list   = row['Ti']   if isinstance(row['Ti'],   list) else ast.literal_eval(row['Ti'])
            grom_list = row['GrOm'] if isinstance(row['GrOm'], list) else ast.literal_eval(row['GrOm'])
            for ti, grom in zip(ti_list, grom_list):
                rows.append({**row.to_dict(), 'Ti': ti, 'GrOm': grom})
        df_exp = pd.DataFrame(rows)

        # ----------------------------------------------------------------
        # STEP 2 — Ciclo su ogni gruppo omogeneo: calcoli e raccolta risultati
        # ----------------------------------------------------------------
        summary_rows = []  # raccoglie una riga di riepilogo per ogni gruppo

        for grp_name, grp in df_exp.groupby('GrOm'):

            # STEP 3 — Verifica che la somma dei Ti sia esattamente T0
            tot_ti = grp['Ti'].sum()
            if tot_ti != T0:
                raise ValueError(
                    f"Gruppo omogeneo '{grp_name}': somma dei Ti = {tot_ti} min "
                    f"!= T0 = {T0} min. Controlla i valori di Ti nel file averaged_data."
                )

            leqa  = grp['LeqA'].values
            ti    = grp['Ti'].values
            u_mis = grp['U'].values

            # STEP 4 — Lex8h  (cella G44 = I44 del foglio Excel)
            #          AG_i = Ti/T0 * 10^(LeqA_i / 10)
            lex8h = 10 * log10(sum(ti / T0 * 10**(leqa / 10)))

            # STEP 5 — Incertezza estesa U  (cella G46 = K44 del foglio Excel)
            #          Z_i = Ti/T0 * 10^((LeqA_i - Lex8h) / 10)   [col. Z]
            #          W_i = max(0, Z_i^2 * (u_i^2 + u2m^2 + u_pos^2))  [col. W]
            #          II termine X_i = 0  (Tmax/Tmin non disponibili in df_avg)
            z     = ti / T0 * 10**((leqa - lex8h) / 10)
            w     = max(z**2 * (u_mis**2 + u2m**2 + u_pos**2), 0)
            U_val = 1.65 * sqrt(sum(w))

            # STEP 6 — Lex_max (O44) e L_picco_C (O45)
            lex_max   = lex8h + U_val
            l_picco_c = max(grp['Ppeak'].values)

            # STEP 7 — Accumulo riga di riepilogo
            summary_rows.append({
                'GrOm':      grp_name,
                'Lex8h':     round(lex8h,     1),
                'U':         round(U_val,     1),
                'Lex_max':   round(lex_max,   1),
                'L_picco_C': round(l_picco_c, 1),
            })

            # STEP 8 — File di dettaglio per il gruppo: ID, LeqA, LeqC, U, Ti
            df_detail = grp[['ID', 'LeqA', 'LeqC', 'U', 'Ti']].reset_index(drop=True)
            detail_csv  = os.path.join(output_dir, f'VR8h_{grp_name}.csv')
            detail_xlsx = os.path.join(output_dir, f'VR8h_{grp_name}.xlsx')

            # Check sull'esistenza della directory
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                print(f'-- {output_dir} created --')

            df_detail.to_csv(detail_csv,   index=True)
            df_detail.to_excel(detail_xlsx, index=True)

            print('\n\n')
            print(f'Gruppo {grp_name}: \nLex8h={round(lex8h,1)} dB(A)\n'
                  f'U={round(U_val,1)} dB \nLex_max={round(lex_max,1)} dB(A)\n'
                  f'L_picco_C={round(l_picco_c,1)} dB(C)')


        # ----------------------------------------------------------------
        # STEP 9 — File riepilogativo unico VR8h.csv e VR8h.xlsx
        # ----------------------------------------------------------------
        df_summary = pd.DataFrame(summary_rows)
        df_summary.to_csv( os.path.join(output_dir, 'VR8h_totale.csv'),  index=True)
        df_summary.to_excel(os.path.join(output_dir, 'VR8h_totale.xlsx'), index=True)

        print(f'\n###\nanalisi_8h completata.\nDati salvati in {output_dir}\n###')


    def VR_8h(self,df_avg_dir):
            '''
            Funzione che analizza i dati avg per restituire i valori nelle 8h.
            NB: è importante aver modificato il file avergaed_data.xlsx in modo da aver inserito
            il numero identificativo del gruppo omogeneo e il tempo di esposizione della mansione
            INPUT:
                df_avg_dir = <str>, directory del file con la media dei valori
            OUTPUT:
                df_VR8h = <dataFrame>, con i valori di valutazione del rischio in 8h
            '''
            df_avg = files.read_csv(df_avg_dir) #lettura del df in csv

            #creo le variabili vuote del dataframe da inserire 
            LeqA8h = []
            Uext = []
            U_comb_std = []
            Peak_max = []
            LeqA_max = []

            #prendo solo i valori unici degli ID del gruppo omogeneo
            grorIDs = df_avg['GrOm'].unique() # ; print(grorIDs)

            #inizio il ciclo su tutti i valori di ID del gruppo omogeneo
            for i in range(len(grorIDs)):
                print('Calcolo su gruppo omogenero:', grorIDs[i]) 

                idx = df_avg['GrOm'] == grorIDs[i] # prendo solo i valori opportuni di GrOm e gl indici
                
                #calcolo il LeqA su 8 ore
                ax1 = self.calcolo_Leq8h(df_GrOm=df_avg[idx]) 
                LeqA8h.append(ax1)

                #calcolo incertezza estesa
                ax2, ax22 = self.calcolo_U_estesa(df_GrOm=df_avg[idx])
                Uext.append(ax2), U_comb_std.append(ax22)
                
                #Calcolo massimo del picco
                ax3 = max(df_avg['Ppeak'][idx])
                Peak_max.append(ax3)

                #Calcolo il massimo di LeqA
                LeqA_max.append(ax1 + ax2)

                # input('--- pausa ---')
            
            # df_VR8h = pd.DataFrame(columns=['GrOm','LeqA', 'U' , 'Ppeak'])
            df_VR8h = pd.DataFrame({'GrOm': grorIDs,
                                    'LeqA_8h': round(LeqA8h,1),
                                    'Peak': round(Peak_max,1),
                                    'U_ext': round(Uext,1),
                                    'Leq_max': round(LeqA_max,1)})

            files.write_exel(df_VR8h, self.main_dir+ '/VR_8h.xlsx')
            files.write_csv(df_VR8h, self.main_dir+ '/VR_8h.csv')
            print(df_VR8h)
            print('Vautazione rischio 8h completata.')
    

    def DPI_HML(self, VR_8h_dir, H, M, L, beta, grom):
        '''
        Funzione che utilizza il metodo HML per il calcolo dei coefficienti di riduzione dei DPI.
        INPUT:
           VR_8h_dir = <str>, directory del file VR_8h.csv
           H = <float>, valore di attenuazione delle alte frequenze
           M = <float>, valore di attenuazione delle medie freq
           L = <float>, valore di attenuazione delle basse freq
           beta = <foat>, coefficiente correttivo dei DPI
           gror = <list>, lista dei gruppi omogenei su cui applicare questi dpi
        OUTPUT:
            df = <pd.DataFrame>, contenente i valori di LeqA_8h con la riduzione e il valore di PNR

        
        Theory:
            1. Calcolo dL = LeqC - LeqA

            2. Calcolo del PNR

                Se dL <= 2 dB:
                    PNR = M - (H - M)/4 * (dL - 2)
                Se dL > 2:
                    PNR = M - (H - M)/8 * (dL - 2)

            3. Calcolo del livello effettivo all'orecchio
                LeqA_eff = LeqA - PNR

            4. Confronta idonietà del DPI

        '''
        # Aggiunta del coefficiente correttivo 
        H = beta * H
        M = beta * M
        L = beta * L


        df_avg = files.read_csv(self.main_dir + '/averaged_data.csv') #lettura del file csv con le medie dei dati
        df_8h = files.read_csv(VR_8h_dir + '/VR_8h.csv') #lettura del file csv dei dati valutazione rischio 8h

        # grorIDs = df_avg['GrOm'].unique() #scelgo solo i valori unici dei gruppi omogenei
        
        # Inizializzo le nuove colonne del df_avg
        df_avg['PNR'] = 0.0
        df_avg['LeqA_eff'] = 0.0
        #Inizializzo le nuove colonne del df_8h
        df_8h['PNR_avg'] = 0.0
        df_8h['LeqA_8h_eff'] = 0.0

        for i in range(len(grom)): #itero sul numero di gruppi omogenei selezionati
            
            print(f'Calcolo HML dei PDI su gruppo omogeneo {grom[i]}') 

            idx_grom = df_avg['GrOm'] == grom[i] # prendo solo i valori opportuni di GrOm e gli indici
            idx_dpi = idx_grom * df_avg['DPI'] # prendo solo i valori degli indici del gruppo omogeneo in cui compaiono i DPI
            
            dL = df_avg['LeqC'][idx_dpi] - df_avg['LeqA'][idx_dpi] #calcolo il valore dL
            
            # Calcolo del PNR in base al valore di dL con la condizione
            df_avg.loc[idx_dpi, 'PNR'] = round(dL.apply(
                lambda x: M - (H - M) / 4 * (x - 2) if x <= 2 else M - (H - M) / 8 * (x - 2)
            ),1)

            df_8h.loc[df_8h['GrOm']==grom[i], 'PNR_avg'] = mean(df_avg['PNR'][idx_dpi])
            # df_8h.loc[df_8h['GrOm'] == i, 'PNR_avg'] = mean(df_avg['PNR'][idx_dpi][df_avg['PNR'][idx_dpi] != 0.0])


        # AVG
        df_avg['LeqA_eff'] = df_avg['LeqA'] - df_avg['PNR'] # calcolo la colonna con la riduzione di LeqA
        files.write_csv(df_avg,self.main_dir + '/averaged_data.csv') #salvo il file avg in csv
        df_avg['DPI'] = df_avg['DPI'].map({True: 'Si', False: 'No'}) #mappo i valori booleani in si e no per exel
        files.write_exel(df_avg,self.main_dir + '/averaged_data.xlsx') #salvo il file avg in xlsx

        # VR 8h
        df_8h['LeqA_8h_eff'] = round(df_8h['LeqA_8h'] - df_8h['PNR_avg'],1)
        files.write_csv(df_8h,VR_8h_dir + '/VR_8h.csv')
        files.write_exel(df_8h,VR_8h_dir + '/VR_8h.xlsx')
        exel_file.color_cell_VR8h(VR_8h_dir + '/VR_8h.xlsx')


    def applica_DPI_HML(self, excel_info_dir, name_excel_info, output_directory, output_dpi, mode = 'both'):
        '''
        Funzione che applica il metodo HML per il calcolo dell'attenuazione dei DPI
        su tutti i file CSV presenti in output_directory e salva i risultati per ogni DPI
        in sottocartelle dedicate dentro output_dpi.

        Per ogni DPI definito nella scheda e per ogni file CSV del gruppo omogeneo:
            1. Legge i parametri H, L, M, beta e calcola H' = beta*H, L' = beta*L, M' = beta*M
            2. Calcola diff_C_A = LeqC - LeqA riga per riga
            3. Calcola PNR con metodo HML:
                - se diff_C_A <= 2 : PNR = M' - (H'-M')/4  * (diff_C_A - 2)
                - se diff_C_A >  2 : PNR = M' - (H'-L')/8  * (diff_C_A - 2)
            4. Aggiunge colonne PNR e LeqA_rid = LeqA - PNR
            5. Salva il file (csv + xlsx) nella cartella output_dpi/DPI_i/
            6. Colora le celle di LeqA_rid nel file xlsx in base al livello di rischio:
                < 65            : arancione  (#FF8C00)
                65 <= v <= 70   : giallo sc. (#FFD700)
                75 <= v <= 80   : giallo sc. (#FFD700)
                70 <  v < 75   : verde      (#008000)
                > 80            : rosso      (#DC143C)

        INPUT:
            excel_info_dir   = <str>, directory del file excel con la scheda DPI
            name_excel_info  = <str>, nome del file excel con la scheda DPI
            output_directory = <str>, directory contenente i file CSV dei gruppi omogenei
            output_dpi       = <str>, directory radice in cui creare le sottocartelle DPI_i
            mode             = <str>, xlsx per salvare in excel, csv per salvare in csv o both per salvare entrmabi

        OUTPUT:
            Nessun valore di ritorno. I file vengono salvati su disco.
        '''
        import glob
        import os
        import re
        import numpy as np
        import openpyxl as ex
        from openpyxl.styles import PatternFill
        from openpyxl.utils import get_column_letter

        # ── Step 1: leggi la scheda DPI ──────────────────────────────────────────
        df_dpi = files.get_scheda_DPI(excel_info_dir, name_excel_info)
        if df_dpi is None:
            print('Errore: impossibile leggere la scheda DPI. Controlla che esista nella directory. Funzione interrotta.')
            return

        # ── Step 2: lista CSV da processare ──────────────────────────────────────
        ESCLUSI = {'VR8h.csv', 'VR8h.xlsx'}
        csv_files = sorted([
            f for f in glob.glob(os.path.join(output_directory, '*.csv'))
            if os.path.basename(f) not in ESCLUSI
        ])

        if not csv_files:
            print(f'Nessun file CSV trovato in: {output_directory}')
            return

        # ── Step 3 + 4: itera su DPI e su file CSV ───────────────────────────────
        for _, dpi_row in df_dpi.iterrows():

            # Estrae il numero identificativo dal codice_dpi (es. "DPI1" -> 1)
            match = re.search(r'\d+', str(dpi_row['codice_dpi']))
            if not match:
                print(f'Codice DPI non riconosciuto: {dpi_row["codice_dpi"]} — riga saltata.')
                continue
            dpi_idx = match.group()

            # Parametri HML con coefficiente correttivo beta
            beta = float(dpi_row['beta'])
            Hp   = beta * float(dpi_row['H'])
            Mp   = beta * float(dpi_row['M'])
            Lp   = beta * float(dpi_row['L'])

            # Crea la cartella di output per questo DPI
            dpi_folder = os.path.join(output_dpi, f'DPI_{dpi_idx}')
            os.makedirs(dpi_folder, exist_ok=True)

            print(f'\n── DPI_{dpi_idx} | {dpi_row["marca"]} {dpi_row["modello"]} '
                  f'| β={beta}  H\'={Hp:.2f}  M\'={Mp:.2f}  L\'={Lp:.2f} ──')

            for csv_path in csv_files:
                df = pd.read_csv(csv_path)

                # ── Calcolo PNR vettorializzato ───────────────────────────────
                diff_C_A = df['LeqC'] - df['LeqA']

                PNR = np.where(
                    diff_C_A <= 2,
                    Mp - (Hp - Mp) / 4 * (diff_C_A - 2),   # diff <= 2
                    Mp - (Hp - Lp) / 8 * (diff_C_A - 2)    # diff >  2
                )

                df['PNR']      = PNR.round(1)
                df['LeqA_rid'] = (df['LeqA'] - df['PNR']).round(1)

                # ── Nomi file di output ───────────────────────────────────────
                base_name = os.path.basename(csv_path)           # es. VR8h_M_01.csv
                stem, _   = base_name.rsplit('.', 1)             # es. VR8h_M_01
                out_stem  = f'{stem}_dpi_{dpi_idx}'              # es. VR8h_M_01_dpi_1
                out_csv   = os.path.join(dpi_folder, out_stem + '.csv')
                out_xlsx  = os.path.join(dpi_folder, out_stem + '.xlsx')

                # ── Salvataggio CSV ───────────────────────────────────────────
                if mode == 'both' or mode == 'csv':
                    files.write_csv(df, out_csv)

                # ── Salvataggio XLSX ──────────────────────────────────────────
                if mode == 'both' or mode == 'xlsx':
                    files.write_exel(df, out_xlsx)

                # ── Colorazione celle LeqA_rid ────────────────────────────────
                wb = ex.load_workbook(out_xlsx)
                ws = wb.active

                # Trova la lettera della colonna LeqA_rid
                leqa_rid_col = None
                for cell in ws[1]:
                    if cell.value == 'LeqA_rid':
                        leqa_rid_col = get_column_letter(cell.column)
                        break

                if leqa_rid_col is None:
                    print(f'  ATTENZIONE: colonna LeqA_rid non trovata in {out_xlsx}')
                    wb.save(out_xlsx)
                    continue

                # Mappa colori per livello di rischio
                def _get_color(val):
                    if val < 65:
                        return 'FF8C00'   # arancione  – iperprotezione
                    elif val > 80:
                        return 'DC143C'   # rosso      – insufficiente
                    elif 65 <= val <= 70 or 75 <= val <= 80:
                        return 'FFD700'   # giallo sc. – accettabile
                    elif 70 < val < 75:
                        return '008000'   # verde      – buona protezione
                    return None           # nessuna colorazione (non dovrebbe accadere)

                for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
                    for cell in row:
                        if get_column_letter(cell.column) == leqa_rid_col \
                                and cell.value is not None:
                            try:
                                color = _get_color(float(cell.value))
                                if color:
                                    cell.fill = PatternFill(
                                        start_color=color,
                                        end_color=color,
                                        fill_type='solid'
                                    )
                            except (TypeError, ValueError):
                                pass

                wb.save(out_xlsx)
                print(f'  Salvato: {out_csv}')
                print(f'  Salvato: {out_xlsx}')

        print('\nApplicazione DPI HML completata.')


                
                
                


                
                










        




        



         
