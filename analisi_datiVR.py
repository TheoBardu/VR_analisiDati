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
            
            
            df_avg['Ti'] = [[10,10]] * len(df_avg) #creo la colonna con i valori di exposure time
            df_avg['GrOm'] = [["M1","M2"]] * len(df_avg) #creo la colonna con gli ID del gruppo omogeneo
            df_avg['DPI'] = [[False,True]] * len(df_avg) #creo la colonna contenente il riferimento all'uso del DPI


            files.write_csv(df_avg, self.main_dir + '/averaged_data.csv')
            files.write_exel(df_avg, self.main_dir + '/averaged_data.xlsx')
            exel_file.adjust_column_lenght(self.main_dir + '/averaged_data.xlsx', 'A')
            print('Averaged data files created')
            return df_avg
        else:
            print('File averaged.csv already exists!')
            df_avg = pd.read_csv(self.main_dir + '/averaged_data.csv')
            return df_avg
            


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
        df_summary.to_csv( os.path.join(output_dir, 'VR8h.csv'),  index=True)
        df_summary.to_excel(os.path.join(output_dir, 'VR8h.xlsx'), index=True)

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





                
                
                


                
                










        




        



         
