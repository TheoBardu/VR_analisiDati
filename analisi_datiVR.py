import pandas as pd
import openpyxl as ex
from openpyxl.styles import PatternFill
from numpy import zeros, arange, mean, std, max, round, ones, log10, sum, dot
from os import chdir, getcwd


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

class files:
    '''
    Classe che serve per leggere e scrivere i file
    '''

    def read_measure_file(file, letter_ID):
        '''
        Funzione che legge il file di misura in txt e restituisce un dataframe pandas.

        INPUT:
            file = <str>, directory del file da leggere
            letter_ID = <str>, lettera della cartella dati (tipicamente D,E,F o W)
        OUTPUT:
            df = <pd.DataFrame>, dataframe contenente i dati della misura
        '''

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
                LeqA_min.append(float(l7[5])); durata.append(l7[8])
                LeqA_min.append(float(l7[9])); durata.append(l7[12])
                LeqA_min.append(float(l7[13])); durata.append(l7[16])
                
                LeqA_max.append(float(l7[6]))
                LeqA_max.append(float(l7[10]))
                LeqA_max.append(float(l7[14]))

                LeqA_eq.append(float(l7[7]))
                LeqA_eq.append(float(l7[11]))
                LeqA_eq.append(float(l7[15]))


                # salvataggio livelli equalizzati C
                l8 = lines[indx+8].split() #questa è la riga dei livelli C
                # print(l8)
                LeqC_min.append(float(l8[5]))
                LeqC_min.append(float(l8[9]))
                LeqC_min.append(float(l8[13])) 

                LeqC_max.append(float(l8[6]))
                LeqC_max.append(float(l8[10]))
                LeqC_max.append(float(l8[14]))

                LeqC_eq.append(float(l8[7]))
                LeqC_eq.append(float(l8[11]))
                LeqC_eq.append(float(l8[15]))  

                #salvataggio picchi C
                l9 = lines[indx+9].split()
                # print(l9)
                PeakC_max.append(float(l9[5]))
                PeakC_max.append(float(l9[8]))
                PeakC_max.append(float(l9[11]))

                PeakC_eq.append(float(l9[6]))
                PeakC_eq.append(float(l9[9]))
                PeakC_eq.append(float(l9[12]))





            
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
        
        print(f'Lista dei file nella directory: {self.file_list}') # stampa la lista dei file nella directory principale


    def iterate_directory(self,file_name = 'LSOURCES - Copia.txt'):
        '''
        Funzione che itera su tutte le directory per salvare i file.
        Crea i file csv e xlsx e colora le colonne opportune del file xlsx mediando tutti i dati.
        '''
        from os import path, mkdir
        for dir in self.file_list:
            print(f'Iterazione sulla directory: {dir}') 
            chdir(self.main_dir + '/' + dir + '/' + dir) 

            self.out_file_dir = self.main_dir + '/' + 'data'
            if path.isdir(self.out_file_dir) == False:
                mkdir(self.out_file_dir)
            
        

            df, num_of_track, n_files = files.read_measure_file(file_name,list(dir)[-1])
            files.write_csv(df, f'{self.out_file_dir}/{dir}.csv')
            files.write_exel(df, f'{self.out_file_dir}/{dir}.xlsx')
            exel_file.adjust_column_lenght(f'{self.out_file_dir}/{dir}.xlsx', ['A'])
            exel_file.color_column(f'{self.out_file_dir}/{dir}.xlsx', ['F','I','K'], ['FFFF00','FFFF00','FFFF00'])


        

        
        
    


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
                    LeqC_mean[i] = round(mean(df['LeqC_eq'][idx]) + std(df['LeqC_eq'][idx],ddof=1),1)
                    Ppeak_mean[i] = round(max(df['PeakC_max'][idx]) + 1.56,1)
                    # print(LeqA_mean)
                    
                    #calcolo l'incertezza sulla misura LeqA (SDOM)
                    U_sdom[i] = round(std(df['LeqA_eq'][idx], ddof=1),1)

                
                new_df = pd.DataFrame({'jobName': fileIDs ,
                                       'ID':letter_IDs ,
                                       'U':U_sdom,
                                       'LeqA' : LeqA_mean, 
                                       'LeqC' : LeqC_mean ,
                                       'Ppeak' : Ppeak_mean})
                df_avg = pd.concat([df_avg,new_df], ignore_index=True)
            
            
            df_avg['Ti'] = [10] * len(df_avg) #creo la colonna con i valori di exposure time
            df_avg['GrOm'] = [1] * len(df_avg) #creo la colonna con gli ID del gruppo omogeneo


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
        LeqA8H = self.calcolo_Leq8h( df_GrOm)
        self.U_comb_std = sum(  (df_GrOm['Ti']/T0  * 10**(0.1*( df_GrOm['LeqA'] - LeqA8H ) ) )**2   * ( df_GrOm['U']**2 + u2m**2 + u_pos **2 ) +
                    ( 4.34 * (1/T0  * 10**(0.1*( df_GrOm['LeqA'] - LeqA8H ) )) * std(df_GrOm['Ti'], ddof=1) )**2 )
        
        self.U_ext = self.U_comb_std * 1.65

        return self.U_ext, self.U_comb_std


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
            df_avg = files.read_csv(df_avg_dir) #lettura del df

            #creo le variabili vuote del dataframe da inserire 
            LeqA8h = []
            Uext = []
            U_comb_std = []
            Peak_max = []

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

                # input('--- pausa ---')
            
            # df_VR8h = pd.DataFrame(columns=['GrOm','LeqA', 'U' , 'Ppeak'])
            df_VR8h = pd.DataFrame({'GrOm': grorIDs,
                                    'LeqA_8h': LeqA8h,
                                    'Peak': Peak_max,
                                    'U_ext': Uext,})
            files.write_exel(df_VR8h, self.main_dir+ '/VR_8h.xlsx')
            files.write_csv(df_VR8h, self.main_dir+ '/VR_8h.csv')
            print(df_VR8h)
            print('Vautazione rischio 8h completata.')
        



        




        



         
