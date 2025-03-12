#! /opt/anaconda3/bin/python3

# This code is the main code for the analysis of the data from the VR acquisitions

import pandas as pd
import openpyxl as ex
from openpyxl.styles import PatternFill
from numpy import zeros, arange


file = '/Users/theo/Desktop/Ermes/Misure/misF/misF/LSOURCES - Copia.txt'


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
            color = <str>, colore in formato esadecimale (es. 'FFFF00' per il giallo)
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

    def read_measure_file(file):
        '''
        Funzione che legge il file di misura e restituisce un dataframe pandas.

        INPUT:
            file = <str>, directory del file da leggere
        OUTPUT:
            df = <pd.DataFrame>, dataframe contenente i dati della misura
        '''

        df = pd.DataFrame(columns=['fileID','nTrack','LeqA_min','LeqA_max','LeqA_eq','LeqC_min','LeqC_max','LeqC_eq','PeakC_max','PeakC_eq','durata','inizio','fine'])
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
        

        inxd_fileIDs = 0
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
                # inxd_fileIDs += 1
                
                l1 = lines[indx+1].split() # questa è la riga dell'inizio
                

                l2 = lines[indx+2].split() #questa è la riga della fine
                
                l3 = lines[indx+3].split() #questa è la riga del numero di sorgenti
                
                ntracks = arange(len(l3)-1) #creo un array con il numero di sorgenti in ordine crescente

                #itero sul numero di sorgenti
                for i in range(len(ntracks)):
                    fileIDs.append(l0[1]) #creo tante copie del nome sorgente quante sono le sorgenti
                    inizio.append(l1[2]) # creo tante copie dell'inizio quante sono le sorgenti
                    fine.append(l2[2]) # creo tante copie della fine quante sono le sorgenti
                    nSorgenti.append(ntracks[i]+1) # creo un numero di sorgenti in ordine crescente

                #salvataggio livelli euqalizzati A e durata tracciati
                l7 = lines[indx+7].split() #questa è la riga dei livelli LeqA
                # print(l7)
                LeqA_min.append(l7[5]); durata.append(l7[8])
                LeqA_min.append(l7[9]); durata.append(l7[12])
                LeqA_min.append(l7[13]); durata.append(l7[16])
                
                LeqA_max.append(l7[6])
                LeqA_max.append(l7[10])
                LeqA_max.append(l7[14])

                LeqA_eq.append(l7[7])
                LeqA_eq.append(l7[11])
                LeqA_eq.append(l7[15])


                # salvataggio livelli equalizzati C
                l8 = lines[indx+8].split() #questa è la riga dei livelli C
                # print(l8)
                LeqC_min.append(l8[5])
                LeqC_min.append(l8[9])
                LeqC_min.append(l8[13]) 

                LeqC_max.append(l8[6])
                LeqC_max.append(l8[10])
                LeqC_max.append(l8[14])

                LeqC_eq.append(l8[7])
                LeqC_eq.append(l8[11])
                LeqC_eq.append(l8[15])  

                #salvataggio picchi C
                l9 = lines[indx+9].split()
                # print(l9)
                PeakC_max.append(l9[5])
                PeakC_max.append(l9[8])
                PeakC_max.append(l9[11])

                PeakC_eq.append(l9[6])
                PeakC_eq.append(l9[9])
                PeakC_eq.append(l9[12])





            
            indx += 1

            

        df['fileID'] = fileIDs
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
        return df


    def write_csv(df,file):
        '''
        Funzione che scrive il dataframe df in un file csv
        '''
        df.to_csv(file)


    def write_exel(df, file):
        '''
        Funzione che scrive il dataframe df df in un file excel
        '''
        df.to_excel(file, index=False)
        


    def read_csv(file):
        '''
        Funzione che legge il file csv grezzo e lo trasforma nel file txt measure file
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
        if ".DS_Store" in self.file_list:
            self.file_list.remove(".DS_Store")
        elif "data" in self.file_list:
            self.file_list.remove("data")
        
        print(f'Lista dei file nella directory: {self.file_list}') # stampa la lista dei file nella directory principale


    def iterate_directory(self,file_name = 'LSOURCES - Copia.txt'):
        '''
        Funzione che itera su tutte le directory per salvare i file
        '''
        from os import path, mkdir
        for dir in self.file_list:
            print(f'Iterazione sulla directory: {dir}')
            chdir(self.main_dir + '/' + dir + '/' + dir)

            self.out_file_dir = self.main_dir + '/' + 'data'
            if path.isdir(self.out_file_dir) == False:
                mkdir(self.out_file_dir)
            
        

            df = files.read_measure_file(file_name)
            files.write_csv(df, f'{self.out_file_dir}/{dir}.csv')
            files.write_exel(df, f'{self.out_file_dir}/{dir}.xlsx')
            exel_file.adjust_column_lenght(f'{self.out_file_dir}/{dir}.xlsx', ['A'])
            exel_file.color_column(f'{self.out_file_dir}/{dir}.xlsx', ['E','H','J'], ['FFFF00','FFFF00','FFFF00'])


    



class analisi:
    '''
    Classe con i metodi per l'analisi delle misure
    '''

        



# MAIN Program

from os import chdir
chdir('/Users/theo/Desktop/Ermes/Misure')

f = manager()
f.iterate_directory()

# df = files.read_measure_file(file)
# files.write_csv(df, '/Users/theo/Desktop/Ermes/test/prova.csv')
# files.write_exel(df, '/Users/theo/Desktop/Ermes/test/prova.xlsx')
# exel_file.color_column('/Users/theo/Desktop/Ermes/test/prova.xlsx', ["E","F"], ['FFFF00','AFFA00'])



