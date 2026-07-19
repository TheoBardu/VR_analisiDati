#! /opt/anaconda3/bin/python3

# This code is the main code for the analysis of the data from the VR acquisitions

from os import chdir, getcwd, system
import sys
from analisi_datiVR import manager, analisi
from config import NAME_RILIEVI_FONOMETRICI

# La main directory deve essere la root (quindi la cartella dell'azienda)
main_directory = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/ALLCOOP/rev/revN/Rumore'
misure_directory = '/misure'
risultati_directory = '/output'
dpi_directory = '/DPI_check'



def main(main_directory):
    # sys.stdout = open('VR_rumore.out', 'w')
    chdir(main_directory + misure_directory) # mi sposto nella cartella delle misure
    
    # ==========================================================================
    # Inizializzazione: lettura file e creazione tabelle =====================================

    m = manager() # inizializzazione delle directory e dei file
    m.iterate_directory(file_name = 'dati.txt',format='csv', versione_lettura='1') #iterazione su tutte le cartelle contenenti i dati .txt per salvare xls e csv
    # ==========================================================================

    # ==========================================================================
    # analisi dei dati acquisiti =====================================

    data_folder = main_directory + misure_directory + '/data' 
    a = analisi(data_folder)
    df_avg = a.average_values() #creazione delle medie delle misure
    # # print(df_avg)
    

    # Copio le info dei GrOm e Ti dal file excel
    df_HEG = a.get_scheda_info(df_avg, excel_info_dir=main_directory)     

    str_print_4input = '''
    §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
    !! -> Prima di Proseguire:\n
    Controlla il file averaged_data.csv e inserisci a mano i Ti e i GrOm.\n
    Poi premi invio per continuare con il codice...
    Puoi anche selezionare direttamente questa parte per continuare se hai interrotto il codice.\n
    §§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§§
    '''
    input(str_print_4input)

    
    

    # ==========================================================================
    # Valutazione Rischio 8 h =====================================
    # ==========================================================================
    a.analisi_8h(main_directory + risultati_directory, df_HEG)
    
    # input('Femi qui per ora, DPI da correggere.')

    # Applicazione medoto HLM per i DPI sulle misure singole
    system('cd ..')

    a.applica_DPI_HML(excel_info_scheda_dpi=main_directory + '/scheda_gruppi_dpi.xlsx' , 
                      excel_total=main_directory + risultati_directory + '/VR8h_totale.xlsx', 
                      excel_output= main_directory + risultati_directory + '/VR8h_riepilogo.xlsx',
                      excel_aggiornato=main_directory + risultati_directory + '/VR8h_totale_aggiornato.xlsx')
    
    # Creazione del file excel per rilievi fonometrici
    system(f"python /Users/theo/Desktop/P.IVA/Aziende/Ermes/Codici/VRR_analisiDati/VRR_analisiDati/utility/crea_excel_dati.py '{data_folder}' '{main_directory}' --output {NAME_RILIEVI_FONOMETRICI}")
    print(f'{NAME_RILIEVI_FONOMETRICI} creato correttamente')



    





if __name__ == "__main__":
    try:
        main(main_directory)
    except KeyboardInterrupt:
        print('End Program')