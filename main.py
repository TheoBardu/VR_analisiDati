#! /opt/anaconda3/bin/python3

# This code is the main code for the analysis of the data from the VR acquisitions

from os import chdir, getcwd
import sys
from analisi_datiVR import manager, analisi

# La main directory deve essere la root (quindi la cartella dell'azienda)
main_directory = '/Users/theo/Desktop/P.IVA/Ermes/Lavori/Ducati Energia'
misure_directory = '/misure'
risultati_directory = '/output'
dpi_directory = '/DPI_check'


def main(main_directory):
    # sys.stdout = open('VR_rumore.out', 'w')

    chdir(main_directory + misure_directory) # mi sposto nella cartella delle misure

    # ==========================================================================
    # Inizializzazione: lettura file e creazione tabelle =====================================

    m = manager() # inizializzazione delle directory e dei file
    m.iterate_directory(file_name = 'dati.txt',format='csv') #iterazione su tutte le cartelle contenenti i dati .txt per salvare xls e csv
    # ==========================================================================

    # ==========================================================================
    # analisi dei dati acquisiti =====================================

    data_folder = main_directory + misure_directory + '/data' 
    a = analisi(data_folder)
    df_avg = a.average_values() #creazione delle medie delle misure
    # # print(df_avg)
    
    # Copio le info dei GrOm e Ti dal file excel
    a.get_scheda_info(df_avg, excel_info_dir=main_directory)     

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
    a.analisi_8h(main_directory + risultati_directory, df_avg)
    

    # Applicazione medoto HLM per i DPI sulle misure singole
    a.applica_DPI_HML(excel_info_dir=main_directory , 
                      name_excel_info='scheda_gruppi_dpi.xlsx', 
                      output_directory= main_directory + risultati_directory, 
                      output_dpi= main_directory + risultati_directory + dpi_directory,
                      mode = 'both')
    




    #Calcolo PNR medio e LeqA ridotto sulla mansione
    a.applica_dpi_HLM_8h(main_dir=main_directory,
                         dpi_dir=main_directory + risultati_directory + dpi_directory,
                         total_VR8h_name="VR8h_totale")
    

    






if __name__ == "__main__":
    try:
        main(main_directory)
    except KeyboardInterrupt:
        print('End Program')