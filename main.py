#! /opt/anaconda3/bin/python3

# This code is the main code for the analysis of the data from the VR acquisitions

from os import chdir, getcwd
from analisi_datiVR import manager, analisi

main_directory = '/Users/theo/Desktop/Ermes/test/dati_test/misure'
#/Users/theo/Desktop/Ermes/Misure
chdir(main_directory)




# ==========================================================================
# Inizializzazione: lettura file e creazione tabelle =====================================

m = manager() # inizializzazione delle directory e dei file
m.iterate_directory() #iterazione su tutte le cartelle contenenti i dati .txt per salvare xls e csv
# ==========================================================================

# ==========================================================================
# analisi dei dati acquisiti =====================================

data_folder = main_directory + '/data' 
a = analisi(data_folder)
df_avg = a.average_values() #creazione delle medie delle misure
# # print(df_avg)


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
m.VR8h_exel(data_folder + '/averaged_data.csv', main_directory + '/VR_8h.xlsx')




