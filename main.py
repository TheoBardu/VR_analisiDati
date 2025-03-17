#! /opt/anaconda3/bin/python3

# This code is the main code for the analysis of the data from the VR acquisitions

from os import chdir, getcwd
from analisi_datiVR import manager, analisi

chdir('/Users/theo/Desktop/Ermes/Misure')

# f = manager() # inizializzazione delle directory e dei file
# f.iterate_directory() #iterazione su tutte le cartelle contenenti i dati

data_folder = '/Users/theo/Desktop/Ermes/Misure/data' 
a = analisi(data_folder)
df_avg = a.average_values()
print(df_avg)