import openpyxl

wb = openpyxl.load_workbook('/Users/theo/Desktop/modello_scheda_gruppi_dpi.xlsx')
ws_mansioni = wb['Scheda_mansioni']
ws_grom = wb['Gruppi_omogenei']

#cell_value = ws_grom['A3'].value

lookup_mansioni = {}
for row in ws_mansioni.iter_rows(min_row=3, values_only=True):
    id_grom    = row[0]   # colonna A: ID_GrOm   (es. 'M_01')
    descrizione = row[1]  # colonna B: Descrizione_GrOm (es. 'Addetto lavaggio')
    if id_grom is not None and descrizione is not None:
        lookup_mansioni[str(descrizione).strip()] = str(id_grom).strip()
    print(id_grom, descrizione)
    print(lookup_mansioni)
    input()