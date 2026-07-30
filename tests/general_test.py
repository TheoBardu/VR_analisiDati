#Code that correct the width of V

import openpyxl

wb = openpyxl.load_workbook('/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/NISI/rev/rev2/Rumore/output/VR8h_totale_aggiornato.xlsx')

for ws in wb.worksheets:
    if ws.title.strip().lower().startswith('scheda'):
        ws.column_dimensions['V'].width = 6

wb.save('/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/NISI/rev/rev2/Rumore/output/VR8h_totale_aggiornato.xlsx')
print('done')