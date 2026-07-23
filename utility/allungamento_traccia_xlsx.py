import math
from copy import copy
from datetime import datetime, timedelta

import openpyxl

file2extend = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/DEFRANCESCHI/rev/rev0/Rumore/misure/misH/730 A10574-26070313-105933.xlsx'
Ttot = 6  # durata minima desiderata in minuti

SHEET_NAME = "Profilo storico"
TIME_FMT = "%Y-%m-%d  %H:%M:%S"
TIME_COL_HEADER = "Data/Tempo"

wb = openpyxl.load_workbook(file2extend)
ws = wb[SHEET_NAME]

header_row = 1
max_col = ws.max_column

# Trova la colonna "Data/Tempo" cercando nell'intestazione
time_col = None
for col in range(1, max_col + 1):
    header_val = ws.cell(row=header_row, column=col).value
    if header_val and TIME_COL_HEADER in header_val:
        time_col = col
        break

if time_col is None:
    raise ValueError(f"Colonna '{TIME_COL_HEADER}' non trovata nel foglio '{SHEET_NAME}'.")

# Trova l'ultima riga dati contigua a partire dalla riga successiva all'intestazione
data_start = header_row + 1
data_end = data_start
while ws.cell(row=data_end, column=time_col).value is not None:
    data_end += 1
data_end -= 1

if data_end < data_start:
    raise ValueError(f"Nessuna riga dati trovata nel foglio '{SHEET_NAME}'.")

data_rows = list(range(data_start, data_end + 1))
footer_rows = list(range(data_end + 1, ws.max_row + 1))

first_time = datetime.strptime(ws.cell(row=data_start, column=time_col).value, TIME_FMT)
last_time = datetime.strptime(ws.cell(row=data_end, column=time_col).value, TIME_FMT)

deltaT = (last_time - first_time).total_seconds()   # usato per calcolare N
deltaT_shift = deltaT + 1                             # stride tra blocchi (nessun overlap)

N = math.ceil(Ttot * 60 / deltaT)

# Salva valori/stili delle righe di footer prima di cancellarle dal foglio
footer_data = []
for r in footer_rows:
    row_values = []
    for col in range(1, max_col + 1):
        cell = ws.cell(row=r, column=col)
        row_values.append((cell.value, cell))
    footer_data.append((row_values, ws.row_dimensions[r].height))

footer_snapshot = []
for row_values, height in footer_data:
    snap = [(value, copy(cell.font), copy(cell.fill), copy(cell.border),
             copy(cell.alignment), copy(cell.protection), cell.number_format)
            for value, cell in row_values]
    footer_snapshot.append((snap, height))

if footer_rows:
    ws.delete_rows(footer_rows[0], len(footer_rows))

# Genera le righe estese modificando solo la colonna Data/Tempo
next_row = data_end + 1
for i in range(1, N + 1):
    for j, src_row in enumerate(data_rows):
        new_time = first_time + timedelta(seconds=i * deltaT_shift + j)
        for col in range(1, max_col + 1):
            src_cell = ws.cell(row=src_row, column=col)
            new_cell = ws.cell(row=next_row, column=col)
            new_cell.value = new_time.strftime(TIME_FMT) if col == time_col else src_cell.value
            new_cell.font = copy(src_cell.font)
            new_cell.fill = copy(src_cell.fill)
            new_cell.border = copy(src_cell.border)
            new_cell.alignment = copy(src_cell.alignment)
            new_cell.protection = copy(src_cell.protection)
            new_cell.number_format = src_cell.number_format
        ws.row_dimensions[next_row].height = ws.row_dimensions[src_row].height
        next_row += 1

# Riscrive le righe di footer originali dopo i nuovi blocchi
for snap, height in footer_snapshot:
    for col, (value, font, fill, border, alignment, protection, number_format) in enumerate(snap, start=1):
        cell = ws.cell(row=next_row, column=col)
        cell.value = value
        cell.font = font
        cell.fill = fill
        cell.border = border
        cell.alignment = alignment
        cell.protection = protection
        cell.number_format = number_format
    ws.row_dimensions[next_row].height = height
    next_row += 1

# Sovrascrive il file originale
wb.save(file2extend)

total_seconds = N * deltaT_shift + deltaT
print(f"File sovrascritto: {file2extend}")
print(f"deltaT = {int(deltaT)}s | copie aggiunte = {N} | durata totale ≈ {total_seconds/60:.1f} min")
