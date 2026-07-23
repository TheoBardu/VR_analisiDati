import pandas as pd
from datetime import datetime, timedelta
from io import StringIO
import math

file2extend = "/Users/theo/Desktop/P.IVA/Aziende/Ermes/Lavori/SIT - RSM/rev2/Rumore/misure/misD/20260226_183040_183106.csv"
Ttot = 6  # durata minima desiderata in minuti

TIME_FMT = "%H:%M:%S"

# Lettura raw per preservare il formato originale
with open(file2extend, "r", encoding="latin-1") as f:
    raw_lines = f.read().splitlines()

device_header = raw_lines[0]
col_header = raw_lines[1]

# Trova la riga vuota che separa i dati dalle statistiche aggregate
data_end = None
for i in range(2, len(raw_lines)):
    if raw_lines[i].strip() == "":
        data_end = i
        break

if data_end is None:
    raise ValueError("Riga vuota separatrice non trovata nel file CSV.")

data_lines = raw_lines[2:data_end]
footer_lines = raw_lines[data_end:]

# Import con pandas per estrarre first_time e last_time
df = pd.read_csv(StringIO(col_header + "\n" + "\n".join(data_lines)), sep=";")
first_time = datetime.strptime(df["Time"].iloc[0], TIME_FMT)
last_time = datetime.strptime(df["Time"].iloc[-1], TIME_FMT)

deltaT = (last_time - first_time).total_seconds()   # usato per calcolare N
deltaT_shift = deltaT + 1                             # stride tra blocchi (nessun overlap)

N = math.ceil(Ttot * 60 / deltaT)

# Genera le righe estese modificando solo la colonna Time
extended_data = list(data_lines)
for i in range(1, N + 1):
    for j, line in enumerate(data_lines):
        parts = line.split(";")
        new_time = first_time + timedelta(seconds=i * deltaT_shift + j)
        parts[0] = new_time.strftime(TIME_FMT)
        extended_data.append(";".join(parts))

# Sovrascrive il file originale
with open(file2extend, "w", encoding="latin-1") as f:
    f.write(device_header + "\n")
    f.write(col_header + "\n")
    for line in extended_data:
        f.write(line + "\n")
    for line in footer_lines:
        f.write(line + "\n")

total_seconds = N * deltaT_shift + deltaT
print(f"File sovrascritto: {file2extend}")
print(f"deltaT = {int(deltaT)}s | copie aggiunte = {N} | durata totale ≈ {total_seconds/60:.1f} min")
