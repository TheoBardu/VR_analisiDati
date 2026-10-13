"""
gestore_tracce_integrative.py

Legge le misure integrative da Excel e le aggiunge al CSV esistente.
Modificare esclusivamente la sezione PARAMETRI UTENTE prima dell'esecuzione.

Dipendenze: pandas, openpyxl
Installazione (se necessaria): pip install pandas openpyxl
"""

from pathlib import Path
import math
import pandas as pd


# =============================================================================
# PARAMETRI UTENTE
# =============================================================================

EXCEL_TRACCE_INTEGRATIVE = r"C:\percorso\FAELUX_SRL_scheda_gruppi_dpi.xlsx"
CSV_data_FILE = r"C:\percorso\averaged_data.csv"

# Nome del foglio Excel contenente le tracce integrative.
FOGLIO_EXCEL = "misure integrative rumore"

# Indici delle colonne Excel: A=0, B=1, C=2, ...
# NOTA: nel file Excel allegato, la colonna C contiene l'ID (es. D2)
# e la D la descrizione; perciò i default sotto corrispondono al file allegato.
# Se il tuo file ha ID in A e descrizione in B, imposta rispettivamente 0 e 1.
COLONNA_ID_MISURA = 2
COLONNA_DESCRIZIONE = 3  # letta per completezza/log, non scritta nel CSV

# Ogni gruppo rappresenta una misurazione: tempo, LAeq, LCeq, Picco_C.
# Inserire fino a 6 gruppi. Gli indici sono zero-based (A=0, B=1, ...).
# Nel file allegato: G:J, M:P, Q:T; lasciare eventuali gruppi mancanti fuori.
GRUPPI_MISURE = [
    {"tempo": 6,  "laeq": 7,  "lceq": 8,  "picco_c": 9},   # G:J
    {"tempo": 12, "laeq": 13, "lceq": 14, "picco_c": 15}, # M:P
    {"tempo": 16, "laeq": 17, "lceq": 18, "picco_c": 19}, # Q:T
    # Esempio per una quarta misura (modificare gli indici secondo il file):
    # {"tempo": 20, "laeq": 21, "lceq": 22, "picco_c": 23},
    # {"tempo": 24, "laeq": 25, "lceq": 26, "picco_c": 27},
    # {"tempo": 28, "laeq": 29, "lceq": 30, "picco_c": 31},
]

# Se True, combina i livelli acustici in modo energetico:
# L_tot = 10 * log10(sum(10 ** (L_i / 10))).
# È la formula acusticamente corretta per livelli riferiti a durate uguali.
# Se le durate differiscono, il codice usa una media energetica ponderata
# per i tempi quando PONDERA_PER_TEMPO è True.
PONDERA_PER_TEMPO = False

# Deviazione standard: "popolazione" oppure "campione".
TIPO_DEVIAZIONE_STANDARD = "campione"

# Il CSV di esempio ha una prima colonna indice "Unnamed: 0".
# Se True, il nuovo record riceve il prossimo indice numerico disponibile.
GESTISCI_COLONNA_INDICE = True


# =============================================================================
# IMPLEMENTAZIONE
# =============================================================================

COLONNE_CSV_ATTESE = ["jobName", "ID", "U", "LeqA", "LeqC", "Ppeak", "Ti", "GrOm"]


def numero(valore):
    """Converte un valore Excel in float; celle vuote/non numeriche diventano None."""
    if pd.isna(valore) or str(valore).strip() == "":
        return None
    try:
        return float(valore)
    except (TypeError, ValueError):
        return None


def combina_livelli(livelli, tempi=None):
    """Combina livelli in dB tramite somma energetica, opzionalmente ponderata."""
    validi = []
    for i, livello in enumerate(livelli):
        if livello is None:
            continue
        peso = 1.0
        if PONDERA_PER_TEMPO:
            tempo = tempi[i] if tempi is not None else None
            if tempo is None or tempo <= 0:
                continue
            peso = tempo
        validi.append((livello, peso))

    if not validi:
        return None

    energia = sum(peso * (10.0 ** (livello / 10.0)) for livello, peso in validi)
    if PONDERA_PER_TEMPO:
        peso_totale = sum(peso for _, peso in validi)
        energia /= peso_totale
    return 10.0 * math.log10(energia)


def deviazione_standard(valori):
    validi = [v for v in valori if v is not None]
    if len(validi) < 2:
        return 0.0 if len(validi) == 1 else None
    ddof = 1 if TIPO_DEVIAZIONE_STANDARD.lower() == "campione" else 0
    if len(validi) <= ddof:
        return None
    return float(pd.Series(validi, dtype=float).std(ddof=ddof))


def main():
    excel_path = Path(EXCEL_TRACCE_INTEGRATIVE)
    csv_path = Path(CSV_data_FILE)

    if not excel_path.is_file():
        raise FileNotFoundError(f"File Excel non trovato: {excel_path}")
    if not csv_path.is_file():
        raise FileNotFoundError(
            f"CSV non trovato: {csv_path}. Il file deve già esistere perché si usa append mode."
        )
    if not GRUPPI_MISURE:
        raise ValueError("Definire almeno un gruppo in GRUPPI_MISURE.")
    if len(GRUPPI_MISURE) > 6:
        raise ValueError("Sono supportati al massimo 6 gruppi di misurazione.")

    tracce = pd.read_excel(excel_path, sheet_name=FOGLIO_EXCEL, header=None)
    csv_esistente = pd.read_csv(csv_path)

    # Il CSV deve avere le colonne attese. È ammessa la colonna indice iniziale.
    colonne = list(csv_esistente.columns)
    if GESTISCI_COLONNA_INDICE and colonne and colonne[0].startswith("Unnamed:"):
        colonna_indice = colonne[0]
        colonne_dati = colonne[1:]
    else:
        colonna_indice = None
        colonne_dati = colonne

    mancanti = [c for c in COLONNE_CSV_ATTESE if c not in colonne_dati]
    if mancanti:
        raise ValueError(
            f"Nel CSV mancano queste colonne richieste: {mancanti}. "
            f"Colonne trovate: {colonne}"
        )

    indici_usati = [COLONNA_ID_MISURA, COLONNA_DESCRIZIONE]
    for gruppo in GRUPPI_MISURE:
        indici_usati.extend(gruppo.values())
    if max(indici_usati) >= tracce.shape[1]:
        raise ValueError(
            f"Una colonna configurata non esiste nel foglio Excel, che ha "
            f"{tracce.shape[1]} colonne (indici da 0 a {tracce.shape[1]-1})."
        )

    nuovi_record = []
    prossimi_indici = []
    if colonna_indice is not None:
        numerici = pd.to_numeric(csv_esistente[colonna_indice], errors="coerce").dropna()
        prossimo_indice = int(numerici.max()) + 1 if not numerici.empty else 0
    else:
        prossimo_indice = None

    for riga_excel, riga in tracce.iterrows():
        id_val = riga.iloc[COLONNA_ID_MISURA]
        descrizione = riga.iloc[COLONNA_DESCRIZIONE]

        if pd.isna(id_val) or str(id_val).strip() == "":
            continue
        id_misura = str(id_val).strip()

        laeq, lceq, picchi, tempi = [], [], [], []
        for gruppo in GRUPPI_MISURE:
            tempo = numero(riga.iloc[gruppo["tempo"]])
            la = numero(riga.iloc[gruppo["laeq"]])
            lc = numero(riga.iloc[gruppo["lceq"]])
            picco = numero(riga.iloc[gruppo["picco_c"]])

            # Una misura è considerata presente se contiene almeno un dato acustico.
            if la is None and lc is None and picco is None:
                continue
            laeq.append(la)
            lceq.append(lc)
            picchi.append(picco)
            tempi.append(tempo)

        if not any(v is not None for v in laeq + lceq + picchi):
            continue

        u = deviazione_standard(laeq)
        leqa_tot = combina_livelli(laeq, tempi)
        leqc_tot = combina_livelli(lceq, tempi)
        picchi_validi = [p for p in picchi if p is not None]
        picco_max = max(picchi_validi) if picchi_validi else None

        record = {
            "jobName": id_misura,
            "ID": id_misura,
            "U": u,
            "LeqA": leqa_tot,
            "LeqC": leqc_tot,
            "Ppeak": picco_max,
            "Ti": "[]",
            "GrOm": "[]",
        }

        if colonna_indice is not None:
            record[colonna_indice] = prossimo_indice
            prossimo_indice += 1

        # Mantiene l'ordine delle colonne del CSV già esistente.
        nuovi_record.append({c: record.get(c, "") for c in colonne})

        print(
            f"Excel riga {riga_excel + 1}: ID={id_misura}, "
            f"descrizione={str(descrizione).strip()}, "
            f"LeqA={leqa_tot}, LeqC={leqc_tot}, Ppeak={picco_max}, U={u}"
        )

    if not nuovi_record:
        print("Nessuna nuova misura valida trovata. Il CSV non è stato modificato.")
        return

    # Append reale: non riscrive le righe già presenti.
    pd.DataFrame(nuovi_record, columns=colonne).to_csv(
        csv_path,
        mode="a",
        header=False,
        index=False,
        encoding="utf-8",
    )
    print(f"\nCompletato: aggiunte {len(nuovi_record)} righe a {csv_path}")


if __name__ == "__main__":
    main()
