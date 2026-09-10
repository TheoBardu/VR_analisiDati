# Frontespizi della relazione Rumore

Guida alla creazione dei file di frontespizio usati da
[`utility/write_docx_Rumore.py`](../utility/write_docx_Rumore.py).

## Come funziona

Il template principale (`Modello_RUM.docx`) non contiene la copertina: al suo
posto ha il segnaposto `{{p frontespizio }}`. La copertina vera è un file `.docx` a sé
stante dentro:

```
/Users/theo/Desktop/P.IVA/Aziende/Ermes/Modelli/docx/frontespizi/RUM/
```

Si sceglie quale usare con la costante `FRONTESPIZIO` in `write_docx_Rumore.py`:

```python
DIR_FRONTESPIZI = path.dirname(DIR_MODELLI) + '/frontespizi/RUM'
FRONTESPIZIO    = 'frontespizio_Relyon_RUM.docx'
```

Il file viene renderizzato **in una passata separata** con lo stesso `context` del
documento principale, e poi inserito come sotto-documento (`doc.new_subdoc(...)`).
Per questo un frontespizio deve essere di per sé un template `.docx` valido: se ha un
errore di sintassi Jinja, la scrittura della relazione si interrompe.

### Intestazione e piè di pagina della prima pagina

`new_subdoc` copia solo il corpo del frontespizio e scarta sezione, intestazione e piè di
pagina. Per questo, dopo il render, `applica_intestazioni_frontespizio()` copia nella
prima sezione della relazione l'intestazione e il piè di pagina **visibili sulla pagina
del frontespizio** (quelli "prima pagina" se il file ha *Diversi per la prima pagina*,
altrimenti quelli normali):

- se il frontespizio ha un'intestazione, la prima pagina della relazione la mostra;
- se non ce l'ha, la prima pagina resta **senza** intestazione (idem per il piè di pagina):
  non si eredita nulla dal modello principale;
- dalla seconda pagina in poi valgono intestazione e piè di pagina di `Modello_RUM.docx`.

Immagini e collegamenti presenti in intestazione/piè di pagina del frontespizio vengono
portati nel documento finale. Anche in queste parti si possono usare i tag Jinja: vengono
risolti con lo stesso `context` del corpo.

---

## Voci obbligatorie

Sono le voci che identificano l'azienda, la revisione e i firmatari. Devono comparire in
ogni frontespizio, perché sono i dati che rendono il documento completo e tracciabile.

| Voce | Contenuto |
|---|---|
| `{{nome_azienda}}` | Ragione sociale dell'azienda |
| `{{indirizzo_azienda}}` | Indirizzo dell'unità produttiva oggetto della valutazione |
| `{{data_emissione}}` | Data di emissione del documento |
| `{{revisione}}` | Numero di revisione (es. `rev.01`) |
| `{{datore_di_lavoro}}` | Firma — Datore di Lavoro |
| `{{RSPP}}` | Firma — Responsabile del Servizio di Prevenzione e Protezione |
| `{{medico_competente}}` | Firma — Medico Competente |
| `{{RLS}}` | Firma — Rappresentante dei Lavoratori per la Sicurezza |
| `{{delegato_sicurezza}}` | Firma — Delegato alla sicurezza |

Queste nove voci sono presenti in tutti i frontespizi esistenti e vanno considerate il
nucleo minimo di un frontespizio nuovo.

## Voci facoltative

| Voce | Contenuto | Nota |
|---|---|---|
| `{{img_logo_azienda}}` | Logo aziendale inserito come immagine | Vedi sotto |
| `{{note_titolo}}` | Sottotitolo/nota (es. "relazione per la sede di Mantova") | Va aggiunta a mano al dizionario `context` se la si usa |
| `{{data_scadenza}}` | Data di scadenza della valutazione | Di norma sta nelle conclusioni, non in copertina |
| `{{sede_legale}}` | Sede legale, se diversa dall'unità produttiva | |

### Due modi di gestire il logo

1. **Immagine incorporata nel `.docx`** — si incolla il logo direttamente nel file del
   frontespizio in Word. È l'approccio usato da `frontespizio_Relyon_RUM.docx`. Consigliato
   quando il frontespizio è già specifico di un'azienda.
2. **Segnaposto `{{img_logo_azienda}}`** — l'immagine viene inserita a run time dal
   percorso indicato in `LOGO_AZIENDA` (larghezza `LARGHEZZA_LOGO_MM`). Se il file non
   esiste, lo script **non si blocca**: stampa un avviso e lascia il campo vuoto.

---

## Regole tecniche

Vincoli che derivano da come docxtpl inserisce il sotto-documento. Se non li rispetti,
il frontespizio viene scritto male o la generazione fallisce.

**Cosa non fare**

- **Niente interruzione di pagina finale.** Il template principale ne ha già una subito
  dopo il segnaposto: aggiungerne un'altra produce una pagina bianca.
- **Intestazione e piè di pagina valgono solo per la prima pagina.** Diventano quelli
  della pagina del frontespizio nella relazione (vedi sopra); se mancano, la prima pagina
  resta senza. Non influenzano le pagine successive.
- **Niente impostazioni di pagina.** Margini, orientamento e formato vengono presi dal
  template principale, non dal file del frontespizio.
- **Mai accedere a campi di variabili non definite.** `{{ogg.campo}}` su un nome assente
  dal `context` solleva `UndefinedError` e blocca tutto. Una variabile semplice assente
  (`{{var_assente}}`) invece viene resa come stringa vuota, senza errore: comodo, ma
  significa che un nome scritto male **non dà errore, dà solo un buco nel documento**.

**Cosa è supportato**

- Tabelle, caselle di testo, immagini e forme: importate con le relative relazioni.
- Il file può contenere una sola pagina o poche pagine; il contenuto viene inserito così
  com'è nel punto del segnaposto.

---

## Creare un frontespizio nuovo

1. Duplica un frontespizio esistente e rinominalo (es. `frontespizio_azienda.docx`).
2. Modifica il layout in Word tenendo **tutte le voci obbligatorie** elencate sopra.
3. Salvalo nella cartella `frontespizi/`.
4. Aggiorna `FRONTESPIZIO` in `write_docx_Rumore.py` con il nome esatto del file.

> **Attenzione al nome.** Su macOS il filesystem di norma non distingue maiuscole e
> minuscole, quindi `frontespizio_relyon_rum.docx` funziona anche se il file si chiama
> `frontespizio_Relyon_RUM.docx`. Su altri sistemi no: conviene scrivere il nome esatto.
> Se il file non viene trovato, lo script elenca i frontespizi disponibili nella cartella.

### Verifica prima dell'uso

Controlla che il file compili e che usi le voci attese:

```bash
python3 -c "
from docxtpl import DocxTemplate
f = '/Users/theo/Desktop/P.IVA/Aziende/Ermes/Modelli/docx/frontespizi/RUM/frontespizio_Relyon_RUM.docx'
print(sorted(DocxTemplate(f).get_undeclared_template_variables()))
"
```

Se il comando solleva `TemplateSyntaxError`, il frontespizio ha un errore di sintassi
Jinja e va corretto prima di usarlo. Se la lista stampata contiene un nome che non
riconosci, è un refuso: verrà reso come stringa vuota senza segnalare nulla.

---

## Elenco completo delle voci disponibili

Il frontespizio riceve l'intero `context` della relazione, quindi oltre alle voci sopra
sono tecnicamente utilizzabili anche tutte le altre definite nella **Sezione 1** di
`write_docx_Rumore.py` — fra cui `attivita_azienda`, `processo_produttivo`,
`sede_operativa`, `sostanze_ototossiche`, `interazione_vib_rum`, `effetti_indesiderati`
e le rispettive misure attuative — oltre alle tabelle caricate automaticamente dagli
Excel (`tabella_dpi`, `tabella_mansioni`, `tabella_HEG`, `HEG_med`, `HEG_alto`).

In pratica non servono in copertina: sono elencate solo per completezza. Se aggiungi una
voce nuova al frontespizio, ricordati di definirla nel dizionario `context`.
