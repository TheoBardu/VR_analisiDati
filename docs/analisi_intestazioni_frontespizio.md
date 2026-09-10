# Intestazione e piè di pagina del frontespizio — backend Rumore

Analisi funzionale (e implementazione) della gestione di intestazione e piè di pagina
della prima pagina di `Relazione_RUM.docx`, in
[`utility/write_docx_Rumore.py`](../utility/write_docx_Rumore.py).

## Problema

La relazione non mostrava l'intestazione di prima pagina del frontespizio scelto (es.
`frontespizio_Ermes_RUM.docx`: `headerReference type="first"` + `titlePg`) e mostrava
sempre il piè di pagina Relyon, che sta in `Modello_RUM.docx` (`footer2.xml`,
`footerReference type="first"` della prima sezione).

Causa: `costruisci_frontespizio()` inserisce il frontespizio con
`documento.new_subdoc()`; docxtpl copia solo il corpo del sottodocumento e scarta
`sectPr` e riferimenti a header/footer. La prima pagina eredita quindi header/footer
della prima sezione del modello.

## Requisito

Dopo il render del documento principale, la **prima sezione** di `Relazione_RUM.docx`
deve avere come intestazione e piè di pagina "di prima pagina" quelli che si vedono
sulla pagina del frontespizio scelto; se il frontespizio non ne ha uno, la parte
corrispondente resta vuota (niente Relyon con Ermes). Dalla seconda pagina in poi
header/footer del modello invariati.

Con `frontespizio_Relyon_RUM.docx` (footer `default`, nessun header, nessun `titlePg`)
il risultato è: prima pagina senza intestazione, piè di pagina Relyon preso dal
frontespizio — identico a prima.

## Vincoli

- Il file è caricato da AnalisiRischio (`runner/runner_relazione.py`, `carica_testa`)
  **solo fino alla riga** `# SEZIONE 3 - COSTRUZIONE DEI CONTEXT AUTOMATICI`: le nuove
  funzioni stanno **prima** di quella sentinella (SEZIONE 2, accanto a `logo_inline` e
  `costruisci_frontespizio`), e quel commento non va cambiato.
- Nessun modello/frontespizio `.docx` viene toccato.
- python-docx 1.2, docxtpl 0.19 (ambiente `/opt/anaconda3`).

## Modifica

1. Import: `copy`, `docx.oxml.OxmlElement`, `docx.oxml.ns.qn`,
   `docx.enum.section.WD_HEADER_FOOTER`.
2. Nuove funzioni in SEZIONE 2, dopo `costruisci_frontespizio`:
   - `_parte_sezione(docx_documento, sezione, quale, prima_pagina)`: restituisce la
     parte header/footer (oggetto `Part`) referenziata dalla sezione, o `None` se il
     riferimento manca.
   - `_parte_visibile(docx_documento, sezione, quale)`: la parte che si vede sulla
     prima pagina della sezione (first se `titlePg`, altrimenti default).
   - `_copia_parte(sorgente, destinazione)`: svuota la parte di destinazione e vi copia
     i figli della sorgente, ricreando le relazioni (`r:embed`, `r:id`, `r:link`:
     immagini e collegamenti) nella parte di arrivo; con sorgente `None` lascia un solo
     paragrafo vuoto.
   - `applica_intestazioni_frontespizio(documento, file_frontespizio, contesto)`:
     renderizza il frontespizio in un `DocxTemplate` a sé (contesto senza la chiave
     `frontespizio`, `img_logo_azienda` legato a quel template), imposta
     `different_first_page_header_footer = True` sulla prima sezione del documento
     principale, crea se mancano le parti "first" (`is_linked_to_previous = False`) e vi
     copia header/footer visibili del frontespizio.
3. SEZIONE 4, fra `doc.render(context_completo)` e `doc.save(OUTPUT_DOCUMENT)`:
   `applica_intestazioni_frontespizio(doc, path.join(DIR_FRONTESPIZI, FRONTESPIZIO), context_completo)`.
   Va dopo il render: docxtpl non rielabora le parti copiate e il render del frontespizio
   non interferisce con quello principale.
4. Commento in testa al file e `docs/frontespizi.md` aggiornati.

### Due trappole di docxtpl 0.19 (trovate in fase di verifica)

- **Non usare `get_docx()` dopo il render.** `DocxTemplate.get_docx()` chiama
  `init_docx()`, che se `is_rendered` è vero **ricarica il template da file** e butta via
  il documento renderizzato (e `save()` salverebbe poi il template vuoto). Si usa
  l'attributo `.docx` direttamente.
- **Non usare `section.header` / `first_page_footer` per leggere o modificare le parti.**
  Nel render docxtpl sostituisce ogni parte header/footer nelle relazioni del documento
  (`rel._target = XmlPart nuova`), ma python-docx tiene una cache (`related_parts`) e
  continua a restituire la parte **originale**: le modifiche fatte lì si perdono al
  salvataggio, e la lettura dal frontespizio restituirebbe i tag Jinja non risolti. Le
  parti si raggiungono da `sectPr.get_headerReference/get_footerReference(tipo).rId` e
  `docx.part.rels[rId].target_part`. Le proprietà python-docx restano utili per creare
  parte e riferimento quando mancano (`is_linked_to_previous = False`).

## Verifica eseguita

Script eseguito su `Lavori/AL7/rev/rev2/Rumore` con `frontespizio_Ermes_RUM.docx` e
`frontespizio_Relyon_RUM.docx`:

- Prima `<w:sectPr>` con `titlePg`, `headerReference type="first"` e
  `footerReference type="first"`; sezioni successive invariate.
- Ermes: header first "missiroli ermes s.r.l. — Acustica Sicurezza Vibrazioni", footer
  first vuoto. Relyon: header first vuoto, footer first "Relyon srl…" con i due
  collegamenti (`mailto:`, `http://www.rely-on.it`) ricreati in `footer2.xml.rels`.
- PDF (`soffice --headless --convert-to pdf`): pagina 1 corretta in entrambi i casi,
  pagina 2 con intestazione azienda/rev/data e "pag. 2 di N" del modello.

Non è stato provato un frontespizio con immagine nell'intestazione (nessun file
disponibile): il percorso `r:embed` è lo stesso usato per gli hyperlink (`r:id`), che è
verificato.
