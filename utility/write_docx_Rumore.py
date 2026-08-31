from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Mm

# VARIABILI GLOBALI E COSE DA IMPORTARE ============
documento_word_template = "/Users/theo/Desktop/P.IVA/Aziende/Ermes/Modelli/Modello_Relazione_RUM.docx"
OUTPUT_DOCUMENT = "/Users/theo/Desktop/Modello_Relazione_RUM_modificato.docx"


# ===============================================
doc = DocxTemplate(documento_word_template)

logo_azienda = "/Users/theo/Desktop/1631305251718.jpeg"
width_logo = 50

context = {
    # Compila i campi vuoti
    "nome_azienda": "PARESA S.R.L.", 
    "indirizzo_azienda": "vicolo malvasia 980 (FC)",
    "note_titolo" : "relazione per la sede di mantova",
    "data_emissione": "16/03/2026",
    "revisione": "rev.01",
    "data_revisione": "16/03/2026",
    "data_scadenza": "16/03/2030",
    "motivo_revisione": "Aggiornamento periodico",    #da modificare
    "img_logo_azienda" : InlineImage(doc, logo_azienda, width=Mm(width_logo)),    #da modificare
    "datore_di_lavoro" : "",    #da modificare 
    "RSPP" : "",    #da modificare
    "medico_competente": "",    #da modificare
    "RLS" : "",    #da modificare
    "delegato_sicurezza": "", #da inserire
    "giornate" : "Nelle giornate 12 e 13 settembre 2026", #modifica con le date giuste,
    #Info generali azienda
    "attivita_azienda": "",    #da modificare
    "processo_produttivo": "", #da modificare
    "gruppo_appartenenza" : "",    #da modificare
    "sede_legale" : "",    #da modificare
    "sede_operativa":"", #da modificare
    "ubicazione_unita_operativa" : "",     #da modificare
    "date_misurazione":["05 dicembre 2025 dalle ore 08:00 alle ore 16:00",
                        "",
                        ""], #da modificare
    "sostanze_ototossiche": "Si", #presenza di sostanze ototossiche o meno
    "misure_attuative_ototossiche": "Si faccia riferimento al documento di valutazione del rischio chimico.",
    "interazione_vib_rum": "Si", #presenza di interazione tra rumore e vibrazione
    "misure_attuative_vib_rum":"Certamente si considerato l’utilizzo di attrezzature elettriche portatili. Vi è dunque trasmissione ossea delle vibrazioni e del rumore all’orecchio medio. Si faccia riferimento alla valutazione del rischio chimico.", 
    "effetti_indesiderati": "Si",
    "misure_attuative_effetti_indesiderati": "Nelle zone/postazioni di lavoro è possibile che gli addetti possano incorrere in tali situazioni. Si consiglia pertanto di utilizzare D.P.I. con grado di protezione SNR come prescritto dalla presente relazione e l’adozione di sistemi alternativi quali segnali oto-acustici.",
    "descrizione_attivita_dettaglio":"boh", #dettaglio della mansione
}

valutazione_context = {
    "base_giornaliera":"",
    "base_settimanale":"",
    "esposizioni_variabili": ""
}

# Dizionari separati per le tabelle
context_tabella_dpi = {
    "tabella_dpi":[
        {"codice_DPI":"",
        "descrizione": "",
        "marca":"",
        "modello":"",
        "snr":"",
        "H":"",
        "L":"",
        "M":"",
        "note":""}
    ],
}

context_tabella_orari = {
    "tabella_orario_lavoro_mansione":[
        {"mansione": "Nome mansione", #da modificare
        "orario_lavoro" : "Lunedì – Venerdì \n 8:00÷12:00 13:00÷17:00" #da modificare
        }
    ],
}

#Pensare se farla prendere da tabella excel
context_tabella_mansioni = {
    "tabella_mansioni": [
        {
            "ID": "", #questo è l'ID_GrOm della tabella mansioni del file excel scheda_gruppi_dpi.xlsx
            "Mansione": " Nome mansione", #Questo è Descrizione_GrOm della tabella mansioni del file excel scheda_gruppi_dpi.xlsx
        }
    ],
}

#pensare se prenderla dai fogli excel
context_tabella_heg = {
    "tabella_HEG":[
        {"gruppo_HEG": "", #Nome del gruppo omogeneo (es: carrellista)
        "numero_scheda":"", # Numero della scheda del gruppo omogeneo (es: 1)
        "lex8h": "",
        "U": "",
        "lexmax":"",
        "peakmax":"",
        "classe_rischio": "", #es: alta, media o bassa
        "vib": "HAV", #WBV, HAV o NO
        "oto": "", #si no
        "imp":"", #si no
        }
    ],
}

# Qui ci vanno solo le classi di rischio medio 
context_tabella_heg_medio = {
    "HEG_med":[
            {"gruppo_HEG": "", #Nome del gruppo omogeneo (es: carrellista)
            "numero_scheda":"", # Numero della scheda del gruppo omogeneo (es: 1)
            "lex8h": "",
            "U": "",
            "lexmax":"",
            "peakmax":"",
            "classe_rischio": "", #es: alta, media o bassa
            },
    ]
}

#qui ci vanno solo le classi con classe di rischio ALTO
context_tabella_heg_alto = {
    "HEG_alto":[
            {"gruppo_HEG": "", #Nome del gruppo omogeneo (es: carrellista)
            "numero_scheda":"", # Numero della scheda del gruppo omogeneo (es: 1)
            "lex8h": "",
            "U": "",
            "lexmax":"",
            "peakmax":"",
            "classe_rischio": "", #es: alta, media o bassa
            },
    ]
}

# Unione di tutti i dizionari
context_completo = context | context_tabella_dpi | context_tabella_orari | context_tabella_mansioni | context_tabella_heg

doc.render(context_completo)
doc.save(OUTPUT_DOCUMENT)
print('Docx scritto')