from docxtpl import DocxTemplate
from docxtpl import InlineImage
from docx.shared import Mm

# VARIABILI GLOBALI E COSE DA IMPORTARE ============
documento_word_template = "/Users/theo/Desktop/P.IVA/Ermes/Modelli/Modello_Relazione_RUM.docx"

logo_azienda = ""
width_logo = 50


# ===============================================
doc = DocxTemplate(documento_word_template)

context = {
    # Compila i campi vuoti
    "nome_azienda": "PARESA S.R.L.", 
    "indirizzo_azienda": "vicolo malvasia 980 (FC) dio can",
    "note_titolo" : "relazione per la sede di mantova",
    "revisione": "rev.01",
    "data_revisione": "16/03/2026",
    "data_scadenza": "16/03/2030",
    "motivo_revisione": "Aggiornamento periodico",    #da modificare
    "img_logo_azienda" : InlineImage(doc, logo_azienda, width=Mm(width_logo)),    #da modificare
    "datore_di_lavoro" : "",    #da modificare 
    "RSPP" : "",    #da modificare
    "medico_competente": "",    #da modificare
    "RLS" : "",    #da modificare
    "giornate" : "Nelle giornate 12 e 13 settembre 2026", #modifica con le date giuste,
    #Info generali azienda
    "attivita_azienda": "",    #da modificare
    "gruppo_appartenenza" : "",    #da modificare
    "sede_legale" : "",    #da modificare
    "ubicazione_unita_operativa" : "",     #da modificare
    "date_misurazione":["05 dicembre 2025 dalle ore 08:00 alle ore 16:00",
                        "",
                        ""], #da modificare
    "strumentazione": "FUSION+DUO+FUSION+WED (01dB) / CAL 21 (01dB)",
    "condizioni_meteo": "Le condizioni meteo non hanno inficiato le misurazioni",
    "sostanze_ototossiche": "Si", #presenza di sostanze ototossiche o meno
    "misure_attuative_ototossiche": "Si faccia riferimento al documento di valutazione del rischio chimico.",
    "interazione_vib_rum": "Si", #presenza di interazione tra rumore e vibrazione
    "misure_attuative_vib_rum":"Certamente si considerato l’utilizzo di attrezzature elettriche portatili. Vi è dunque trasmissione ossea delle vibrazioni e del rumore all’orecchio medio. Si faccia riferimento alla valutazione del rischio chimico.", 
    "effetti_indesiderati": "Si",
    "misure_attuative_effetto_indesiderati": "Nelle zone/postazioni di lavoro è possibile che gli addetti possano incorrere in tali situazioni. Si consiglia pertanto di utilizzare D.P.I. con grado di protezione SNR come prescritto dalla presente relazione e l’adozione di sistemi alternativi quali segnali oto-acustici.",
    "descrizione_attivita_dettaglio":"", #dettaglio della mansione
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
    "tabella_mansioni_numero_lavoratori": [
        {
            "ID": "",
            "Mansione": "",
            "N_lavoratori" : ""
        }
    ],
}

#pensare se prenderla dai fogli excel
context_tabella_heg = {
    "tabella_HEG":[
        {"gruppo_HEG": "", #Nome del gruppo omogeneo (es: carrellista)
        "numero_scheda":"", # Numero della scheda del gruppo omogeneo (es: 1)
        "codice_HEG": "" ,#codice del gruppo omogeneo (es: M01)
        "parametro_riferimento": "Lex,8h",
        "lex8h": "",
        "incertezza": "",
        "lexmax":"",
        "peakmax":"",
        "classe_rischio": "", #es: alta, media o bassa
        "esposizione_vibrazioni": "HAV", #WBV, HAV o NO
        "esposizione_ototossici": "", #si no
        "rumori_impulsivi":"", #si no
        }
    ],
}

# Unione di tutti i dizionari
context_completo = context | context_tabella_dpi | context_tabella_orari | context_tabella_mansioni | context_tabella_heg

doc.render(context_completo)
doc.save("/Users/theo/Downloads/Modello_Relazione_RUM.docx")
print('Docx scritto')