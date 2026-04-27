from docxtpl import DocxTemplate

doc = DocxTemplate("/Users/theo/Downloads/test.docx")

context = {
    "azienda": "PARESA S.R.L."
}

doc.render(context)
doc.save("/Users/theo/Downloads/output.docx")
print('Docx scritto')