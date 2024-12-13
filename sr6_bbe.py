import sys
import csv
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches

# TODO: liste à puces

def openDocument(filename):
    f = open(input_filename, 'rb')
    document = Document(f)
    f.close()
    return document

def delete_paragraph(paragraph):
    p = paragraph._element
    p.getparent().remove(p)
    p._p = p._element = None

def shadowtalk(paragraph, document):
    paragraph.insert_paragraph_before(paragraph.text.replace("> ",""),document.styles["SR6 - Shadowtalk"])
    delete_paragraph(paragraph)


def shadowtalks(document):
    st_begin = False
    for paragraph in document.paragraphs:
        if paragraph.text.startswith(">"):
            st_begin = not st_begin
            shadowtalk(paragraph, document)
        else:
            if st_begin:
                shadowtalk(paragraph, document)

def encart(paragraph, text, document, style):
    paragraph.insert_paragraph_before(text,style)
    delete_paragraph(paragraph)

def encarts(document):
    encart_begin = False
    for paragraph in document.paragraphs:
        if paragraph.text.startswith("Encart:"):
            print(paragraph.text)
            encart_begin = not encart_begin
            encart(paragraph,paragraph.text.replace("Encart:",""), document, document.styles['SR6 - Encart titre 1'])
        else:
            if encart_begin:
                encart(paragraph, paragraph.text, document, document.styles['SR6 - Encart texte'])
        if paragraph.text.startswith("Fin Encart"):
            #encart(paragraph, " ", document, document.styles['SR6 - Encart texte'])
            print(paragraph.text)
            encart_begin = not encart_begin

    for paragraph in document.paragraphs:
        if paragraph.text.startswith("Fin Encart"):
            delete_paragraph(paragraph)

def map_gdoc_styles_to_sr6(document):
    for paragraph in document.paragraphs:
        if paragraph.style.name in index.keys():
            paragraph.style.name = index[paragraph.style.name]



if len(sys.argv) != 3:
    raise ValueError("Invalid number of parameters")

input_filename = sys.argv[1]
output_filename = sys.argv[2]

index = {
	"LO-normal": "SR6 - Texte",
    "normal": "SR6 - Texte",
    "normal1": "SR6 - Texte",
	"Heading 1": "SR6 - Titre 1",
	"Heading 2": "SR6 - Titre 2",
	"Heading 3": "SR6 - Titre 3",
    "Heading 4": "SR6 - Titre 4",
    "Title": "SR6 - Titre chapitre",
}

document = openDocument(input_filename)
document.styles.add_style('SR5 - Maquette', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Encart texte', WD_STYLE_TYPE.PARAGRAPH).paragraph_format.left_indent = Inches(0.5)
document.styles.add_style('SR6 - Encart titre 1', WD_STYLE_TYPE.PARAGRAPH).paragraph_format.left_indent = Inches(0.5)
document.styles.add_style('SR6 - Encart texte (liste)', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Encart titre 2', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Stat décroché armes', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Stat décroché (bold :)', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Stat titre 1', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Stat titre 2', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Table en-tête', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Table texte', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Shadowtalk', WD_STYLE_TYPE.PARAGRAPH)
document.styles.add_style('SR6 - Posté par', WD_STYLE_TYPE.PARAGRAPH)
document.styles['SR6 - Shadowtalk'].font.italic = True

map_gdoc_styles_to_sr6(document)

for paragraph in document.paragraphs:
    if paragraph.text.startswith("Posté par"):
        paragraph.insert_paragraph_before(paragraph.text,document.styles["SR6 - Posté par"])
        delete_paragraph(paragraph)

shadowtalks(document)
encarts(document)
document.save(output_filename)
