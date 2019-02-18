# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 13:42:50 2019

@author: Julie
"""

import docx
from StyleProt1 import Style
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#revoir titre1 et texte encardé gris

def Page13():
    'Creation de la page 13 du protcole de catégorie 1'
    document = docx.Document()
    styles = document.styles

    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml

#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
 
    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w'))) #CREER LE FOND GRIS

    
#   definition du style Titre1 --> AJOUTER LA BORDURE EN BAS
    styleTitre1 = styles.add_style('Titre1', WD_STYLE_TYPE.PARAGRAPH, WD_ALIGN_PARAGRAPH.CENTER)
    styleTitre1.base_style = styles['Heading1']
    fontTitre1 = styleTitre1.font
    fontTitre1.name = 'Times New Roman' #police
    fontTitre1.size = docx.shared.Pt(12) #taille
    fontTitre1.all_caps = True #toujours en majuscule
    fontTitre1.bold= True #en gras
    fontTitre1.color.rgb = RGBColor(0x0,0x70,0xC0) #couleur bleu, en base 16
    #ecriture du premier titre 
    paragraph=document.add_paragraph('3	CRITERES DE JUGEMENT\n', style='Titre1') #titre
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER #centrer
                

    
    #Definition du Titre2, correspond par exemple au 1.1 ou 1.2
    styleTitre2 = styles.add_style('Titre2', WD_STYLE_TYPE.PARAGRAPH)
    styleTitre2.base_style = styles['Heading2']
    fontTitre2 = styleTitre2.font
    fontTitre2.name = 'Times New Roman'
    fontTitre2.size = docx.shared.Pt(14)
    fontTitre2.bold= True
    fontTitre2.color.rgb = RGBColor(0x0,0x0,0x0)
    styleTitre2.paragraph_format.left_indent = Inches(0.59)
    
   # Ecriture du 3.1 
    document.add_paragraph('3.1	Critère d’évaluation principal\n', style='Titre2') 
    
    # Ecriture du 3.2  
    document.add_paragraph('3.2	Critères d’évaluation secondaires\n', style='Titre2') 
    
#    
#    tableTEST = document.add_table(rows = 1, cols = 1)
#    tableTEST.style = "Table Grid"
#    row = tableTEST.rows[0]
#    cell = row.cells[0]
#    cell.text = "text"

    
    
    document.save("page13.docx")   


