# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 13:34:49 2019

@author: Julie
"""

import docx
from StyleProt1 import Style
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#revoir titre1 et texte encardé gris

def Page12():
    'Creation de la page 12 du protcole de catégorie 1'
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
    paragraph=document.add_paragraph('2	OBJECTIFS DE LA RECHERCHE\n', style='Titre1') #titre
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER #centrer
                

    #Definition style texte surligné en gris   --> SUPPRIMER ESPACE EN BAS
    styles = document.styles
    styleBackgroundGrey = styles.add_style('BackgroundGrey', WD_STYLE_TYPE.CHARACTER)
    styleBackgroundGrey.base_style = styles['No Spacing']
    fontBackgroundGrey = styleBackgroundGrey.font
    fontBackgroundGrey.name = 'Times New Roman'
    fontBackgroundGrey.size = docx.shared.Pt(11)
    fontBackgroundGrey.bold = True
    fontBackgroundGrey.small_caps = True
    
    #Texte sur fond gris   
    table = document.add_table(rows = 1, cols = 1)
    row = table.rows[0].cells
    para_text = 'prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ce chapitre'
    cell = row[0]
    pt = cell.paragraphs[0]
    t = pt.text = ''
    p = pt.add_run(para_text)
    cell._tc.get_or_add_tcPr().append(shading_elm)
    p.style='BackgroundGrey'
    pt.alignment=WD_ALIGN_PARAGRAPH.CENTER
    
    
    #Definition du Titre2, correspond par exemple au 1.1 ou 1.2
    styleTitre2 = styles.add_style('Titre2', WD_STYLE_TYPE.PARAGRAPH)
    styleTitre2.base_style = styles['Heading2']
    fontTitre2 = styleTitre2.font
    fontTitre2.name = 'Times New Roman'
    fontTitre2.size = docx.shared.Pt(14)
    fontTitre2.bold= True
    fontTitre2.color.rgb = RGBColor(0x0,0x0,0x0)
    styleTitre2.paragraph_format.left_indent = Inches(0.59)
    
   # Ecriture du 2.1  
    document.add_paragraph('2.1	Objectif principal\n', style='Titre2') 
    
    # Ecriture du 2.2  
    document.add_paragraph('2.1	Objectifs secondaires\n', style='Titre2') 
    
    document.save("page12.docx")   


