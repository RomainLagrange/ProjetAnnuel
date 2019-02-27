# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 10:22:29 2019

@author: Asuspc
"""

# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 10:13:09 2019

@author: Asuspc
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Feb 18 11:51:47 2019

@author: Asuspc
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1, Titre2, Titre3, TexteGris, TexteGrisJustif
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)

def Partie10(document):
    'Creation de la partie 10 du protcole de catégorie 3'
   # document = docx.Document()


#   Marge de la page
#    sections = document.sections
#    for section in sections:
#        section.top_margin = Cm(2)
#        section.bottom_margin = Cm(2)
#        section.left_margin = Cm(2)
#        section.right_margin = Cm(2)

#---------------------------DEFINITIONS DES STYLES
 

   # Style(document)


#    
#---------------------------------------------------------------ECRITURE
    
    
    #ecriture du premier titre 
    Titre1('10	CONTROLE ET ASSURANCE DE LA QUALITE',document)
    
    
   # Ecriture du 10.1  
    Titre2('10.1	Consignes pour le recueil des données',document)
    
    # Ecriture du 10.2  
    Titre2('10.2	Suivi de la recherche',document)
    
    
    # Ecriture du 10.3  
    Titre2('10.3	Contrôle de Qualité',document)
    
    TexteGris('a completer uniquement si applicable', document)

    
    #Ecriture du titre 10.4
    Titre2('10.4	Gestion des données',document)

#----------------gestion des données ----------
 #style    
    styles = document.styles
    styleBackgroundGrey = styles.add_style('CRF', WD_STYLE_TYPE.CHARACTER)
    styleBackgroundGrey.base_style = styles['No Spacing']
    fontBackgroundGrey = styleBackgroundGrey.font
    fontBackgroundGrey.name = 'Times New Roman'
    fontBackgroundGrey.size = docx.shared.Pt(11)
    fontBackgroundGrey.bold = True
    
    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
    table = document.add_table(rows = 1, cols = 1)
    row = table.rows[0].cells
    para_text ='Gestion des données pour une étude e-CRF'
    cell = row[0]
    pt = cell.paragraphs[0]
    t = pt.text = ''
    p = pt.add_run(para_text)
    cell._tc.get_or_add_tcPr().append(shading_elm)
    p.style='CRF'
    pt.alignment=WD_ALIGN_PARAGRAPH.CENTER
    
    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
    table = document.add_table(rows = 1, cols = 1)
    row = table.rows[0].cells
    para_text ='Gestion des données pour une étude CRF papier'
    cell = row[0]
    pt = cell.paragraphs[0]
    t = pt.text = ''
    p = pt.add_run(para_text)
    cell._tc.get_or_add_tcPr().append(shading_elm)
    p.style='CRF'
    pt.alignment=WD_ALIGN_PARAGRAPH.CENTER
#-----------------------------------------------------------

    #Ecriture du titre 10.5
    Titre2('10.5	Audit et inspection',document)


    
    #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
  
  #  document.save("Cat3Part10.docx")   