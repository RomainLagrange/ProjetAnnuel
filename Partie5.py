# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 14:07:53 2019

@author: Julie
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1, Titre2, Titre3, TexteGris, TexteGrisJustif
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)


def Partie5():
    'Creation de la partie 5 du protcole de catégorie 1'
    document = docx.Document()


#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
 
    Style(document)

    
    #ecriture du premier titre 
    Titre1('5	CRITERES D’ELIGIBILITE',document)
    
   # Ecriture du 5.1  
    Titre2('5.1	Critères d’inclusion',document)
    
    # Ecriture du 5.2  
    Titre2('5.2	Critères de non inclusion',document)
    
    # Ecriture du 5.3  
    Titre2('5.3	Faisabilité et modalités de recrutement',document)

   


    
    
    document.save("Partie5.docx")   


