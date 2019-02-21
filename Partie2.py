# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 13:34:49 2019

@author: Julie
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1,Titre2, Titre3,TexteGris, TexteGrisJustif
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)

def Partie2():
    'Creation de la partie 2 du protcole de catégorie 1'
    document = docx.Document()



#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
 
    Style(document)

    Titre1('2	OBJECTIFS DE LA RECHERCHE',document)
    
    #Texte sur fond gris   
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ce chapitre', document)

    
   # Ecriture du 2.1  
    Titre2('2.1	Objectif principal',document)
    
    # Ecriture du 2.2  
    Titre2('2.1	Objectifs secondaires',document)
    
    document.save("Partie2.docx")   


