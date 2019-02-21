# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 16:59:51 2019

@author: Asuspc
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

def Partie15():
    'Creation de la partie 15 du protcole de catégorie 2'
    document = docx.Document()


#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

#---------------------------DEFINITIONS DES STYLES
 

    Style(document)

   
#---------------------------------------------------------------ECRITURE
    
    
    #ecriture du premier titre 
    Titre1('15	REGLES RELATIVES A LA PUBLICATION',document)
    
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ces chapitres', document)

    
   # Ecriture du 15.1  
    Titre2('15.1	Communications scientifiques',document)
    
    
    # Ecriture du 15.2  
    Titre2('15.2	Communication des résultats aux participants',document)
    
     # Ecriture du 15.3 
    Titre2('15.3	Cession des données',document)
    
  
 
    document.save("Cat2Partie15.docx")