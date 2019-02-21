# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 14:05:20 2019

@author: Julie
"""

import docx
import StyleProt1
from StyleProt1 import Style, TexteGris, Titre1, Titre2,Titre3
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    StyleProt1.Titre1('num + texte du protocole',document)
#    StyleProt1.Titre2('num + texte du protocole',document)
#    StyleProt1.Titre3('numero','texte',document)

def Partie4():
    'Creation de la partie 4 du protcole de catégorie 1'
    document = docx.Document()

    from docx.oxml.ns import nsdecls
    from docx.oxml import parse_xml

#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
 
    StyleProt1.Style(document)

    StyleProt1.Titre1('4	CONCEPTION DE LA RECHERCHE',document)
    
    #Texte sur fond gris  
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ce chapitre',document)

   # Ecriture du 4.1  
    StyleProt1.Titre2('4.1	Schéma de la recherche',document)
    
    # Ecriture du 4.2  
    StyleProt1.Titre2('4.2	Méthode pour la randomisation',document)
    
    document.save("Partie4.docx")   

