# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 16:52:34 2019

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

def Partie12():
    'Creation de la partie 12 du protcole de catégorie 2'
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
    Titre1('12	CONTROLE ET ASSURANCE DE LA QUALITE',document)
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

    
   # Ecriture du 12.1  
    Titre2('12.1	Consignes pour le recueil des données',document)
    
    
    # Ecriture du 12.2  
    Titre2('12.2	Contrôle de la qualité',document)
    
     # Ecriture du 12.3 
    Titre2('12.3	Gestion des données',document)
    
    #Texte gris justifié
    TexteGrisJustif('Gestion des données pour une étude e-CRF',document)
    
    #Texte gris justifié
    TexteGrisJustif('Gestion des données pour une étude CRF papier',document)
    
    
     # Ecriture du 12.4
    Titre2('12.4	Audits et inspections',document)
    
 
    document.save("Cat2Partie12.docx")