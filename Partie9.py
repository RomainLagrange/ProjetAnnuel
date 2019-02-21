# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 14:54:04 2019

@author: Asuspc
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1, Titre2, Titre3, TexteGris, TexteGrisJustif
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document) --> écrire en minuscule !!!
#    TexteGrisJustif(texte,document)

def Partie9(document):
    'Creation de la partie 9 du protcole de catégorie 1'
  #  document = docx.Document()


#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

#---------------------------DEFINITIONS DES STYLES
 

#    Style(document)


#    
#---------------------------------------------------------------ECRITURE
    
    
     #ecriture du premier titre 
    Titre1('9	EVALUATION DE LA SECURITE',document)
    
     #Texte gris centré
    TexteGris('prendre contact avec l\'unite de vigilance des essais cliniques \n pour aide a la redaction de ce chapitre',document)
    
    
    #Ecriture du 9.1  
    Titre2('9.1	Définitions',document)
    
    #Ecriture du 9.2  
    Titre2('9.2	Description des événements indésirables graves attendus',document)
    
    #Ecriture du 9.3  
    Titre2('9.3	Conduite à tenir par l’investigateur en cas d’événement indésirable, de fait nouveau ou de grossesse',document)
    
    #Ecriture du titre 9.3.1
    Titre3('9.3.1','Recueil des événements indésirables (EvI)',document)
    
    #Ecriture du titre 9.3.2
    Titre3('9.3.2','Déclaration des événements indésirables graves (EvIG), des événements indésirables d’intérêt et des faits nouveaux ',document)
    
    #Ecriture du titre 9.3.3
    Titre3('9.3.3','Déclaration des grossesses',document)
    
    #Ecriture du titre 9.3.4
    Titre3('9.3.4','Tableau récapitulatif du circuit de déclaration par type d’événement',document)
    
    #Ecriture du 9.4
    Titre2('9.4 Déclaration par le promoteur des effets indésirables graves inattendus, des faits nouveaux et autres évènements',document)
    
    #Ecriture du 9.5
    Titre2('9.5	Essai chez un volontaire sain',document)
    
    #Ecriture du 9.6
    Titre2('9.6 Rapport annuel de sécurité',document)
    
        #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
    #document.save("Partie7.docx") 
  #  document.save("Partie9.docx")
    
    
    