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

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)

def Partie6(document):
    'Creation de la partie 6 du protcole de catégorie 3'
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
    Titre1('6	DEROULEMENT DE LA RECHERCHE',document)
    
    
   # Ecriture du 6.1  
    Titre2('6.1	Calendrier de la recherche',document)
    
    # Ecriture du 6.2  
    Titre2('6.2	Tableau récapitulatif du suivi participant',document)
    
    #AJOUTER TABLEAU
    
    # Ecriture du 6.3  
    Titre2('6.3	Visites de pré-inclusion / inclusion = Visite V06.3	Information des personnes concernées',document)
    
    
    #Ecriture du titre 6.4
    Titre2('6.4	Visites de suivi',document)


    #Ecriture du titre 6.5
    Titre2('6.5	Visite de fin de la recherche',document)

    
    #Ecriture du titre 6.6
    Titre2('6.6	Collection d’échantillons biologiques',document)

    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

	
    #Ecriture du titre6.6.1
    Titre3('6.6.1','Objectifs',document)


    #Ecriture du titre6.6.2
    Titre3('6.6.2','Description de(s) la collection(s)',document)

    
    #Ecriture du titre6.6.3
    Titre3('6.6.3','Conservation',document)

    
    #Ecriture du titre6.6.4
    Titre3('6.6.4','Devenir de la collection',document)


    
    #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
  
  #  document.save("Cat3Part6.docx")   