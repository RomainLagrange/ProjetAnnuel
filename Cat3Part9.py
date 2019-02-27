# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 10:20:01 2019

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

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)

def Partie9(document):
    'Creation de la partie 9 du protcole de catégorie 3'
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
    Titre1('9	DROITS D’ACCES AUX DONNEES ET AUX DOCUMENTS SOURCE',document)
    
    
   # Ecriture du 9.1  
    Titre2('9.1	Accès aux données',document)
    
    # Ecriture du 9.2  
    Titre2('9.2	Données sources',document)
    
    #AJOUTER TABLEAU
    
    # Ecriture du 9.3  
    Titre2('9.3	Confidentialité des données',document)
    
    
    #Ecriture du titre 9.4
    Titre2('9.4	Origine et nature des données recueillies :',document)


    #Ecriture du titre 9.5
    Titre2('9.5	Mode de circulation des données',document)

    
    
    #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
  
  #  document.save("Cat3Part9.docx")   