# -*- coding: utf-8 -*-
"""
Created on Mon Feb 18 11:51:47 2019

@author: Asuspc
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1, Titre2, Titre3, TexteGris
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    StyleProt1.Titre1('num + texte du protocole',document)
#    StyleProt1.Titre2('num + texte du protocole',document)
#    StyleProt1.Titre3('numero','texte',document)

def Partie6():
    'Creation de la partie 6 du protcole de catégorie 1'
    document = docx.Document()


#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)

#---------------------------DEFINITIONS DES STYLES
 

    StyleProt1.Style(document)


#    
#---------------------------------------------------------------ECRITURE
    
    #AJOUTER SCHEMA DEROULEMENT
    
    #ecriture du premier titre 
    StyleProt1.Titre1('6	DEROULEMENT DE LA RECHERCHE',document)
    
    
   # Ecriture du 6.1  
    StyleProt1.Titre2('6.1	Calendrier de la recherche',document)
    
    # Ecriture du 6.2  
    StyleProt1.Titre2('6.2	Tableau récapitulatif du suivi d’un participant à la recherche',document)
    
    #AJOUTER TABLEAU
    
    # Ecriture du 6.3  
    StyleProt1.Titre2('6.3	Visites de pré-inclusion / inclusion = Visite V0',document)
    
    
    
    #Ecriture du titre6.3.1
    StyleProt1.Titre3('6.3.1','Recueil du consentement',document)
   
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)
    

    #TEXTE
    
    #Ecriture du titre6.3.2
    StyleProt1.Titre3('6.3.2','Déroulement de la visite',document)

    
    #Ecriture du titre 6.4
    StyleProt1.Titre2('6.4	Visite de randomisation = Visite (Vx, ou Jx, ou Mx…)',document)

    #Ecriture du titre6.4.1
    StyleProt1.Titre3('6.4.1','Description des examens',document)

    
    #Ecriture du titre6.4.2
    StyleProt1.Titre3('6.4.2','Randomisation du patient',document)



    #Ecriture du titre 6.5
    StyleProt1.Titre2('6.5	Visites de suivi = visite (Vx, ou Jx ou Sx ou Mx…)',document)

    #Ecriture du titre6.5.1
    StyleProt1.Titre3('6.5.1','Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)

    
    #Ecriture du titre6.5.2
    StyleProt1.Titre3('6.5.2','Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)

    
    #Ecriture du titre 6.6
    StyleProt1.Titre2('6.6	Visite de fin de la recherche',document)
    
    #Ecriture du titre 6.7
    StyleProt1.Titre2('6.7	Règles d’arrêt de la participation d’une personne à la recherche',document)

    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

	
    #Ecriture du titre6.7.1
    StyleProt1.Titre3('6.7.1','Arrêt de participation définitif ou temporaire d’un patient dans l’étude)',document)


    #Ecriture du titre6.7.2
    StyleProt1.Titre3('6.7.2','Modalités de remplacement des patients exclus, le cas échéant',document)

    
    #Ecriture du titre6.7.3
    StyleProt1.Titre3('6.7.3','Modalités et calendrier de recueil pour ces données',document)

    
    #Ecriture du titre6.7.4
    StyleProt1.Titre3('6.7.4','Modalités de suivi de ces personnes',document)

    #Ecriture du titre 6.8
    StyleProt1.Titre2('6.8	Contraintes liées à la recherche et indemnisation éventuelle des participants',document)
    
    #Ecriture du titre 6.9
    StyleProt1.Titre2('6.9	Collection d’échantillons biologiques',document)
    
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

    
    #Ecriture du titre6.9.1
    StyleProt1.Titre3('6.9.1','Objectifs',document)
    
    #Ecriture du titre6.9.2
    StyleProt1.Titre3('6.9.2','Description de(s) (la) collection(s) ',document)
    
    #Ecriture du titre6.9.3
    StyleProt1.Titre3('6.9.3','Conservation',document)
    
    #Ecriture du titre6.9.4
    StyleProt1.Titre3('6.9.4','Devenir de la collection',document)
    
    #Ecriture du titre 6.10
    StyleProt1.Titre2('6.10	Arrêt d’une partie ou de la totalité de la recherche',document)
    

    
    
    document.save("Partie6.docx")   