# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 16:28:06 2019

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

def Partie5():
    'Creation de la partie 5 du protcole de catégorie 2'
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


#    
#---------------------------------------------------------------ECRITURE
    
    
    #ecriture du premier titre 
    Titre1('5	DEROULEMENT DE LA RECHERCHE',document)
    
    
   # Ecriture du 5.1  
    Titre2('5.1	Schéma de la recherche',document)
    
    # Ecriture du 5.2  
    Titre2('5.2	Méthode pour la randomisation',document)
    
    
    # Ecriture du 5.3  
    Titre2('5.3	Calendrier de la recherche',document)
    
    #Ecriture du titre 5.4
    Titre2('5.4	Tableau récapitulatif du suivi d’un participant à la recherche',document)


    #Ecriture du titre 5.5
    Titre2('5.5	Visites de pré-inclusion / inclusion = Visite V0',document)

    #Ecriture du titre5.5.1
    Titre3('5.5.1','Recueil du consentement',document)

    
    #Ecriture du titre5.5.2
    Titre3('5.5.2','Déroulement de la visite',document)

    
    #Ecriture du titre 5.6
    Titre2('5.6	Visite de randomisation = Visite (Vx, ou Jx, ou Mx…)',document)
    
    #Ecriture du titre5.6.1
    Titre3('5.6.1','Description des examens',document)
    
    #Ecriture du titre5.6.2
    Titre3('5.6.2','Randomisation du patient',document)
    
    #Ecriture du titre 5.7
    Titre2('5.7	Visites de suivi = visite (Vx, ou Jx ou Sx ou Mx…)',document)

	
    #Ecriture du titre5.7.1
    Titre3('5.7.1','Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)


    #Ecriture du titre5.7.2
    Titre3('5.7.2	Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)


    #Ecriture du titre 5.8
    Titre2('5.8	Visite de fin de la recherche',document)
    
    #Ecriture du titre 5.9
    Titre2('5.9	Règles d’arrêt de la participation d’une personne à la recherche',document)
    
    
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ce chapitre', document)

    
    #Ecriture du titre5.9.1
    Titre3('5.9.1','Arrêt de participation définitif ou temporaire d’un patient dans l’étude',document)
    
    #Ecriture du titre5.9.2
    Titre3('5.9.2	Modalités de remplacement des patients exclus, le cas échéant',document)
    
    #Ecriture du titre5.9.3
    Titre3('5.9.3','Modalités et calendrier de recueil pour ces données',document)
    
    #Ecriture du titre 5.9.4
    Titre3('5.9.4','Modalités de suivi de ces personnes',document)
    
    #Ecriture du titre 5.10
    Titre2('5.10	Contraintes liées à la recherche et indemnisation éventuelle des participants',document)
    
     #Ecriture du titre 5.11
    Titre2('5.11	Collection d’échantillons biologiques',document)
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

    #Ecriture du titre5.11.1
    Titre3('5.11.1','Objectifs',document)
    
    #Ecriture du titre5.11.2
    Titre3('5.11.2','Description de(s) (la) collection(s) ',document)
    
    #Ecriture du titre5.11.3
    Titre3('5.11.3','Conservation',document)
    
    #Ecriture du titre5.11.4
    Titre3('5.11.4','Devenir de la collection',document)    
    
     #Ecriture du titre 5.12
    Titre2('5.12	Arrêt d’une partie ou de la totalité de la recherche',document)
    
    
    document.save("Cat2Partie5.docx")   