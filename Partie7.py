# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 14:01:09 2019

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

def Partie7():
    'Creation de la partie 7 du protcole de catégorie 1'
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
    Titre1('7	TRAITEMENT(S) / STRATEGIE(S) / PROCEDURES DE LA RECHERCHE',document)
    
    
   # Ecriture du 7.1  
    Titre2('7.1	Traitement / stratégie / procédure expérimental(e)',document)
    
    #Ecriture des trois textes gris justifiés
    TexteGrisJustif('Pour un traitement de type médicament',document)
    TexteGrisJustif('Pour un placebo',document)
    TexteGrisJustif('Pour un traitement de type dispositif médical (DM)',document)
    TexteGrisJustif('Pour une stratégie/procédure',document)
    
    # Ecriture du 7.2
    Titre2('7.2	Traitement / Stratégie / Procédure de comparaison',document)
    
    #Ecriture des deux textes gris justifiés
    TexteGrisJustif('Pour un traitement de type dispositif médical (DM)',document)
    TexteGrisJustif('Pour une stratégie/procédure',document)
    
    # Ecriture du 7.3
    Titre2('7.3	Circuit des produits',document)
    
    #Texte gris
    TexteGris('prendre contact avec la pharmacie du chu de poitiers \n pour aide a la redaction de ces chapitres',document)
    #Ecriture du 7.3.1
    Titre3('7.3.1','Libération et distribution des produits',document)
    #Ecriture du 7.3.2
    Titre3('7.3.2','Fourniture des produits',document)
    #Ecriture du 7.3.3
    Titre3('7.3.3','Conditionnement des produits',document)
    #Ecriture du 7.3.4
    Titre3('7.3.4','Etiquetage des produits',document)
    #Ecriture du 7.3.5
    Titre3('7.3.5','Expédition et gestion des produits',document)
    #Ecriture du 7.3.6
    StyleProt1.Titre3('7.3.6','Dispensation des produits et observance',document)
    #Ecriture du 7.3.7
    Titre3('7.3.7','Stockage ',document)
    #Ecriture du 7.3.8
    Titre3('7.3.8','Retour et destruction des produits non utilisés',document)
    
    # Ecriture du 7.4
    Titre2('7.4	Insu',document)
    
    #Texte gris centré
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction de ce chapitre', document)

    #Ecriture du 7.4.1
    Titre3('7.4.1','Organisation de l’insu',document)
    #Ecriture du 7.4.2
    Titre3('7.4.2','Levée de l’insu',document)
   
    # Ecriture du 7.5
    Titre2('7.5	Réductions et ajustements de dose',document)
    
    #Ecriture du 7.5.1
    Titre3('7.5.1','Réductions/ajustements de doses',document)
    #Ecriture du 7.5.2
    Titre3('7.5.2','Réductions de dose pour les toxicités hématologiques',document)
    #Ecriture du 7.5.3
    Titre3('7.5.3','Réductions de dose pour les toxicités non hématologiques',document)
    
    document.save("Partie7.docx")   
    
    