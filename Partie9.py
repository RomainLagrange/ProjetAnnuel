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
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Evénement indésirable ')
    run1.style='Paragraphe'
    run1.font.bold= True
    run2=p.add_run('(article R1123-46 du code de la santé publique)\nToute manifestation nocive survenant chez une personne qui se prête à une recherche impliquant la personne humaine, que cette manifestation soit liée ou non à la recherche ou au produit sur lequel porte cette recherche.')
    run2.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Effet indésirable ')
    run1.style='Paragraphe'
    run1.font.bold= True
    run2=p.add_run('(article R1123-46 du code de la santé publique)\nEvénement indésirable survenant chez une personne qui se prête à une recherche impliquant la personne humaine, lorsque cet événement est lié à la recherche ou au produit sur lequel porte cette recherche.')
    run2.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Evénement ou effet indésirable grave ')
    run1.style='Paragraphe'
    run1.font.bold= True
    run2=p.add_run('(article R1123-46 du code de la santé publique et guide ICH E2B)\nTout événement ou effet indésirable qui :')
    run2.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('entraîne la mort,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('met en danger la vie de la personne qui se prête à la recherche,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('nécessite une hospitalisation ou la prolongation de l\'hospitalisation,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('provoque une incapacité ou un handicap important ou durable,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('se traduit par une anomalie ou une malformation congénitale,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('ou tout événement considéré comme médicalement grave,')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('et s\'agissant du médicament, quelle que soit la dose administrée.')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('L’expression « mise en jeu du pronostic vital » est réservée à une menace vitale immédiate, au moment de l’événement indésirable.')
    run1.style='Paragraphe'
    
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Effet indésirable inattendu ')
    run1.style='Paragraphe'
    run1.font.bold= True
    run2=p.add_run('(article R1123-46 du code de la santé publique)')
    run2.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('Pour les recherches portant sur un médicament, effet indésirable inattendu : tout effet indésirable du produit dont la nature, la sévérité, la fréquence ou l\'évolution ne concorde pas avec les informations de référence sur la sécurité mentionnées dans le résumé des caractéristiques du produit ou dans la brochure pour l’investigateur lorsque le produit n’est pas autorisé.')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('Pour les recherches portant sur un dispositif médical ou sur un dispositif médical de diagnostic in vitro, effet indésirable inattendu : tout effet du dispositif dont la nature, la sévérité ou l’évolution ne concordent pas avec les informations de référence figurant respectivement dans la notice d’instruction ou dans la notice d’utilisation du dispositif lorsque celui-ci fait l’objet d’un marquage CE, et dans le protocole ou la brochure pour l’investigateur lorsqu’il ne fait pas l’objet d’un tel marquage.')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.line_spacing_rule = WD_LINE_SPACING.SINGLE
    p.style='List Bullet 2'
    run1=p.add_run('Pour les autres recherches impliquant la personne humaine, effet indésirable inattendu : tout effet indésirable dont la nature, la sévérité ou l’évolution ne concorde pas avec les informations relatives aux produits, actes pratiqués et méthodes utilisées au cours de la recherche')
    run1.style='Paragraphe'
    

 
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
    
    
    