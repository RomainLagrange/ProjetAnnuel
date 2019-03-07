# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 14:01:09 2019

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

def Partie7(document):
    'Creation de la partie 7 du protcole de catégorie 1'
 #   document = docx.Document()


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
    Titre1('7	TRAITEMENT(S) / STRATEGIE(S) / PROCEDURES DE LA RECHERCHE',document)
    
    
   # Ecriture du 7.1  
    Titre2('7.1	Traitement / stratégie / procédure expérimental(e)',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Médicament expérimental : médicament expérimenté ou utilisé comme référence, y compris en tant que placebo, lors d’un essai clinique (article 2 du règlement européen).')
    run1.style='Paragraphe'
    
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
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Chaque patient se verra remettre ')
    run1.style='Paragraphe'
    run2=p.add_run('XXXX boîtes, flacons ')
    run2.style='Paragraphe'
    run2.font.italic= True
    run3=p.add_run('pour la totalité de la durée du traitement.')
    run3.style='Paragraphe'
    
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
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('La pharmacie est destinataire de la liste de randomisation.')
    run1.style='Paragraphe'
    
    #Ecriture du 7.4.2
    Titre3('7.4.2','Levée de l’insu',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('En situation d’urgence médicale nécessitant une levée d’aveugle, la procédure DRC-VIGI-003 du promoteur sera suivie. ')
    run1.style='Paragraphe'
   
    # Ecriture du 7.5
    Titre2('7.5	Réductions et ajustements de dose',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Les retards et modifications de dose seront effectués selon les recommandations suivantes. L’évaluation des toxicités se fera selon la classification CTCAE (Common Terminology Criteria for Adverse Events) du NCI (National Cancer Institute).')
    run1.style='Paragraphe'
    
    #Ecriture du 7.5.1
    Titre3('7.5.1','Réductions/ajustements de doses',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Les tableaux suivants résument les modifications de dose du médicament 1, médicament 2,… pour gérer d’éventuelles toxicités.')
    run1.style='Paragraphe'
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Tableau 1 : diminutions de dose pour ')
    run1.style='Paragraphe'
    run2=p.add_run('médicament 1')
    run2.style='Paragraphe'
    run2.font.italic= True
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Tableau 2 : diminutions de dose pour ')
    run1.style='Paragraphe'
    run2=p.add_run('médicament 2')
    run2.style='Paragraphe'
    run2.font.italic= True
    
    
    #Ecriture du 7.5.2
    Titre3('7.5.2','Réductions de dose pour les toxicités hématologiques',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Les tableaux suivants décrivent les recommandations de réduction de dose pour ')
    run1.style='Paragraphe'
    run2=p.add_run('médicament 1 / médicament… ')
    run2.style='Paragraphe'
    run2.font.italic= True
    run3=p.add_run('en cas de thrombopénie, neutropénie et anémie.')
    run3.style='Paragraphe'

    #Ecriture du 7.5.3
    Titre3('7.5.3','Réductions de dose pour les toxicités non hématologiques',document)
    
    p=document.add_paragraph()
    p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run1=p.add_run('Les lignes directrices d’ajustement de dose pour ')
    run1.style='Paragraphe'
    run2=p.add_run('médicament 1 / médicament 2… ')
    run2.style='Paragraphe'
    run2.font.italic= True
    run3=p.add_run('en cas de toxicités non hématologiques sont résumées comme suit :')
    run3.style='Paragraphe'
    
    
        #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
    #document.save("Partie7.docx")   
    
    