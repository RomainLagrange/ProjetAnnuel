# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 18:17:55 2019

@author: Marion
"""

import pandas as pd
import docx
from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
#from docx.shared import RGBColor

#docmuents du cpp pour les medoc cat 1

def cpp_medoc():
    
    document = docx.Document()
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    '''Titre CPP'''
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run('Comité de Protection des Personnes')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sentence.font.name = 'Book Antiqua'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(20) 
    
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run('OUEST III')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    sentence.font.name = 'Book Antiqua'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(20)
    
    #ajouter le trai et les ombres
    
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run("Agréé par arrêté ministériel en date du 31 mai 2012, \n Constitué selon l'arrêté du Directeur Général de l'ARS Poitou Charentes en date du 25 juin 2012.")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sentence.font.name = 'Book Antiqua'
    sentence.italic = True
    sentence.font.size = docx.shared.Pt(10)
    
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run("\nC.H.U La Milétrie\nPavillon Administratif - Porte 213\n "
                                 "2 rue de le milétrie - CS 90 577 - 86021 POITIERS CEDEX\n"
                                 "Tel : 05.49.45.21.57\nFax : 05.49.46.12.62 \nE-mail : ")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sentence.font.name = 'Book Antiqua'
    sentence.italic = True
    sentence2 = paragraph.add_run(" cpp-ouest3@chu-poitiers.fr")
    sentence2.font.name = 'Book Antiqua'
    sentence2.italic = True
  #  sentence2.underline = True
 #   sentence2.font.color.rgb = RGB(0x0, 0x0, 0xFF)
    sentence.font.size = docx.shared.Pt(10)
    sentence2.font.size = docx.shared.Pt(10)
    
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run("\n Demande d'avis au CPP")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence.font.name = 'Arial'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(11.5)
    sentence2 = paragraph.add_run(" (arrêté du 2 décembre 2016)\n")
    sentence2.font.name = 'Arial'
    sentence2.font.size = docx.shared.Pt(11.5)
    sentence3 = paragraph.add_run("sur un projet de recherche mentionnée au 1° de l'article L. 1121-1 du CSP\nportant sur un ")
    sentence3.font.name = 'Arial'
    sentence3.bold = True
    sentence3.font.size = docx.shared.Pt(11.5)
    sentence4 = paragraph.add_run("médicament à usage humain\n")
    sentence4.font.name = 'Arial'
    sentence4.bold = True
    sentence4.underline = True
    sentence4.font.size = docx.shared.Pt(11.5)
    
 ###########################################   
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run("Préalablement au dépôt du dossier le promoteur obtient un numéro d’enregistrement de la recherche dans la base de données "
                                 "européenne des essais cliniques de médicaments à usage humain (EudraCT) et établie par l’Agence européenne des "
                                 "médicaments. Ce numéro EudraCT identifie chaque recherche conduite dans un ou plusieurs lieux de recherches situés sur le ")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
    sentence.font.name = 'Arial'
    sentence.italic = True
    sentence.font.size = docx.shared.Pt(9)
    sentence2 = paragraph.add_run("territoire de l’Union européenne.\n")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sentence2.font.name = 'Arial'
    sentence2.italic = True
    sentence2.font.size = docx.shared.Pt(9)
    
 ##############################

    #Tableau central
    table=document.add_table(rows=16, cols=1, style='Table Grid')

    table.cell(0,0).text=("DOSSIER ADMINISTRATIF")
    table.cell(1,0).text=("Courrier de demande d’avis daté et signé")
    table.cell(2,0).text=("Formulaire de demande d’avis (site internet de la base de données EudraCT)")
    table.cell(3,0).text=("Document additionnel (annexe 1) + supports pour recrutement des personnes")
    table.cell(4,0).text=("Le cas échéant, la copie de la ou des autorisations de lieux de recherches mentionnées à l’article L. 1121-13 du CSP")
    table.cell(5,0).text=("DOSSIER SUR LA RECHERCHE")
    table.cell(6,0).text=("Protocole de recherche (daté + numéro de version)")
    table.cell(7,0).text=("Résumé du protocole (daté + numéro de version)")
    table.cell(8,0).text=("La brochure pour l’investigateur\nou le résumé des caractéristiques du produit pour tout ME disposant d’une AMM en France.\n"
                          "ou dans un autre Etat membre de l’U.E accompagné, s’il est utilisé dans des conditions différentes de celles prévues par cette autorisation, de la synthèse des données justifiant l’utilisation et la sécurité d’emploi du médicament dans la recherche"
                          "\nSi la brochure pour l’investigateur appartient à un tiers, l’autorisation du tiers délivrée au promoteur pour l’utiliser")
    table.cell(9,0).text=("Le document d’information destiné aux personnes qui se prêtent à la recherche prévu à l’article L. 1122-1 du CSP.\n"
                          "Si le ME dispose d’une AMM en France, le dossier comprend une comparaison et, le cas échéant, la description et la justification des divergences pertinentes en terme de sécurité des personnes entre le document d’information destiné aux personnes qui se prêtent à la recherche et la notice prévue à l’article R. 5121-148 du CSP, au regard des contre- indications et des effets indésirables graves ou des mises en garde ou précautions d’emploi particulières).")
    table.cell(10,0).text=("Le formulaire de recueil du consentement des personnes se prêtant à la recherche")
    table.cell(11,0).text=("Attestation d’assurance (Décret n°2016-1537 du 16 novembre 2016 - art. 3)")
    table.cell(12,0).text=("Le cas échéant, l’avis d’un comité scientifique consulté par le promoteur")
    table.cell(13,0).text=("Une justification de l’adéquation des moyens humains, matériels et techniques au projet de recherche et de leur compatibilité avec les impératifs de sécurité des personnes qui s’y prêtent, sauf si le lieu bénéficie de l’autorisation mentionnée à l’article L. 1121-13 du CSP")
    table.cell(14,0).text=("Curriculum vitae signé du ou des investigateurs datant d’un an maximum")
    table.cell(15,0).text=("La nature de la décision finale de l’ANSM, si disponible.")
    n=1
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    if n==1 or n==6:
                        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run.bold = True
                    font = run.font
                    font.size= docx.shared.Pt(11)
                    font.name = 'Arial'
                    n=n+1
        
    
    document.save("soumission-cpp-medicaments.docx")