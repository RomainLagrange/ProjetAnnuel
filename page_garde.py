#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec  6 16:44:16 2018

@author: romain
"""
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm


def PageGarde(document):
    
 #   document = docx.Document()
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    '''Logos de l'en-tete'''
    p = document.add_paragraph()
    r = p.add_run()
    r.add_picture('imageGauche.png')
    r.add_text('                                                                                                                                     ')
    r.add_picture('imageDroite.png')
    
    
    '''Titre de la recherche'''
    paragraph = document.add_paragraph()
    sentence = paragraph.add_run('  \n\nTitre de la recherche')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence.font.name = 'Times New Roman'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(14) 
    
    '''Acronyme'''
    paragraph2 = document.add_paragraph()
    paragraph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence = paragraph2.add_run('ACRONYME')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(22) 
    
    '''Version protocole'''
    paragraph2 = document.add_paragraph()
    paragraph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence = paragraph2.add_run('Protocole version n° X en date du XX/XX/201X\n\n')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.font.size = docx.shared.Pt(14) 
    sentence.bold = False
    
    '''Promoteur'''
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = paragraph.add_run('PROMOTEUR :\n')
    run1.font.name = 'Times New Roman'
    run1.font.size = docx.shared.Pt(11) 
    run1.bold = True
    run1.underline = True
    run2 = paragraph.add_run('Centre Hospitalier Universitaire de Poitiers - 2 rue de la Milétrie\n86021 POITIERS cedex\nTél : 05 49 44 33 89 / Fax : 05 49 44 30 58\n')
    run2.font.name = 'Times New Roman'
    run2.font.size = docx.shared.Pt(11) 
    
    '''investigateur'''
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = paragraph.add_run('INVESTIGATEUR COORDONNATEUR :\n')
    run1.font.name = 'Times New Roman'
    run1.font.size = docx.shared.Pt(11) 
    run1.bold = True
    run1.underline = True
    run2 = paragraph.add_run('Nom Investigateur\nService de : indiquer le nom du service\nCentre Hospitalier Universitaire de Poitiers - 2 rue de la Milétrie – CS 90577\n86021 Poitiers cedex\nTél : 05 49 44 xx xx / Fax : 05 49 44 xx xx\nE-mail : xxxxxxxx@chu-poitiers.fr')
    run2.font.name = 'Times New Roman'
    run2.font.size = docx.shared.Pt(11) 
    
    '''GIRCI SOHO'''
    paragraph2 = document.add_paragraph()
    paragraph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence = paragraph2.add_run('Ce protocole a été conçu et rédigé à partir de la version 3.0 du 01/02/2017\ndu protocole-type du GIRCI SOHO\n')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.font.size = docx.shared.Pt(12) 
    sentence.bold = True
    
    '''Confidentiel'''
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = paragraph.add_run('CE DOCUMENT CONFIDENTIEL')
    run1.font.name = 'Times New Roman'
    run1.font.size = docx.shared.Pt(10) 
    run1.underline = True
    run2 = paragraph.add_run(' EST LA PROPRIETE DU CHU DE POITIERS.\nAUCUNE INFORMATION NON PUBLIEE FIGURANT DANS CE DOCUMENT NE PEUT ETRE DIVULGUEE SANS AUTORISATION ECRITE PREALABLE DU CHU DE POITIERS')
    run2.font.name = 'Times New Roman'
    run2.font.size = docx.shared.Pt(10) 
    
   
  #  document.save("page_garde.docx")                   #sauvegarde
   
def PageSignature(document):
    
  #  document = docx.Document()
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.5)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    '''Logos de l'en-tete'''
    header = section.header
    p = header.paragraphs[0]
    r = p.add_run() 
    r.add_picture('imageGauche.png')
    r.add_text('                                                                                                                                     ')
    r.add_text('ACRONYME')
        
    '''Titre'''
    paragraph2 = document.add_paragraph()
    paragraph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence = paragraph2.add_run('PAGE DE SIGNATURE DU PROTOCOLE')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(16) 
    
    '''Signature investigateur'''
    paragraph2 = document.add_paragraph()
    sentence = paragraph2.add_run('Signature de l’investigateur')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.font.size = docx.shared.Pt(11) 
    sentence.bold = True
    sentence.underline = True
    
    '''Premiere case'''
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    text1 = ' \nJ\'ai lu ce protocole d’essai clinique dont le CHU de Poitiers est le promoteur. Je confirme qu\'il contient toutes les informations nécessaires à la conduite de l’essai. Je m\'engage à mener cet essai en respectant ses directives et les termes et conditions qui y sont définis.\n'
    text2 = 'Je m\'engage à réaliser l’essai en respectant :\n\n'
    text3 = '    -  les principes de la “Déclaration d’Helsinki”, \n\
    -  les règles et recommandations de bonnes pratiques cliniques internationales (ICH-E6) et française      (règles de bonnes pratiques cliniques pour les recherches portant sur des médicaments à usage humain - décisions du 24 novembre 2006), \n\
    -  la législation nationale et la réglementation relative aux essais cliniques,\n\
    -  la conformité avec la Directive Essais Cliniques de l’UE [2001/20/EC].\n\n\n'
    text4 = "Je m'engage également à ce que les investigateurs et les autres membres qualifiés de mon équipe aient accès au protocole et aux documents relatifs à la conduite de l’essai pour leur permettre de travailler dans le respect des dispositions figurant dans ces documents.\n"
    text5 = "Investigateur : Dr/ Pr XXXXX\n(Prénom NOM)\n\n\n\n"
    text6 = "Signature : ……………………………………………..                          Date : ___________________\n"       

    table.cell(0,0).text = text1 +text2 + text3 + text4 +text5+text6
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size= docx.shared.Pt(11)
                    font.name = 'Times New Roman'
                    
    '''Signature investigateur coordonnateur'''
    paragraph2 = document.add_paragraph()
    sentence = paragraph2.add_run(' \nSignature de l’Investigateur Coordonnateur')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.font.size = docx.shared.Pt(11) 
    sentence.bold = True
    sentence.underline = True
    
    '''Deuxieme case'''
    table2 = document.add_table(rows=1, cols=1, style='Table Grid')
    text5 = "Investigateur Coordonnateur : Dr/ Pr XXXXX\n(Prénom NOM)\n\n\n\n"
    text6 = "Signature : ……………………………………………..                          Date : ___________________\n" 
    table2.cell(0,0).text = text5 + text6
    for row in table2.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size= docx.shared.Pt(11)
                    font.name = 'Times New Roman'
                    
                    
    '''Signature investigateur coordonnateur'''
    paragraph2 = document.add_paragraph()
    sentence = paragraph2.add_run(' \nSignature de l’Investigateur Coordonnateur')
    '''Then format the sentence'''
    sentence.font.name = 'Times New Roman'
    sentence.font.size = docx.shared.Pt(11) 
    sentence.bold = True
    sentence.underline = True
    
    '''Troisieme case'''
    table3 = document.add_table(rows=1, cols=1, style='Table Grid')
    text5 = "Promoteur : Jean-Pierre DEWITTE\nPour le Directeur Général et par délégation\nle Directeur de la Recherche,\n\n\n\n"
    text6 = "Signature : ……………………………………………..                          Date : ___________________\n" 
    table3.cell(0,0).text = text5 + text6
    for row in table3.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    font = run.font
                    font.size= docx.shared.Pt(11)
                    font.name = 'Times New Roman'
  
    '''Pied de page'''
    footer = section.footer
    p = footer.paragraphs[0]
    r = p.add_run('Version n°X du XX/XX/201X	                               CONFIDENTIEL                                                Page 3 sur 14') 
    r.font.name = 'Times New Roman'
    r.font.size = docx.shared.Pt(11)

    
  #  document.save("page_signature.docx")                   #sauvegarde