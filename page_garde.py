#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Created on Thu Dec  6 16:44:16 2018

@author: romain
"""
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm


def PageGarde():
    
    document = docx.Document()
    
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
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
    
    '''investigateur'''
    paragraph = document.add_paragraph()
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run1 = paragraph.add_run('CE DOCUMENT CONFIDENTIEL')
    run1.font.name = 'Times New Roman'
    run1.font.size = docx.shared.Pt(10) 
    run1.underline = True
    run2 = paragraph.add_run(' EST LA PROPRIETE DU CHU DE POITIERS.\nAUCUNE INFORMATION NON PUBLIEE FIGURANT DANS CE DOCUMENT NE PEUT ETRE DIVULGUEE SANS AUTORISATION ECRITE PREALABLE DU CHU DE POITIERS')
    run2.font.name = 'Times New Roman'
    run2.font.size = docx.shared.Pt(10) 
    
       
    document.save("page_garde.docx")                   #sauvegarde
   