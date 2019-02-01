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
    sentence = paragraph.add_run("\n C.H.U La Milétrie \n Pavillon Administratif - Porte 213 \n "
                                 "2 rue de le milétrie - CS 90 577 - 86021 POITIERS CEDEX \n"
                                 " Tel : 05.49.45.21.57 \n Fax : 05.49.46.12.62 \n E-mail : cpp-ouest3@chu-poitiers.fr")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    sentence.font.name = 'Book Antiqua'
    sentence.italic = True
    sentence.font.size = docx.shared.Pt(10)
        
    document.save("soumission-cpp-medicaments.docx")