# -*- coding: utf-8 -*-
"""
Created on Mon Mar  4 22:56:32 2019

@author: Marion
"""


import pandas as pd
import docx
from docx.api import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Cm
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt, RGBColor
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml
from docx.oxml import OxmlElement
import qn

def main_ansm_pb():
     document = docx.Document()
     partie_A_B(document)
     #partie_C(document)
     #partie_D(document)
     #partie_E(document)
     #partie_F_G(document)
     #partie_H_I(document)
     document.save("soumission-ansm-medicament.docx")

def partir_A_B(document):
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)