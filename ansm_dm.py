# -*- coding: utf-8 -*-
"""
Created on Wed Feb 20 22:04:40 2019

@author: Utilisateur
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

def main_ansm_dm():
    document = docx.Document()
    partie_une_ansm_dm(document)
    
    document.save("soumission-ansm-dm.docx")

def partie_une_ansm_dm(document):
    doc=document
    tableIdx=document.add_table(rows=15, cols=1, style='Table Grid')
    
    x = doc.tables[tableIdx].cell(0,1)._element.xpath('.//w:checkBox')
