# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 18:26:32 2019

@author: Asuspc
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import Cat2Part1,Cat2Part2,Cat2Part3,Cat2Part4,Cat2Part5,Cat2Part6,Cat2Part7,Cat2Part8,Cat2Part9,Cat2Part10,Cat2Part11,Cat2Part12,Cat2Part13,Cat2Part14,Cat2Part15,Cat2Part16,Cat2Part17,Cat2Part18
from Cat2Part1 import Partie1
from Cat2Part2 import Partie2
from Cat2Part3 import Partie3
from Cat2Part4 import Partie4
from Cat2Part5 import Partie5
from Cat2Part6 import Partie6
from Cat2Part7 import Partie7
from Cat2Part8 import Partie8
from Cat2Part9 import Partie9
from Cat2Part10 import Partie10
from Cat2Part11 import Partie11
from Cat2Part12 import Partie12
from Cat2Part13 import Partie13
from Cat2Part14 import Partie14
from Cat2Part15 import Partie15
from Cat2Part16 import Partie16
from Cat2Part17 import Partie17
from Cat2Part18 import Partie18

#import page_garde A FAIRE POUR CAT2
import extraction
import docx
from docx.shared import Cm
from docx import Document

#def extraction_info():

# construire un document 
def construit_doc():
    
    
    document = docx.Document()
  #  extract=extraction.extraction()
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    
    #page_garde.PageGarde(document)
    #historique des mises a jour
  #  page_garde.PageSignature(document)
    #principaux correspondants
    #sommaire
 #   page_garde.liste_abreviation(document,extract)
    #resume du proto version XX
    #abstract
    Partie1(document)
    Partie2(document)
    Partie3(document)
    Partie4(document)
    Partie5(document)
    Partie6(document)
    Partie7(document)
    Partie8(document)
    Partie9(document)
    Partie10(document)
    Partie11(document)
    Partie12(document)
    Partie13(document)
    Partie14(document)
    Partie15(document)
    Partie16(document)
    Partie17(document)
    Partie18(document)
#    Cat2Part1(document)
#    Cat2Part2(document)
#    Cat2Part3(document)
#    Cat2Part4(document)
#    Cat2Part5(document)
#    Cat2Part6(document)
#    Cat2Part7(document)
#    Cat2Part8(document)
#    Cat2Part9(document)
#    Cat2Part10(document)
#    Cat2Part11(document)
#    Cat2Part12(document)
#    Cat2Part13(document)
#    Cat2Part14(document)
#    Cat2Part15(document)
#    Cat2Part16(document)
#    Cat2Part17(document)
#    Cat2Part18(document)
    document.save("ProtocoleCat2.docx")