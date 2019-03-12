# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 15:43:54 2019

@author: Marion
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import Partie1,Partie2,Partie3,Partie4,Partie5,Partie6,Partie7,Partie8,Partie9,Partie10,Partie11,Partie12,Partie13,Partie14,Partie15,Partie16,Partie17,Partie18,Partie19,Partie20

from Partie1 import Partie1
from Partie2 import Partie2
from Partie3 import Partie3
from Partie4 import Partie4
from Partie5 import Partie5
from Partie6 import Partie6
from Partie7 import Partie7
from Partie8 import Partie8
from Partie9 import Partie9
from Partie10 import Partie10
from Partie11 import Partie11
from Partie12 import Partie12
from Partie13 import Partie13
from Partie14 import Partie14
from Partie15 import Partie15
from Partie16 import Partie16
from Partie17 import Partie17
from Partie18 import Partie18
from Partie19 import Partie19
from Partie20 import Partie20
import page_garde
import extraction
import docx
from docx.shared import Cm
from docx import Document

#def extraction_info():

# construire un document 
def construit_doc():
    
    
    document = docx.Document()
    extract=extraction.extract1()
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    
    page_garde.PageGarde(document,extract)
    #historique des mises a jour
    page_garde.PageSignature(document)
    #principaux correspondants
    #sommaire
    page_garde.liste_abreviation(document,extract)
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
    Partie19(document)
    Partie20(document)
    #grand 21 annexes (pas mal de trucs a faire)
    document.save("ProtocoleCat1.docx")
