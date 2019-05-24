# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 18:26:32 2019

@author: Asuspc
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import page_garde_cat2, cpp_dm, cpp_hps, Cat2Part1,Cat2Part2,Cat2Part3,Cat2Part4,Cat2Part5,Cat2Part6,Cat2Part7,Cat2Part8,Cat2Part9,Cat2Part10,Cat2Part11,Cat2Part12,Cat2Part13,Cat2Part14,Cat2Part15,Cat2Part16,Cat2Part17,Cat2Part18
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
from cpp_dm import main_cpp_dm
from cpp_hps import main_cpp_hps
import time
from time import gmtime, strftime
import extraction
import docx
from docx.shared import Cm
from docx import Document

#def extraction_info():

# construire un document 
def construit_doc(dico):
    
    
    document = docx.Document()
    extract=extraction.extract2(dico)
    if dico['le_type_recherche']=='7':
        main_cpp_dm(extract)
    else:
        main_cpp_hps(extract)
  #  extract=extraction.extraction()
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    #page de garde
    page_garde_cat2.PageGarde(document,extract)
    #historique des mises a jour
    page_garde_cat2.Page_version(document,extract)
    #signature
    page_garde_cat2.PageSignature(document,extract)
    #principaux correspondants
    page_garde_cat2.PageCorespondant(document,extract)
    #abreviations
    page_garde_cat2.liste_abreviation(document,extract)
    #resume du proto version XX
    page_garde_cat2.resume_protocole(document,extract)
    #abstract
    page_garde_cat2.Page_abstract(document,extract)
    
    Partie1(document,extract)
    Partie2(document,extract)
    Partie3(document,extract)
    Partie4(document,extract)
    Partie5(document,extract)
    Partie6(document,extract)
    Partie7(document)
    Partie8(document)
    Partie9(document,extract)
    Partie10(document,extract)
    Partie11(document)
    Partie12(document)
    Partie13(document)
    Partie14(document)
    Partie15(document)
    Partie16(document,extract)
    Partie17(document)
    Partie18(document)
    
    sentence = extract['titre_abrege']
    date = (strftime('%d-%m-%Y',time.localtime()))

    document.save("ProtocoleCat2_"+ sentence+"_"+date+".docx")