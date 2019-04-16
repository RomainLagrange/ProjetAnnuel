# -*- coding: utf-8 -*-
"""
Created on Wed Feb 27 10:39:54 2019

@author: Asuspc
"""

# -*- coding: utf-8 -*-
"""
Created on Thu Feb 21 18:26:32 2019

@author: Asuspc
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import cpp_rni, Cat3Part1,Cat3Part2,Cat3Part3,Cat3Part4,Cat3Part5,Cat3Part6,Cat3Part7,Cat3Part8,Cat3Part9,Cat3Part10,Cat3Part11,Cat3Part12,Cat3Part13,Cat3Part14,Cat3Part15
from Cat3Part1 import Partie1
from Cat3Part2 import Partie2
from Cat3Part3 import Partie3
from Cat3Part4 import Partie4
from Cat3Part5 import Partie5
from Cat3Part6 import Partie6
from Cat3Part7 import Partie7
from Cat3Part8 import Partie8
from Cat3Part9 import Partie9
from Cat3Part10 import Partie10
from Cat3Part11 import Partie11
from Cat3Part12 import Partie12
from Cat3Part13 import Partie13
from Cat3Part14 import Partie14
from Cat3Part15 import Partie15
from cpp_rni import main_cpp_RNI
import page_garde_cat3
import time
from time import gmtime, strftime

#import page_garde A FAIRE POUR CAT3
import extraction
import docx
from docx.shared import Cm
from docx import Document

#def extraction_info():

# construire un document 
def construit_doc(dico):
    
    
    document = docx.Document()
    extract=extraction.extract3(dico)
    main_cpp_RNI(extract)
  #  extract=extraction.extraction()
    
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
#page de garde
    #page_garde_cat3.PageGarde(document)
#historique des mises a jour
    #page_garde_cat3.Page_version(document,extract)
#signatures
    #page_garde_cat3.PageSignature(document)
#principaux correspondants
    #page_garde_cat3.PageCorespondant(document,extract)
#sommaire
#abbreviations
    #page_garde_cat3.liste_abreviation(document,extract)
#resume 
    #page_garde_cat3.resume_protocole(document,extract)
    

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

    sentence = extract['titre_abrege']
    date = (strftime('%d-%m-%Y',time.localtime()))

    document.save("ProtocoleCat3_"+sentence+"_"+date+".docx")