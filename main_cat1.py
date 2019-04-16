# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 15:43:54 2019

@author: Marion
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import cpp_medoc, cpp_hps, cpp_dm, Partie1,Partie2,Partie3,Partie4,Partie5,Partie6,Partie7,Partie8,Partie9,Partie10,Partie11,Partie12,Partie13,Partie14,Partie15,Partie16,Partie17,Partie18,Partie19,Partie20

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
import time
from time import gmtime, strftime
import page_garde
import extraction
import docx
from docx.shared import Cm
from docx import Document
from cpp_medoc import main_cpp_medoc
from cpp_dm import main_cpp_dm
from cpp_hps import main_cpp_hps
from ansm_dm import main_ansm_dm
from ansm_hps import main_ansm_hps
from ansm_pb import main_ansm_pb
from ansm_medoc import main_ansm_medoc
import test_sommaire

#def extraction_info():

# construire un document 
def construit_doc(dico):
      
    
    document = docx.Document()
    extract=extraction.extract1(dico)
    if dico['le_type_recherche']=='6':
        main_cpp_medoc(extract)
        main_ansm_medoc(extract)
    elif dico['le_type_recherche']=='8':
        main_cpp_hps(extract)
        main_ansm_hps(extract)
    elif dico['le_type_recherche']=='7':
        main_cpp_dm(extract)
        main_ansm_dm(extract)
    else:
        main_ansm_pb(extract)
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    
    page_garde.PageGarde(document,extract)
    page_garde.Page_version(document,extract)
    page_garde.PageSignature(document,extract)
    page_garde.PageCorespondant(document,extract)
    #sommaire
    test_sommaire.sommaire(document)
    page_garde.liste_abreviation(document,extract)
    page_garde.resume_protocole(document,extract)

    Partie1(document,extract)
    Partie2(document,extract)
    Partie3(document,extract)
    Partie4(document,extract)
    Partie5(document,extract)
    Partie6(document,extract)
    Partie7(document,extract)
    Partie8(document)
    Partie9(document)
    Partie10(document,extract)
    Partie11(document,extract)
    Partie12(document)
    Partie13(document)
    Partie14(document,extract)
    Partie15(document)
    Partie16(document)
    Partie17(document)
    Partie18(document,extract)
    Partie19(document)
    Partie20(document)
    
    sentence = extract['titre_abrege']

    
    date = (strftime('%d-%m-%Y',time.localtime()))
    
    
    #grand 21 annexes (pas mal de trucs a faire)
    document.save("ProtocoleCat1_"+ sentence +"_" + date + ".docx")
