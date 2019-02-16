# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 15:43:54 2019

@author: Marion
"""

#ce document va construire le protocole de categorie 1

#import gestion_tableau
import page_garde
import extraction
import docx
from docx.shared import Cm
from docx import Document

#def extraction_info():

# construire un document 
def construit_doc():
    
    
    document = docx.Document()
    extract=extraction.extraction()
    '''Marge des page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    
    page_garde.PageGarde(document)
    #historique des mises a jour
    page_garde.PageSignature(document)
    #principaux correspondants
    #sommaire
    page_garde.liste_abreviation(document,extract)
    #resume du proto version XX
    #abstract
    #grand I justification
    #grand 2 objectifs de la recherche
    #grand 3 criteres de jusgement
    #♥grand 4 conception de la recherche
    #♦grand 5 critere d'eligibilite
    #grand 6 deroulement de la recherche
    #♠grand 7 traitement strategies procedures de la recheche
    #•grand 8 traitement et procedures associees
    #◘grand 9 evaluationde la securite
    #○grand 10 surveillance de la recherche
    #grand 11 aspect statistuqe
    #grand 12 droit d'acces aux donnees et documents source
    #grand 13 controle et assurance de la qualite
    #grand 14 considerations ethiques et reglementaires
    #grand 15 conservation des documents et des donnees relatifs a la recherche
    #grand 16 rapport final
    #grand 17 regles relatives a la publication
    #grand 18 faisaabilite de l'etude
    #grand 19 biblio
    #grand 20 liste des annexes
    #grand 21 annexes (pas mal de trucs a faire)
    document.save("essai.docx")
