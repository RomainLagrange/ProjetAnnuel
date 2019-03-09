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
     partie_C(document)
     partie_D(document)
     #partie_E(document)
     #partie_F_G(document)
     #partie_H_I(document)
     document.save("soumission-ansm-medicament.docx")

def partie_A_B(document):
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.2)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        
    '''Introduction'''
    table = document.add_table(rows=1, cols=6, style='Table Grid')
    table.cell(0,0).text=("Formulaire de demande d’autorisation auprès de l’ANSM et de demande d’avis à un Comité de protection des personnes d'une recherche mentionnée au 1° de l’article L. 1121-1 du code de la santé publique portant sur un médicament à usage humain ")
    table.cell(0,5).text=("FAEC")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.bold = True
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(11)
    a=table.cell(0,0)
    b=table.cell(0,4)
    a.merge(b)
    
    styles= document.styles
    style=styles.add_style('debut_page', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10) 
    
    paragraph=document.add_paragraph("\nPARTIE A COMPLETER PAR L’ANSM / LE COMITE DE PROTECTION DES PERSONNES (CPP)\n", style='debut_page')
    
    table = document.add_table(rows=8, cols=3, style='Table Grid')
    table.cell(0,0).text=("Date de réception de la demande :     \n"
                          "Date de demande d’information pour validation :      ")
    table.cell(0,1).text=("Date de demande d’informations complémentaires :      ")
    table.cell(0,2).text=("Refus d’autorisation / avis défavorable :    □\nDate :      ")
    table.cell(5,0).text=("Date d’enregistrement du dossier complet :      \nDate du début d'évaluation :      ")
    table.cell(5,1).text=("Date de réception des informations complémentaires / amendées :      ")
    table.cell(5,2).text=("Autorisation / avis favorable :      □\nDate :      ")
    table.cell(6,0).text=("Référence attribuée par l'ANSM :      \nRéférence attribuée par le CPP :      ")
    table.cell(6,1).text=("Retrait de la demande    □\nDate :      ")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
    a=table.cell(6,1)
    b=table.cell(6,2)
    a.merge(b)
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nPARTIE A COMPLETER PAR LE DEMANDEUR\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.underline=True
    sentence=paragraph.add_run("\nDEMANDE D’AUTORISATION À L’ANSM : 	                                                                         □\n"
                               "DEMANDE D’AVIS AU CPP	                                                                                                   □\n"
                               "\nA. IDENTIFICATION DE L’ESSAI CLINIQUE")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("A.1 Etat membre dans lequel la demande est soumise : FRANCE\n"
                          "A.2 Numéro EudraCT  :      \n"
                          "A.3 Titre complet de l’essai clinique :      \n"
                          "A.4 Numéro de code du protocole de l’essai attribué par le promoteur, version et date  :      \n"
                          "A.5 Nom ou titre abrégé de l’essai, le cas échéant :      \n"
                          "A.6 Numérotation ISRCTN , le cas échéant :      \n"
                          "A.7 S'agit-il d'une resoumission de la demande ?  □ oui  □ non  Si oui, indiquer la lettre de resoumission  :      ")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
    
    '''Partie B'''
    paragraph=document.add_paragraph("\nB. IDENTIFICATION DU PROMOTEUR RESPONSABLE DE LA DEMANDE", style='debut_page')
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.1	Promoteur ")
    table.cell(1,0).text=("B.1.1	Organisme :      \n"
                          "B.1.2	Nom de la personne à contacter :     \n" 
                          "B.1.3	Adresse :     \n" 
                          "B.1.4	Numéro de téléphone :      \n"
                          "B.1.5	Numéro de télécopie :    \n"  
                          "B.1.6	Mél :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.2 Représentant légal  du promoteur dans l’Union européenne pour l’essai concerné")
    table.cell(1,0).text=("B.2.1	Organisme :      \n"
                          "B.2.2	Nom de la personne à contacter :    \n"  
                          "B.2.3	Adresse :      \n"
                          "B.2.4	Numéro de téléphone :  \n"    
                          "B.2.5	Numéro de télécopie :   \n"   
                          "B.2.6	Mail :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.3 Statut du promoteur")
    table.cell(1,0).text=("B.3.1    Privé (commercial)                                                                      	□\n"
                          "B.3.2    Institutionnel  (non commercial)                                              	□")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    
def partie_C(document):
    
    '''Partie C'''
    paragraph=document.add_paragraph("\nC. IDENTIFICATION DU DEMANDEUR (cocher les cases appropriées)\n", style='debut_page')
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("C.1	Demande auprès de l’ANSM	                                                                      □")
    table.cell(1,0).text=("C.1.1       Promoteur                                                                                                                   □\n"
                          "C.1.2       Représentant légal du promoteur                                                                               □\n"
                          "C.1.3	Personne ou organisme délégué par le promoteur pour soumettre la demande	     □\n"
                          "C.1.4 	Préciser ci-après les informations relatives au demandeur, même si elles figurent ailleurs dans le formulaire : Si promoteur, partie B1, si représentant légal du promoteur, partie B2\n"
                          "C.1.4.1 	Organisme :      \n"
                          "C.1.4.2 	Nom de la personne à contacter :      \n"
                          "C.1.4.3 	Adresse :      \n"
                          "C.1.4.4 	Numéro de téléphone :      \n"
                          "C.1.4.5 	Numéro de télécopie :      \n"
                          "C.1.4.6	Mail :      
                          "C.1.5 Demande d'envoi d'une copie des données du formulaire sous format xml :\n"
                          "C.1.5.1 Souhaitez-vous recevoir une copie du fichier xml des données du formulaire sauvegardées sur la base EudraCT ?	□ oui    □ non \n"
                          "C.1.5.1.1 Si oui, indiquer les adresses mél auxquelles cette copie doit être adressée (5 adresses maximum) :      \n"
                          "C.1.5.1.2 Souhaitez-vous que cet envoi soit sécurisé  ?	□ oui    □ non\n"
                          "Si non à la question C.1.5.1.2, le fichier xml vous sera transmis par courrier électronique non sécurisé.")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("C.2	DEMANDE AUPRÈS DU CPP                                                                                      □")
    table.cell(1,0).text=("C.2.1       Promoteur                                                                                                                   □\n"
                          "C.2.2       Représentant légal du promoteur                                                                               □\n"
                          "C.2.3	Personne ou organisme délégué par le promoteur pour soumettre la demande	     □\n"
                          "C.2.4 Investigateur chargé de soumettre la demande, si applicable  :
                          "•	Investigateur coordonnateur (en cas d'essai multicentrique)	           □\n"
                          "•	Investigateur principal (en cas d'essai monocentrique)	               □\n"
                          "C.2.5 	Préciser ci-après les informations relatives au demandeur, même si elles figurent ailleurs dans le formulaire : Si promoteur, partie B1, si représentant légal du promoteur, partie B2\n"
                          "C.2.5.1	Organisme :      \n"
                          "C.2.5.2	Nom de la personne à contacter :      \n"
                          "C.2.5.3	Adresse :      \n"
                          "C.2.5.4	Numéro de téléphone :      \n"
                          "C.2.5.5	Numéro de télécopie : \n"     
                          "C.2.5.6	Mél :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
def Partie_D(document):
    
    '''Partie D'''
    paragraph=document.add_paragraph("\nD. DONNEES RELATIVES A CHAQUE MEDICAMENT EXPERIMENTAL", style='debut_page')
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Les informations concernant chaque "produit vrac" [c’est-à-dire avant toute opération pharmaceutique spécifique à l’essai (mise en insu, conditionnement et étiquetage)], doivent être indiquées dans cette section, pour chaque médicament expérimental (ME) étudié, y compris pour chaque médicament utilisé comme comparateur et pour chaque placebo, le cas échéant. Si l’essai clinique porte sur plusieurs ME, répéter cette section, en attribuant à chaque ME un numéro d’ordre à l'item D.1.1. Si le médicament est une association, les informations doivent être données pour chaque substance active concernée.\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    
    
    
    