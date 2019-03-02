# -*- coding: utf-8 -*-
"""
Created on Fri Mar  1 22:35:39 2019

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
     document.save("soumission-ansm-pb.docx")
     
def partie_A_B(document):
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(0.95)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1.8)
        section.right_margin = Cm(1.8)
        
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("\n Demande d’autorisation auprès de l’ANSM et demande d’avis a un comité de protection des personnes d'une recherche biomédicale portant sur un produit sanguin labile, un organe, un tissu d’origine humaine ou animale ou une préparation de thérapie cellulaire\n")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.bold = True
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(11)
                    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nCe formulaire est commun pour la demande d’autorisation auprès de l’ANSM et pour la demande d’avis au CPP. Certains items peuvent ne pas être applicables à tous les produits, dans ce cas ne pas en tenir compte.")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    
    '''Partie ANSM/CPP'''
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nPartie réservée à l’ANSM / au Comité de protection des personnes (CPP) \n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.underline=True
    fontdebut.color.rgb = RGBColor(0x0,0x70,0xC0)
                    
    table = document.add_table(rows=9, cols=3, style='Table Grid')
    table.cell(0,0).text=("RECEVABILITE")
    table.cell(0,1).text=("ÉVALUATION")
    table.cell(0,2).text=("DÉCISION/AVIS")
    table.cell(1,0).text=("Date de réception de la demande :")
    table.cell(1,1).text=("Date de passage devant le groupe d’experts : ")
    table.cell(1,2).text=("Refus d’autorisation / avis défavorable 	□")
    for i in range(2,7):
        for j in range(0,3):
            if i==2 or i==4 or i==6:
                if j==2:
                    table.cell(i,j).text=("Date :   /  /  ")
                else:
                    table.cell(i,j).text=("  /  /  ")
    table.cell(3,0).text=("Date de demande de documents manquants : ")
    table.cell(3,1).text=("Date de demande d’informations complémentaires / objections motivées : ")
    table.cell(3,2).text=("Autorisation / avis favorable 	                □")
    table.cell(5,0).text=("Date d’enregistrement du dossier complet : ")
    table.cell(5,1).text=("Date de réception des informations complémentaires / amendées : ")
    table.cell(5,2).text=("Retrait de la demande 	                □")
    table.cell(7,0).text=("Date du début d'évaluation (J0) :   /  /  ")
    table.cell(8,0).text=("Référence attribuée par l'ANSM : 	     \nRéférence attribuée par le CPP : 	     ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n<3 or n==7 or n==6 or n==13 or n==12 or n==18 or n==19:
                    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n<3:
                        fontdebut.bold = True
                    n=n+1
    for i in range(0,3):
        a=table.cell(1,i)
        b=table.cell(6,i)
        a.merge(b)
    a=table.cell(7,0)
    b=table.cell(7,2)
    a.merge(b)
    a=table.cell(8,0)
    b=table.cell(8,2)
    a.merge(b)    
    
    '''Partie A'''
    
    styles= document.styles
    style=styles.add_style('debut_page', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10) 
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nPARTIE A COMPLETER PAR LE DEMANDEUR\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.underline=True
    sentence=paragraph.add_run("\nDEMANDE D’AUTORISATION À L’ANSM : 	                                                                         □\n"
                               "DEMANDE D’AVIS AU CPP	                                                                                                   □\n"
                               "A. IDENTIFICATION DE LA RECHERCHE BIOMÉDICALE\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10)
    
    table = document.add_table(rows=3, cols=1, style='Table Grid')
    table.cell(0,0).text=("A.1	    Etat membre dans lequel la demande est soumise : FRANCE")
    table.cell(1,0).text=("A.2	    Numéro d’enregistrement de la recherche en France (ID RCB)  :\n"
                          "A.3	    Titre complet de la recherche :")
    table.cell(2,0).text=("A.4	    Numéro de code du protocole attribué par le promoteur, version et date  : \n"
                          "A.5	    Nom ou titre abrégé de la recherche, le cas échéant : \n"
                          "A.6	    Numérotation ISRCTN , le cas échéant :      \n"
                          "A.7	    S'agit-il d'une resoumission de la demande ?	□ oui 	□ non\n"
                          "A.7.1	Si oui, indiquer la lettre de resoumission  :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n<2:
                        fontdebut.bold = True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(2,0)
    a.merge(b)
                    
    '''Partie B'''
    paragraph=document.add_paragraph("B. IDENTIFICATION DU PROMOTEUR RESPONSABLE DE LA DEMANDE", style='debut_page')
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.1	PROMOTEUR ")
    table.cell(1,0).text=("B.1.1	Organisme :      \n"
                          "B.1.2	Nom de la personne à contacter :     \n" 
                          "B.1.3	Adresse :     \n" 
                          "B.1.4	Numéro de téléphone :      \n"
                          "B.1.5	Numéro de télécopie :    \n"  
                          "B.1.6	Mail :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.2	REPRÉSENTANT LÉGAL DU PROMOTEUR  DANS LA COMMUNAUTÉ EUROPÉENNE POUR LA RECHERCHE CONCERNÉE")
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
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("B.3	STATUT DU PROMOTEUR")
    table.cell(1,0).text=("B.3.1    Privé (commercial)                                                                      	□\n"
                          "B.3.2	Institutionnel  (non commercial)                                              	□")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
                    
def partie_C(document):
    
    '''Partie C'''
    paragraph=document.add_paragraph("\nC. IDENTIFICATION DU DEMANDEUR (cocher les cases appropriées)\n", style='debut_page')
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("C.1	DEMANDE AUPRÈS DE L’ANSM	                                               □")
    table.cell(1,0).text=("C.1.1       Promoteur                                                                                     	□\n"
                          "C.1.2       Représentant légal du promoteur                                                	□\n"
                          "C.1.3	Personne ou organisme délégué par le promoteur pour soumettre la demande	□\n"
                          "C.1.4 	Préciser ci-après les informations relatives au demandeur, même si elles figurent ailleurs dans le formulaire : Si promoteur, partie B1, si représentant légal du promoteur, partie B2\n"
                          "C.1.4.1 	Organisme :      \n"
                          "C.1.4.2 	Nom de la personne à contacter :      \n"
                          "C.1.4.3 	Adresse :      \n"
                          "C.1.4.4 	Numéro de téléphone :      \n"
                          "C.1.4.5 	Numéro de télécopie :      \n"
                          "C.1.4.6	Mail :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    paragraph=document.add_paragraph()
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("C.2	DEMANDE AUPRÈS DU CPP                                                                                   □")
    table.cell(1,0).text=("C.2.1       Promoteur                                                                                                                   □\n"
                          "C.2.2       Représentant légal du promoteur                                                                               □\n"
                          "C.2.3	Personne ou organisme délégué par le promoteur pour soumettre la demande	     □\n"
                          "C.2.4 	Préciser ci-après les informations relatives au demandeur, même si elles figurent ailleurs dans le formulaire : Si promoteur, partie B1, si représentant légal du promoteur, partie B2\n"
                          "C.2.4.1	Organisme :      \n"
                          "C.2.4.2	Nom de la personne à contacter :      \n"
                          "C.2.4.3	Adresse :      \n"
                          "C.2.4.4	Numéro de téléphone :      \n"
                          "C.2.4.5	Numéro de télécopie : \n"     
                          "C.2.4.6	Mail :      ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
                    
    document.add_page_break()
    
def partie_D(document):
    
    '''Partie D1'''
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nD. DONNEES RELATIVES A CHAQUE PRODUIT SUR LEQUEL PORTE LA RECHERCHE")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(10) 
    sentence=paragraph.add_run("Les informations concernant chaque produit doivent être indiquées dans cette section :\n"
                               "       -	pour chaque produit sur lequel porte la recherche\n"
                               "       -	pour chaque produit utilisé comme comparateur \n"
                               "       -	et pour chaque placebo, le cas échéant.\n"
                               "Si la recherche biomédicale porte sur plusieurs produits, répéter cette section, en attribuant à chaque produit un numéro d’ordre à l'item D.1.1. \n"
                               "Si le produit est une association, les informations doivent être données pour chaque substance active ou produit concerné.\n\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10) 
    
    table = document.add_table(rows=4, cols=1, style='Table Grid')
    table.cell(0,0).text=("D.1	 IDENTIFICATION DU PRODUIT SUR LEQUEL PORTE LA RECHERCHE ")
    table.cell(1,0).text=("Indiquer ci-dessous quel produit est décrit dans cette section D. Le cas échéant, répéter cette section autant de fois qu'il y a de produits utilisés dans la recherche (numéroter chaque produit de 1 à n)")
    table.cell(2,0).text=("D.1.1         Cette section concerne le produit numéro :	     \n"
                          "D.1.2         Produit étudié                                                                       	□\n"
                          "D.1.3         Produit utilisé comme comparateur                                   	□")
    table.cell(3,0).text=("Pour le placebo, aller directement en section D.7")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0 or n==2:
                        fontdebut.bold = True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(1,0)
    a.merge(b)
    a=table.cell(2,0)
    b=table.cell(3,0)
    a.merge(b)
    
    paragraph=document.add_paragraph()
    
    '''Partie D2'''
    
    table = document.add_table(rows=10, cols=6, style='Table Grid')
    table.cell(0,0).text=("D.2	   STATUT DU PRODUIT SUR LEQUEL PORTE LA RECHERCHE ")
    table.cell(1,0).text=("D.2.1	   Le produit utilisé dans la recherche dispose-t-il d'une autorisation en France ou est-il enregistré en France  ?")
    table.cell(2,0).text=("D.2.1.1	   Si oui en D.2.1, préciser pour le produit utilisé dans la recherche :")
    table.cell(3,0).text=("D.2.1.1.1    Nom du produit autorisé ou enregistré ou nom commercial, le cas échéant :      ")
    table.cell(4,0).text=("D.2.1.1.2    Nom du titulaire de l’autorisation :      ")
    table.cell(5,0).text=("D.2.1.1.3    Numéro d’autorisation ou d’enregistrement :      ")
    table.cell(6,0).text=("D.2.1.1.4    Le produit sur lequel porte la recherche est-il modifié par rapport à son autorisation ? ")
    table.cell(7,0).text=("D.2.1.1.4.1 Si oui, veuillez préciser :      ")
    table.cell(8,0).text=("D.2.1.2       Le produit dispose-t-il d’une autorisation ou d’un enregistrement dans un autre pays ?")
    table.cell(9,0).text=("D.2.1.2.1    Si oui, veuillez préciser le pays et le nom de l’autorité qui a autorisé le produit :      ")
    for i in range(1,10):
        if i==1:
            table.cell(i,5).text=("□ oui  □ non\n")
        elif i==8 or i==6:
            table.cell(i,5).text=("□ oui  □ non")
        else:
            table.cell(i,5).text=(" ")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n==2 or n==12 or n==16:
                    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0 or n==1:
                        fontdebut.bold = True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(0,5)
    a.merge(b)
    a=table.cell(1,0)
    b=table.cell(9,4)
    a.merge(b)
    a=table.cell(1,5)
    b=table.cell(9,5)
    a.merge(b)
    
    paragraph=document.add_paragraph()
    
    table = document.add_table(rows=2, cols=6, style='Table Grid')
    table.cell(0,0).text=("D.2.2	 Dossier du produit sur lequel porte la recherche :")
    table.cell(1,0).text=("D.2.2.1     Dossier complet\n"
                          "D.2.2.2	 Dossier simplifié ")
    table.cell(0,5).text=(" ")
    table.cell(1,5).text=("□ oui  □ non\n□ oui  □ non")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n==3:
                    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0:
                        fontdebut.bold = True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(1,4)
    a.merge(b)
    a=table.cell(0,5)
    b=table.cell(1,5)
    a.merge(b)
    
    
    
    
    
    
    
    
    
    