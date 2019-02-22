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
    partie_B_C(document)
    partie_D(document)
    document.save("soumission-ansm-dm.docx")

def partie_une_ansm_dm(document):
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.7)
        section.bottom_margin = Cm(1)
        section.left_margin = Cm(1)
        section.right_margin = Cm(1)
        
    '''Titre du document'''
    styles= document.styles
    style=styles.add_style('debut_page', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(11) 
    
    paragraph=document.add_paragraph("Formulaire de demande d’autorisation auprès de l'ANSM et demande d’avis du comite de protection des personnes (CPP) pour une recherche mentionnée au 1° ou au 2° de l’article L. 1121-1 du code de la santé publique portant sur un dispositif médical (DM) ou un dispositif médical de diagnostic in vitro (DMDIV)\n", style='debut_page')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Toutes les rubriques du formulaire doivent être complétées.\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.italic=True
    fontdebut.size = docx.shared.Pt(10.5)
    sentence=paragraph.add_run("Partie réservée à l’ANSM / au CPP :\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.italic=True
    fontdebut.size = docx.shared.Pt(10)
    
    table = document.add_table(rows=3, cols=3, style='Table Grid')
    a=table.cell(1,0)
    b=table.cell(2,0)
    a.merge(b)
    c=table.cell(1,1)
    d=table.cell(2,1)
    c.merge(d)
    table.cell(0,0).text=("Date de réception de la demande :\n\n\n     /     /     ")
    table.cell(0,1).text=("Date de demande d’informations complémentaires :\n\n     /     /     ")
    table.cell(0,2).text=("Refus d’autorisation / avis défavorable :	□\n\n	Préciser la  date :\n     /     /     ")
    table.cell(1,0).text=("\nDate du début de la procédure :\n\n     /     /     ")
    table.cell(1,1).text=("Date de réception des informations complémentaires :\n\n     /     /     ")
    table.cell(1,2).text=("Autorisation / avis favorable :	□\n\nPréciser la date : \n     /     /     ")
    table.cell(2,2).text=("Retrait de la demande :	□\nDate :      /     /     ")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Partie à compléter par le demandeur :\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.italic=True
    fontdebut.size = docx.shared.Pt(10)
    sentence=paragraph.add_run("Ce formulaire est commun pour la demande d’autorisation auprès de l’ANSM et pour la demande d’avis auprès du CPP. Veuillez cocher ci-après la case correspondant à l’objet de la demande.\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    sentence=paragraph.add_run("□ Recherche interventionnelle mentionnée au 1° de l’article L. 1121-1 du code de la santé publique\n\n"
                               "□ Recherche interventionnelle ")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(12)
    sentence=paragraph.add_run("comportant des risques et contraintes minimes ")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold=True
    fontdebut.size = docx.shared.Pt(12)
    sentence=paragraph.add_run("mentionnée au 2° de l’article L. 1121-1 du code de la santé publique\n\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(12)
    sentence=paragraph.add_run("Demande d’autorisation à l’ANSM :  □				Demande d’avis au CPP :  □\n\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold=True
    fontdebut.size = docx.shared.Pt(11)
    
    
    '''A identification de la recherche'''
    paragraph=document.add_paragraph("A. Identification de la recherche \n", style='debut_page')
    table = document.add_table(rows=9, cols=2, style='Table Grid')
    table.cell(0,0).text=("Numéro d’enregistrement de la recherche auprès de l'ANSM (n°IDRCB)")
    table.cell(1,0).text=("Numéro EUDAMED  (le cas échéant)")
    table.cell(2,0).text=("Titre complet de la recherche ")
    table.cell(3,0).text=("Numéro de code du protocole attribué par le promoteur, version et date")
    table.cell(4,0).text=("Nom ou titre abrégé (le cas échéant) ")
    table.cell(5,0).text=("S’agit-il d’une resoumission ?")
    table.cell(5,1).text=("□ oui       □ non")
    table.cell(6,0).text=("Si oui, indiquer la lettre de resoumission  ")
    table.cell(7,0).text=("Préciser par ailleurs si les documents précédemment versés ont été modifiés ?")
    table.cell(7,1).text=("□ oui       □ non")
    a=table.cell(8,0)
    b=table.cell(8,1)
    a.merge(b)
    table.cell(8,0).text=("(Si oui, joindre au dossier de demande d’AEC un tableau comparatif)")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==0 or n==2:
                        fontdebut.bold=True
                    n=n+1
 
    
# Parties B et C du document
def partie_B_C(document):
    
    paragraph=document.add_paragraph("\n\nB. Identification du promoteur responsable de la recherche \n", style='debut_page')
    table = document.add_table(rows=7, cols=2, style='Table Grid')
    a=table.cell(0,0)
    b=table.cell(0,1)
    a.merge(b)
    table.cell(0,0).text=("B1. Promoteur")
    table.cell(1,0).text=("Nom de l'organisme")
    table.cell(2,0).text=("Nom de la personne à contacter ")
    table.cell(3,0).text=("Adresse")
    table.cell(4,0).text=("Numéro de téléphone")
    table.cell(5,0).text=("Numéro de télécopie")
    table.cell(6,0).text=("Courriel")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    if n==0:
                        fontdebut.bold=True
                        fontdebut.size = docx.shared.Pt(11)
                    else:
                        fontdebut.size = docx.shared.Pt(10)
                    n=n+1
    #separe les deux tableaux
    paragraph=document.add_paragraph()
    
    table = document.add_table(rows=7, cols=2, style='Table Grid')
    a=table.cell(0,0)
    b=table.cell(0,1)
    a.merge(b)
    table.cell(0,0).text=("B2. Représentant légal  du promoteur dans la Communauté européenne pour la recherche (si différent du promoteur)")
    table.cell(1,0).text=("Nom de l'organisme")
    table.cell(2,0).text=("Nom de la personne à contacter ")
    table.cell(3,0).text=("Adresse")
    table.cell(4,0).text=("Numéro de téléphone")
    table.cell(5,0).text=("Numéro de télécopie")
    table.cell(6,0).text=("Courriel")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    if n==0:
                        fontdebut.bold=True
                        fontdebut.size = docx.shared.Pt(11)
                    else:
                        fontdebut.size = docx.shared.Pt(10)
                    n=n+1
    #separe les deux tableaux
    paragraph=document.add_paragraph()
    
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("B3. Statut du promoteur \nCommercial	□	Non commercial	□")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.bold=True
                    fontdebut.size = docx.shared.Pt(11)
    
    '''Partie C'''
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\n\nC. Identification du demandeur ")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(11)
    fontdebut.bold=True
    sentence=paragraph.add_run("(cocher les cases appropriées)\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(11)
    
    table = document.add_table(rows=8, cols=2, style='Table Grid')
    a=table.cell(1,0)
    b=table.cell(1,1)
    a.merge(b)
    table.cell(0,0).text=("C1. Demandeur pour l'ANSM 		□")
    table.cell(0,1).text=("C2. Demandeur pour le CPP  		□")
    table.cell(1,0).text=("\nPromoteur ……………………………………………………………………………………………□\n"
                          "\nReprésentant légal du promoteur …………………………………………………………………□\n"
                          "\nPersonne ou organisme délégué par le promoteur pour soumettre la demande…………….□\n"
                          "\nDans ce cas, compléter ci-après :")
    table.cell(2,0).text=("Nom de l'organisme")
    table.cell(3,0).text=("Nom de la personne à contacter")
    table.cell(4,0).text=("Adresse")
    table.cell(5,0).text=("Numéro de téléphone")
    table.cell(6,0).text=("Numéro de télécopie")
    table.cell(7,0).text=("Courriel")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    if n==0:
                        fontdebut.bold=True
                        fontdebut.size = docx.shared.Pt(11)
                    else:
                        fontdebut.size = docx.shared.Pt(10)

#partie D
def partie_D(document):
    
    paragraph=document.add_paragraph("\n\nD. Fiche de données sur le(s) DM (s)/ DM-DIV (s) faisant l'objet de la recherche, y compris les comparateurs \n", style='debut_page')
    table = document.add_table(rows=2, cols=1, style='Table Grid')
    table.cell(0,0).text=("Indiquer ici quel DM / DM-DIV est concerné par cette section D (utiliser une fiche pour chaque DM / DM-DIV)\n")
    table.cell(1,0).text=("Dispositif sur lequel porte la recherche ……………………………………………….□\nDispositif utilisé comme comparateur………………………………….………………□")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    if n==0:
                        fontdebut.italic=True
                        fontdebut.size = docx.shared.Pt(10)
                    else:
                        fontdebut.bold=True
                        fontdebut.size = docx.shared.Pt(10)
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(1,0)
    a.merge(b)
    
    