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
    
    paragraph=document.add_paragraph("\nD1. Statut du DM / DM-DIV\n", style='debut_page')
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Le dispositif est-il marqué CE ?			□ oui   □ non\n\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.bold=True
    sentence=paragraph.add_run("Si le dispositif est marqué CE, compléter la rubrique ci-dessous ")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.italic=True
    
    table = document.add_table(rows=10, cols=2, style='Table Grid')
    table.cell(0,0).text=("Nom de l'Organisme Notifié")
    table.cell(1,0).text=("Numéro de l'Organisme Notifié")
    table.cell(2,0).text=("Date du marquage CE")
    table.cell(2,1).text=("     /     /     ")
    table.cell(3,0).text=("Destination(s) du marquage CE (telles que mentionnées dans la notice) ")
    table.cell(4,0).text=("Destination(s) du dispositif dans l’essai  ")
    table.cell(5,0).text=("Est-ce que l’utilisation du dispositif, dans le cadre de la recherche, se fait dans l’indication de son marquage CE ?")
    table.cell(5,1).text=("□ oui   □ non")
    table.cell(6,0).text=("La destination du dispositif dans l’essai est-elle conforme à des recommandations publiées (HAS, ANSM, Sociétés savantes, etc..) ?")
    table.cell(6,1).text=("□ oui   □ non   □ NA")
    table.cell(7,0).text=("Si oui, citer les références : ")
    table.cell(8,0).text=("Le dispositif est-il commercialisé dans un Etat membre de la Communauté européenne ou dans un pays tiers ? ")
    table.cell(8,1).text=("□ oui   □ non")
    table.cell(9,0).text=("Si oui, citer les pays concernés  ")
    n=0
    for col in table.columns:
        for cell in col.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n>9:
                    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    n=n+1
    
    paragraph=document.add_paragraph("\nD2. Identification du dispositif à étudier \n", style='debut_page')
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Rubriques à compléter dans tous les cas")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.italic=True
    
    table = document.add_table(rows=10, cols=2, style='Table Grid')
    table.cell(0,0).text=("Dénomination commune \n(exemple : stent artériel…)")
    table.cell(1,0).text=("Dénomination commerciale")
    table.cell(2,0).text=("Modèle")
    table.cell(3,0).text=("Version (y compris version du logiciel) ")
    table.cell(4,0).text=("Classification européenne")
    table.cell(5,0).text=("Classe du DM :\n")
    table.cell(5,1).text=("Classe du DMDIV :\n")
    table.cell(6,0).text=("Classe I                                                    □\nClasse IIa invasif à long terme                 □\n"
                          "Autres IIa                                                  □\nClasse IIb                                                 □\n"
                          "Classe III                                                  □\nDMIA                                                        □\n")
    table.cell(6,1).text=("Hors annexe II      		□\nAnnexe II liste A    		□\n"
                          "Annexe II liste B    		□\nAutotest                              	□\n")
    a=table.cell(7,0)
    b=table.cell(7,1)
    a.merge(b)
    table.cell(7,0).text=("En cas de dispositif non pourvu du marquage CE, renseigner la classe du dispositif et joindre une justification de la classification ")
    table.cell(8,0).text=("S’agit-il d’un dispositif implantable ?")
    table.cell(8,1).text=("□ oui   □ non")
    table.cell(9,0).text=("S’agit-il d’un dispositif « sur mesure » ?")
    table.cell(9,1).text=("□ oui   □ non")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n==12 or n==14:
                    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    if n==5 or n==6:
                        fontdebut.bold=True
                    elif n==9:
                        fontdebut.italic=True
                    n=n+1
    c=table.cell(5,0)
    d=table.cell(6,0)
    c.merge(d)
    e=table.cell(5,1)
    f=table.cell(6,1)
    e.merge(f)
    
    paragraph=document.add_paragraph()
    
    table = document.add_table(rows=6, cols=2, style='Table Grid')
    table.cell(0,0).text=("Fabricant du dispositif à étudier ")
    table.cell(0,1).text=("(à compléter quel que soit le statut du promoteur) ")
    table.cell(1,0).text=("Nom")
    table.cell(2,0).text=("Adresse")
    table.cell(3,0).text=("Numéro de téléphone")
    table.cell(4,0).text=("Numéro de télécopie")
    table.cell(5,0).text=("Courriel")
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
                        fontdebut.bold=True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(0,1)
    a.merge(b)
    
    paragraph=document.add_paragraph()
    
    table = document.add_table(rows=5, cols=3, style='Table Grid')
    table.cell(0,0).text=("Le dispositif sur lequel porte la recherche contient-il une des substances suivantes :")
    table.cell(0,1).text=("oui")
    table.cell(0,2).text=("non")
    table.cell(1,0).text=("-	Substance qui, si elle est utilisée séparément, est susceptible d'être considérée comme un médicament ? ")    
    table.cell(2,0).text=("-	Produits d’origine biologique (DMOA) ou dans la fabrication duquel interviennent de tels produits ?")
    table.cell(3,0).text=("-	OGM ?")
    table.cell(4,0).text=("-	Radioélément ?")
    table.cell(1,1).text=("□")
    table.cell(1,2).text=("□")
    table.cell(2,1).text=("□")
    table.cell(2,2).text=("□")
    table.cell(3,1).text=("□")
    table.cell(3,2).text=("□")
    table.cell(4,1).text=("□")
    table.cell(4,2).text=("□")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n!=0 or n!=3 or n!=6 or n!=9 or n!=12:
                   paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER 
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    if n==0:
                        fontdebut.bold=True
                        fontdebut.size = docx.shared.Pt(11)
                    elif n==1 or n==2:
                        fontdebut.size = docx.shared.Pt(11)
                    elif n==3 or n==6 or n==9 or n==12:
                        fontdebut.size = docx.shared.Pt(10)
                    else:
                        fontdebut.size = docx.shared.Pt(10.5)
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(1,0)
    a.merge(b)
    c=table.cell(2,0)
    a.merge(c)
    d=table.cell(3,0)
    a.merge(d)
    e=table.cell(4,0)
    a.merge(e)
    
    paragraph=document.add_paragraph("\nD3. Cas particulier : utilisation de dispositifs à étudier commercialisés et ayant la même dénomination commune, dans le cadre d’un essai dont le protocole n’impose pas l’utilisation d’un dispositif en particulier", style='debut_page')
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("\nEst-ce que ce cas particulier est applicable à l’essai concerné ?			□ oui   □ non\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    sentence=paragraph.add_run("Si oui,  compléter la rubrique ci-dessous")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.italic=True
    
    table = document.add_table(rows=3, cols=3, style='Table Grid')
    table.cell(0,1).text=("oui")
    table.cell(0,2).text=("non")
    table.cell(1,0).text=("DM")
    table.cell(1,1).text=("□")
    table.cell(1,2).text=("□")
    table.cell(2,0).text=("DMDIV")
    table.cell(2,1).text=("□")
    table.cell(2,2).text=("□")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                if n==0 or n==1 or n==3 or n==4 or n==6 or n==7:
                   paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER 
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(11)
                    n=n+1
    
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Dans l’affirmative, préciser les informations mentionnées ci-après")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(10)
    fontdebut.italic=True
    
    table = document.add_table(rows=5, cols=4, style='Table Grid')
    
    
    