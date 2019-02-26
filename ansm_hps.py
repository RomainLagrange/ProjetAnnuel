# -*- coding: utf-8 -*-
"""
Created on Tue Feb 26 12:13:42 2019

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

def main_ansm_hps():
    document = docx.Document()
    parties_ABC(document)
    partie_D_a_G(document)
    
    document.save("soumission-ansm-hps.docx")
    
def parties_ABC(document):
    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(1.59)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(2.5)
        section.right_margin = Cm(1.63)
        
    
    '''Titre du document en bleu'''
    
    
    
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("FORMULAIRE DE DEMANDE D'AUTORISATION AUPRES DE L’AGENCE NATIONALE DE SECURITE DU MEDICAMENT ET DES PRODUITS DE SANTE "
                          "OU DE DEMANDE D’AVIS AU COMITÉ DE PROTECTION DES PERSONNES POUR UNE RECHERCHE MENTIONNEE AU 1° OU AU 2° DE L’ARTICLE L. 1121-1 DU CODE DE LA "
                          "SANTE PUBLIQUE NE PORTANT PAS SUR UN PRODUIT MENTIONNE A L’ARTICLE L. 5311-1 DU CODE DE LA SANTE PUBLIQUE - ")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.bold = True
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(10)
                    fontdebut.color.rgb = RGBColor(0x0,0x70,0xC0)
                    
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
    
    paragraph=document.add_paragraph("\nPartie réservée à l’ANSM / au Comité de protection des personnes (CPP) ", style='debut_page')
    table = document.add_table(rows=4, cols=3, style='Table Grid')
    table.cell(0,0).text=("Date de réception de la demande : ")
    table.cell(0,1).text=("Date de demande d’informations complémentaires : ")
    table.cell(0,2).text=("Autorisation de l'ANSM:")
    table.cell(1,2).text=("    Oui         Non\n")
    table.cell(2,2).text=("Date :")
    table.cell(3,0).text=("Date d’enregistrement de la demande considérée complète : \n\nDate du début de la procédure :")
    table.cell(3,1).text=("Date de réception des informations complémentaires / amendées :")
    table.cell(3,2).text=("Avis du CPP :")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(9)
                    if n==3:
                        fontdebut.bold = True
                    n=n+1
    for i in range(0,3):
        a=table.cell(0,i)
        b=table.cell(2,i)
        a.merge(b)
    
    paragraph=document.add_paragraph()
    sentence=paragraph.add_run("Partie à compléter par le demandeur :\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(11)
    sentence=paragraph.add_run("Recherche interventionnelle mentionnée au 1° de l’article L. 1121-1 du code de la santé publique :")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run(":    OUI      NON\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run("Recherche interventionnelle ne comportant que des risques et contraintes minimes mentionnée au 2° de l’article L. 1121-2 du code de la santé publique :")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run(":    OUI      NON\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run("\nDEMANDE D'AUTORISATION A L’ANSM :")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run(":   ouiI      non\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run("DEMANDE D’AVIS AU CPP :")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.size = docx.shared.Pt(9)
    sentence=paragraph.add_run(":   ouiI      non\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(9)
    

    '''Partie A'''
    paragraph=document.add_paragraph("\nA. IDENTIFICATION DE LA RECHERCHE ", style='debut_page')
    table = document.add_table(rows=6, cols=6, style='Table Grid')
    #merge des cellules
    for i in range(0,6):
        if i==4:
            a=table.cell(i,0)
            b=table.cell(i,2)
            a.merge(b)
            a=table.cell(i,3)
            b=table.cell(i,5)
            a.merge(b)
        elif i==2 or i==3:
            a=table.cell(i,0)
            b=table.cell(i,1)
            a.merge(b)
            a=table.cell(i,2)
            b=table.cell(i,3)
            a.merge(b)
            a=table.cell(i,4)
            b=table.cell(i,5)
            a.merge(b)
        else:
            a=table.cell(i,0)
            b=table.cell(i,5)
            a.merge(b)
    table.cell(0,0).text=("Titre complet de la recherche : \n")
    table.cell(1,0).text=("Numéro IDRCB d'enregistrement de la recherche :")
    table.cell(2,0).text=("Numéro de code du protocole de la recherche donné par le promoteur")
    table.cell(2,2).text=("Version")
    table.cell(2,4).text=("Date :")
    table.cell(4,0).text=("Nom ou titre abrégé de la recherche, le cas échéant :  ")
    table.cell(5,0).text=("Inscription au fichier VRB           oui            non")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    '''Partie B'''
    paragraph=document.add_paragraph("\nB. IDENTIFICATION DU PROMOTEUR RESPONSABLE DE LA DEMANDE\n      B1. Promoteur  ", style='debut_page')
    table = document.add_table(rows=3, cols=2, style='Table Grid')
    for i in range(0,2):
        a=table.cell(i,0)
        b=table.cell(i,1)
        a.merge(b)
    table.cell(0,0).text=("Nom de l'organisme :")
    table.cell(1,0).text=("Nom de la personne à contacter :")
    table.cell(2,0).text=("Adresse : ")
    table.cell(2,1).text=("Numéro de téléphone :\n\nNuméro de télécopie :\n\nCourriel :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)

    paragraph=document.add_paragraph("\n       B2. Représentant légal  du promoteur dans l’Union européenne pour la recherche (si différent du promoteur) ",style='debut_page')
    table = document.add_table(rows=4, cols=2, style='Table Grid')
    for i in range(0,4):
        if not(i==2):
            a=table.cell(i,0)
            b=table.cell(i,1)
            a.merge(b)
    table.cell(0,0).text=("Nom de l'organisme :")
    table.cell(1,0).text=("Nom de la personne à contacter :")
    table.cell(2,0).text=("Adresse : ")
    table.cell(2,1).text=("Numéro de téléphone :\n\nNuméro de télécopie :\n\nCourriel :")
    table.cell(3,0).text=("Statut du promoteur :        commercial          non commercial")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)

    '''Partie C'''
    paragraph=document.add_paragraph("\n C. IDENTIFICATION DU DEMANDEUR \nC1. Demande pour l’ANSM", style='debut_page')
    table = document.add_table(rows=4, cols=2, style='Table Grid')
    for i in range(0,3):
        a=table.cell(i,0)
        b=table.cell(i,1)
        a.merge(b)
    table.cell(0,0).text=("Si promoteur, partie B1, si représentant légal du promoteur, partie B2 ")
    table.cell(1,0).text=("Nom de l'organisme :")
    table.cell(2,0).text=("Nom de la personne à contacter :")
    table.cell(3,0).text=("Adresse : ")
    table.cell(3,1).text=("Numéro de téléphone :\n\nNuméro de télécopie :\n\nCourriel :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    paragraph=document.add_paragraph("\n  C2. Demande pour le Comité de protection des personnes", style='debut_page')
    table = document.add_table(rows=4, cols=2, style='Table Grid')
    for i in range(0,3):
        a=table.cell(i,0)
        b=table.cell(i,1)
        a.merge(b)
    table.cell(0,0).text=("Si promoteur, partie B1, si représentant légal du promoteur, partie B2 ")
    table.cell(1,0).text=("Nom de l'organisme :")
    table.cell(2,0).text=("Nom de la personne à contacter :")
    table.cell(3,0).text=("Adresse : ")
    table.cell(3,1).text=("Numéro de téléphone :\n\nNuméro de télécopie :\n\nCourriel :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
def partie_D_a_G(document):
    
    '''Partie D'''
    paragraph=document.add_paragraph("\n D. DONNÉES SUR LE(S) PRODUITS(S) EXPÉRIMENTAL (AUX) UTILISÉ(S) DANS LA RECHERCHE : PRODUITS(S) ÉTUDIÉ(S) OU UTILISÉ(S) COMME COMPARATEUR(S) ", style='debut_page')
    table = document.add_table(rows=4, cols=1, style='Table Grid')
    table.cell(0,0).text=("Indiquer ici quel PE est concerné par cette section D ; si nécessaire, utiliser d’autres fiches pour chaque PE utilisé dans l’essai (à numéroter de 1 à n) :")
    table.cell(1,0).text=("Cette section concerne le PE numéro :  ")
    table.cell(2,0).text=("PE étudié               oui              non")
    table.cell(3,0).text=("PE utilisé comme comparateur            oui           non")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    '''Partie D1'''
    paragraph=document.add_paragraph("\n  D.1. DESCRIPTION DU PRODUIT EXPÉRIMENTAL", style='debut_page')
    table = document.add_table(rows=5, cols=1, style='Table Grid')
    table.cell(0,0).text=("Nom du produit, le cas échéant :")
    table.cell(1,0).text=("Nom de code, le cas échéant :")
    table.cell(2,0).text=("Voie d’administration (utiliser les termes standard) :")
    table.cell(3,0).text=("Dosage (préciser tous les dosages utilisés) : \n- Concentration (nombre) : \n- Unité de concentration :")
    table.cell(4,0).text=("Le produit expérimental contient-il une substance active :\n\n- d’origine chimique ?            oui          non	\n- d’origine biologique ?          oui          non\n\n"
                          "Est-ce :\n- un produit à base de plantes ?         oui         non\n- un produit contenant des organismes génétiquement modifiés ?		 oui  	 non\n"
                          "- un autre type de produit ?   oui        non	\n\n• Si oui, préciser :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    
    paragraph=document.add_paragraph("\n Type de produit ", style='debut_page')
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    paragraph=document.add_paragraph("\n Fabricant du produit utilisé", style='debut_page')
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("Fabricant\n- Nom de l’établissement : \n- Adresse :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    '''Partie E'''
    paragraph=document.add_paragraph("\n E. INFORMATIONS SUR LE PLACEBO (le cas échéant) (répéter la section si nécessaire) ", style='debut_page')
    table = document.add_table(rows=6, cols=1, style='Table Grid')
    table.cell(0,0).text=("Cette section se rapporte au placebo n° :  ")
    table.cell(1,0).text=("Un placebo est-il utilisé ? 		   oui           non")
    table.cell(2,0).text=("De quel produit expérimental est-ce un placebo ?")
    table.cell(3,0).text=("Préciser le(s) numéro(s) de PE selon la section D.")
    table.cell(4,0).text=("Voie d’administration :")
    table.cell(5,0).text=("Composition, hormis la (les) substance(s) active(s) : \n- est-elle identique à celle du produit expérimental étudié?       oui           non\n\n•  Si non, préciser les principaux composants :   ")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    paragraph=document.add_paragraph("\n FABRICANT DU PLACEBO", style='debut_page')
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=("Fabricant\n- Nom de l’établissement : \n- Adresse :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
    
    '''Partie G'''
    paragraph=document.add_paragraph("\n G. INFORMATIONS GÉNÉRALES SUR L’ESSAI", style='debut_page')
    table = document.add_table(rows=7, cols=1, style='Table Grid')
    table.cell(0,0).text=("Condition médicale ou pathologie étudiée")
    table.cell(1,0).text=("Préciser la condition médicale :\nClassification CIM  : 	\n\nEst-ce une maladie rare ?       oui           non\n\nObjectif(s) de l’essai\nObjectif principal : p13, 2.1\nObjectifs secondaires :")
    table.cell(2,0).text=("Principaux critères d’inclusion (énumérer les plus importants) ")
    table.cell(3,0).text=("Principaux critères de non inclusion (énumérer les plus importants")
    table.cell(4,0).text=("Critère(s) d’évaluation principal (aux) ")
    table.cell(5,0).text=("Domaine(s) d’étude :\n")
    table.cell(6,0).text=("- Physiologie\n- Physiopathologie\n- Epidémiologie\n- Génétique\n- Science du comportement\n- Produits à visée nutritionnelle\n- Stratégies diagnostiques\n- Stratégies thérapeutiques et préventives\n\n                • Si autres préciser :")
    n=0
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial'
                    fontdebut.size = docx.shared.Pt(9)
                    if n==0 or n==5:
                        fontdebut.bold=True
                    n=n+1
    a=table.cell(0,0)
    b=table.cell(1,0)
    a.merge(b)
    a=table.cell(5,0)
    b=table.cell(6,0)
    a.merge(b)
    
    
    
    
    
    
    
    
    
    
    