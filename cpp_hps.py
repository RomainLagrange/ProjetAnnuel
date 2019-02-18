# -*- coding: utf-8 -*-
"""
Created on Fri Feb  1 18:17:55 2019

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
#from docx.shared import RGBColor

#docmuents du cpp pour les dispositifs médicaux

def main_cpp_hps():
    document = docx.Document()
    cpp_hps(document)
    page2_cpp_hps(document)
    document.save("soumission-cpp-hps.docx")

def cpp_hps(document):
    

    
    '''Marge de la page'''
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
    
    '''Titre CPP'''
  #  paragraph = document.add_paragraph()
    styles= document.styles
    style1 = styles.add_style('Debut', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style1.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style1.font
    fontdebut.name = 'Book Antiqua'
    fontdebut.bold = True
    fontdebut.size = docx.shared.Pt(20) 
    
    
    
    paragraph1 = document.add_paragraph('Comité de Protection des Personnes', style='Debut')
    paragraph1.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    #Sous titre
     
    
    paragraph = document.add_paragraph('OUEST III', style='Debut')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    #ajouter le trait et les ombres
    
    #Infos promoteur
    
    style2=styles.add_style('Promoteur', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style2.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style2.font
    fontdebut.name = 'Book Antiqua'
    fontdebut.italic = True
    fontdebut.size = docx.shared.Pt(10) 
      
    paragraph = document.add_paragraph("Agréé par arrêté ministériel en date du 31 mai 2012, \nConstitué selon l'arrêté du Directeur Général de l'ARS Poitou Charentes en date du 25 juin 2012.\n\n"
                                       "C.H.U La Milétrie\nPavillon Administratif - Porte 213\n "
                                       "2 rue de le milétrie - CS 90 577 - 86021 POITIERS CEDEX\n"
                                       "Tel : 05.49.45.21.57\nFax : 05.49.46.12.62 \nE-mail : cpp-ouest3@chu-poitiers.fr \n", style='Promoteur')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    
    #titre du milieu qui dit pour quel proto
    
    paragraph = document.add_paragraph()
    paragraph_format.space_before = Pt(0)
    paragraph_format.space_after = Pt(0)
    sentence = paragraph.add_run("Demande d'avis au CPP")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sentence.font.name = 'Arial Narrow'
    sentence.bold = True
    sentence.font.size = docx.shared.Pt(12)
    sentence2 = paragraph.add_run(" (arrêté du 2 décembre 2016)\n")
    sentence2.font.name = 'Arial Narrow'
    sentence2.font.size = docx.shared.Pt(12)
    sentence3 = paragraph.add_run("sur un projet de recherche mentionnée au 1o ou au 2o de l’article L. 1121-1 du code de la santé publique ne portant pas sur un produit mentionné à l’article L. 5311-1 du même code.")
    sentence3.font.name = 'Arial Narrow'
    sentence3.bold = True
    sentence3.font.size = docx.shared.Pt(12)
    sentence4 = paragraph.add_run("(les médicaments, les produits contraceptifs, les biomatériaux et les dispositifs médicaux …)\n")
    sentence4.font.name = 'Arial Narrow'
    sentence4.font.size = docx.shared.Pt(12)
    
 ###########################################   
 
    #ENtre le titre et le tableau
 
    style3=styles.add_style('Avant_tableau', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style3.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
#    paragraph_format.left_indent = Inches(10)
    fontdebut = style3.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.italic = True
    fontdebut.size = docx.shared.Pt(10) 
    
    paragraph = document.add_paragraph("Préalablement au dépôt du dossier le promoteur obtient un numéro d’enregistrement sur le site internet de l’ANSM. Ce numéro identifie chaque recherche réalisée en France.", style='Avant_tableau')
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
  
 ##############################

    #Tableau central
    table=document.add_table(rows=15, cols=1, style='Table Grid')
    

    table.cell(0,0).text=("DOSSIER ADMINISTRATIF")
    table.cell(1,0).text=("1 Courrier de demande d’avis daté et signé")
    table.cell(2,0).text=("2 Formulaire de demande d’avis (annexe 1)")
    table.cell(3,0).text=("3 Document additionnel (annexe 2) + supports pour recrutement des personnes")
    table.cell(4,0).text=("8.2 Pour les recherches mentionnées au 1o de l’article L. 1121-1, si nécessaire, la copie de la ou des autorisations de lieux de recherches mentionnées à l’article L. 1121-13 du CSP")
    table.cell(5,0).text=("DOSSIER SUR LA RECHERCHE")
    table.cell(6,0).text=("4 Protocole de recherche (daté + numéro de version)")
    table.cell(7,0).text=("5 Résumé du protocole (daté + numéro de version)")
    table.cell(8,0).text=("Le cas échéant, la brochure pour l’investigateur mentionnée à l’article R. 1123-20 du code de la santé publique, datée et comportant un numéro de version, lorsque la recherche porte sur un produit autre que ceux mentionnés à l’article L. 5311-1 du CSP")
    table.cell(9,0).text=("6.1 Document d’information sauf situation art. L. 1122-1-4")
    table.cell(10,0).text=("6.2 Formulaire de consentement sauf situation art. L. 1122-1-4")
    table.cell(11,0).text=("7 Attestation d’assurance (Décret n°2016-1537 du 16 novembre 2016 - art. 3)")
    table.cell(12,0).text=("8 Une justification de l’adéquation des moyens humains, matériels et techniques au projet de recherche et de leur compatibilité avec les impératifs de sécurité des personnes qui s’y prêtent, sauf si le lieu bénéficie de l’autorisation mentionnée à l’article L. 1121-13 du CSP")
    table.cell(13,0).text=("8.1 Curriculum vitae signé du ou des investigateurs datant d’un an maximum")
    table.cell(14,0).text=("Le cas échéant, la nature de la décision finale de l’ANSM, si disponible.")
    n=1
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                for run in paragraph.runs:
                    if n==1 or n==6:
                        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        #run.bold = True
                        font = run.font
                        font.size= docx.shared.Pt(11)
                        font.name = 'Arial'
                    else:
                        font = run.font
                        font.size= docx.shared.Pt(10)
                        font.name = 'Arial Narrow'
                        #paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY_LOW
                    n=n+1
    shading_elm_1 = parse_xml(r'<w:shd {} w:fill="AFAFAF"/>'.format(nsdecls('w')))
    table.rows[0].cells[0]._tc.get_or_add_tcPr().append(shading_elm_1)
    shading_elm_2 = parse_xml(r'<w:shd {} w:fill="AFAFAF"/>'.format(nsdecls('w')))
    table.rows[5].cells[0]._tc.get_or_add_tcPr().append(shading_elm_2)
    
    style5=styles.add_style('fin_tableau', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style5.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style5.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.italic = True
    fontdebut.size = docx.shared.Pt(10) 
    document.add_paragraph('Forme : 4 dossiers complets + 1 version électronique\n\n', style='fin_tableau')
    
    
def page2_cpp_hps(document):
    document.add_page_break()
    styles= document.styles
    style=styles.add_style('debut_page', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.size = docx.shared.Pt(10) 
    
    paragraph = document.add_paragraph("Annexe 1\nFormulaire de damande d'avis au comité de protection des personnes pour une recherche\nmentionnée au 1° ou au 2° de l'article L.1121-1 du code de la santé publique et ne portant pas\nsur un produit de santé\n", style="debut_page")
    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    sentence=("Demande d'avis au comite de protection des personnes pour une recherche\nmentionnee au 1° ou 2° de l'article L.1121-1 du code de la sante publique et ne\nportant pas sur un produit mentionne a\nl'article L. 5311-1 du code de la sante publique\n")
    sentence.upper()
    
    table = document.add_table(rows=1, cols=1, style='Table Grid')
    table.cell(0,0).text=sentence
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.bold = True
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
                    fontdebut.color.rgb = RGBColor(0x0,0x70,0xC0)
    
    style=styles.add_style('gras_tableau', WD_STYLE_TYPE.PARAGRAPH)
    paragraph_format = style.paragraph_format
    paragraph_format.space_before
    paragraph_format.space_after
    fontdebut = style.font
    fontdebut.bold = True
    fontdebut.name = 'Arial Narrow'
    fontdebut.size = docx.shared.Pt(10) 
    
    paragraph = document.add_paragraph("Partie réservée au Comité de protection des personnes (CPP)", style="gras_tableau")
               
    table = document.add_table(rows=2, cols=3, style='Table Grid')
    table.cell(0,0).text=("Date d'enregistrement de la\ndemande considérée complète :")
    table.cell(0,1).text=("Date de réception des informations\ncomplémentaires / amendées :")
    table.cell(0,2).text=("Avis du CPP :")
    a=table.cell(1,0)
    b=table.cell(1,2)
    a.merge(b)
    table.cell(1,0).text=("Date du début de procédure :")
    for row in table.rows:
        for cell in row.cells:
            paragraphs = cell.paragraphs
            for paragraph in paragraphs:
                paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
                for run in paragraph.runs:
                    fontdebut = run.font
                    fontdebut.name = 'Arial Narrow'
                    fontdebut.size = docx.shared.Pt(10)
    
    paragraph = document.add_paragraph("Partie à compléter par le demandeur :", style="gras_tableau")
    paragraph=document.add_paragraph()
    sentence = paragraph.add_run("RECHERCHE MENTIONNEE AU 1° de l'article L.1121-1 □             RECHERCHE MENTIONNEE AU 2° DE L'ARTICLE L.1121-1□\nDEMANDE D'AUTORISATION A L'ANSM :")
    fontdebut = sentence.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.size = docx.shared.Pt(10) 
    sentence = paragraph.add_run("     oui        non\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.bold=True
    fontdebut.size = docx.shared.Pt(10)
    sentence = paragraph.add_run("DEMANDE D'AVIS AU CPP :     ")
    fontdebut = sentence.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.size = docx.shared.Pt(10) 
    sentence = paragraph.add_run("     oui        non\n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.bold=True
    fontdebut.size = docx.shared.Pt(10)
    
    paragraph = document.add_paragraph("A. IDENTIFICATION DE LA RECHERCHE", style="gras_tableau")
    paragraph=document.add_paragraph()
    sentence = paragraph.add_run("Titre complet de la recherche :\n \nNuméro d'enregistrement de la recherche (délivré par l'ANSM) : \n")
    fontdebut = sentence.font
    fontdebut.name = 'Arial Narrow'
    fontdebut.size = docx.shared.Pt(10) 
    
    
    
 #   modifyBorder(table)


#def modifyBorder(table):
#    tbl = table._tbl # get xml element in table
#    for cell in tbl.iter_tcs():
#        tcPr = cell.tcPr # get tcPr element, in which we can define style of borders
#        tcBorders = OxmlElement('w:tcBorders')
#        top = OxmlElement('w:top')
#        top.set(qn('w:val'), 'nil')
#        
#        left = OxmlElement('w:left')
#        left.set(qn('w:val'), 'nil')
#        
#        bottom = OxmlElement('w:bottom')
#        bottom.set(qn('w:color'), 'blue')
#
#        right = OxmlElement('w:right')
#        right.set(qn('w:color'), 'blue')
#        
#        left = OxmlElement('w:right')
#        left.set(qn('w:color'), 'blue')
#        
#        top = OxmlElement('w:right')
#        top.set(qn('w:color'), 'blue')
#
#        tcBorders.append(top)
#        tcBorders.append(left)
#        tcBorders.append(bottom)
#        tcBorders.append(right)
#        tcPr.append(tcBorders)
#    
#    '''Marge de la page'''
#    sections = document.sections
#    for section in sections:
#        section.top_margin = Cm(1)
#        section.bottom_margin = Cm(2)
#        section.left_margin = Cm(2)
#        section.right_margin = Cm(2)
#        
#    styles= document.styles
#    style1 = styles.add_style('Debut_page2', WD_STYLE_TYPE.PARAGRAPH)
#    paragraph_format = style1.paragraph_format
#    paragraph_format.space_before
#    paragraph_format.space_after
#    fontdebut = style1.font
#    fontdebut.name = 'Arial Narrow'
#    fontdebut.size = docx.shared.Pt(12) 
#    
#    paragraph=document.add_paragraph()
#    sentence=paragraph.add_run('Annexe 1\n')
#    sentence.font.name = 'Arial Narrow'
#    sentence.font.size = docx.shared.Pt(10.5)
#    sentence2=paragraph.add_run('DOCUMENT ADDITIONNEL\n')
#    sentence2.bold = True
#    sentence2.font.name = 'Arial Narrow'
#    sentence2.font.size = docx.shared.Pt(12)
#    sentence3=paragraph.add_run('À LA DEMANDE D’AVIS AU COMITÉ DE PROTECTION DES PERSONNES SUR UN PROJET DE \nRECHERCHE MENTIONNÉE AU 1° OU AU 2° DE L’ARTICLE L. 1121-1 PORTANT SUR UN\n')
#    sentence3.font.name = 'Arial Narrow'
#    sentence3.font.size = docx.shared.Pt(12)
#    sentence4=paragraph.add_run('DISPOSITIF MÉDICAL OU UN DISPOSITIF MÉDICAL DE DIAGNOSTIC')
#    sentence4.bold = True
#    sentence4.font.name = 'Arial Narrow'
#    sentence4.font.size = docx.shared.Pt(12)
#    sentence5=paragraph.add_run(' IN VITRO')
#    sentence5.bold = True
#    sentence5.italic = True
#    sentence5.font.name = 'Arial Narrow'
#    sentence5.font.size = docx.shared.Pt(12)
#    paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER   
#    
#    style3 = styles.add_style('page2_normal', WD_STYLE_TYPE.PARAGRAPH)
#    paragraph_format = style3.paragraph_format
#    paragraph_format.space_before
#    paragraph_format.space_after
#    fontdebut = style3.font
#    fontdebut.name = 'Arial Narrow'
#    fontdebut.size = docx.shared.Pt(11)
#    
#    paragraph=document.add_paragraph('\nCe document doit être complété de façon claire, compréhensible et en français.\n', style='page2_normal')
#  
#    
#    paragraph=document.add_paragraph('1. Numéro d’enregistrement de la recherche :', style='page2_normal')
#    paragraph=document.add_paragraph('2. Titre complet de la recherche :', style='page2_normal')
#    paragraph=document.add_paragraph('3. Justification de la recherche :', style='page2_normal')
#    paragraph=document.add_paragraph('4. Hypothèse principale de la recherche et objectifs :', style='page2_normal')
#    paragraph=document.add_paragraph('5. Evaluation des bénéfices et des risques que présente la recherche, notamment les bénéfices escomptés pour les personnes qui se prêtent à la recherche et les risques prévisibles liés à l’utilisation des produits et aux procédures d’investigation de la recherche (incluant notamment la douleur, l’inconfort, l’atteinte à l’intégrité physique des personnes se prêtant à la recherche, les mesures visant à éviter et/ou prendre en charge les événements) :', style='page2_normal')
#    paragraph=document.add_paragraph('6. Justifications de l’inclusion de personnes visées aux articles L. 1121-5 à L. 1121-8 et L. 1122-1-2 du code de la santé publique (notamment mineurs, majeurs protégés, recherches mises en oeuvre dans des situations d’urgence) et procédure mise en oeuvre afin d’informer et recueillir le consentement de ces personnes ou de leurs représentants légaux :', style='page2_normal')
#    paragraph=document.add_paragraph('7. Description des modalités de recrutement des personnes (joindre notamment tous les supports publicitaires utilisés pour la recherche en vue du recrutement des personnes) :', style='page2_normal')
#    paragraph=document.add_paragraph('8. Procédures d’investigation menées et différences par rapport aux conditions habituelles d’utilisation du dispositif médical ou dispositif médical de diagnostic in vitro, le cas échéant :', style='page2_normal')
#    paragraph=document.add_paragraph('9. Justification de l’existence ou non : i) d’une interdiction de participer simultanément à une autre recherche ; ii) d’une période d’exclusion pendant laquelle la participation à une autre recherche est interdite.', style='page2_normal')   
#    paragraph=document.add_paragraph('10. Modalités et montant de l’indemnisation des personnes se prêtant à la recherche, le cas échéant :', style='page2_normal')
#    paragraph=document.add_paragraph('11. Motifs de constitution ou non d’un comité de surveillance indépendant :', style='page2_normal')
#    paragraph=document.add_paragraph('12. Nombre prévu de personnes à inclure dans la recherche :\n', style='page2_normal')
#    paragraph=document.add_paragraph('Par la présente, j’atteste/j’atteste au nom du promoteur (rayer la mention inutile) que les informations fournies ci-dessus à l’appui de la demande d’avis sont exactes.\n', style='page2_normal')
#    paragraph=document.add_paragraph('Nom :\nPrénom :\nAdresse :\nFonction :\nDate :\nSignature :', style='page2_normal')
#    