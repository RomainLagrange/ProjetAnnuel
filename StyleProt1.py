# -*- coding: utf-8 -*-
"""
Created on Fri Feb 15 17:50:09 2019

@author: Julie
"""
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


def Titre1(texte, document):
    paragraph=document.add_paragraph(texte+'\n', style='Titre1') #titre
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER #centrer
    
def Titre2(texte,document):
    document.add_paragraph(texte+'\n', style='Titre2')

def Titre3(num, texte, document):
    p=document.add_heading()
    p.paragraph_format.left_indent = Inches(0.98) #indentation en pouce, ici 1,5cm
    run1=p.add_run()
    run1.text=num+'	    '
    run1.style='ListeTitre3'
    run2=p.add_run()
    run2.text=texte+'\n'
    run2.style='Titre3'

def TexteGris(texte,document):
    shading_elm = parse_xml(r'<w:shd {} w:fill="D9D9D9"/>'.format(nsdecls('w')))
    table = document.add_table(rows = 1, cols = 1)
    row = table.rows[0].cells
    para_text =texte
    cell = row[0]
    pt = cell.paragraphs[0]
    t = pt.text = ''
    p = pt.add_run(para_text)
    cell._tc.get_or_add_tcPr().append(shading_elm)
    p.style='BackgroundGrey'
    pt.alignment=WD_ALIGN_PARAGRAPH.CENTER

def Style(document):
    'Défintion des styles du protocole de catégorie 1'
    
# TROUVER COMMENT PARTAGER LES STYLES A TOUTES LES PAGES SONT COPIER TOUT
# LE CODE    
    
    styles=document.styles

#   definition du style Titre1 --> AJOUTER LA BORDURE EN BAS
    styleTitre1 = styles.add_style('Titre1', WD_STYLE_TYPE.PARAGRAPH, WD_ALIGN_PARAGRAPH.CENTER)
    styleTitre1.base_style = styles['Heading1']
    fontTitre1 = styleTitre1.font
    fontTitre1.name = 'Times New Roman' #police
    fontTitre1.size = docx.shared.Pt(12) #taille
    fontTitre1.all_caps = True #toujours en majuscule
    fontTitre1.bold= True #en gras
    fontTitre1.color.rgb = RGBColor(0x0,0x70,0xC0) #couleur bleu, en base 16
    

    #Definition du Titre2, correspond par exemple au 1.1 ou 1.2
    styleTitre2 = styles.add_style('Titre2', WD_STYLE_TYPE.PARAGRAPH)
    styleTitre2.base_style = styles['Heading2']
    fontTitre2 = styleTitre2.font
    fontTitre2.name = 'Times New Roman'
    fontTitre2.size = docx.shared.Pt(14)
    fontTitre2.bold= True
    fontTitre2.color.rgb = RGBColor(0x0,0x0,0x0)
    styleTitre2.paragraph_format.left_indent = Inches(0.59)
    
    
     #Definition du Titre3; correspond aux 1.1.1 ou 1.1.2...
    styleTitre3 = styles.add_style('Titre3', WD_STYLE_TYPE.CHARACTER)
  #  styleTitre3 = styles.add_style('Titre3', WD_STYLE_TYPE.CHARACTER)
    styleTitre3.base_style = styles['Heading3']
    fontTitre3 = styleTitre3.font
    fontTitre3.name = 'Times New Roman'
    fontTitre3.size = docx.shared.Pt(12)
    fontTitre3.bold= False
    fontTitre3.underline= True
#    styleTitre3.paragraph_format.left_indent = Inches(0.98) #indentation en pouce, ici 2,5cm
    fontTitre3.color.rgb = RGBColor(0x0,0x0,0x0)
    
    
   #Définition du ListeTitre3, correspond au nom du titre après le 1.1.1 ou 1.2.1    
    styleTitreListe3 = styles.add_style('ListeTitre3', WD_STYLE_TYPE.CHARACTER)
    styleTitreListe3.base_style = styles['Heading3']
    fontTitreListe3 = styleTitreListe3.font
    fontTitreListe3.name = 'Times New Roman'
    fontTitreListe3.size = docx.shared.Pt(12)
    fontTitreListe3.bold= True
    fontTitreListe3.underline= False
    fontTitreListe3.color.rgb = RGBColor(0x0,0x0,0x0)  
    
    #Definition style texte surligné en gris   --> SUPPRIMER ESPACE EN BAS
    styles = document.styles
    styleBackgroundGrey = styles.add_style('BackgroundGrey', WD_STYLE_TYPE.CHARACTER)
    styleBackgroundGrey.base_style = styles['No Spacing']
    fontBackgroundGrey = styleBackgroundGrey.font
    fontBackgroundGrey.name = 'Times New Roman'
    fontBackgroundGrey.size = docx.shared.Pt(11)
    fontBackgroundGrey.bold = True
    fontBackgroundGrey.small_caps = True
    
        #definition du style pour le texte indicatif -->  ESPACEMENT LIGNES 
    styleIndic = styles.add_style('TexteItalic', WD_STYLE_TYPE.PARAGRAPH)
    styleIndic.base_style = styles['Normal']
    fontIndic = styleIndic.font
    fontIndic.name = 'Times New Roman'   
    fontIndic.size = docx.shared.Pt(11)
    fontIndic.italic= True
    

