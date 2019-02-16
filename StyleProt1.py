# -*- coding: utf-8 -*-
"""
Created on Fri Feb 15 17:50:09 2019

@author: Julie
"""
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches


def Style():
    'Défintion des styles du protocole de catégorie 1'
    
# TROUVER COMMENT PARTAGER LES STYLES A TOUTES LES PAGES SONT COPIER TOUT
# LE CODE    


    document = docx.Document()
    styles = document.styles

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
    
    
    #Definition du Titre3; correspond aux 1.1.1 ou 1.1.2... EN DEUX PARTIES car deux styles sur une ligne...
#      1) 
    styleTitre3 = styles.add_style('Titre3', WD_STYLE_TYPE.CHARACTER)
    styleTitre3.base_style = styles['Heading3']
    fontTitre3 = styleTitre3.font
    fontTitre3.name = 'Times New Roman'
    fontTitre3.size = docx.shared.Pt(12)
    fontTitre3.bold= False
    fontTitre3.underline= True
    fontTitre3.color.rgb = RGBColor(0x0,0x0,0x0)
    
#    2)
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


    return (styleTitre1,styleTitre2,styleTitre3,styleTitreListe3)

