# -*- coding: utf-8 -*-
"""
Created on Thu Jan 31 13:32:03 2019

@author: Julie
"""

import docx
import StyleProt1
from StyleProt1 import Style, Titre1,Titre2, Titre3, TexteGris
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#revoir titre1 et texte encardé gris
#MEMO POUR ECRIRE LES TITRES :
#    StyleProt1.Titre1('num + texte du protocole',document)
#    StyleProt1.Titre2('num + texte du protocole',document)
#    StyleProt1.Titre3('numero','texte',document)
    

def Partie1():
    'Creation de la partie 1 du protcole de catégorie 1'
    document = docx.Document()


#   Marge de la page
    sections = document.sections
    for section in sections:
        section.top_margin = Cm(2)
        section.bottom_margin = Cm(2)
        section.left_margin = Cm(2)
        section.right_margin = Cm(2)
        
    StyleProt1.Style(document)

 
    StyleProt1.Titre1('1	JUSTICATION SCIENTIFIQUE ET DESCRIPTION GENERALE',document)
    #ecriture du premier titre 
#    paragraph=document.add_paragraph('1	JUSTICATION SCIENTIFIQUE ET DESCRIPTION GENERALE\n', style='Titre1') #titre
#    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER #centrer


    #Texte indicatif en italique 
    paragraph1 = document.add_paragraph ('Présentation du problème et justification étayée par les connaissances actuelles avec leurs références à la littérature scientifique et aux données pertinentes.\
    Indiquer en quoi l’objectif est nouveau et utile, pour le progrès des connaissances médicales et/ou de la prise en charge des malades. Les retombées attendues et perspectives peuvent également être développées dans ce chapitre.\n\
    C’est dans ce paragraphe que vous devez justifier la pertinence de votre étude.', style ='TexteItalic') 
    paragraph1.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

    
    # Ecriture du 1.1  
    StyleProt1.Titre2('1.1	Etat actuel des connaissances',document)
  
    
    
    #Ecriture du titre1.1.1
    StyleProt1.Titre3('1.1.1','Sur la pathologie',document)

    
    #Texte indicatif en italique TEST    #UTILISER, style ='TexteItalic'
    paragraph2 = document.add_paragraph() 
    paragraph2.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sentence = paragraph2.add_run('Epidémiologie de la pathologie traitée comportant fréquence et gravité, facteurs de risque, conséquences individuelles et médico-économiques')
    sentence.font.italic=True
    sentence.font.size = docx.shared.Pt(11)
    
    
    
    #ecriture du titre1.1.2
    StyleProt1.Titre3('1.1.2','Sur les traitements, stratégies et procédures de référence et à l’étude',document)
#
#    p=document.add_heading()
#    p.paragraph_format.left_indent = Inches(0.98) #indentation en pouce, ici 1,5cm
#    run1=p.add_run()
#    run1.text='1.1.2     '
#    run1.style='ListeTitre3'
#    run2=p.add_run()
#    run2.text='Sur les traitements, stratégies et procédures de référence et à l’étude\n'
#    run2.style='Titre3'
    

    #Texte indicatif en italique
    paragraph3 = document.add_paragraph ('Décrire les traitements/stratégies/ \
    procédures déjà existants et leurs limites. \
    Décrire les traitements/stratégies/procédures à l’étude\
    pharmacologique et/ou physiopathologique,…)\n Justifier notamment le \
    choix du groupe de comparaison : utilisation d’un traitement/stratégie\
    /procédure de référence, d’un placebo en se fondant sur les connaissances \
    scientifiques actuelles, en particulier les recommandations pour le \
    traitement de la maladie ou de l’état de santé à l’étude.\n Si la recherche\
    porte sur un médicament,préciser et justifier le choix de mener la \
    recherche sur des volontaires sains ou sur des patients, le choix de la \
    forme pharmaceutique, la posologie, le schéma d’administration,la durée du\
    traitement et de la voie d’administration.\n Attention les modalités \
    précises d’administration, le schéma d’administration ne sont pas à \
    détailler ici ; ils seront détaillés dans le paragraphe 7.\n', 
    style='TexteItalic')
    paragraph4 = document.add_paragraph() 
    paragraph4.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sentence = paragraph4.add_run('Exemple : Les voies d’administration, \
    posologie, schéma d’administration et durée de traitement sont ceux de \
    l’AMM de [produit à l’étude] et sont détaillés dans les Résumés des \
    Caractéristiques Produits de cette molécule en annexe de ce protocole.')
    sentence.font.size = docx.shared.Pt(11)
                                        
                                        
    #ecriture du titre1.2
    StyleProt1.Titre2('1.2	Hypothèse de la recherche et résultats attendus',document)

    #Texte indicatif en italique
    paragraph5 = document.add_paragraph ('Définir précisément l’hypothèse, \
    physiopathologique ou autre, qui justifie la mise en place de la recherche\
    , en mentionnant le traitement/ la stratégie/la procédure à l’étude, \
    la population cible (justifier le choix de mener la recherche sur des \
    volontaires sains/patients) et le critère sur lequel il sera jugé.', 
    style='TexteItalic')


    #Ecriture du titre1.3
    StyleProt1.Titre2('1.3 Justification des choix méthodologiques',document)

    
    #Texte sur fond gris  
    TexteGris('prendre contact avec la plateforme de methodologie \n pour aide a la redaction du paragraphe 2.3', document)


    
    #Texte indicatif en italique
    paragraph7 = document.add_paragraph() 
    paragraph7.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    sentence = paragraph7.add_run('Si nécessaire, justifier les choix réalisés\
    pour le schéma de la recherche, le critère de jugement principal, le type \
    de comparaison et la conduite de la recherche.')
    sentence.font.italic=True
    sentence.font.size = docx.shared.Pt(12)

   
     #Ecriture du titre1.4
    StyleProt1.Titre2('1.4 Rapport bénéfices / risques prévisibles',document)
     

    
    #Ecriture du titre1.4.1
    StyleProt1.Titre3('1.4.1','Bénéfices',document)


    #Texte indicatif en italique
    paragraph8 = document.add_paragraph('Expliquer quel(s) est (sont) le(s) \
    bénéfice(s) individuel(s) et collectif(s), le(s) risque(s) prévisible(s) et\
    les contraintes liées à la recherche. Juger le rapport qui permet de \
    proposer ce protocole à l’étude. Indiquer les bénéfices et les risques que \
    présente la recherche, notamment les bénéfices escomptés pour les personnes\
    qui se prêtent à la recherche.', style='TexteItalic')

   
    #Ecriture du titre1.4.2
    StyleProt1.Titre3('1.4.2','Risques',document)


#    #Texte indicatif en italique
    paragraph9 = document.add_paragraph('Décrire les risques prévisibles liés \
    au traitement et aux procédures d’investigation de la recherche (incluant \
    notamment la douleur, l’inconfort, l’atteinte à l’intégrité physique des\
    personnes se prêtant à la recherche, les mesures visant à éviter et/ou \
    prendre en charge les événements inattendus).', style='TexteItalic')
#    
    
     #Ecriture du titre1.5
    StyleProt1.Titre2('1.5 Retombées attendues',document)

     
     #Texte indicatif en italique
    paragraph10 = document.add_paragraph('Description détaillée des retombées \
    attendues par cette recherche (en terme d’amélioration des connaissances \
    sur une pathologie, d’augmentation de l’arsenal thérapeutique,…)..',
    style='TexteItalic')



    document.save("Partie1.docx")   
    