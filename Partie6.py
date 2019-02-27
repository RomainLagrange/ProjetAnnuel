# -*- coding: utf-8 -*-
"""
Created on Mon Feb 18 11:51:47 2019

@author: Asuspc
"""

import docx
import StyleProt1
from StyleProt1 import Style,Titre1, Titre2, Titre3, TexteGris, TexteGrisJustif
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_UNDERLINE, WD_LINE_SPACING, WD_COLOR_INDEX, WD_BREAK
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Cm, Pt, RGBColor, Inches

#MEMO POUR ECRIRE LES TITRES :
#    Titre1('num + texte du protocole',document)
#    Titre2('num + texte du protocole',document)
#    Titre3('numero','texte',document)
#    TexteGris(texte,document)
#    TexteGrisJustif(texte,document)

#REVOIR STYLE CHARACTER OU PRAGRAPHE POUR PARAGRAPHE 


def Partie6(document):
#def Partie6():
    'Creation de la partie 6 du protcole de catégorie 1'
  #  document = docx.Document()


#   Marge de la page
#    sections = document.sections
#    for section in sections:
#        section.top_margin = Cm(2)
#        section.bottom_margin = Cm(2)
#        section.left_margin = Cm(2)
#        section.right_margin = Cm(2)

#---------------------------DEFINITIONS DES STYLES
 

    #Style(document)


#    
#---------------------------------------------------------------ECRITURE
    
    
    #ecriture du premier titre 
    Titre1('6	DEROULEMENT DE LA RECHERCHE',document)
    
    
   # Ecriture du 6.1  
    Titre2('6.1	Calendrier de la recherche',document)
    
    # Ecriture du 6.2  
    Titre2('6.2	Tableau récapitulatif du suivi d’un participant à la recherche',document)
    
    #AJOUTER TABLEAU
#    
#    p=document.add_paragraph('(*) V-X : unité de temps à adapter en fonction de la recherche : A (année), M (mois), S (semaine), J (jour), H (heure)', style='Normal')
#    p.font.italic=True
#    
#    p=document.add_paragraph('Examen clinique : détail de ce que comporte l’examen clinique ', style='Normal')
#    
#    p=document.add_paragraph()
#    
#    p=document.add_paragraph()
    
    # Ecriture du 6.3  
    Titre2('6.3	Visites de pré-inclusion / inclusion = Visite V0',document)
    
#    p=document.add_paragraph()
#    run1=p.add_run()
#    run1.text='Lors de la visite de '
#    run1.style='Paragraphe'
#    run2=p.add_run()
#    run2.text='pré-inclusion (voir selon l’étude si visite d’inclusion),'
#    run2.style='Paragraphe'
#    run2.font.italic=True
#    run3=p.add_run()
#    run3.text='le médecin investigateur informe le patient de la possibilité de participer à cet essai clinique et répond à toutes ses questions concernant l\'objectif, la nature des contraintes, les risques prévisibles et les bénéfices attendus de la recherche. Il précise également les droits du patient dans le cadre d’une recherche et vérifie les critères d’éligibilité. '
#    run3.style='Paragraphe'
#    
#    
#    document.add_paragraph('Un exemplaire de la note d’information et du formulaire de consentement est alors remis au participant par le médecin investigateur.', style='Paragraphe')
#    
#    
#    p=document.add_paragraph()
#    run1=p.add_run('Après cette séance d’information, le participant dispose d’un délai de réflexion. Le médecin investigateur est responsable de l’obtention du consentement éclairé écrit du participant.\n Si le participant donne son accord de participation, ce dernier et l’investigateur inscrivent leurs noms et prénoms en clair, datent et signent le formulaire de consentement. Celui-ci ')
#    run1.style='Paragraphe'
#    run2=p.add_run('doit être signé avant la réalisation de tout examen ')
#    run2.style='Paragraphe'
#    run2.font.bold= True
#    run2.font.underline= True
#    run3=p.add_run('clinique ou biologique ou para-clinique nécessité par la recherche. ')
#    run3.style='Paragraphe'
#    
#    p=document.add_paragraph()
#    run1=p.add_run('L’exemplaire ')
#    run1.style='Paragraphe'
#    run2=p.add_run('original ')
#    run2.style='Paragraphe'
#    run2.font.underline= True
#    document.add_paragraph('sera conservé dans le classeur de l’investigateur. Un exemplaire (un autre original ou une copie) sera remis au patient. ', style = 'Paragraphe')
#    document.add_paragraph('L’investigateur précisera dans le dossier médical du patient sa participation à la recherche, les modalités du recueil du consentement ainsi que celle de l’information. ', style = 'Paragraphe')
#    
    #A FINIR 
    
    #Ecriture du titre6.3.1
    Titre3('6.3.1','Recueil du consentement',document)
   
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)
    

    #TEXTE
    
    #Ecriture du titre6.3.2
    Titre3('6.3.2','Déroulement de la visite',document)

    
    #Ecriture du titre 6.4
    Titre2('6.4	Visite de randomisation = Visite (Vx, ou Jx, ou Mx…)',document)

    #Ecriture du titre6.4.1
    Titre3('6.4.1','Description des examens',document)

    
    #Ecriture du titre6.4.2
    Titre3('6.4.2','Randomisation du patient',document)



    #Ecriture du titre 6.5
    Titre2('6.5	Visites de suivi = visite (Vx, ou Jx ou Sx ou Mx…)',document)

    #Ecriture du titre6.5.1
    Titre3('6.5.1','Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)

    
    #Ecriture du titre6.5.2
    Titre3('6.5.2','Visite (Vx, ou Sx, ou Jx, ou Mx…)',document)

    
    #Ecriture du titre 6.6
    Titre2('6.6	Visite de fin de la recherche',document)
    
    #Ecriture du titre 6.7
    Titre2('6.7	Règles d’arrêt de la participation d’une personne à la recherche',document)

    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

	
    #Ecriture du titre6.7.1
    Titre3('6.7.1','Arrêt de participation définitif ou temporaire d’un patient dans l’étude)',document)


    #Ecriture du titre6.7.2
    Titre3('6.7.2','Modalités de remplacement des patients exclus, le cas échéant',document)

    
    #Ecriture du titre6.7.3
    Titre3('6.7.3','Modalités et calendrier de recueil pour ces données',document)

    
    #Ecriture du titre6.7.4
    Titre3('6.7.4','Modalités de suivi de ces personnes',document)

    #Ecriture du titre 6.8
    Titre2('6.8	Contraintes liées à la recherche et indemnisation éventuelle des participants',document)
    
    #Ecriture du titre 6.9
    Titre2('6.9	Collection d’échantillons biologiques',document)
    
    
    TexteGris('prendre contact avec la promotion interne \n pour aide a la redaction de ce chapitre', document)

    
    #Ecriture du titre6.9.1
    Titre3('6.9.1','Objectifs',document)
    
    #Ecriture du titre6.9.2
    Titre3('6.9.2','Description de(s) (la) collection(s) ',document)
    
    #Ecriture du titre6.9.3
    Titre3('6.9.3','Conservation',document)
    
    #Ecriture du titre6.9.4
    Titre3('6.9.4','Devenir de la collection',document)
    
    #Ecriture du titre 6.10
    Titre2('6.10	Arrêt d’une partie ou de la totalité de la recherche',document)
    
    
    #FIN DU DOC 
    paragraph = document.add_paragraph()
    run = paragraph.add_run()
    run.add_break(WD_BREAK.PAGE)
  
   # document.save("Partie6.docx")   