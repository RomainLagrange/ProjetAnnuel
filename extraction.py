# -*- coding: utf-8 -*-
"""
Created on Sat Feb 16 14:09:42 2019

@author: Romain
"""


import re
from docx import Document


def extraction():
    f1 = open('Trame-simplifiée-cat-1.docx', 'rb') #ouvre le premier fichier
    doc = Document(f1)
    fullText=[]
    for para in doc.paragraphs:
          fullText.append(para.text)    
    f1.close()
  #  print(fullText)
    infos={}
    for i in range(len(fullText)):
        if re.search(r"Titre complet de la recherche",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["titre_complet"]=x2
        if re.search(r"Nom ou titre abrégé",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["titre_abrege"]=x2  
        if re.search(r"N° de code du protocole attribué par le promoteur",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["code_protocole"]=x2
        if re.search(r"N° de code du protocole attribué par le promoteur",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["code_protocole"]=x2
        if re.search(r"N°EudraCT",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["num_eudract"]=x2
        if re.search(r"N°IDRCB",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["num_idrcb"]=x2
        if re.search(r"Classification CIM",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["classification_cim"]=x2
        if re.search(r"Préciser la condition médicale ou pathologie étudiée",fullText[i]):
            x=fullText[i+1]
            x2=x.replace("\xa0","")
            infos["pathologie_etudiee"]=x2
    return infos

def extract():
    
    #dabord on extrait tout qu'on ajoute dans une liste
    f1 = open('Trame-simplifiée-cat-1 (test).docx', 'rb') #ouvre le premier fichier
    doc = Document(f1)
    fullText=[]
    for para in doc.paragraphs:
          fullText.append(para.text)   
    f1.close()
    
    #puis on met tous les éléments de la liste bout a bout dans un immense string
    texte=""
    for i in fullText:
        texte+=i
    texte=texte.replace("\xa0","")
    texte=texte.replace("\n","")
    print(texte)
    #creation du dico de donnes
    infos={}
    #on ajoute les éléments 1 par 1
    x=re.search(r"(?<=Titre complet de la recherche: ).*(?=Nom ou titre)",texte).group()
    infos["titre_complet"]=x
    x=re.search(r"(?<=Nom ou titre abrégé: ).*(?=N° de code du protocole)",texte).group()
    infos["titre_abrege"]=x
    x=re.search(r"(?<=Protocole ... version n°… du .../.../…\): ).*(?=N°EudraCT)",texte).group()
    infos["code_protocole"]=x
    x=re.search(r"(?<=N°EudraCT: ).*(?=N°IDRCB:)",texte).group()
    infos["num_eudract"]=x
    x=re.search(r"(?<=N°IDRCB: ).*(?=Classification CIM: )",texte).group()
    infos["num_idrcb"]=x
    x=re.search(r"(?<=Classification CIM: ).*(?=Préciser la condition médicale)",texte).group()
    infos["classification_cim"]=x
    x=re.search(r"(?<=condition médicale ou pathologie étudiée: ).*(?=Identification du promoteur responsable)",texte).group()
    infos["pathologie_etudiee"]=x
    
    #en premier on récupère tout le bloc promoteur
    x=re.search(r"(?<=Identification du promoteur responsable de la demande : Promoteur ).*(?=Représentant légal du promoteur dans l’UE)",texte).group()
    #puis tous les éléments du promoteur 1 par 1
    y=re.search(r"(?<=Nom de l’organisme: ).*(?=Nom de la personne à contacter:)",x).group()
    infos['promoteur_nom_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter: ).*(?=Adresse:)",x).group()
    infos['promoteur_nom_personne_contact']=y
    y=re.search(r"(?<=Adresse: ).*(?=N° téléphone:)",x).group()
    infos['promoteur_adresse']=y
    y=re.search(r"(?<=N° téléphone: ).*(?=N° télécopie:)",x).group()
    infos['promoteur_num_telephone']=y
    y=re.search(r"(?<=N° télécopie: ).*(?=Courriel:)",x).group()
    infos['promoteur_num_telecopie']=y
    y=re.search(r"(?<=Courriel: ).*",x).group()
    infos['promoteur_courriel']=y
    
    x=re.search(r"(?<=Représentant légal du promoteur dans l’UE ).*(?=Identification des investigateurs)",texte).group()
    #puis tous les éléments du promoteur 1 par 1
    y=re.search(r"(?<=Nom de l’organisme: ).*(?=Nom de la personne à contacter:)",x).group()
    infos['promoteur_UE_nom_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter: ).*(?=Adresse:)",x).group()
    infos['promoteur_UE_nom_personne_contact']=y
    y=re.search(r"(?<=Adresse: ).*(?=N° téléphone:)",x).group()
    infos['promoteur_UE_adresse']=y
    y=re.search(r"(?<=N° téléphone: ).*(?=N° télécopie:)",x).group()
    infos['promoteur_UE_num_telephone']=y
    y=re.search(r"(?<=N° télécopie: ).*(?=Courriel:)",x).group()
    infos['promoteur_UE_num_telecopie']=y
    y=re.search(r"(?<=Courriel: ).*",x).group()
    infos['promoteur_UE_courriel']=y
    
    #idem pour investigateur coordinateur
    x=re.search(r"(?<=Investigateur coordinateur: ).*(?=Autres investigateurs: )",texte).group()
    #puis tous les éléments de l'investigateur coordinateur
    y=re.search(r"(?<=Nom: ).*(?=Prénom:)",x).group()
    infos['investigateur_coordinateur_nom']=y
    y=re.search(r"(?<=Prénom: ).*(?=Qualification, spécialité: )",x).group()
    infos['investigateur_coordinateur_prenom']=y
    y=re.search(r"(?<=Qualification, spécialité: ).*(?=Adresse professionnelle:)",x).group()
    infos['investigateur_coordinateur_qualification']=y
    y=re.search(r"(?<=Adresse professionnelle: ).*(?=Nom de l’établissement: )",x).group()
    infos['investigateur_coordinateur_adresse_professionnelle']=y
    y=re.search(r"(?<=Nom de l’établissement: ).*(?=Service: )",x).group()
    infos['investigateur_coordinateur_nom_etablissement']=y
    y=re.search(r"(?<=Service: ).*(?=Adresse: )",x).group()
    infos['investigateur_coordinateur_service']=y
    y=re.search(r"(?<=Adresse: ).*(?=N° téléphone:)",x).group()
    infos['investigateur_coordinateur_adresse']=y
    y=re.search(r"(?<=N° téléphone: ).*(?=N° télécopie:)",x).group()
    infos['investigateur_coordinateur_telephone']=y
    y=re.search(r"(?<=N° télécopie: ).*(?=Courriel:)",x).group()
    infos['investigateur_coordinateur_telecopie']=y
    y=re.search(r"(?<=Courriel: ).*",x).group()
    infos['investigateur_coordinateur_courriel']=y
    
    #pour les autres investigateurs
    x=re.search(r"(?<=Autres investigateurs: ).*(?=Identification du demandeur: )",texte).group()
    #ressemble au reste sauf que pour gerer le fait qu'il puisse exister plusieurs autres investigateurs
    #on utilise findall plutot que search, permet de stocker les differents points des investigateurs
    #dans une liste
    y=re.findall(r"(?<=Nom: ).*?(?=Prénom:)",x)
    infos['autre_investigateur_nom']=y
    y=re.findall(r"(?<=Prénom: ).*?(?=Qualification, spécialité: )",x)
    infos['autre_investigateur_prenom']=y
    y=re.findall(r"(?<=Qualification, spécialité: ).*?(?=Adresse professionnelle:)",x)
    infos['autre_investigateur_qualification']=y
    y=re.findall(r"(?<=Adresse professionnelle: ).*?(?=Nom de l’établissement: )",x)
    infos['autre_investigateur_adresse_professionnelle']=y
    y=re.findall(r"(?<=Nom de l’établissement: ).*?(?=Service: )",x)
    infos['autre_investigateur_nom_etablissement']=y
    y=re.findall(r"(?<=Service: ).*?(?=Adresse: )",x)
    infos['autre_investigateur_service']=y
    y=re.findall(r"(?<=Adresse: ).*?(?=N° téléphone:)",x)
    infos['autre_investigateur_adresse']=y
    y=re.findall(r"(?<=N° téléphone: ).*?(?=N° télécopie:)",x)
    infos['autre_investigateur_telephone']=y
    y=re.findall(r"(?<=N° télécopie: ).*?(?=Courriel:)",x)
    infos['autre_investigateur_telecopie']=y
    y=re.findall(r"(?<=Courriel: ).*?(?=Nom)",x)
    z=re.findall(r"(?<=Courriel: ).*?(?=Identification)",x)
    infos['autre_investigateur_courriel']=y+z
    
    
    return infos
    
    
    
    
    
    
    
    
    