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
    print(fullText)
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
    f1 = open('Trame-simplifiée-cat-1.docx', 'rb') #ouvre le premier fichier
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
    
    
    return infos
    
    
    
    
    
    
    
    
    