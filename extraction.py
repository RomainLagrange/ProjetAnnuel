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