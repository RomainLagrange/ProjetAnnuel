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
    f1 = open('Trame-simplifiée-cat-1(test).docx', 'rb') #ouvre le premier fichier
    doc = Document(f1)
    fullText=[]
    for para in doc.paragraphs:
          fullText.append(para.text)   
    f1.close()
    
    #puis on met tous les éléments de la liste bout a bout dans un immense string
    texte1=""
    for i in fullText:
        texte1+=i
    texte1=texte1.replace("\xa0","")
    texte=texte1.replace("\n","")

    #creation du dico de donnes
    infos={}
    #on ajoute les éléments 1 par 1
    x=re.search(r"(?<=Titre complet de la recherche:).*(?=Nom ou titre)",texte).group()
    infos["titre_complet"]=x
    x=re.search(r"(?<=Nom ou titre abrégé:).*(?=N° de code du protocole)",texte).group()
    infos["titre_abrege"]=x
    x=re.search(r"(?<=Protocole ... version n°… du .../.../…\):).*(?=N°EudraCT)",texte).group()
    infos["code_protocole"]=x
    x=re.search(r"(?<=N°EudraCT:).*(?=N°IDRCB:)",texte).group()
    infos["num_eudract"]=x
    x=re.search(r"(?<=N°IDRCB:).*(?=Classification CIM: )",texte).group()
    infos["num_idrcb"]=x
    x=re.search(r"(?<=Classification CIM:).*(?=Préciser la condition médicale)",texte).group()
    infos["classification_cim"]=x
    x=re.search(r"(?<=condition médicale ou pathologie étudiée:).*(?=Identification du promoteur responsable)",texte).group()
    infos["pathologie_etudiee"]=x
    
    #en premier on récupère tout le bloc promoteur
    x=re.search(r"(?<=Identification du promoteur responsable de la demande : Promoteur ).*(?=Représentant légal du promoteur dans l’UE)",texte).group()
    #puis tous les éléments du promoteur 1 par 1
    y=re.search(r"(?<=Nom de l’organisme:).*(?=Nom de la personne à contacter:)",x).group()
    infos['promoteur_nom_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter:).*(?=Adresse:)",x).group()
    infos['promoteur_nom_personne_contact']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['promoteur_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['promoteur_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['promoteur_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['promoteur_courriel']=y
    
    x=re.search(r"(?<=Représentant légal du promoteur dans l’UE ).*(?=Identification des investigateurs)",texte).group()
    #puis tous les éléments du promoteur 1 par 1
    y=re.search(r"(?<=Nom de l’organisme:).*(?=Nom de la personne à contacter:)",x).group()
    infos['promoteur_UE_nom_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter:).*(?=Adresse:)",x).group()
    infos['promoteur_UE_nom_personne_contact']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['promoteur_UE_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['promoteur_UE_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['promoteur_UE_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['promoteur_UE_courriel']=y
    
    #idem pour investigateur coordinateur
    x=re.search(r"(?<=Investigateur coordinateur:).*(?=Autres investigateurs: )",texte).group()
    #puis tous les éléments de l'investigateur coordinateur
    y=re.search(r"(?<=Nom:).*(?=Prénom:)",x).group()
    infos['investigateur_coordinateur_nom']=y
    y=re.search(r"(?<=Prénom:).*(?=Qualification, spécialité: )",x).group()
    infos['investigateur_coordinateur_prenom']=y
    y=re.search(r"(?<=Qualification, spécialité:).*(?=Adresse professionnelle:)",x).group()
    infos['investigateur_coordinateur_qualification']=y
    y=re.search(r"(?<=Adresse professionnelle:).*(?=Nom de l’établissement: )",x).group()
    infos['investigateur_coordinateur_adresse_professionnelle']=y
    y=re.search(r"(?<=Nom de l’établissement:).*(?=Service: )",x).group()
    infos['investigateur_coordinateur_nom_etablissement']=y
    y=re.search(r"(?<=Service:).*(?=Adresse: )",x).group()
    infos['investigateur_coordinateur_service']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['investigateur_coordinateur_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['investigateur_coordinateur_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['investigateur_coordinateur_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['investigateur_coordinateur_courriel']=y
    
    #pour les autres investigateurs, plus subtile
    x=re.search(r"(?<=Autres investigateurs: ).*(?=Identification du demandeur: )",texte).group()
    #je m'en sers pour les regex
    x+="Nom: "

    #initialise avec des listes vides
    infos['autre_investigateur_nom']=[]
    infos['autre_investigateur_prenom']=[]
    infos['autre_investigateur_qualification']=[]
    infos['autre_investigateur_adresse_professionnelle']=[]
    infos['autre_investigateur_nom_etablissement']=[]
    infos['autre_investigateur_service']=[]
    infos['autre_investigateur_adresse']=[]
    infos['autre_investigateur_telephone']=[]
    infos['autre_investigateur_telecopie']=[]
    infos['autre_investigateur_courriel']=[]
    #compte le nombre d'autres investigateurs
    n=x.count('Nom: ')
    #on boucle pour chaque investigateur
    for i in range(n-1): 
        infos['autre_investigateur_nom'].append(re.search(r"(?<=Nom:).*?(?=Prénom:)",x).group())
        infos['autre_investigateur_prenom'].append(re.search(r"(?<=Prénom:).*?(?=Qualification, spécialité: )",x).group())
        infos['autre_investigateur_qualification'].append(re.search(r"(?<=Qualification, spécialité:).*?(?=Adresse professionnelle:)",x).group())
        infos['autre_investigateur_adresse_professionnelle'].append(re.search(r"(?<=Adresse professionnelle:).*?(?=Nom de l’établissement: )",x).group())
        infos['autre_investigateur_nom_etablissement'].append(re.search(r"(?<=Nom de l’établissement:).*?(?=Service: )",x).group())
        infos['autre_investigateur_service'].append(re.search(r"(?<=Service:).*?(?=Adresse: )",x).group())
        infos['autre_investigateur_adresse'].append(re.search(r"(?<=Adresse:).*?(?=N° téléphone:)",x).group())
        infos['autre_investigateur_telephone'].append(re.search(r"(?<=N° téléphone:).*?(?=N° télécopie:)",x).group())
        infos['autre_investigateur_telecopie'].append(re.search(r"(?<=N° télécopie:).*?(?=Courriel:)",x).group())
        infos['autre_investigateur_courriel'].append(re.search(r"(?<=Courriel:).*?(?=Nom: )",x).group())
        #on recup la taille du premier bloc investigateur
        z=len(re.search(r"(?<=Nom:).*?(?=Nom: )",x).group())
        #on enleve ce bloc a x qui contient tous les investigateurs, ainsi a la prochaine boucle la regex se fera sur l'investigateur suivant
        x=x[(z-1):]
        
    #idem pour demande
    x=re.search(r"(?<=Identification du demandeur:).*(?=Justification de l’étude : )",texte).group()
    #puis tous les éléments du demandeur
    y=re.search(r"(?<=Nom de l’organisme:).*(?=Nom de la personne à contacter:)",x).group()
    infos['demandeur_nom_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter:).*(?=Adresse:)",x).group()
    infos['demandeur_nom_personne_contact']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['demandeur_UE_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['demandeur_UE_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['demandeur_UE_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['demandeur_UE_courriel']=y  
    
    #justification de létude
    table = doc.tables[0]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage dans courte
    courte=re.search(r"(?<=Bref rappel \(données de la littérature scientifique, pathologie, domaine d’étude\)).*",courte).group()
    infos['justification_etude_courte']=courte
    infos['justification_etude_longue']=longue
    
    #benefices de l'étude
    benefice=re.search(r"(?<=notamment les bénéfices escomptés pour les personnes qui se prêtent à la recherche\.).*(?=Risques:)",texte1).group()
    infos['benefices']=benefice
    
    #risques de l'étude
    risque=re.search(r"(?<=visant à éviter et/ou prendre en charge les événements inattendus\)\.).*(?=Retombées attendues)",texte1).group()
    infos['risques']=risque
    
    #retombées attendues
    table = doc.tables[1]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=Description des retombées attendues par cette recherche\n).*",courte).group()
    longue=re.search(r"(?<=d’augmentation de l’arsenal thérapeutique,…\)\.\n).*",longue).group()
    infos['retombee_attenduees_courte']=courte
    infos['retombee_attenduees_longue']=longue
    
    #objectif principal
    principal=re.search(r"(?<=Objectif Principal:).*(?=Objectif secondaires:)",texte1).group()
    infos['objectif_principal']=principal
    
    #objectif secondaire
    secondaire=re.search(r"(?<=Objectif secondaires:).*(?=Critères de Jugement)",texte1).group()
    infos['objectif_secondaire']=secondaire
    
    #critères de jugement principal
    table = doc.tables[2]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=Un seul critère correspondant à l’objectif principal \n).*",courte).group()
    longue=re.search(r"(?<=Il permettra également le calcul de l’effectif de l’étude\. \n).*",longue).group()
    infos['critere_jugement_principal_courte']=courte
    infos['critere_jugement_principal_longue']=longue
    
    #critères de jugement secondaire
    table = doc.tables[3]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=Liste de tous les critères de jugement secondaires\n).*",courte).group()
    longue=re.search(r"(?<=répondant aux objectifs secondaires\.\n).*",longue).group()
    infos['critere_jugement_secondaire_courte']=courte
    infos['critere_jugement_secondaire_longue']=longue
    
    #critères d'inclusion
    table = doc.tables[4]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=à la partie correspondante dans le corps du protocole § 6\.1\)\n).*",courte).group()
    infos['critere_inclusion_courte']=courte
    infos['critere_inclusion_secondaire_longue']=longue
    
    #critères de non inclusion
    table = doc.tables[5]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=à la partie correspondante dans le corps du protocole § 6\.2\)\n).*",courte).group()
    infos['critere_non_inclusion_courte']=courte
    infos['critere_non_inclusion_secondaire_longue']=longue
    
    #justification inclusion
    justif=re.search(r"(?<=Justifications de l’inclusion de personnes visées:).*(?=Modalités de recrutements)",texte1).group()
    infos['justification_inclusion']=justif
    
    #modalités_recrutement
    recru=re.search(r"(?<=Modalités de recrutements:).*(?=Traitement et stratégie)",texte1).group()
    infos['modalite_recrutement']=recru
    
    #traitement et stratégie
    table = doc.tables[6]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=traitements/stratégies/procédures\n).*",courte).group()
    longue=re.search(r"(?<=la durée du traitement et de la voie d’administration\.\n).*",longue).group()
    infos['traitement_strategie_courte']=courte
    infos['traitement_strategie_longue']=longue
    
    #fabriquant du dispositif
    x=re.search(r"(?<=Fabriquant du dispositif étudié:).*(?=Fabriquant du placebo)",texte).group()
    #puis tous les éléments du fabriquant
    y=re.search(r"(?<=Nom:).*(?=Adresse)",x).group()
    infos['fabriquant_dispositif_nom']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['fabriquant_dispositif_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['fabriquant_dispositif_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['fabriquant_dispositif_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['fabriquant_dispositif_courriel']=y  
    
    #fabriquant du placebo
    x=re.search(r"(?<=Fabriquant du placebo:).*(?=Description du produit/médicament)",texte).group()
    #puis tous les éléments du fabriquant
    y=re.search(r"(?<=Nom:).*(?=Adresse)",x).group()
    infos['fabriquant_placebo_nom']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone:)",x).group()
    infos['fabriquant_placebo_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie:)",x).group()
    infos['fabriquant_placebo_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel:)",x).group()
    infos['fabriquant_placebo_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*",x).group()
    infos['fabriquant_placebo_courriel']=y  
    
    #description produit
    x=re.search(r"(?<=Description du produit/médicament expérimental:).*(?=Informations sur le placebo)",texte).group()
    #puis tous les éléments du produit
    y=re.search(r"(?<=Nom du produit :).*(?=Nom de code)",x).group()
    infos['produit_nom']=y
    y=re.search(r"(?<=Nom de code:).*(?=Voie d’administration)",x).group()
    infos['produit_nom_code']=y
    y=re.search(r"(?<=Voie d’administration:).*(?=Dosageconcentration)",x).group()
    infos['produit_voie_administration']=y
    y=re.search(r"(?<=Dosageconcentration :).*(?=Dosage unité de concentration)",x).group()
    infos['produit_dosage_concentration']=y
    y=re.search(r"(?<=Dosage unité de concentration:).*",x).group()
    infos['produit_dosage_unite_concentration']=y 
    
    #description placebo
    x=re.search(r"(?<=Informations sur le placebo).*(?=Etude)",texte).group()
    #puis tous les éléments du placebo
    y=re.search(r"(?<=Numéro du placebo:).*(?=De quel produit expérimental)",x).group()
    infos['placebo_numero']=y
    y=re.search(r"(?<=préciser le numéro du ME:).*(?=Voie d’administration)",x).group()
    infos['placebo_numero_ME']=y
    y=re.search(r"(?<=Voie d’administration:).*",x).group()
    infos['placebo_voie_administration']=y
   
    #taille de l'étude
    table = doc.tables[7]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=Nombre de personnes à inclure:).*",courte).group()
    infos['taille_etude_courte']=courte
    infos['taille_etude_longue']=longue
    
    #modalités de l'indemnisation
    indem=re.search(r"(?<=Modalités et montant de l’indemnisation des personnes se prêtant à la recherche:).*(?=Justification de l’existence)",texte1).group()
    infos['indemnisation']=indem
    
    #justification existence
    justi=re.search(r"(?<=Justification de l’existence:).*?(?=Durée)",texte1).group()
    infos['justification_existence']=justi
    
    #durée des inclusions
    x=re.search(r"(?<=Durée prévue des inclusions:).*(?=Durée de participation pour une personne se prêtant à la recherche)",texte1).group()
    infos['duree_inclusion']=x
    
    #durée de participation
    x=re.search(r"(?<=c'est-à-dire la dernière visite du dernier patient inclus.).*(?=Durée totale de l’étude)",texte1).group()
    infos['duree_participation']=x
    
    #durée totale
    x=re.search(r"(?<=Durée totale de l’étude:).*(?=Analyse statistiques des données)",texte1).group()
    infos['duree_totale_etude']=x
    
    #analyse statistique
    table = doc.tables[8]  
    #créé une liste des valeurs des cases
    data = [] 
    keys = None
    #on récupère toutes les valeur dans une liste dont chaque valeur est un dico
    #la clé est la valeur de la colonne de gauche et la valeur celle de la cellule de droite
    #chaque valeur du dico est de la fçon suivante: key, value où key=cellule de gauche et value= cellule de droite
    for i, column in enumerate(table.columns):
        text = (cell.text for cell in column.cells)
    #créé le dictionnaire
        if i == 0:
            keys = tuple(text)
            continue
        row_data = dict(zip(keys, text))
        data.append(row_data)
    tab=data[0]
    for key, value in tab.items():
        courte=key
        longue=value
    courte=courte.replace("\xa0","")
    longue=longue.replace("\xa0","")
    #on ajoute ' ' pour eviter l'erreur avec les regex en cas de non remplissage par l'investigateur
    courte+=" "
    longue+=" "
    #on retire l'aide au remplissage 
    courte=re.search(r"(?<=Bref rappel des méthodes statistiques\n).*",courte).group()
    longue=re.search(r"(?<=données manquantes, inutilisées ou non valides\.\n).*",longue).group()
    infos['analyse_statistique_courte']=courte
    infos['analyse_statistique_longue']=longue
    
    #lieu de recherche
    x=re.search(r"(?<=dans un lieu nécessitant une autorisation de l’ARS\)).*(?=Plateau technique)",texte).group()
    #puis tous les éléments du placebo
    y=re.search(r"(?<=Intitulé du lieu:).*(?=N° d’autorisation:)",x).group()
    infos['lieu_recherche_intitule']=y
    y=re.search(r"(?<=N° d’autorisation:).*(?=Délivré le:)",x).group()
    infos['lieu_recherche_num_autorisation']=y
    y=re.search(r"(?<=Délivré le:).*(?=Date de limite de validité:)",x).group()
    infos['lieu_recherche_delivre_le']=y
    y=re.search(r"(?<=Date de limite de validité:).*(?=Nom et adresse:)",x).group()
    infos['lieu_recherche_date_limite_validite']=y
    y=re.search(r"(?<=Nom et adresse:).*",x).group()
    infos['lieu_recherche_nom_adresse']=y
    
    #plateau technique
    x=re.search(r"(?<=Plateau technique).*(?=Nom du CPP:)",texte).group()
    #puis tous les éléments du placebo
    y=re.search(r"(?<=Organisme:).*(?=Nom de la personne à contacter)",x).group()
    infos['plateau_technique_organisme']=y
    y=re.search(r"(?<=Nom de la personne à contacter:).*(?=Adresse)",x).group()
    infos['plateau_technique_personne_contact']=y
    y=re.search(r"(?<=Adresse:).*(?=N° téléphone)",x).group()
    infos['plateau_technique_adresse']=y
    y=re.search(r"(?<=N° téléphone:).*(?=N° télécopie)",x).group()
    infos['plateau_technique_num_telephone']=y
    y=re.search(r"(?<=N° télécopie:).*(?=Courriel)",x).group()
    infos['plateau_technique_num_telecopie']=y
    y=re.search(r"(?<=Courriel:).*(?=Tâches confiées)",x).group()
    infos['plateau_technique_courriel']=y
    y=re.search(r"(?<=Tâches confiées:).*(?=CPP)",x).group()
    infos['plateau_technique_taches_confiees']=y
    
    #CPP
    x=re.search(r"(?<=Nom du CPP:).*(?=Modalités de constitution ou non d’un comité de surveillance indépendant)",texte1).group()
    infos['CPP']=x
    
    #CPP
    x=re.search(r"(?<=Modalités de constitution ou non d’un comité de surveillance indépendant:).*",texte1).group()
    infos['comite_surveillance_independant']=x
            
    return infos
    
    
    
    
    
    
    
    
    