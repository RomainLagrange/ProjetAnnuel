# -*- coding: utf-8 -*-
"""
Created on Wed Apr  3 10:48:05 2019

@author: Utilisateur
"""

import main_cat1, main_cat2, main_cat3
from main_cat1 import *
from main_cat2 import *
from main_cat3 import *


dico = {}

def fct_gen(dic):

    dico = dic
    
    print (dico)
    
    if dico['la_categorie']=='1':
        main_cat1.construit_doc(dico)
    elif dico['la_categorie']=='2':
        main_cat2.construit_doc(dico)
    else:
        main_cat3.construit_doc(dico)