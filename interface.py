# -*- coding: utf-8 -*-
"""
Created on Tue Feb  5 11:02:39 2019

@author: Anaïs
"""

from tkinter import * 
from tkinter.filedialog import askopenfilename
from functools import partial
import main_gen
from main_gen import *

fenetre = Tk()

    
from tkinter import Frame, Tk, BOTH, Text, Menu, END
from tkinter import filedialog

mon_dictionnaire = {}

class Example(Frame):
  
    def __init__(self):
        super().__init__()   
         
        self.initUI()
        
    
        
    def initUI(self):
      
        self.master.title("File dialog")
        self.pack(fill=BOTH, expand=1)
        
        menubar = Menu(self.master)
        self.master.config(menu=menubar)
        
        fileMenu = Menu(menubar)
        fileMenu.add_command(label="Ouvrir", command=self.onOpen)
        menubar.add_cascade(label="Fichier", menu=fileMenu)        
        
        
    def onOpen(self):
      
        ftypes = [('text files', '*.docx'), ('All files', '*')]
        dlg = filedialog.Open(self, filetypes = ftypes)
        fl = dlg.show()
        print(fl)   
        
        mon_dictionnaire['le_chemin'] = fl
        print (ftypes)

        return ftypes
        print(dlg)

ex = Example()


#==============================================================================
# retourne la valeur du bouton du nombre d'investigateur

def validebutton(): 
    P =entree.get()
    mon_dictionnaire['le_nb_investigateur'] = P
    print(P)
    #return (P)
    label4 = Label(fenetre, text="Combien de plateaux techniques sont intégrés dans cette étude ? ")
    label4.pack()
    entree2.pack()
    
    #Bouton qui récupère la valeur du nb de plateau technique
    boutonChif2 = Button(fenetre, text="Valider", command = validebutton2)
    boutonChif2.pack()
    
    

def selec_les_doc():
    main_gen.fct_gen(mon_dictionnaire)


#==============================================================================
# retourne la valeur du bouton du nombre de plateau technique
def validebutton2(): 
    
    P2 =entree2.get()
    mon_dictionnaire['le_plateau'] = P2
    print(P2)
    selec_les_doc()
    return (P2)
    

#==============================================================================  
# retourne la valeur correspondant au nombre d'investigateur    
def valeurB6():
    y = varia.get()  
    mon_dictionnaire['le_type_recherche'] =y
    print(y)
    #return (y)  
    label3 = Label(fenetre, text="Combien d’investigateurs participent à cette étude ? ")
    label3.pack()
    entree.pack()
    
    # bouton donnant le nombre d'investigateur
    boutonChif = Button(fenetre, text="Valider", command = validebutton)
    boutonChif.pack()
#==============================================================================  
# récupère la valeur de la deuxième série de boutons radios   
def valeurB4(): 
    x = var.get()
    mon_dictionnaire['la_categorie'] =x
    print(x)
    
    label2 = Label(fenetre, text="Il s’agit d’une recherche portant sur : ")
    label2.pack()
    bouton6.pack()
    bouton7.pack()
    bouton9.pack()
    bouton8.pack()
    bouton8bis.pack()
    
    y=varia.get()
    
    #bouton donnant la valeur du type de recherche
    boutonRecherche = Button(fenetre, text=" Valider ", command = valeurB6)
    boutonRecherche.pack()
    
    if x == "1":     
        bouton8bis.configure(state = DISABLED)
    elif x == "2": 
        bouton8bis.configure(state = DISABLED)
        bouton6.configure(state = DISABLED)
        bouton9.configure(state = DISABLED)
    elif x == "3": 
        bouton6.configure(state = DISABLED)
        bouton7.configure(state = DISABLED)
        bouton8.configure(state = DISABLED)
        bouton9.configure(state = DISABLED)

    return y
   


#==============================================================================
# première série de boutons radios
label = Label(fenetre, text="Veuillez choisir la catégorie du protocole ")
label.pack()

var = StringVar() 
bouton1 = Radiobutton(fenetre, text="Catégorie 1", variable=var, value=1, tristatevalue=0)
bouton2 = Radiobutton(fenetre, text="Catégorie 2", variable=var, value=2, tristatevalue=0)
bouton3 = Radiobutton(fenetre, text="Catégorie 3", variable=var, value=3, tristatevalue=0)
bouton1.pack()
bouton2.pack()
bouton3.pack()
#x = var.get()

#deuxième de boutons séries
varia = StringVar()
bouton6 = Radiobutton(fenetre, text="Médicaments", variable=varia, value=6, tristatevalue=0)
bouton7 = Radiobutton(fenetre, text="Dispositifs médicaux", variable=varia, value=7, tristatevalue=0)
bouton9 = Radiobutton(fenetre, text="Produits biologiques", variable=varia, value=9, tristatevalue=0)
bouton8 = Radiobutton(fenetre, text="Hors produits de santé", variable=varia, value=8, tristatevalue=0)
bouton8bis = Radiobutton(fenetre, text="Recherche non interventionnelle", variable=varia, value=10, tristatevalue=0)

#bouton qui retourne la valeur de la catégorie
bouton45 = Button(fenetre, text="Valider", command = valeurB4)
bouton45.pack()



#==============================================================================



#saisie du nombre d'investigateur 
valuesaisie = IntVar() 
valuesaisie.set("texte par défaut")
entree = Entry(fenetre, textvariable=int, width=30)




#==============================================================================


# saisie du nb de plateau technique
valuesaisie2 = IntVar() 
valuesaisie2.set("texte par défaut")
entree2 = Entry(fenetre, textvariable=int, width=30)



fenetre.mainloop()


#if __name__ == '__main__':
#    main()   

