# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
      
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook('Code_IMB_INEXISTANTS/' + classeur + '.xlsx')
                        
    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
    
         
    def ecrireResultat(self, resultat, classeur, feuille) :
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):

                if c in (8, 10, 12, 14, 15, 22,  25, 26):

                    self.feuille.write(r+3, c, col, formatDate)

                else :

                    self.feuille.write(r+3, c, col, formatCellule)


    def ecrireVal(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'}) 

        self.feuille.write(cellule, valeur, formatCellule)


    def ecrireResultatCPL(self, resultat, classeur, feuille) :
                
        formatCelCondiR = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#FF8F8F'})
        formatCelCondiV = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#92D050'})
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        formatCelCondiP = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#FFC000'})


        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            # print("r = ", r)
            # print("row = ", row)
            for c, col in enumerate(row):
                # print("c = ", c)
                # print("col = ", col)

                if c in (17, 18, 19, 20, 21, 22,  23, 25, 26, 29, 30, 31, 32, 33):
                    self.feuille.write(r+3, c, col, formatDate)
                else :
                    self.feuille.write(r+3, c, col, formatCellule)       



    def ecrireValSyn(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center', 'num_format': '#,##0' }) 

        self.feuille.write(cellule, valeur, formatCellule)    


    def fermeture(self,classeur):

        self.classeur.close()
        
    def ajoutOnglet_IXI(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:C3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 3, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
