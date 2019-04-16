# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
      
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook('Fichiers_complementaires/' + classeur + '.xlsx')
                        
    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
    
    
    def ajoutOnglet_infos(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:L3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 13, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Infos Etudes', formatCellule)       
        self.feuille.write('E1', 'Centre Etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', "Centre d'études", formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'PB Terrain', formatFichier)
        self.feuille.write('G3', 'Adduction PB vers PA ', formatFichier)
        self.feuille.write('H3', 'Type PB Qualif ETD', formatFichier)
        self.feuille.write('I3', 'Numéro support site', formatFichier)
        self.feuille.write('J3', 'Adduction aval PB vers EL PB IPON', formatFichier)
        self.feuille.write('K3', 'PT IPON', formatFichier)
        self.feuille.write('L3', 'Commentaire RGT', formatFichier)


    def ajoutOnglet_CommentaireZN(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:F3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 6, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Commentaire ZN', formatCellule)       
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Type commentaire', formatFichier)
        self.feuille.write('B3', 'Code IMB', formatFichier)
        self.feuille.write('C3', 'ZN', formatFichier)
        self.feuille.write('D3', "Date commentaire", formatFichier)
        self.feuille.write('E3', 'Resp. commentaire', formatFichier)
        self.feuille.write('F3', 'Commentaire', formatFichier)


    def ajoutOnglet_contact_orange(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:M3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 13, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Contacts Orange', formatCellule)       
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'ZN', formatFichier)
        self.feuille.write('B3', 'Type contact', formatFichier)
        self.feuille.write('C3', 'Civilité', formatFichier)
        self.feuille.write('D3', "Prénom contact", formatFichier)
        self.feuille.write('E3', 'Nom contact', formatFichier)
        self.feuille.write('F3', 'Téléphone fixe', formatFichier)
        self.feuille.write('G3', 'Téléphone portable', formatFichier)
        self.feuille.write('H3', 'Fax', formatFichier)
        self.feuille.write('I3', 'Adresse mail', formatFichier)
        self.feuille.write('J3', 'Num_Adr_Cont', formatFichier)
        self.feuille.write('K3', 'Compl_Adr_Cont', formatFichier)
        self.feuille.write('L3', 'Type_Adr_Cont', formatFichier)
        self.feuille.write('M3', 'Voie_Adr_Cont', formatFichier)
        

    def ajoutOnglet_contact_projet(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:N3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 15, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Contacts projet', formatCellule)       
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'ZN', formatFichier)
        self.feuille.write('B3', 'Type contact', formatFichier)
        self.feuille.write('C3', 'Civilité', formatFichier)
        self.feuille.write('D3', "Prénom contact", formatFichier)
        self.feuille.write('E3', 'Nom contact', formatFichier)
        self.feuille.write('F3', 'Ville contact', formatFichier)
        self.feuille.write('G3', 'Code postal contact', formatFichier)
        self.feuille.write('H3', 'Code INSEE contact', formatFichier)
        self.feuille.write('I3', 'Téléphone fixe', formatFichier)
        self.feuille.write('J3', 'Téléphone portable', formatFichier)
        self.feuille.write('K3', 'Fax', formatFichier)
        self.feuille.write('L3', 'Adresse mail', formatFichier)
        self.feuille.write('M3', 'Adresse contact', formatFichier)
        self.feuille.write('N3', 'Commentaire', formatFichier)
         
    def ecrireResultat(self, resultat, classeur, feuille) :
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):

                if c in (10, 16, 17, 19, 25, 27, 28, 29, 30, 31, 33, 45, 46, 47):

                    self.feuille.write(r+3, c, col, formatDate)

                else :

                    self.feuille.write(r+3, c, col, formatCellule)


    def ecrireVal(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'}) 

        self.feuille.write(cellule, valeur, formatCellule)


    def fermeture(self,classeur):

        self.classeur.close()



    def ajoutOnglet_qualificateur(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:B3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 6, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Qualificateur', formatCellule)       
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Login Qualificateur', formatFichier)

    def ajoutOnglet_negociateur(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:B3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 6, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)    
        self.feuille.write('D1', 'Négociateur', formatCellule)       
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'ID ZN', formatFichier)
        self.feuille.write('B3', 'Login Négociateur', formatFichier)