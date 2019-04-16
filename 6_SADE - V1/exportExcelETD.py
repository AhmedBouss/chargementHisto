# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
      
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook('Fichier_DPL_OPT/' + classeur + '.xlsx')
                        
    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
    
    
    def ajoutOnglet_D(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('B5:AZ5')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(1, 51, 29)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
        

        self.feuille.write('B5', 'Code IMB', formatFichier)
        self.feuille.write('C5', 'Adresse', formatFichier)
        self.feuille.write('D5', 'Code postal', formatFichier)
        self.feuille.write('E5', 'Localité', formatFichier)
        self.feuille.write('F5', 'Identifiant Processus', formatFichier)
        self.feuille.write('G5', 'Code Syndic', formatFichier)
        self.feuille.write('H5', 'Nb logements', formatFichier)
        self.feuille.write('I5', 'Identifiant PB', formatFichier)
        self.feuille.write('J5', 'Identifiant PA', formatFichier)
        self.feuille.write('K5', 'Identifiant PM', formatFichier)
        self.feuille.write('L5', 'Date envoi pré-étude au syndic - OPT', formatFichier)
        self.feuille.write('M5', 'Escalier', formatFichier)
        self.feuille.write('N5', 'Nb logements P - OPT', formatFichier)
        self.feuille.write('O5', 'Nb logements R - OPT', formatFichier)
        self.feuille.write('P5', 'Type site - OPT', formatFichier)
        self.feuille.write('Q5', 'Résultat Etude D2', formatFichier)
        self.feuille.write('R5', 'Date pose du PB', formatFichier)
        self.feuille.write('S5', 'Date mise à disposition pré-étude', formatFichier) 
        self.feuille.write('T5', 'Etat Négociation', formatFichier) 
        self.feuille.write('U5', 'Date programmée raccordement ', formatFichier) 
        self.feuille.write('V5', 'Nom Syndic', formatFichier) 
        self.feuille.write('W5', 'Code regroup. Syndic', formatFichier)
        self.feuille.write('X5', 'Nom contact Syndic', formatFichier) 
        self.feuille.write('Y5', 'Type de regroupement ', formatFichier) 
        self.feuille.write('Z5', 'Etat réseau T/D ', formatFichier) 
        self.feuille.write('AA5', 'Date réalisation réseau T/D', formatFichier) 
        self.feuille.write('AB5', 'AG non pertinente', formatFichier)
        self.feuille.write('AC5', 'Date AG', formatFichier) 
        self.feuille.write('AD5', 'Date envoi accord syndic', formatFichier) 
        self.feuille.write('AE5', 'Date retour accord syndic', formatFichier) 
        self.feuille.write('AF5', 'Date de saisie accord syndic', formatFichier) 
        self.feuille.write('AG5', 'Date Retour Convention', formatFichier) 
        self.feuille.write('AH5', 'Attente Probation', formatFichier) 
        self.feuille.write('AI5', 'Date validation pré-étude par syndic', formatFichier) 
        self.feuille.write('AJ5', 'Présence DTA', formatFichier) 
        self.feuille.write('AK5', 'Civilité Pdt conseil syndical', formatFichier) 
        self.feuille.write('AL5', 'Nom Pdt conseil syndical', formatFichier) 
        self.feuille.write('AM5', 'Prénom Pdt conseil syndical', formatFichier) 
        self.feuille.write('AN5', 'N° Tel. Pdt conseil syndical', formatFichier) 
        self.feuille.write('AO5', 'N° Mobile Pdt conseil syndical', formatFichier) 
        self.feuille.write('AP5', 'Fax Pdt conseil syndical', formatFichier) 
        self.feuille.write('AQ5', 'Email Pdt conseil syndical', formatFichier) 
        self.feuille.write('AR5', 'Prénom Contact Syndic', formatFichier) 
        self.feuille.write('AS5', 'N° Tel. Contact Syndic', formatFichier) 
        self.feuille.write('AT5', 'Mobile Contact Syndic', formatFichier)
        self.feuille.write('AU5', 'Date récept. bon pré-étude', formatFichier)
        self.feuille.write('AV5', 'Date souhaitée pré-étude', formatFichier)
        self.feuille.write('AW5', 'Date fin etd ZA PA-OPT', formatFichier)
        # self.feuille.write('AX5', 'Regroupement_ZN', formatFichier)
        # self.feuille.write('AY5', 'Regroupement_ZE', formatFichier)
        # self.feuille.write('AZ5', 'Regroupement_RGT', formatFichier)




         
    def ecrireResultat(self, resultat, classeur, feuille) :
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):

                if c in (10, 16, 17, 19, 25, 27, 28, 29, 30, 31, 33, 45, 46, 47):

                    self.feuille.write(r+5, c+1, col, formatDate)

                else :

                    self.feuille.write(r+5, c+1, col, formatCellule)


    def ecrireVal(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'}) 

        self.feuille.write(cellule, valeur, formatCellule)


    def fermeture(self,classeur):

        self.classeur.close()