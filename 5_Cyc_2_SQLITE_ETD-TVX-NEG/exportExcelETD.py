# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
      
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook('Fichiers_Chargement_cycle_2_ETD_TVX/' + classeur + '.xlsx')
                        
    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
    
    def ajoutOnglet(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:R3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)

    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Etudes - Travaux lancés
    
    def ajoutOnglet_Q(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:X3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 24, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)
        self.feuille.write('R3', 'Nombre d’escaliers', formatFichier) 
        self.feuille.write('S3', 'Nombre d’étages', formatFichier) 
        self.feuille.write('T3', 'Nombre de logements P', formatFichier) 
        self.feuille.write('U3', 'Nombre de logements R', formatFichier) 
        self.feuille.write('V3', 'Type de site', formatFichier)

     # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Etudes
    
    def ajoutOnglet_E(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:X3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 24, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)
        self.feuille.write('R3', 'Nombre d’escaliers', formatFichier) 
        self.feuille.write('S3', 'Nombre d’étages', formatFichier) 
        self.feuille.write('T3', 'Nombre de logements P', formatFichier) 
        self.feuille.write('U3', 'Nombre de logements R', formatFichier) 
        self.feuille.write('V3', 'Type de site', formatFichier)
        self.feuille.write('W3', 'Date de Fin EZA', formatFichier) 
        self.feuille.write('X3', 'Résultat Etude D2', formatFichier)  
        
     # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers  Travaux lancés
    
    def ajoutOnglet_TL(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:Y3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 24, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)
        self.feuille.write('R3', 'Nombre d’escaliers', formatFichier) 
        self.feuille.write('S3', 'Nombre d’étages', formatFichier) 
        self.feuille.write('T3', 'Nombre de logements P', formatFichier) 
        self.feuille.write('U3', 'Nombre de logements R', formatFichier) 
        self.feuille.write('V3', 'Type de site', formatFichier)
        self.feuille.write('W3', 'Date de Fin EZA', formatFichier) 
        self.feuille.write('X3', 'Résultat Etude D2', formatFichier) 
        self.feuille.write('Y3', 'Numéro FCI d’accès', formatFichier) 

     # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Etudes - Travaux lancés
    
    def ajoutOnglet_D(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:AA3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 28, 29)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)
        self.feuille.write('R3', 'Nombre d’escaliers', formatFichier) 
        self.feuille.write('S3', 'Nombre d’étages', formatFichier) 
        self.feuille.write('T3', 'Nombre de logements P', formatFichier) 
        self.feuille.write('U3', 'Nombre de logements R', formatFichier) 
        self.feuille.write('V3', 'Type de site', formatFichier)
        self.feuille.write('W3', 'Date de Fin EZA', formatFichier) 
        self.feuille.write('X3', 'Résultat Etude D2', formatFichier) 
        self.feuille.write('Y3', 'Numéro FCI d’accès', formatFichier) 
        self.feuille.write('Z3', 'Date de Pose PB', formatFichier) 
        self.feuille.write('AA3', 'Numéro FCI fin de travaux', formatFichier) 


       # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Travaux
   
    def ajoutOngletTvx_D(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:N3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
        formatDate = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center','num_format': 'number','num_format': 'dd/mm/yyyy'})
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre de travaux', formatFichier)
        self.feuille.write('E3', 'Conducteur de travaux', formatFichier)
        self.feuille.write('F3', 'Localité', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'PA', formatFichier)
        self.feuille.write('J3', 'PM', formatFichier)
        self.feuille.write('K3', 'Date réalisation action', formatFichier)
        self.feuille.write('L3', 'Prod interne/Presta', formatFichier)
        self.feuille.write('M3', 'Entreprise', formatFichier) 
        self.feuille.write('N3', 'Numéro FCI d’accès', formatFichier) 

       # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Travaux
   

    def ajoutOnglet_S(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:R3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
                 
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'ID Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'FI', formatFichier)
        self.feuille.write('J3', 'PA', formatFichier)
        self.feuille.write('K3', 'PM', formatFichier)
        self.feuille.write('L3', 'Type adduction', formatFichier)
        self.feuille.write('M3', 'Date réalisation action', formatFichier)
        self.feuille.write('N3', 'Nature Code IMB', formatFichier)
        self.feuille.write('O3', 'Code Père', formatFichier)
        self.feuille.write('P3', 'Nature Façade', formatFichier)
        self.feuille.write('Q3', 'Type regroupement', formatFichier)
        self.feuille.write('R3', 'Type Suspension', formatFichier)  


    def ajoutOngletTvx_B(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:O3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
        formatDate = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center','num_format': 'number','num_format': 'dd/mm/yyyy'})
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre de travaux', formatFichier)
        self.feuille.write('E3', 'Conducteur de travaux', formatFichier)
        self.feuille.write('F3', 'Localité', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'PA', formatFichier)
        self.feuille.write('J3', 'PM', formatFichier)
        self.feuille.write('K3', 'Date réalisation action', formatFichier)
        self.feuille.write('L3', 'Prod interne/Presta', formatFichier)
        self.feuille.write('M3', 'Entreprise', formatFichier) 
        self.feuille.write('N3', 'Type suspension', formatFichier)         
        self.feuille.write('O3', 'Numéro FCI d’accès', formatFichier) 


       # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Travaux
   
    def ajoutOngletTvx_R(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:O3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 31)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
        formatDate = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center','num_format': 'number','num_format': 'dd/mm/yyyy'})
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre de travaux', formatFichier)
        self.feuille.write('E3', 'Conducteur de travaux', formatFichier)
        self.feuille.write('F3', 'Localité', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'PB', formatFichier)
        self.feuille.write('I3', 'PA', formatFichier)
        self.feuille.write('J3', 'PM', formatFichier)
        self.feuille.write('K3', 'Date réalisation action', formatFichier)
        self.feuille.write('L3', 'Prod interne/Presta', formatFichier)
        self.feuille.write('M3', 'Entreprise', formatFichier) 
        self.feuille.write('N3', 'Numéro FCI d’accès', formatFichier) 
        self.feuille.write('O3', 'Date de Pose PB', formatFichier) 
        
         
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

              


 

    def ajoutOngletCA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:AI3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 35, 20)
        self.feuille.freeze_panes(3, 4)

        self.feuille.data_validation('S4:S0000',  {'validate': 'list',
                                'source': ['0-AG pertinente', '1-AG non pertinente']})



        self.feuille.data_validation('R4:R10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})


        self.feuille.data_validation('T4:X10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})

        self.feuille.data_validation('Y4:Y10000',  {'validate': 'list',
                                'source': ['1-DTA présent', '0-DTA non présent et nécessaire', '2-DTA inutile']})

        self.feuille.data_validation('Z4:Z10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})

        self.feuille.data_validation('AA4:AA10000',  {'validate': 'list',
                                'source': ['C-Pour un immeuble', 'S-Pour une maison individuelle']})



        self.feuille.data_validation('AB4:AC10000',  {'validate': 'integer',
                                  'criteria': 'between',
                                  'minimum': 0,
                                  'maximum': 1000,
                                  # 'input_title': 'Entrez un entier:',
                                  # 'input_message': 'Valeur supérieur ou égale à 0',
                                  'error_title': 'Valeur non valide',
                                  'error_message': 'La valeur doit être supérieure ou égale à 0'})
        
        self.feuille.data_validation('AD4:AD10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})

        self.feuille.data_validation('AE4:AE10000',  {'validate': 'list',
                                'source': ['OK']})

        self.feuille.data_validation('AG4:AG10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})  

        self.feuille.data_validation('AH4:AH10000',  {'validate': 'list',
                                'source': ['Oui', 'Non'],
                                'input_title': 'Suspension',
                                'input_message': 'Mettre à Oui si une information a été supprimée ou Non sinon'})


        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        formatFichier1 = self.classeur.add_format({'border': True, 'bg_color': '#8DB4E2', 'align' : 'center', 'text_wrap': 'True'})
    
        self.feuille.write('O2', 'Nego_non_deb', formatFichier)
        self.feuille.write('P2', 'Qualif', formatFichier1)
        self.feuille.write('Q2', 'Qualif - ImportZN', formatFichier)
        self.feuille.write('R2', 'PrevAccord', formatFichier1)
        self.feuille.write('S2', 'PrevAccord', formatFichier1)
        self.feuille.write('T2', 'Accord', formatFichier)
        self.feuille.write('U2', 'Accord', formatFichier)
        self.feuille.write('V2', 'BPE', formatFichier1)
        self.feuille.write('W2', 'Piquetage OK', formatFichier)
        self.feuille.write('X2', 'Piquetage OK', formatFichier)
        self.feuille.write('Y2', 'BPT', formatFichier1)
        self.feuille.write('Z2', 'BPT', formatFichier1)
        self.feuille.write('AA2', 'BPT', formatFichier1)
        self.feuille.write('AB2', 'BPT', formatFichier1)
        self.feuille.write('AC2', 'BPT', formatFichier1)
        self.feuille.write('AD2', 'BPT', formatFichier1)
        self.feuille.write('AE2', 'Etudes', formatFichier)
        self.feuille.write('AF2', 'Travaux lancés', formatFichier1)
        self.feuille.write('AG2', 'DOE', formatFichier)


        self.feuille.write('A1', 'Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('E1', 'Centre d’études', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Nb logements - OPT', formatFichier) 
        self.feuille.write('D3', 'Jalon en cours', formatFichier)
        self.feuille.write('E3', 'Qualifie', formatFichier)
        self.feuille.write('F3', 'ImportZN', formatFichier)
        self.feuille.write('G3', 'PrevAccord', formatFichier)
        self.feuille.write('H3', 'Accord', formatFichier)
        self.feuille.write('I3', 'BPE', formatFichier)
        self.feuille.write('J3', 'Piquetage', formatFichier)
        self.feuille.write('K3', 'BPT', formatFichier)
        self.feuille.write('L3', 'Etudes', formatFichier)
        self.feuille.write('M3', 'Travaux lancés', formatFichier)
        self.feuille.write('N3', 'DOE', formatFichier)
        self.feuille.write('O3', 'Etat Négociation - OPT', formatFichier)
        self.feuille.write('P3', 'Code Syndic - OPT', formatFichier)
        self.feuille.write('Q3', 'Code regroup. Syndic - OPT', formatFichier)
        self.feuille.write('R3', 'Date envoi accord syndic - OPT', formatFichier)
        self.feuille.write('S3', 'AG non pertinente - OPT', formatFichier)
        self.feuille.write('T3', 'Date Retour Convention - OPT', formatFichier)
        self.feuille.write('U3', 'Date de saisie de l’accord syndic - OPT', formatFichier)
        self.feuille.write('V3', 'Date récept. bon pré-étude - OPT', formatFichier)
        self.feuille.write('W3', 'Date mise à disposition pré-étude - OPT', formatFichier)
        self.feuille.write('X3', 'Date fin etd ZA PA-OPT', formatFichier)
        self.feuille.write('Y3', 'Présence DTA - OPT', formatFichier)
        self.feuille.write('Z3', 'Date envoi pré-étude au syndic - OPT', formatFichier)
        self.feuille.write('AA3', 'Type site - OPT', formatFichier)
        self.feuille.write('AB3', 'Nb logements P - OPT', formatFichier)
        self.feuille.write('AC3', 'Nb logements R - OPT', formatFichier)
        self.feuille.write('AD3', 'Date validatipn pré-étude', formatFichier)
        self.feuille.write('AE3', 'Résultat de l’étude site - OPT', formatFichier)
        self.feuille.write('AF3', 'Date programmée raccordement - OPT', formatFichier)
        self.feuille.write('AG3', 'Date pose du PB - OPT', formatFichier)
        self.feuille.write('AH3', 'A suspendre', formatFichier)
        self.feuille.write('AI3', 'Type suspension', formatFichier)

    def ajoutOngletSansPA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:H3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 7, 25)
        self.feuille.freeze_panes(3, 4)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})

        self.feuille.write('A1', 'Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('E1', 'Centre d’études', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Nb logements - OPT', formatFichier) 
        self.feuille.write('D3', 'Jalon en cours', formatFichier)
        self.feuille.write('E3', 'Etat Négociation - OPT', formatFichier)
        self.feuille.write('F3', 'Code Syndic - OPT', formatFichier)
        self.feuille.write('G3', 'Code regroup. Syndic - OPT', formatFichier)
        self.feuille.write('H3', 'ID PA', formatFichier)



    def ecrireValSyn(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center', 'num_format': '#,##0' }) 

        self.feuille.write(cellule, valeur, formatCellule)    

    def ajoutOngletSyntheseETD(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        

        self.feuille.hide_gridlines(2)
        self.feuille.set_column(1, 1, 59)
        self.feuille.set_column(2, 3, 20)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichierBC = self.classeur.add_format({'border': True, 'bg_color': '#00B0F0', 'align' : 'center'})
        formatFichierOR = self.classeur.add_format({'border': True, 'bg_color': '#FFC000', 'align' : 'center'})
        formatFichierVC = self.classeur.add_format({'border': True, 'bg_color': '#92D050', 'align' : 'center'})

        self.feuille.write('B3', 'Général', formatFichierBC)
        self.feuille.write('C3', 'Qt Code IMB', formatFichierBC)
        self.feuille.write('D3', 'Qt EL', formatFichierBC)
        self.feuille.write('B4', 'Qt dans Code plaque', formatCellule)
        self.feuille.write('B5', 'Etat négo Déclaré', formatCellule)
        self.feuille.write('B6', 'Etat négo pas de syndic', formatCellule)


        self.feuille.write('B9', 'Répartition en fonction de l’état du PA', formatFichierBC)
        self.feuille.write('C9', 'Qt Code IMB', formatFichierBC)
        self.feuille.write('D9', 'Qt EL', formatFichierBC)
        self.feuille.write('B10', 'Avec PA renseigné sans avancement - Non débuté                 (liste A)', formatCellule)
        self.feuille.write('B11', 'Avec PA renseigné  avec avancement                                         (liste B)', formatCellule)
        self.feuille.write('B12', 'Avec PA renseigné et Prod terminée                                           (liste C)', formatCellule)
        self.feuille.write('B13', 'Avec PA non renseigné  sans avancement                                 (liste D)', formatCellule)
        self.feuille.write('B14', 'Avec PA non renseigné  avec avancement                                 (liste E)', formatCellule)


        self.feuille.write('B17', 'Contrôle cohérence', formatFichierOR)
        self.feuille.write('C17', 'Qt Code IMB', formatFichierOR)
        self.feuille.write('D17', 'Qt EL', formatFichierOR)
        self.feuille.write('B18', 'Valide', formatCellule)
        self.feuille.write('B19', 'Avec incohérence avancement                                                   (liste F)', formatCellule)
        self.feuille.write('B20', 'Incohérence etudes avec Nok nego                                                              ', formatCellule)
        self.feuille.write('B21', 'Sans avancement étude                                                               (liste G)', formatCellule)

        self.feuille.write('B23', 'Répartition des valides', formatFichierVC)
        self.feuille.write('C23', 'Qt Code IMB', formatFichierVC)
        self.feuille.write('D23', 'Qt EL', formatFichierVC)
        self.feuille.write('B24', 'Piquetage en cours                                                                       (liste H1)', formatCellule)
        self.feuille.write('B25', 'Piquetage OK                                                                                 (liste H2)', formatCellule)
        self.feuille.write('B26', 'Qualif OK                                                                                        (liste H3)', formatCellule)
        self.feuille.write('B27', 'Etude OK                                                                                         (liste H4)', formatCellule)
        self.feuille.write('B28', 'Lancement travaux OK                                                                (liste H5)', formatCellule)
        self.feuille.write('B29', 'Travaux débutés OK                                                                     (liste H6)', formatCellule)
        self.feuille.write('B30', 'Travaux réalisés OK                                                                      (liste H7)', formatCellule)
        self.feuille.write('B31', 'DOE OK                                                                                           (liste H8)', formatCellule)

        self.feuille.write('B33', 'Informations non existantes dans Optimum et obligatoires', formatFichierVC)
        self.feuille.write('C33', 'Qt Code IMB', formatFichierVC)
        self.feuille.write('D33', 'Qt EL', formatFichierVC)
        self.feuille.write('B34', 'Etude réalisée - Date Fin EZA à saisir                                       (Liste I1)', formatCellule)
        self.feuille.write('B35', 'Etude réalisée-Travaux lancés/ DOE réalisé - Numéro FCI commande d’accès à saisir      (Liste I2)', formatCellule)
        self.feuille.write('B36', 'DOE réalisé - N° FCI Fin travaux à saisir                                  (Liste I3)', formatCellule)
        


    def ajoutOngletlistes(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:D3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 5, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Pavillon sans FIS avec état négo déclaré', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'NB EL', formatFichier)
        self.feuille.write('D3', 'PA', formatFichier)


    def ajoutOngletEZA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:D3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 5, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Pavillon sans FIS avec état négo déclaré', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'NB EL', formatFichier)
        self.feuille.write('D3', 'Date fin EZA', formatFichier)

    def ajoutOngletFCIA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:D3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 5, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Pavillon sans FIS avec état négo déclaré', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'NB EL', formatFichier)
        self.feuille.write('D3', 'Numéro FCI commande accès', formatFichier)

    def ajoutOngletFCIF(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:D3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 5, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Pavillon sans FIS avec état négo déclaré', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'NB EL', formatFichier)
        self.feuille.write('D3', 'FCI fin de travaux', formatFichier)

    def fermeture(self,classeur):

        self.classeur.close()
        

    def ajoutOngletAVCT(self, classeur, feuille):

        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('B2:G2')
        self.feuille.hide_gridlines(2)
        self.feuille.freeze_panes(2, 4)
        self.feuille.set_zoom(80)

        self.feuille.set_column(0, 0, 7)
        self.feuille.set_column(1, 2, 10)
        self.feuille.set_column(3, 3, 25)
        self.feuille.set_column(4, 4, 93)
        self.feuille.set_column(5, 5, 47)
        self.feuille.set_column(6, 6, 31)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'left', 'valign' : 'vcenter', 'text_wrap': 'True'})
        formatCellule1 = self.classeur.add_format({'border': True, 'align' : 'center', 'valign' : 'vcenter', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#00B0F0', 'align' : 'center', 'valign' : 'vcenter', 'text_wrap': 'True'})

         
        self.feuille.write('B2', 'N°', formatFichier)
        self.feuille.write('C2', 'BF', formatFichier)
        self.feuille.write('D2', 'Descriptif', formatFichier)
        self.feuille.write('E2', 'Règle de déclaration par jalon', formatFichier)
        self.feuille.write('F2', 'Données obligatoires pour enregistrement du jalon', formatFichier)
        self.feuille.write('G2', 'Données non obligatoires pour enregistrement du jalon', formatFichier)
        self.feuille.write('B3', 1, formatCellule1)
        self.feuille.write('C3', 'NEGO', formatCellule1)
        self.feuille.write('D3', 'Négo non débuté', formatCellule1)
        self.feuille.write('E3', 'Etat Négociation <> « PAS DE SYNDIC » ET Code Syndic = Vide', formatCellule)
        self.feuille.write('F3', '', formatCellule)
        self.feuille.write('G3', '', formatCellule)
        self.feuille.write('B4', 2, formatCellule1)
        self.feuille.write('C4', 'NEGO', formatCellule1)
        self.feuille.write('D4', 'Qualifié', formatCellule1)
        self.feuille.write('E4', 'Code Syndic – OPT Non vide ET Code regroup Syndic vide', formatCellule)
        self.feuille.write('F4', 'Code Syndic', formatCellule)
        self.feuille.write('G4', 'Type syndic\nNom contact\nType contact\nAdresse contact\nTel contact\nMail contact\nEtat réseau T/D\nDate réalisation réseau T/D\nQualificateur', formatCellule)
        self.feuille.write('B5', 3, formatCellule1)
        self.feuille.write('C5', 'NEGO', formatCellule1)
        self.feuille.write('D5', 'Import ZN OK', formatCellule1)
        self.feuille.write('E5', 'Code regroup. Syndic – OPT non vide ET Code Syndic – OPT Non vide', formatCellule)
        self.feuille.write('F5', '', formatCellule)
        self.feuille.write('G5', '', formatCellule)
        self.feuille.write('B6', 4, formatCellule1)
        self.feuille.write('C6', 'NEGO', formatCellule1)
        self.feuille.write('D6', 'Prev Accord OK', formatCellule1)
        self.feuille.write('E6', '[Envoi accord syndic renseigné ET AG non pertinente = 0-AG pertinente]\n\nOU\n\n[Envoi accord syndic renseigné ET AG non pertinente = 1-AG non pertinente]', formatCellule)
        self.feuille.write('F6', 'Code Syndic\nAG Non-pertinente', formatCellule)
        self.feuille.write('G6', 'Statut d‘affectation Syndic\nDate d’AG\nDate d’Envoi Accord Type\nSite en copropriété\nFlag immeuble neuf', formatCellule)
        self.feuille.write('B7', 5, formatCellule1)
        self.feuille.write('C7', 'NEGO', formatCellule1)
        self.feuille.write('D7', 'Accord OK', formatCellule1)
        self.feuille.write('E7', 'Date Retour Convention – OPT non vide OU  Date de saisie de l’accord syndic', formatCellule)
        self.feuille.write('F7', 'Code Syndic\nAG Non-pertinente', formatCellule)
        self.feuille.write('G7', 'Statut d’affectation Syndic\nDate d’AG\nDate d’Envoi Accord Type\nSite en copropriété\nFlag immeuble neuf', formatCellule)
        self.feuille.write('B8', 6, formatCellule1)
        self.feuille.write('C8', 'NEGO', formatCellule1)
        self.feuille.write('D8', 'DTA OK', formatCellule1)
        self.feuille.write('E8', 'Présence DTA == 1-DTA présent\n\nOU\n\nPrésence DTA == 2-DTA inutile', formatCellule)
        self.feuille.write('F8', 'Code Syndic\nAG Non-pertinente\nPrésence DTA', formatCellule)
        self.feuille.write('G8', 'Type DTA', formatCellule)
        self.feuille.write('B9', 7, formatCellule1)
        self.feuille.write('C9', 'NEGO', formatCellule1)
        self.feuille.write('D9', 'BPE OK', formatCellule1)
        self.feuille.write('E9', '[Date récept. bon pré-étude renseignée]', formatCellule)
        self.feuille.write('F9', 'Code Syndic\nAG Non-pertinente\nPrésence DTA', formatCellule)
        self.feuille.write('G9', 'Statut d’affectation Syndic\nDate d’AG\nDate d’Envoi Accord Type\nSite en copropriété\nFlag immeuble neuf\nAttente Probation\nDate de Retour Convention', formatCellule)
        self.feuille.write('B10', 8, formatCellule1)
        self.feuille.write('C10', 'ETD-TVX', formatCellule1)
        self.feuille.write('D10', 'PIQUETAGE OK', formatCellule1)
        self.feuille.write('E10', '[Etat Négociation <> « PAS DE SYNDIC » ET Info « Date mise à disposition pré-étude » non vide]\n\nOU\n\n[Etat Négociation = « PAS DE SYNDIC » ET Résultat étude site non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide]', formatCellule)
        self.feuille.write('F10', 'Date mise à disposition pré-étude', formatCellule)
        self.feuille.write('G10', '', formatCellule)
        self.feuille.write('B11', 9, formatCellule1)
        self.feuille.write('C11', 'NEGO', formatCellule1)
        self.feuille.write('D11', 'BPT OK', formatCellule1)
        self.feuille.write('E11', '[Date validation pré-étude par syndic – OPT » non vide ET Date de saisie de l’accord (colonne FA) syndic non vide ET Date envoi pré-étude au syndic non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide et Présence DTA == 1-DTA présent]\n\nOU\n\n[Date validation pré-étude par syndic – OPT » non vide ET Date de saisie de l’accord (colonne FA) syndic non vide ET Date envoi pré-étude au syndic non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide et Présence DTA == 2-DTA inutile]', formatCellule)
        self.feuille.write('F11', 'Code Syndic\nAG Non-pertinente\nPrésence DTA\nDate d’Accord Syndic\nDate de validation P-E par le syndic (DVPE)\nDate envoi pré-étude au syndic\nNombre de logements P\nNombre de logements R', formatCellule)
        self.feuille.write('G11', 'Statut d’affectation Syndic\nDate d’AG\nDate d’Envoi Accord Type\nSite en copropriété\nFlag immeuble neuf\nAttente Probation\nDate de Retour Convention\nNombre d’escaliers\nNombre d’étages', formatCellule)
        self.feuille.write('B12', 10, formatCellule1)
        self.feuille.write('C12', 'ETD-TVX', formatCellule1)
        self.feuille.write('D12', 'ETUDES OK', formatCellule1)
        self.feuille.write('E12', '[Etat Négociation <> « PAS DE SYNDIC » ET Résultat étude site non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide]\n\nOU\n\n[Etat Négociation = « PAS DE SYNDIC » ET Résultat étude site non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide]', formatCellule)
        self.feuille.write('F12', 'Date de Fin EZA\nNombre de logements P\nNombre de logements R\nType de site', formatCellule)
        self.feuille.write('G12', 'Nombre d’escaliers\nNombre d’étages', formatCellule)
        self.feuille.write('B13', 11, formatCellule1)
        self.feuille.write('C13', 'ETD-TVX', formatCellule1)
        self.feuille.write('D13', 'ETUDES - TVX LANCES OK', formatCellule1)
        self.feuille.write('E13', 'Date programmée raccordement non vide', formatCellule)
        self.feuille.write('F13', 'Date envoi pré-étude au syndic\nDate de Fin EZA\nNombre de logements P\nNombre de logements R\nNuméro FCI commande accès', formatCellule)
        self.feuille.write('G13', 'Nombre d’escaliers\nNombre d’étages', formatCellule)
        self.feuille.write('B14', 12, formatCellule1)
        self.feuille.write('C14', 'ETD-TVX', formatCellule1)
        self.feuille.write('D14', 'DOE OK', formatCellule1)
        self.feuille.write('E14', 'Date du Pose PB non vide', formatCellule)    
        self.feuille.write('F14', 'Date envoi pré-étude au syndic\nDate de Fin EZA\nNombre de logements P\nNombre de logements R\nNuméro FCI d’accès\nRésultat Etude D2\nDate de Pose PB\nNuméro FCI commande accès\nNuméro FCI fin de travaux', formatCellule)
        self.feuille.write('G14', 'Nombre d’escaliers\nNombre d’étages', formatCellule)


    def ajoutOnglet_EZA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:F3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 6, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'Date de Fin EZA', formatFichier) 


    def ajoutOnglet_FCI(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:I3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 8, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Localité', formatFichier)
        self.feuille.write('F3', 'N° ENR GCB 1 / N° FCI FIN TVX', formatFichier)
        self.feuille.write('G3', 'Type commande', formatFichier)
        self.feuille.write('H3', 'ID interne commande', formatFichier) 
        self.feuille.write('I3', 'Date demande CMD accès', formatFichier) 


    def ecrireResultatCPLC(self, resultat, classeur, feuille) :
                
        formatCelCondiR = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#FF8F8F'})
        formatCelCondiV = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#92D050'})
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        formatCelCondiP = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#CCE9FF'})


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

                if row[c] =='Non' and c < 15 and c > 3:
                    self.feuille.write(r+3, c, col, formatCelCondiR)                  

                elif row[c] =='Oui' and c < 15 and c > 3:
                    self.feuille.write(r+3, c, col, formatCelCondiV)      


            if row[4] == 'Non':

                if row[15] == '':
                    self.feuille.write(r+3, 15, row[15], formatCelCondiP)

            if row[5] == 'Non':

                if row[16] == '':
                    self.feuille.write(r+3, 16, row[16], formatCelCondiP)
 
            if row[6] == 'Non':

                if row[17] == '':
                    self.feuille.write(r+3, 17, row[17], formatCelCondiP)
                if row[18] == '':
                    self.feuille.write(r+3, 18, row[18], formatCelCondiP)

            if row[7] == 'Non':

                if row[19] == '':
                    self.feuille.write(r+3, 19, row[19], formatCelCondiP)
 
                if row[20] == '':
                    self.feuille.write(r+3, 20, row[20], formatCelCondiP)

            if row[8] == 'Non':

                if row[21] == '':
                    self.feuille.write(r+3, 21, row[21], formatCelCondiP)

            if row[9] == 'Non':

                if row[22] == '':
                    self.feuille.write(r+3, 22, row[22], formatCelCondiP)
 
                if row[23] == '':
                    self.feuille.write(r+3, 23, row[23], formatCelCondiP)

            if row[10] == 'Non':

                if row[24] == '' or row[24] == '0-DTA non présent et nécessaire':
                    self.feuille.write(r+3, 24, row[24], formatCelCondiP)
 
                if row[25] == '':
                    self.feuille.write(r+3, 25, row[25], formatCelCondiP)

                if row[26] == '':
                    self.feuille.write(r+3, 26, row[26], formatCelCondiP)

                if row[27] == '':
                    self.feuille.write(r+3, 27, row[27], formatCelCondiP)

                if row[28] == '':
                    self.feuille.write(r+3, 28, row[28], formatCelCondiP)

                if row[29] == '':
                    self.feuille.write(r+3, 29, row[29], formatCelCondiP)

                if row[20] == '':
                    self.feuille.write(r+3, 20, row[20], formatCelCondiP)

            if row[11] == 'Non':

                if row[30] == '':
                    self.feuille.write(r+3, 30, row[30], formatCelCondiP)


            if row[12] == 'Non':

                if row[31] == '':
                    self.feuille.write(r+3, 31, row[31], formatCelCondiP)  



    def ajoutOnglet_Suspendre(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:E3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 6, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre etudes', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Code plaque', formatFichier)
        self.feuille.write('D3', 'Centre études', formatFichier)
        self.feuille.write('E3', 'Type Suspension', formatFichier)

    def ajoutOnglet_Admin_PA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:B3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 2, 31)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
     
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'PA', formatFichier)