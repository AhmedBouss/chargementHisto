# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
    
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook('Fichiers_Chargement_cycle_2_NEGO/' + classeur + '.xlsx')

    # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Qualif - Non Qualif---------------------
    
    def ajoutOngletQ(self, classeur, feuille):

        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:T3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 19, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
         
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Type syndic', formatFichier)
        self.feuille.write('L3', 'Nom contact', formatFichier)
        self.feuille.write('M3', 'Type contact', formatFichier)
        self.feuille.write('N3', 'Adresse contact', formatFichier)
        self.feuille.write('O3', 'Tel contact', formatFichier)
        self.feuille.write('P3', 'Mail contact', formatFichier)
        self.feuille.write('Q3', 'Etat réseau T/D', formatFichier)
        self.feuille.write('R3', 'Date réalisation réseau T/D', formatFichier)
        self.feuille.write('S3', 'Qualificateur', formatFichier)
                        
    # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Import ZN - Accord - DTA---------------------
    
    def ajoutOngletI(self, classeur, feuille):

        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:J3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 9, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Import ZN', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)      

       # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Prev accord---------------------
   
    def ajoutOngletPA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:P3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 16, 25)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('D1', 'Prev Accord', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Statut d’affectation Syndic', formatFichier)
        self.feuille.write('L3', 'AG Non-pertinente', formatFichier)
        self.feuille.write('M3', 'Date d’AG', formatFichier)
        self.feuille.write('N3', 'Date d’Envoi Accord Type', formatFichier)
        self.feuille.write('O3', 'Site en copropriété', formatFichier)
        self.feuille.write('P3', 'Flag immeuble neuf', formatFichier)

       # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers BPE---------------------

    def ajoutOngletE(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:S3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 19, 25)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True', 'text_wrap': 'True'})

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('D1', 'BPE', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Statut d’affectation Syndic', formatFichier)
        self.feuille.write('L3', 'AG Non-pertinente', formatFichier)
        self.feuille.write('M3', 'Date d’AG', formatFichier)
        self.feuille.write('N3', 'Date d’Envoi Accord Type', formatFichier)
        self.feuille.write('O3', 'Site en copropriété', formatFichier)
        self.feuille.write('P3', 'Flag immeuble neuf', formatFichier)
        self.feuille.write('Q3', 'Attente Probation', formatFichier)
        self.feuille.write('R3', 'Date de Retour Convention', formatFichier)
        self.feuille.write('S3', 'Présence DTA', formatFichier)


       # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers BPT---------------------

    def ajoutOngletT(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:AJ3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 36, 25)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        formatNum = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'num_format' : '0.00%', 'text_wrap': 'True'})
        
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('D1', 'BPT', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Statut d’affectation Syndic', formatFichier)
        self.feuille.write('L3', 'AG Non-pertinente', formatFichier)
        self.feuille.write('M3', 'Date d’AG', formatFichier)
        self.feuille.write('N3', 'Date d’Envoi Accord Type', formatFichier)
        self.feuille.write('O3', 'Site en copropriété', formatFichier)
        self.feuille.write('P3', 'Flag immeuble neuf', formatFichier)
        self.feuille.write('Q3', 'Attente Probation', formatFichier)
        self.feuille.write('R3', 'Date de Retour Convention', formatFichier)
        self.feuille.write('S3', 'Présence DTA', formatFichier)
        self.feuille.write('T3', 'Date d’Accord Syndic', formatFichier)
        self.feuille.write('U3', 'Date de validation P-E par le syndic (DVPE)', formatFichier)
        self.feuille.write('V3', 'Civilité Président', formatFichier)
        self.feuille.write('W3', 'Email PdtConseilSyndical', formatFichier)
        self.feuille.write('X3', 'N° fax PdtConseilSyndical', formatNum)
        self.feuille.write('Y3', 'Nom PdtConseilSyndical', formatFichier)
        self.feuille.write('Z3', 'Prénom PdtConseilSyndical', formatFichier)
        self.feuille.write('AA3', 'N° téléphone PdtConseilSyndical', formatNum)
        self.feuille.write('AB3', 'N° mobile PdtConseilSyndical', formatNum)
        self.feuille.write('AC3', 'Civilité Responsable', formatFichier)
        self.feuille.write('AD3', 'Email Syndic', formatNum)
        self.feuille.write('AE3', 'N°Fax Syndic', formatNum)
        self.feuille.write('AF3', 'Nom Responsable', formatFichier)
        self.feuille.write('AG3', 'Prénom Responsable', formatFichier)
        self.feuille.write('AH3', 'N°Tel Syndic', formatNum)
        self.feuille.write('AI3', 'N°Mobile Syndic', formatNum)
        self.feuille.write('AJ3', 'Date envoi pré-étude au syndic', formatNum)

       # ---------------------Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers DTA---------------------

    def ajoutOngletD(self, classeur, feuille):

        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:L3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 11, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'DTA', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Présence DTA', formatFichier)
        self.feuille.write('L3', 'Type DTA', formatFichier)

       # ---------------------Ecriture des résultats depuis la base de données vers les fichiers excel---------------------
                        
    def ecrireResultat(self, resultat, classeur, feuille) :
        

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})


        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):
                if c in (5, 7,8,9, 10, 11, 12, 13, 16, 17, 18, 19, 20, 21, 24, 28, 29, 30, 36):
                    self.feuille.write(r+3, c, col, formatDate)

                else :
                    self.feuille.write(r+3, c, col, formatCellule)

    def ecrireResultatCPL(self, resultat, classeur, feuille) :
                
        formatCelCondiR = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#FF8F8F'})
        formatCelCondiV = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#92D050'})
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})


        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):
                if c in (6, 8, 9, 10, 11, 13, 14, 15, 16, 17, 19, 20, 26, 27, 28, 36):
                    self.feuille.write(r+3, c, col, formatDate)
                else :
                    self.feuille.write(r+3, c, col, formatCellule)       

                if row[c] =='Non':
                    self.feuille.write(r+3, c, col, formatCelCondiR)

                elif row[c] =='Oui':
                    self.feuille.write(r+3, c, col, formatCelCondiV)


    def ecrireVal(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'}) 

        self.feuille.write(cellule, valeur, formatCellule)

   
    def ecrireValSyn(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center', 'num_format': '#,##0' }) 

        self.feuille.write(cellule, valeur, formatCellule)    

       # ---------------------Fermeture du classeur - enregistre le fichier---------------------


    def fermeture(self,classeur):

        self.classeur.close()
        

    def ajoutOngletInfosCompl(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:AK3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 36, 25)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        formatNum = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'num_format' : '0.00%', 'text_wrap': 'True'})
        
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('D1', 'BPT', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Localité', formatFichier)
        self.feuille.write('D3', 'Code Postal', formatFichier)
        self.feuille.write('E3', 'ID SYNDIC', formatFichier)
        self.feuille.write('F3', 'Nom Syndic', formatFichier)
        self.feuille.write('G3', 'Nb EL', formatFichier)
        self.feuille.write('H3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('I3', 'Type regroupement', formatFichier)
        self.feuille.write('J3', 'Date réalisation', formatFichier)
        self.feuille.write('K3', 'Statut d’affectation Syndic', formatFichier)
        self.feuille.write('L3', 'AG Non-pertinente', formatFichier)
        self.feuille.write('M3', 'Date d’AG', formatFichier)
        self.feuille.write('N3', 'Date d’Envoi Accord Type', formatFichier)
        self.feuille.write('O3', 'Site en copropriété', formatFichier)
        self.feuille.write('P3', 'Flag immeuble neuf', formatFichier)
        self.feuille.write('Q3', 'Attente Probation', formatFichier)
        self.feuille.write('R3', 'Date de Retour Convention', formatFichier)
        self.feuille.write('S3', 'Présence DTA', formatFichier)
        self.feuille.write('T3', 'Type DTA', formatFichier)
        self.feuille.write('U3', 'Date d’Accord Syndic', formatFichier)
        self.feuille.write('V3', 'Date de validation P-E par le syndic (DVPE)', formatFichier)
        self.feuille.write('W3', 'Civilité Président', formatFichier)
        self.feuille.write('X3', 'Email PdtConseilSyndical', formatFichier)
        self.feuille.write('Y3', 'N° fax PdtConseilSyndical', formatNum)
        self.feuille.write('Z3', 'Nom PdtConseilSyndical', formatFichier)
        self.feuille.write('AA3', 'Prénom PdtConseilSyndical', formatFichier)
        self.feuille.write('AB3', 'N° téléphone PdtConseilSyndical', formatNum)
        self.feuille.write('AC3', 'N° mobile PdtConseilSyndical', formatNum)
        self.feuille.write('AD3', 'Civilité Responsable', formatFichier)
        self.feuille.write('AE3', 'Email Syndic', formatNum)
        self.feuille.write('AF3', 'N°Fax Syndic', formatNum)
        self.feuille.write('AG3', 'Nom Responsable', formatFichier)
        self.feuille.write('AH3', 'Prénom Responsable', formatFichier)
        self.feuille.write('AI3', 'N°Tel Syndic', formatNum)
        self.feuille.write('AJ3', 'N°Mobile Syndic', formatNum)
        self.feuille.write('AK3', 'Date envoi pré-étude au syndic', formatNum)


    def ajoutOngletCI(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:K3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 13, 20)
        self.feuille.freeze_panes(3, 4)
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        formatFichierB = self.classeur.add_format({'border': True, 'bg_color': '#538DD5', 'align' : 'center', 'text_wrap': 'True'})
        formatFichierC = self.classeur.add_format({'border': True, 'bg_color': '#92D050', 'align' : 'center', 'text_wrap': 'True'})

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichierB)
        self.feuille.write('B3', 'Adresse', formatFichierB)
        self.feuille.write('C3', 'Nb logements - OPT', formatFichierB) 
        self.feuille.write('D3', 'Code regroup. Syndic - OPT', formatFichierC)
        self.feuille.write('E3', 'Code Syndic - OPT', formatFichier)
        self.feuille.write('F3', 'Etat Négociation - OPT', formatFichier)
        self.feuille.write('G3', 'Date envoi accord syndic - OPT', formatFichier)
        self.feuille.write('H3', 'AG non pertinente - OPT', formatFichier)
        self.feuille.write('I3', 'Date Retour Convention - OPT', formatFichier)
        self.feuille.write('J3', 'Date de saisie de l’accord syndic - OPT', formatFichier)
        self.feuille.write('K3', 'Date récept. bon pré-étude - OPT', formatFichier)


    def ajoutOngletCA(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:Z3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 28, 25)
        self.feuille.freeze_panes(3, 4)

        self.feuille.data_validation('O4:O10000',  {'validate': 'list',
                                'source': ['0-AG pertinente', '1-AG non pertinente']})



        self.feuille.data_validation('P4:R10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})


        self.feuille.data_validation('N4:N10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})

        self.feuille.data_validation('S4:S10000',  {'validate': 'list',
                                'source': ['1-DTA présent', '0-DTA non présent et nécessaire', '2-DTA inutile']})

        self.feuille.data_validation('T4:U10000',  {'validate': 'date',
                                'criteria': 'between',
                                'minimum': date(2000, 1, 1),
                                'maximum': date(2050, 12, 12),
                                # 'input_title': 'Entrez une date:',
                                # 'input_message': 'Date postérieure au 01/01/2000',
                                'error_title': 'Entrez une date valide!',
                                'error_message': 'La date doit être comprise entre le 01/01/2000 et le 12/12/2050'})

        self.feuille.data_validation('X4:X10000',  {'validate': 'list',
                                'source': ['C-Pour un immeuble', 'S-Pour une maison individuelle']})


        self.feuille.data_validation('Y4:Y10000',  {'validate': 'list',
                                'source': ['Oui', 'Non'],
                                'input_title': 'Suspension',
                                'input_message': 'Mettre à Oui si une information a été supprimée ou Non sinon'})

        self.feuille.data_validation('V4:W10000',  {'validate': 'integer',
                                  'criteria': 'between',
                                  'minimum': 0,
                                  'maximum': 1000,
                                  # 'input_title': 'Entrez un entier:',
                                  # 'input_message': 'Valeur supérieur ou égale à 0',
                                  'error_title': 'Valeur non valide',
                                  'error_message': 'La valeur doit être supérieure ou égale à 0'})


        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        formatFichier1 = self.classeur.add_format({'border': True, 'bg_color': '#8DB4E2', 'align' : 'center', 'text_wrap': 'True'})
    
        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)


        self.feuille.write('K2', 'Nego_non_deb', formatFichier)
        self.feuille.write('L2', 'Qualif', formatFichier1)
        self.feuille.write('M2', 'Qualif - ImportZN', formatFichier)
        self.feuille.write('N2', 'PrevAccord', formatFichier1)
        self.feuille.write('O2', 'PrevAccord', formatFichier1)
        self.feuille.write('P2', 'Accord', formatFichier)
        self.feuille.write('Q2', 'Accord', formatFichier)
        self.feuille.write('R2', 'BPE', formatFichier1)
        self.feuille.write('S2', 'BPT', formatFichier)
        self.feuille.write('T2', 'BPT', formatFichier)
        self.feuille.write('U2', 'BPT', formatFichier)
        self.feuille.write('V2', 'BPT', formatFichier)
        self.feuille.write('W2', 'BPT', formatFichier)
        self.feuille.write('X2', 'BPT', formatFichier)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Adresse', formatFichier)
        self.feuille.write('C3', 'Nb logements - OPT', formatFichier) 
        self.feuille.write('D3', 'Jalon en cours', formatFichier)
        self.feuille.write('E3', 'Qualifie', formatFichier)
        self.feuille.write('F3', 'ImportZN', formatFichier)
        self.feuille.write('G3', 'PrevAccord', formatFichier)
        self.feuille.write('H3', 'Accord', formatFichier)
        self.feuille.write('I3', 'BPE', formatFichier)
        self.feuille.write('J3', 'BPT', formatFichier)
        self.feuille.write('K3', 'Etat Négociation - OPT', formatFichier)
        self.feuille.write('L3', 'Code Syndic - OPT', formatFichier)
        self.feuille.write('M3', 'Code regroup. Syndic - OPT', formatFichier)
        self.feuille.write('N3', 'Date envoi accord syndic - OPT', formatFichier)
        self.feuille.write('O3', 'AG non pertinente - OPT', formatFichier)
        self.feuille.write('P3', 'Date Retour Convention - OPT', formatFichier)
        self.feuille.write('Q3', 'Date de saisie de l’accord syndic - OPT', formatFichier)
        self.feuille.write('R3', 'Date récept. bon pré-étude - OPT', formatFichier)
        self.feuille.write('S3', 'Présence DTA - OPT', formatFichier)
        self.feuille.write('T3', 'Date envoi pré-étude au syndic - OPT', formatFichier)
        self.feuille.write('U3', 'Date validation pré-étude par syndic - OPT', formatFichier)
        self.feuille.write('V3', 'Nombre de logements P', formatFichier)
        self.feuille.write('W3', 'Nombre de logements R', formatFichier)
        self.feuille.write('X3', 'Type de site', formatFichier)
        self.feuille.write('Y3', 'A suspendre', formatFichier)
        self.feuille.write('Z3', 'Type suspension', formatFichier)
        


    def ajoutOngletSynthese(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        

        self.feuille.hide_gridlines(2)
        self.feuille.set_column(1, 1, 75)
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

        self.feuille.write('B9', 'Contrôle cohérence', formatFichierOR)
        self.feuille.write('C9', 'Qt Code IMB', formatFichierOR)
        self.feuille.write('D9', 'Qt EL', formatFichierOR)
        self.feuille.write('B10', 'Valide', formatCellule)
        self.feuille.write('B11', 'Non débuté', formatCellule)
        self.feuille.write('B12', 'Avec incohérence avancement                                                                                   (liste A)', formatCellule)
        self.feuille.write('B13', 'Complétude Infos FIS non conforme                                                                         (liste B)', formatCellule)
        self.feuille.write('B14', 'Complétude Infos FIS non conforme avec état négo pas de syndic                    (liste B)', formatCellule)
        self.feuille.write('B15', 'Codes IMB Pavillon (<4 EL) avec état négo déclaré sans FIS                                 (liste C)', formatCellule)


        self.feuille.write('B17', 'Répartition des valides', formatFichierVC)
        self.feuille.write('C17', 'Qt Code IMB', formatFichierVC)
        self.feuille.write('D17', 'Qt EL', formatFichierVC)
        self.feuille.write('B18', 'Non qualifié                                                                                                                     (liste D1)', formatCellule)
        self.feuille.write('B19', 'Qualifié                                                                                                                             (liste D2)', formatCellule)
        self.feuille.write('B20', 'Importé ZN                                                                                                                       (liste D3)', formatCellule)
        self.feuille.write('B21', 'Prev Accord OK                                                                                                                (liste D4)', formatCellule)
        self.feuille.write('B22', 'Accord OK                                                                                                                         (liste D5)', formatCellule)
        self.feuille.write('B23', 'DTA OK                                                                                                                              (liste D6)', formatCellule)
        self.feuille.write('B24', 'BPE OK                                                                                                                               (liste D7)', formatCellule)
        self.feuille.write('B25', 'BPT OK                                                                                                                               (liste D8)', formatCellule)




    def ajoutOngletPavSansFIS(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:C3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 5, 25)

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
        formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center', 'text_wrap': 'True'})
        
        self.feuille.write('D1', 'Pavillon sans FIS avec état négo déclaré', formatCellule)

        self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
        self.feuille.write('C1', 'Type', formatCellule)      
        self.feuille.write('E1', 'Centre négo', formatCellule)
        self.feuille.write('F1', None, formatCellule)

        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'Code regroupement syndic', formatFichier)
        self.feuille.write('C3', 'NB EL', formatFichier)


    # def ajoutOngletCA(self, classeur, feuille):
        
    #     self.feuille = self.classeur.add_worksheet(feuille)
        
    #     self.feuille.autofilter('A3:AC3')
    #     self.feuille.hide_gridlines(2)
    #     self.feuille.set_column(0, 28, 20)
    #     self.feuille.freeze_panes(3, 4)

    #     formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'})
    #     formatFichier = self.classeur.add_format({'border': True, 'bg_color': '#D9D9D9', 'align' : 'center'})

    #     self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
    #     self.feuille.write('C1', 'Type', formatCellule)
    #     self.feuille.write('E1', 'Centre négo', formatCellule)
    #     self.feuille.write('F1', None, formatCellule)

    #     self.feuille.write('A3', 'Code IMB', formatFichier)
    #     self.feuille.write('B3', 'Adresse', formatFichier)
    #     self.feuille.write('C3', 'Nb logements - OPT', formatFichier) 
    #     self.feuille.write('D3', 'Jalon en cours', formatFichier)
    #     self.feuille.write('E3', 'Qualifie', formatFichier)
    #     self.feuille.write('F3', 'ImportZN', formatFichier)
    #     self.feuille.write('G3', 'PrevAccord', formatFichier)
    #     self.feuille.write('H3', 'Accord', formatFichier)
    #     self.feuille.write('I3', 'DTA', formatFichier)
    #     self.feuille.write('J3', 'BPE', formatFichier)
    #     self.feuille.write('K3', 'BPT', formatFichier)
    #     self.feuille.write('L3', 'Etat Négociation - OPT', formatFichier)
    #     self.feuille.write('M3', 'Code Syndic - OPT', formatFichier)
    #     self.feuille.write('N3', 'Code regroup. Syndic - OPT', formatFichier)
    #     self.feuille.write('O3', 'Date envoi accord syndic - OPT', formatFichier)
    #     self.feuille.write('P3', 'AG non pertinente - OPT', formatFichier)
    #     self.feuille.write('Q3', 'Date Retour Convention - OPT', formatFichier)
    #     self.feuille.write('R3', 'Date de saisie de l’accord syndic - OPT', formatFichier)
    #     self.feuille.write('S3', 'Présence DTA - OPT', formatFichier)
    #     self.feuille.write('T3', 'Date mise à disposition pré-étude - OPT', formatFichier)
    #     self.feuille.write('U3', 'Date fin etd ZA PA-OPT', formatFichier)
    #     self.feuille.write('V3', 'Résultat de l’étude site - OPT', formatFichier)
    #     self.feuille.write('W3', 'Date envoi pré-étude au syndic - OPT', formatFichier)
    #     self.feuille.write('X3', 'Type site - OPT', formatFichier)
    #     self.feuille.write('Y3', 'Nb logements P - OPT', formatFichier)
    #     self.feuille.write('Z3', 'Nb logements R - OPT', formatFichier)
    #     self.feuille.write('AA3', 'Date récept. bon pré-étude - OPT', formatFichier)
    #     self.feuille.write('AB3', 'Date théorique raccordement - OPT', formatFichier)
    #     self.feuille.write('AC3', 'Date pose du PB - OPT', formatFichier)

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
        self.feuille.write('E11', '[Date validation pré-étude par syndic – OPT » non vide ET Date de saisie de l’accord (colonne FA) syndic non vide ET Date envoi pré-étude au syndic non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide et Présence DTA == 1-DTA présent]\n\nOU\n\n[Date validation pré-étude par syndic – OPT » non vide ET Date de saisie de l’accord (colonne FA) syndic non vide ET Date envoi pré-étude au syndic non vide ET Type de site non vide ET Nb logement P non vide ET Nb logements R non vide et Présence DTA == 2-DTA inutile]', formatCellule)
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
 


    def ecrireResultatCPLC(self, resultat, classeur, feuille) :
                
        formatCelCondiR = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#FF8F8F'})
        formatCelCondiV = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#92D050'})
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        formatDate = self.classeur.add_format({'border': True, 'align' : 'bottom', 'num_format': 'dd/mm/yyyy'})

        formatCelCondiP = self.classeur.add_format({'border': True, 'align' : 'bottom', 'bg_color' : '#CCE9FF'})


        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):
                if c in (6, 8, 9, 10, 11, 13, 14, 15, 16, 17, 19, 20, 26, 27, 28, 36):
                    self.feuille.write(r+3, c, col, formatDate)
                else :
                    self.feuille.write(r+3, c, col, formatCellule)       

                if row[c] =='Non':
                    self.feuille.write(r+3, c, col, formatCelCondiR)

                elif row[c] =='Oui':
                    self.feuille.write(r+3, c, col, formatCelCondiV)


            if row[4] == 'Non':

                if row[11] == '':
                    self.feuille.write(r+3, 11, row[11], formatCelCondiP)

            if row[5] == 'Non':

                if row[12] == '':
                    self.feuille.write(r+3, 12, row[12], formatCelCondiP)
 
            if row[6] == 'Non':

                if row[13] == '':
                    self.feuille.write(r+3, 13, row[13], formatCelCondiP)
                if row[14] == '':
                    self.feuille.write(r+3, 14, row[14], formatCelCondiP)

            if row[7] == 'Non':

                if row[15] == '':
                    self.feuille.write(r+3, 15, row[15], formatCelCondiP)
 
                if row[16] == '':
                    self.feuille.write(r+3, 16, row[16], formatCelCondiP)

            if row[8] == 'Non':

                if row[17] == '':
                    self.feuille.write(r+3, 17, row[17], formatCelCondiP)

            if row[9] == 'Non':

                if row[18] == '' or row[18] == '0-DTA non présent et nécessaire':
                    self.feuille.write(r+3, 18, row[18], formatCelCondiP)
 
                if row[19] == '':
                    self.feuille.write(r+3, 19, row[19], formatCelCondiP)

                if row[20] == '':
                    self.feuille.write(r+3, 20, row[20], formatCelCondiP)

                if row[21] == '':
                    self.feuille.write(r+3, 21, row[21], formatCelCondiP)

                if row[22] == '':
                    self.feuille.write(r+3, 22, row[22], formatCelCondiP)

                if row[23] == '':
                    self.feuille.write(r+3, 23, row[23], formatCelCondiP)