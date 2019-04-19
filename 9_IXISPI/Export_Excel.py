# -*- coding: utf-8 -*-
import xlsxwriter

class ajoutClasseur :
    
    # Création du fichier Excel
    
    def __init__(self, classeur):
        
        self.classeur  = xlsxwriter.Workbook(classeur + '.xlsx')
                        
    # Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
    
    def ajoutOnglet(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.autofilter('A3:R3')
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 17, 21)
        
        formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
         
        
        self.feuille.write('A3', 'Code IMB', formatFichier)
        self.feuille.write('B3', 'ID Chantier', formatFichier)
        self.feuille.write('C3', 'NB_EL', formatFichier)
        self.feuille.write('D3', 'Clé site', formatFichier)
        self.feuille.write('E3', 'Statut Syndic', formatFichier)
        self.feuille.write('F3', 'Action', formatFichier)
        self.feuille.write('G3', 'Réalisation', formatFichier)
        self.feuille.write('H3', 'Lancement', formatFichier)
        self.feuille.write('I3', 'BPT', formatFichier)
        self.feuille.write('J3', 'Etude', formatFichier)
        self.feuille.write('K3', 'Probation', formatFichier)
        self.feuille.write('L3', 'Qualif', formatFichier)


    def ajoutOngletSyn(self, classeur, feuille):
        
        self.feuille = self.classeur.add_worksheet(feuille)
        
        self.feuille.hide_gridlines(2)
        self.feuille.set_column(0, 11, 18)
        
        formatFichier = self.classeur.add_format({'border': True, 'align' : 'center', 'valign' : 'center', 'text_wrap' : True})
        formatFichierB = self.classeur.add_format({'border': True,'bg_color': '#97E4FF', 'align' : 'center','valign' : 'top', 'text_wrap' : True})
        formatFichierO = self.classeur.add_format({'border': True, 'bg_color': '#FFE285', 'align' : 'center','valign' : 'top', 'text_wrap' : True})
        formatFichierG = self.classeur.add_format({'border': True, 'bg_color': '#BFBFBF', 'align' : 'center','valign' : 'top', 'text_wrap' : True})
                          
        self.feuille.merge_range('B2:F2', 'Nom code plaque', formatFichierB)
        self.feuille.merge_range('G2:L2', 'Contrôles de cohérences', formatFichierO)  
        self.feuille.write('B3', 'Table contrôle des codes IMB IMM', formatFichier)
        self.feuille.write('C3', 'Nb code IMB publiés par action', formatFichierB)
        self.feuille.write('D3', 'Nb EL publiés par action', formatFichierB)
        self.feuille.write('E3', 'Nb code IMB publiés par action avec anomalie', formatFichierB)
        self.feuille.write('F3', 'Nb EL publiés par action avec anomalie', formatFichierB)

        self.feuille.write('G3', 'Nb code IMB non publiés Qualif négo', formatFichierO)
        self.feuille.write('H3', 'Nb code IMB non publiés Probation', formatFichierO)
        self.feuille.write('I3', 'Nb code IMB non publiés Réalisation étude', formatFichierO)
        self.feuille.write('J3', 'Nb code IMB non publiés BPT', formatFichierO)
        self.feuille.write('K3', 'Nb code IMB non publiés Lancement travaux', formatFichierO)
        self.feuille.write('L3', 'Nb code IMB non publiés Réalisation travaux', formatFichierO)

        self.feuille.write('B4', 'Qualif négo', formatFichier)
        self.feuille.write('B5', 'Probation', formatFichier)
        self.feuille.write('B6', 'Réalisation étude', formatFichier)
        self.feuille.write('B7', 'BPT', formatFichier)
        self.feuille.write('B8', 'Lancement travaux', formatFichier)
        self.feuille.write('B9', 'Réalisation travaux', formatFichier)
        self.feuille.write('B10', 'TFX', formatFichier)


        self.feuille.merge_range('B14:F14', 'Nom code plaque', formatFichierB)
        self.feuille.merge_range('G14:L14', 'Contrôles de cohérences', formatFichierO) 
        self.feuille.write('B15', 'Table contrôle des codes IMB PAV (<= 3EL)', formatFichier)
        self.feuille.write('C15', 'Nb code IMB publiés par action', formatFichierB)
        self.feuille.write('D15', 'Nb EL publiés par action', formatFichierB)
        self.feuille.write('E15', 'Nb code IMB publiés par action avec anomalie', formatFichierB)
        self.feuille.write('F15', 'Nb EL publiés par action avec anomalie', formatFichierB)

        self.feuille.write('G15', 'Nb code IMB non publiés Qualif négo', formatFichierO)
        self.feuille.write('H15', 'Nb code IMB non publiés Probation', formatFichierO)
        self.feuille.write('I15', 'Nb code IMB non publiés Réalisation étude', formatFichierO)
        self.feuille.write('J15', 'Nb code IMB non publiés BPT', formatFichierO)
        self.feuille.write('K15', 'Nb code IMB non publiés Lancement travaux', formatFichierO)
        self.feuille.write('L15', 'Nb code IMB non publiés Réalisation travaux', formatFichierO)

        self.feuille.write('B16', 'Qualif négo', formatFichier)
        self.feuille.write('B17', 'Probation', formatFichier)
        self.feuille.write('B18', 'Réalisation étude', formatFichier)
        self.feuille.write('B19', 'BPT', formatFichier)
        self.feuille.write('B20', 'Lancement travaux', formatFichier)
        self.feuille.write('B21', 'Réalisation travaux', formatFichier)
        self.feuille.write('B22', 'TFX', formatFichier)

        self.feuille.write('E4', None, formatFichierG)
        self.feuille.write('F4', None, formatFichierG)
        self.feuille.write('G4', None, formatFichierG)
        self.feuille.write('H4', None, formatFichierG)
        self.feuille.write('I4', None, formatFichierG)
        self.feuille.write('J4', None, formatFichierG)
        self.feuille.write('K4', None, formatFichierG)
        self.feuille.write('L4', None, formatFichierG)

        self.feuille.write('H5', None, formatFichierG)
        self.feuille.write('I5', None, formatFichierG)
        self.feuille.write('J5', None, formatFichierG)
        self.feuille.write('K5', None, formatFichierG)
        self.feuille.write('L5', None, formatFichierG)

        self.feuille.write('I6', None, formatFichierG)
        self.feuille.write('J6', None, formatFichierG)
        self.feuille.write('K6', None, formatFichierG)
        self.feuille.write('L6', None, formatFichierG)

        self.feuille.write('J7', None, formatFichierG)
        self.feuille.write('K7', None, formatFichierG)
        self.feuille.write('L7', None, formatFichierG)

        self.feuille.write('K8', None, formatFichierG)
        self.feuille.write('L8', None, formatFichierG)
        self.feuille.write('L9', None, formatFichierG)

        self.feuille.write('C16', None, formatFichierG)
        self.feuille.write('D16', None, formatFichierG)
        self.feuille.write('C17', None, formatFichierG)
        self.feuille.write('D17', None, formatFichierG)
        self.feuille.write('C19', None, formatFichierG)
        self.feuille.write('D19', None, formatFichierG)

        self.feuille.write('G20', None, formatFichierG)
        self.feuille.write('H20', None, formatFichierG)
        self.feuille.write('J20', None, formatFichierG)
        self.feuille.write('K20', None, formatFichierG)
        self.feuille.write('L20', None, formatFichierG)               
        self.feuille.write('G21', None, formatFichierG)
        self.feuille.write('H21', None, formatFichierG)
        self.feuille.write('J21', None, formatFichierG)
        self.feuille.write('L21', None, formatFichierG)
        self.feuille.write('G22', None, formatFichierG)
        self.feuille.write('H22', None, formatFichierG)
        self.feuille.write('J22', None, formatFichierG)

        row = 15
        while row < 19:
            col = 4
            while col < 12:
                self.feuille.write(row, col, None, formatFichierG)
                col = col + 1
            row = row + 1




    def ecrireResultat(self, resultat, classeur, feuille) :
        
        formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom'}) 
        self.resultat = resultat
        
        for r, row in enumerate(resultat):
            for c, col in enumerate(row):
                self.feuille.write(r+3, c, col, formatCellule)


    def ecrireVal(self, cellule, valeur, classeur, feuille):

        formatCellule = self.classeur.add_format({'border': True, 'align' : 'center'}) 

        self.feuille.write(cellule, valeur, formatCellule)
                
    def ecrireVal2(self, cellule, valeur, classeur, feuille):

        formatFichierB = self.classeur.add_format({'border': True,'bg_color': '#97E4FF', 'align' : 'center','valign' : 'top', 'text_wrap' : True})
        self.feuille.write(cellule, valeur, formatFichierB)
        
    def fermeture(self,classeur):

        self.classeur.close()
        