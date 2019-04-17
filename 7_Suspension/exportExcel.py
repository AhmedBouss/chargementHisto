# -*- coding: utf-8 -*-
import xlsxwriter
from datetime import date

class ajoutClasseur :
	  
	# Création du fichier Excel
	
	def __init__(self, classeur):
		
		self.classeur  = xlsxwriter.Workbook('Fichiers_Suspension/' + classeur + '.xlsx')
						
	# Ajout d'un nouvel ongelt avec initialisation du format pour les fichiers Piquetage non réalisé , Piquetage réalisé, Qualif
	


	def ajoutOnglet_SuspendreE(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:D3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 4, 21)
		self.feuille.set_column(5, 6, 40)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)
		self.feuille.write('D1', 'Suspension RGT', formatCellule)          
		self.feuille.write('E1', 'Centre etudes', formatCellule)
		self.feuille.write('F1', None, formatCellule)

		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'Code plaque', formatFichier)
		self.feuille.write('C3', 'Type Suspension', formatFichier)
		self.feuille.write('D3', 'Commentaire suspension', formatFichier)


	def ajoutOnglet_SuspendreN(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:C3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)    
		self.feuille.write('D1', 'Suspension ZN', formatCellule)       
		self.feuille.write('E1', 'Centre Nego', formatCellule)
		self.feuille.write('F1', None, formatCellule)

		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'Type suspension', formatFichier)
		self.feuille.write('C3', 'Commentaire suspension', formatFichier)


	def ajoutOnglet_Commentaire(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:F3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)    
		self.feuille.write('D1', 'Suspension ZN', formatCellule)       
		self.feuille.write('E1', 'Centre Nego', formatCellule)
		self.feuille.write('F1', None, formatCellule)

		self.feuille.write('A3', 'Type commentaire', formatFichier)
		self.feuille.write('B3', 'Code IMB', formatFichier)
		self.feuille.write('C3', 'ZN', formatFichier)
		self.feuille.write('D3', 'Date commentaire', formatFichier)
		self.feuille.write('E3', 'Resp. commentaire', formatFichier)
		self.feuille.write('F3', 'Commentaire', formatFichier)


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

		
	def fermeture(self,classeur):

		self.classeur.close()
		


	def ajoutOnglet_action(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:G3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)
        

		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'Code plaque', formatFichier)
		self.feuille.write('C3', 'Action', formatFichier)
		self.feuille.write('D3', 'ID_ZN', formatFichier)
		self.feuille.write('E3', 'ID_ZE', formatFichier)
		self.feuille.write('F3', 'ID_RGT', formatFichier)
		self.feuille.write('G3', 'Statut code IMB', formatFichier)

	def ajoutOnglet_action_neg(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:F3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)
        

		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'Code plaque', formatFichier)
		self.feuille.write('C3', 'Existence SPI', formatFichier)
		self.feuille.write('D3', 'Statut code IMB', formatFichier)
		self.feuille.write('E3', 'ID_ZN', formatFichier)
		self.feuille.write('F3', 'Erreur', formatFichier)




	def ajoutOnglet_action_etd(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:G3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A1', 'Nom Code Plaque', formatCellule)
		self.feuille.write('C1', 'Type', formatCellule)
        

		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'Code plaque', formatFichier)
		self.feuille.write('C3', 'Existence SPI', formatFichier)
		self.feuille.write('D3', 'Statut code IMB', formatFichier)
		self.feuille.write('E3', 'ID_ZE', formatFichier)
		self.feuille.write('F3', 'ID_RGT', formatFichier)
		self.feuille.write('G3', 'Statut code IMB', formatFichier)



	def ajoutOnglet_CEM_OPT_ZN(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:C3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		
		self.feuille.write('A3', 'Code IMB', formatFichier)
		self.feuille.write('B3', 'ID ZN', formatFichier)
		self.feuille.write('C3', 'Type processus', formatFichier)

	def ajoutOnglet_CEM_OPT_ZE(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:B3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A3', 'ID ZE', formatFichier)
		self.feuille.write('B3', 'Type processus', formatFichier)


	def ajoutOnglet_CEM_OPT_RGT(self, classeur, feuille):
		
		self.feuille = self.classeur.add_worksheet(feuille)
		
		self.feuille.autofilter('A3:B3')
		self.feuille.hide_gridlines(2)
		self.feuille.set_column(0, 6, 31)
		
		formatCellule = self.classeur.add_format({'border': True, 'align' : 'bottom', 'text_wrap': 'True'})
	 
		formatFichier = self.classeur.add_format({'border': True, 'font_color': 'white', 'bg_color': '#0066CC', 'align' : 'center'})
		 
		self.feuille.write('A3', 'ID RGT', formatFichier)
		self.feuille.write('B3', 'Type processus', formatFichier)