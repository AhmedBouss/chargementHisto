# -*- coding: utf-8 -*-
import sqlite3
import xlrd
import bddSqlite
from mesRequetesETD import *
from datetime import datetime
import os

import os
os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Mise à jour données DPL          #")
os.system("@echo 						#                   FCI - EZA                 #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")

debut =  datetime(1900, 1, 1, 0, 0, 0)
fin = datetime.today()
auj = fin.strftime('%d/%m/%Y')
cde = datetime.today().strftime('%d%m%y')

cde = 'F99999' + str(cde)



for element in os.listdir('data'):

	if element[0:42] == 'Numéro FCI Fin travaux à saisir - Liste I3':

		
		ETD = bddSqlite.BaseSqlite()
		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		print('mise à jour du fichier ' + element)
		for row in range (3, sheet.nrows) :

			liste = (sheet.row_values(row))
			a = liste[3]

			if liste[3] != '' :

				date = datetime(int('20'+a[10]+a[11]), int(a[8]+a[9]), int(a[6]+a[7]))

				date_fci_nb = (date - debut ).days + 2

				ETD.rechercher("Update C set fci_fin = " + "'" + liste[3] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set fci_fin = " + "'" + liste[3] + "'" + " where code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set date_fci = " + "'" + str(date_fci_nb) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set date_fci = " + "'" + str(date_fci_nb) + "'" + " where code_IMB = " + "'" + liste[0] + "'")				

			elif liste[3] == '':

				date = datetime.today()

				date_fci_nb = (date - debut ).days + 2

				ETD.rechercher("Update C set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and  code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and  code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and jalon in ('Etude', 'Travaux lancés') and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and jalon in ('Etude', 'Travaux lancés') and  code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set fci_fin = " + "'" + str(cde) + "'"  + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set fci_fin = " + "'" + str(cde) + "'"  + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set date_fci = " + "'" + str(date_fci_nb) + "'" + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set date_fci = " + "'" + str(date_fci_nb) + "'" + "where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

		ETD.fermerBdd()

# ---------------------------------------------------------------------------------------------------------------------------------------------

	elif element[0:47] == "Numéro FCI commande d'accès à saisir - Liste I2":

		
		ETD = bddSqlite.BaseSqlite()
		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		print('mise à jour du fichier ' + element)
		for row in range (3, sheet.nrows) :

			liste = (sheet.row_values(row))
			a = liste[3]
			
			
			if liste[3] != '' :

				date = datetime(int('20'+a[10]+a[11]), int(a[8]+a[9]), int(a[6]+a[7]))

				date_fci_nb = (date - debut ).days + 2

				ETD.rechercher("Update C set fci_acces = " + "'" + liste[3] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set fci_acces = " + "'" + liste[3] + "'" + " where code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set date_fci = " + "'" + str(date_fci_nb) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set date_fci = " + "'" + str(date_fci_nb) + "'" + " where code_IMB = " + "'" + liste[0] + "'")

			elif liste[3] == '' :

				date = datetime.today()

				date_fci_nb = (date - debut ).days + 2

				ETD.rechercher("Update C set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and  code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and  code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and jalon in ('Etude', 'Travaux lancés') and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and jalon in ('Etude', 'Travaux lancés') and  code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set fci_acces = " + "'" + str(cde) + "'"  + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set fci_acces = " + "'" + str(cde) + "'"  + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set date_fci = " + "'" + str(date_fci_nb) + "'" + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set date_fci = " + "'" + str(date_fci_nb) + "'" + "where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

		ETD.fermerBdd()


# ---------------------------------------------------------------------------------------------------------------------------------------------

	elif element[0:49] == "Etude réalisée - Date Fin EZA à saisir - Liste I1":

		
		ETD = bddSqlite.BaseSqlite()

		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		print('mise à jour du fichier ' + element)
		for row in range (3, sheet.nrows) :

			liste = (sheet.row_values(row))

			date = datetime.today()

			date_EZA = (date - debut ).days + 2

			if liste[3] != '' :

				ETD.rechercher("Update C set Date_fin_EZA = " + "'" + str(liste[3]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Date_fin_EZA = " + "'" + str(liste[3]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")

			elif liste[3] == '' :

				ETD.rechercher("Update C set Resultat_etude_site = '', Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Resultat_etude_site = '', Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Oui' and code_IMB = " + "'" + liste[0] + "'")

				ETD.rechercher("Update C set Resultat_etude_site = '', Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				ETD.rechercher("Update B1 set Resultat_etude_site = '', Date_prog_racco = '', Date_pose_PB = '' where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

				# ETD.rechercher("Update C set Date_fin_EZA = " + "'" + str(date_EZA) + "'" + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")
				# ETD.rechercher("Update B1 set Date_fin_EZA = " + "'" + str(date_EZA) + "'" + " where SPI = 'Non' and code_IMB = " + "'" + liste[0] + "'")

		ETD.fermerBdd()

ETD = bddSqlite.BaseSqlite()

ETD.rechercher("Update C set  Etudes = 'Non', Travaux_Lances = 'Non', Travaux_realises = 'Non', DOE = 'Non'")


ETD.rechercher(updateEtudes)
ETD.rechercher(updateTravauxLances)
ETD.rechercher(updateTravauxRealises)
ETD.rechercher(updateDOE)

ETD.rechercher(JalonPiquetage)
ETD.rechercher(JalonBPT)
ETD.rechercher(JalonEtude)
ETD.rechercher(JalonTL)
ETD.rechercher(JalonDOE)


ETD.fermerBdd()

os.system("pause")