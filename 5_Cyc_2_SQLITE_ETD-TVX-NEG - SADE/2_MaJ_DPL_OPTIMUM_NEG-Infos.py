# -*- coding: utf-8 -*-
import sqlite3
import xlrd
import bddSqlite
import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Mise à jour données DPL          #")
os.system("@echo 						#            Infos et avancement Négo         #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")



for element in os.listdir('data'):

	if element[2:23] == 'Completudes_infos_NOK':

		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		ETD = bddSqlite.BaseSqlite()
		print('mise à jour du fichier ' + element)
		
		for row in range (3, sheet.nrows) :
			liste = (sheet.row_values(row))
			
			ETD.rechercher("Update OPT set CRS = " + "'" + liste[3] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Code_Syndic = " + "'" + liste[4] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set ETAT_NEGO = " + "'" + liste[5] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_envoi_AS = " + "'" + str(liste[6]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set AG_non_pertinente = " + "'" + liste[7] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_retour_convention = " + "'" + str(liste[8]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_saisie_AS = " + "'" + str(liste[9]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_reception_bon_PE = " + "'" + str(liste[10]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("update OPT set attente_probation = '7-Autres'  where date_retour_convention <> '' and attente_probation in ('', 'Optionnelle')")
		
		ETD.fermerBdd()

	elif element[2:24] == 'completude_avc_neg_NOK':

		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		ETD = bddSqlite.BaseSqlite()

		print('mise à jour du fichier ' + element)
		for row in range (3, sheet.nrows) :
			liste = (sheet.row_values(row))
			# print(liste)

			ETD.rechercher("Update OPT set ETAT_NEGO = " + "'" + liste[10] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Code_Syndic = " + "'" + liste[11] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set CRS = " + "'" + liste[12] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_envoi_AS = " + "'" + str(liste[13]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set AG_non_pertinente = " + "'" + liste[14] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_retour_convention = " + "'" + str(liste[15]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_saisie_AS = " + "'" + str(liste[16]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_reception_bon_PE = " + "'" + str(liste[17]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Presence_DTA = " + "'" + liste[18] + "'" + " where code_IMB = " + "'" + liste[0] + "'")	
			ETD.rechercher("Update OPT set Date_env_pe_syndic = " + "'" + str(liste[19]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Date_valid_PE = " + "'" + str(liste[20]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set NB_EL_P = " + "'" + str(liste[21]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set NB_EL_R = " + "'" + str(liste[22]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Type_Site = " + "'" + liste[23] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Suspendre = " + "'" + liste[24] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update OPT set Type_suspension = " + "'" + liste[25] + "'" + " where code_IMB = " + "'" + liste[0] + "'")


		ETD.rechercher("Update OPT set AG_non_pertinente = '0-AG pertinente' where AG_non_pertinente = 'Non'")
		ETD.rechercher("Update OPT set AG_non_pertinente = '1-AG non pertinente' where AG_non_pertinente = 'Oui'")

		ETD.rechercher("Update OPT set Presence_DTA = '1-DTA présent' where Presence_DTA = 'Présent'")
		ETD.rechercher("Update OPT set Presence_DTA = '0-DTA non présent et nécessaire' where Presence_DTA = 'Pas présent'")
		ETD.rechercher("Update OPT set Presence_DTA = '2-DTA inutile' where Presence_DTA = 'Pas nécessaire'")

		ETD.rechercher("update OPT set attente_probation = '7-Autres'  where date_retour_convention <> '' and attente_probation in ('', 'Optionnelle')")

		ETD.fermerBdd()
		
os.system("pause")
