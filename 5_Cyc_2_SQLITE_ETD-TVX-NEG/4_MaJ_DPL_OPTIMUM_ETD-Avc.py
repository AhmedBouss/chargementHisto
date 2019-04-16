# -*- coding: utf-8 -*-
import sqlite3
import xlrd
import bddSqlite
from mesRequetesETD import *
import os
os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Mise à jour données DPL          #")
os.system("@echo 						#              avancement études              #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")

for element in os.listdir('data'):

	if element[2:28] == 'completude_avc_etd_tvx_NOK':

		
		ETD = bddSqlite.BaseSqlite()
		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		print('mise à jour du fichier ' + element)
		for row in range (3, sheet.nrows) :

			liste = (sheet.row_values(row))

			

			ETD.rechercher("Update C set ETAT_NEGO = " + "'" + liste[14] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Code_Syndic = " + "'" + liste[15] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set CRS = " + "'" + liste[16] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_envoi_AS = " + "'" + str(liste[17]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set AG_non_pertinente = " + "'" + liste[18] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_retour_convention = " + "'" + str(liste[19]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_saisie_AS = " + "'" + str(liste[20]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_reception_bon_PE = " + "'" + str(liste[21]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_mad_PE = " + "'" + str(liste[22]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			# ETD.rechercher("Update C set Date_fin_EZA = " + "'" + str(liste[23]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Presence_DTA = " + "'" + liste[24] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_env_pe_syndic = " + "'" + str(liste[25]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Type_Site = " + "'" + liste[26] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set NB_EL_P = " + "'" + str(liste[27]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set NB_EL_R = " + "'" + str(liste[28]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_valid_PE = " + "'" + str(liste[29]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Resultat_etude_site = " + "'" + liste[30] + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_prog_racco = " + "'" + str(liste[31]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Date_pose_PB = " + "'" + str(liste[32]) + "'" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Suspendre = " + "\""+ liste[33] + "\"" + " where code_IMB = " + "'" + liste[0] + "'")
			ETD.rechercher("Update C set Type_suspension = " + "\"" + liste[34] + "\"" + " where code_IMB = " + "'" + liste[0] + "'")


		ETD.rechercher("Update C set AG_non_pertinente = '0-AG pertinente' where AG_non_pertinente = 'Non'")
		ETD.rechercher("Update C set AG_non_pertinente = '1-AG non pertinente' where AG_non_pertinente = 'Oui'")

		ETD.rechercher("Update C set Presence_DTA = '1-DTA présent' where Presence_DTA = 'Présent'")
		ETD.rechercher("Update C set Presence_DTA = '0-DTA non présent et nécessaire' where Presence_DTA = 'Pas présent'")
		ETD.rechercher("Update C set Presence_DTA = '2-DTA inutile' where Presence_DTA = 'Pas nécessaire'")

		ETD.rechercher("Update C set Piquetage = 'Non', Etudes = 'Non',  BPT = 'Non', Travaux_Lances = 'Non', DOE = 'Non', BPE = 'Non'")

		
		ETD.rechercher(updateNondebute)
		ETD.rechercher(updateQualifie)
		ETD.rechercher(updateImportZN)
		ETD.rechercher(updatePrevAccord)
		ETD.rechercher(updateAccord)
		ETD.rechercher(updateDTA)
		ETD.rechercher(updateBPE)
		ETD.rechercher(updatePiquetage)
		ETD.rechercher(updateBPT)
		ETD.rechercher(updateEtudes)
		ETD.rechercher(updateTravauxLances)
		ETD.rechercher(updateTravauxRealises)
		ETD.rechercher(updateDOE)


		ETD.rechercher(JalonSansProd)
		ETD.rechercher(JalonQualif)
		ETD.rechercher(JalonImportZN)
		ETD.rechercher(JalonPrevAccord)
		ETD.rechercher(JalonAccord)
		# ETD.rechercher(JalonDTA)
		ETD.rechercher(JalonBPE)
		ETD.rechercher(JalonPiquetage)
		ETD.rechercher(JalonBPT)
		ETD.rechercher(JalonEtude)
		ETD.rechercher(JalonTL)
		ETD.rechercher(JalonDOE)


		ETD.fermerBdd()

os.system("pause")