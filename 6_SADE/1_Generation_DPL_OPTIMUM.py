# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcelETD
from mesRequetesNEG import *

import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#     Génération des fichiers cycle 2 ETD     #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")


if os.path.exists("Fichier_DPL_OPT"):
	try :
		os.system("rmdir Fichier_DPL_OPT /S /Q")
		os.mkdir("Fichier_DPL_OPT")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:
	
	try :
		os.mkdir("Fichier_DPL_OPT")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()

if DPL <= Dep:

	ETD = bddSqlite.BaseSqlite()

	ETD.rechercher("SELECT name FROM sqlite_master")
	tables = ETD.resultat()

	print("Suppression des tables existantes en cours")

	for table in tables:

		ETD.rechercher("DROP TABLE " + table[0])
		print("Table " + table[0] + " supprimée")


	ETD.fermerBdd()

	
	if os.path.exists("ma_base.db"):

		os.remove("ma_base.db")

	print("Insertion des données OPTIMUM dans la table OPT")
	    
	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("OPT", OPT_request)
	ETD.insererDonneesOPT()
	        
	print("Le fichier : DPL_SADE a été chargé dans la table OPT")
	print("--------------------------------------------------------")


	ETD.fermerBdd()


	ETD = bddSqlite.BaseSqlite()

	ETD.rechercher("Update OPT set Resultat_etude_d2 = 'OK' where Resultat_etude_d2 = '1-OK'")
	ETD.rechercher("Update OPT set Resultat_etude_d2 = 'IMPOSSIBLE' where Resultat_etude_d2 = '3-IMPOSSIBLE'")
	ETD.rechercher("Update OPT set Resultat_etude_d2 = 'NO GO OPGC' where Resultat_etude_d2 = '2-NO GO OPGC'")

	ETD.rechercher("Update OPT set Etat_NEGO = 'PAS DE SYNDIC' where Dernier_Jalon in ('0_Négo non débuté', '') and NB_EL < 4")
	
	ETD.rechercher("update OPT set Date_reception_BPE = Date_d_accord_syndic + 10 where dernier_jalon = '6_BPE OK' and Presence_DTA <> ''")
	ETD.rechercher("update OPT set Date_reception_BPE = Date_de_validation_PE - 10 where dernier_jalon = '7_BPT OK' and Presence_DTA <> ''")

	ETD.rechercher("update OPT set Date_prog_racco = Date_pose_PB where Numero_fci_acces <> '' and Date_pose_PB <> ''")
	ETD.rechercher("update OPT set Date_prog_racco = " + "'" + str(DPL) + "'" + " where Numero_fci_acces <> '' and Date_pose_PB = ''")

	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()

	OPT = exportExcelETD.ajoutClasseur("DPL_SADE")
	OPT.ajoutOnglet_D("DPL", "DPL")

	ETD.rechercher(DPL_OPT)

	v = ETD.resultat()
	OPT.ecrireResultat(v, "OPT", "OPT")


	OPT.fermeture("DPL_SADE")

	ETD.fermerBdd()
	
else:
	
	ETD = bddSqlite.BaseSqlite()
	ETD.rechercher("Update OPT set AG_Non_pertinente = '0 AG Pertinente'  where AG_Non_pertinente = 'Oui'")
	
os.system("pause")