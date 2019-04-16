# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcelNEG
import exportExcelETD
from mesRequetesNEG import *
from mesRequetesETD import *
import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Insertion données DPL            #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")


if DPL <= Dep:
	print("\x1b[2;41;77mSelectionner le type d'ingénierie de la plaque 1:ZTD HD\x1b[0m")
	print("\x1b[2;41;77m                                               2:ZMD-ZTD BD\x1b[0m")

	ING = input("Valeur 1 ou 2 : ")

	if ING not in ('1', '2') :
		print("Choix incorrect")
		exit()

	else:

		if ING == '1' :

			print("--------------------------------------------------------------------------------")
			print("Type d'ingénierie : ZTD HD ")
			print("--------------------------------------------------------------------------------")

		elif ING == '2' :

			print("--------------------------------------------------------------------------------")
			print("Type d'ingénierie : ZMD-ZTD BD ")
			print("--------------------------------------------------------------------------------")

	if ING	in ('1'):
		print("Le PM sera pris en compte pour la définition des ensembles et des regroupements - un fichier admin_PA sera généré ")
		print("--------------------------------------------------------------------------------")


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

	print("Insertion des données DPL dans la table OPT")
	    
	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("OPT", OPT_request)
	ETD.insererDonneesOPT()
	        
	print("Le fichier DPL a été chargé dans la table")
	print("--------------------------------------------------------")

	if ING == '1' :

		ETD.rechercher("Update OPT set Table_etd = 'ZTD HD' ")

	elif ING == '2' :
		ETD.rechercher("Update OPT set Table_etd = 'ZMD-ZTD BD' ")


	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()

	ETD.rechercher("Update OPT set PA = '' where PA = 'N.D.'")
	ETD.rechercher("Update OPT set PM = '' where PM = 'N.D.'")
	ETD.rechercher("Update OPT set PA_INI = PA")
	ETD.rechercher("Update OPT set Escalier = 0 where Escalier = 'Optionnelle'")
	ETD.rechercher("Update OPT set Resultat_etude_site = '' where Resultat_etude_site = 'Optionnelle'")
	ETD.rechercher("Update OPT set AG_non_pertinente = '' where AG_non_pertinente = 'Optionnelle'")
	ETD.rechercher("Update OPT set attente_probation = '' where attente_probation = 'Optionnelle'")


	ETD.rechercher("Update OPT set AG_non_pertinente = '0-AG pertinente' where AG_non_pertinente = 'Non'")
	ETD.rechercher("Update OPT set AG_non_pertinente = '1-AG non pertinente' where AG_non_pertinente = 'Oui'")

	ETD.rechercher("Update OPT set Presence_DTA = '1-DTA présent' where Presence_DTA = 'Présent'")
	ETD.rechercher("Update OPT set Presence_DTA = '0-DTA non présent et nécessaire' where Presence_DTA = 'Pas présent'")
	ETD.rechercher("Update OPT set Presence_DTA = '2-DTA inutile' where Presence_DTA = 'Pas nécessaire'")

	ETD.rechercher("Update OPT set Attente_Probation = '1-T/D1 non effectué' where Attente_Probation = 'T/D1 non effectué'")
	ETD.rechercher("Update OPT set Attente_Probation = '2-Attente DTA' where Attente_Probation = 'Attente DTA'")
	ETD.rechercher("Update OPT set Attente_Probation = '3-Raccordement aérien' where Attente_Probation = 'Raccordement aérien'")
	ETD.rechercher("Update OPT set Attente_Probation = '4-Imb -12logmts qu’on ne sait pas produire' where Attente_Probation = 'Imb -12logmts qu’on ne sait pas produire'")
	ETD.rechercher("Update OPT set Attente_Probation = '5-Accord Free' where Attente_Probation = 'Accord Free'")
	ETD.rechercher("Update OPT set Attente_Probation = '6-PME non déployé' where Attente_Probation = 'PME non déployé'")
	ETD.rechercher("Update OPT set Attente_Probation = '7-Autres' where Attente_Probation = 'Autres'")

	ETD.rechercher("update OPT set attente_probation = '7-Autres'  where date_retour_convention <> '' and attente_probation in ('', 'Optionnelle')")

	ETD.rechercher("Update OPT set Type_DTA = 'Sans objet' where Presence_DTA = '0-DTA non présent et nécessaire'")
	ETD.rechercher("Update OPT set Type_DTA = 'Attestation' where Presence_DTA = '2-DTA inutile'")
	ETD.rechercher("Update OPT set Type_DTA = 'DTA' where Presence_DTA = '1-DTA présent'")

	ETD.rechercher("Update OPT set NB_EL = (NB_EL_P + NB_EL_R) where NB_EL <> (NB_EL_P + NB_EL_R)")
	
	ETD.rechercher("Update OPT set  Type_site = 'C-Pour un immeuble' where NB_EL > 2")
	ETD.rechercher("Update OPT set  Type_site = 'S-Pour une maison individuelle' where NB_EL < 3")


	ETD.rechercher("Update OPT set PB = 'N.D.' where PB = ''")



	
	if ING == '1':
		ETD.rechercher("Update OPT set PA = PM ")


	ETD.rechercher(Update_statut_affectation)

	ETD.rechercher(Update_01)
	ETD.rechercher(Update_02)
	ETD.rechercher(Update_03)
	ETD.rechercher(Update_04)
	ETD.rechercher(Update_05)

	ETD.rechercher(Update_Mob_contact_syndic)
	ETD.rechercher(Update_Tel_contact_syndic)
	ETD.rechercher(Update_Tel_PDT_SYN)
	ETD.rechercher(Update_Fax_PDT_SYN)
	ETD.rechercher(Update_Mob_PDT_SYN)

	ETD.fermerBdd()

else:
	
	ETD = bddSqlite.BaseSqlite()
	ETD.rechercher("Drop table")
	
os.system("pause")