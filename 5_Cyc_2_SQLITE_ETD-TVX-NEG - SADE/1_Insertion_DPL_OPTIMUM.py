# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import IXI
from mesRequetesNEG import *
from mesRequetesETD import *
import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Insertion données DPL            #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")


num_plaque = input("Entrer le numéro de plaque : ")

print("Le numéro de plaque est :" , num_plaque)

if os.path.exists("Code_IMB_INEXISTANTS"):

	try :
		os.system("rmdir Code_IMB_INEXISTANTS /S /Q")
		os.mkdir("Code_IMB_INEXISTANTS")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:

	try :
		os.mkdir("Code_IMB_INEXISTANTS")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()

if DPL <= Dep:

	ING = input("Entrer le type d'ingénierie de la plaque ZTD ou ZMD : ")

	if ING not in ('ZTD', 'ZMD') :
		print("Choix incorrect")
		exit()

	else:

		print("--------------------------------------------------------------------------------")
		print("Type d'ingénierie :", ING)
		print("--------------------------------------------------------------------------------")

	if ING	in ('ZTD'):
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
	        
	print("Le fichier DPL SADE a été chargé dans la table")
	print("--------------------------------------------------------")


	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTableIXI()
	ETD.insererDonneesIXI()
	        
	print("Le fichier DPL IXIN a été chargé dans la table")
	print("--------------------------------------------------------")


	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()

	ETD.rechercher("Update OPT set Table_etd = " + "'" + ING + "'")
	ETD.rechercher("Update OPT set plaque = " + "'" + num_plaque + "'")

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
	
	
	ETD.rechercher("Update OPT set FCI_ACCES = Numero_FCI_Acces")
	ETD.rechercher("Update OPT set FCI_Fin = Numero_FCI_Fin")

	

	if ING == 'ZTD':
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



	IMB_inexistant = IXI.ajoutClasseur("IMB_inexistant")
	IMB_inexistant.ajoutOnglet_IXI("IMB_inexistant", "IMB_inexistant")

	ETD.rechercher("select code_IMB, Adresse, localite from OPT where code_IMB not in (select code_imb from IXI)")
	v = ETD.resultat()
	IMB_inexistant.ecrireResultat(v, "IMB_inexistant", "IMB_inexistant")

	IMB_inexistant.fermeture("IMB_inexistant")

	ETD.rechercher("Delete from OPT where code_IMB not in (select code_imb from IXI) ")

	ETD.fermerBdd()

else:
	
	ETD = bddSqlite.BaseSqlite()
	ETD.rechercher("Drop table")
	
os.system("pause")