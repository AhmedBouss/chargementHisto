# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcel
from mesRequetes import *
import time

import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#     Génération des fichiers cycle 2 ETD     #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")


if os.path.exists("Fichiers_complementaires"):
	try :
		os.system("rmdir Fichiers_complementaires /S /Q")
		os.mkdir("Fichiers_complementaires")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:
	
	try :
		os.mkdir("Fichiers_complementaires")

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
	ETD.creerTableIXI()
	ETD.insererDonneesIXI()
	        
	print("Le fichier : IXI a été chargé dans la table IXI")
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

	

	ETD.rechercher(infos)

	ETD.rechercher("Update infos set date_action = " + "'" + str(time.strftime("%d/%m/%Y %H:%M:%S")) + "'" )
	ETD.rechercher("Update infos set Nom_contact = 'Histo' where Nom_contact = '' and (tel_contact <> '' OR  mail_contact <> '' OR adresse_contact <> '')" )

	ETD.rechercher("Update infos set Adduction_pb_vers_PA = 'Immeuble' where  upper(replace(Adduction_pb_vers_PA,'é','e')) = 'IMMEUBLE' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_PA = 'Aéro-souterrain' where  upper(replace(Adduction_pb_vers_PA,'é','e')) = 'AERO-SOUTERRAIN' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_PA = 'Souterrain' where  upper(replace(Adduction_pb_vers_PA,'é','e')) = 'SOUTERRAIN' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_PA = 'Façade' where  upper(replace(Adduction_pb_vers_PA,'ç','c')) = 'FACADE' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_PA = 'Inconnu' where  upper(replace(Adduction_pb_vers_PA,'é','e')) = 'INCONNU' " )

	ETD.rechercher("Update infos set Adduction_pb_vers_EL = 'Immeuble' where  upper(replace(Adduction_pb_vers_EL,'é','e')) = 'IMMEUBLE' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_EL = 'Aéro-souterrain' where  upper(replace(Adduction_pb_vers_EL,'é','e')) = 'AERO-SOUTERRAIN' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_EL = 'Souterrain' where  upper(replace(Adduction_pb_vers_EL,'é','e')) = 'SOUTERRAIN' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_EL = 'Façade' where  upper(replace(Adduction_pb_vers_EL,'ç','c')) = 'FACADE' " )
	ETD.rechercher("Update infos set Adduction_pb_vers_EL = 'Inconnu' where  upper(replace(Adduction_pb_vers_EL,'é','e')) = 'INCONNU' " )

	ETD.rechercher("Update infos set Type_qualif_ETD = 'Immeuble' where  upper(replace(Type_qualif_ETD,'é','e')) = 'IMMEUBLE' " )
	ETD.rechercher("Update infos set Type_qualif_ETD = 'Façade' where  upper(replace(Type_qualif_ETD,'ç','c')) = 'FACADE' " )
	ETD.rechercher("Update infos set Type_qualif_ETD = 'Inconnu' where  upper(replace(Type_qualif_ETD,'é','e')) = 'INCONNU' " )

	ETD.rechercher("Update infos set Type_qualif_ETD = 'Chambre' where  upper(replace(Type_qualif_ETD,'é','e')) = 'CHAMBRE' " )
	ETD.rechercher("Update infos set Type_qualif_ETD = 'Sur appui FT' where  upper(replace(Type_qualif_ETD,'é','e')) = 'APPUIS FT'" )
	ETD.rechercher("Update infos set Type_qualif_ETD = 'Sur appui ERDF' where  upper(replace(Type_qualif_ETD,'é','e')) = 'APPUIS ERDF' " )
	ETD.rechercher("Update infos set Type_qualif_ETD = 'Sur appui Régie' where  upper(replace(Type_qualif_ETD,'é','e')) = 'APPUIS REGIE' " )

	ETD.fermerBdd()



	ETD = bddSqlite.BaseSqlite()

	qualificateur = exportExcel.ajoutClasseur("Attribution Qualificateur")
	qualificateur.ajoutOnglet_qualificateur("Attribution Qualificateur", "Attribution Qualificateur")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	qualificateur.ecrireVal("B1", p, "Attribution Qualificateur", "Attribution Qualificateur")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	qualificateur.ecrireVal("F1", p, "Attribution Qualificateur", "Attribution Qualificateur")
	ETD.rechercher(Qualificateur)
	v = ETD.resultat()

	qualificateur.ecrireResultat(v, "Attribution Qualificateur", "Attribution Qualificateur")

	qualificateur.fermeture("Attribution Qualificateur")


	negociateur = exportExcel.ajoutClasseur("Attribution Negociateur")
	negociateur.ajoutOnglet_qualificateur("Attribution Negociateur", "Attribution Negociateur")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	negociateur.ecrireVal("B1", p, "Attribution Negociateur", "Attribution Negociateur")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	negociateur.ecrireVal("F1", p, "Attribution Negociateur", "Attribution Negociateur")
	ETD.rechercher(Negociateur)
	v = ETD.resultat()

	negociateur.ecrireResultat(v, "Attribution Negociateur", "Attribution Negociateur")

	negociateur.fermeture("Attribution Negociateur")


	Commentaire = exportExcel.ajoutClasseur("Commentaire ZN")
	Commentaire.ajoutOnglet_CommentaireZN("Commentaire ZN", "Commentaire ZN")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	Commentaire.ecrireVal("B1", p, "Commentaire ZN", "Commentaire ZN")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	Commentaire.ecrireVal("F1", p, "Commentaire ZN", "Commentaire ZN")
	ETD.rechercher(commentaire)
	v = ETD.resultat()

	Commentaire.ecrireResultat(v, "Commentaire ZN", "Commentaire ZN")

	Commentaire.fermeture("Commentaire ZN")



	orange_pressyndic = exportExcel.ajoutClasseur("Contacts Orange Pressyndic")
	orange_pressyndic.ajoutOnglet_contact_orange("Contacts Orange Pressyndic", "Contacts Orange Pressyndic")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	orange_pressyndic.ecrireVal("B1", p, "Contacts Orange Pressyndic", "Contacts Orange Pressyndic")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	orange_pressyndic.ecrireVal("F1", p, "Contacts Orange Pressyndic", "Contacts Orange Pressyndic")
	ETD.rechercher(contact_orange_pressyndic)
	v = ETD.resultat()

	orange_pressyndic.ecrireResultat(v, "Contacts Orange Pressyndic", "Contacts Orange Pressyndic")

	orange_pressyndic.fermeture("Contacts Orange Pressyndic")


	orange_pressyndic = exportExcel.ajoutClasseur("Contacts Orange Gestimm")
	orange_pressyndic.ajoutOnglet_contact_orange("Contacts Orange Gestimm", "Contacts Orange Gestimm")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	orange_pressyndic.ecrireVal("B1", p, "Contacts Orange Gestimm", "Contacts Orange Gestimm")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	orange_pressyndic.ecrireVal("F1", p, "Contacts Orange Gestimm", "Contacts Orange Gestimm")
	ETD.rechercher(contact_orange_gestimm)
	v = ETD.resultat()

	orange_pressyndic.ecrireResultat(v, "Contacts Orange Gestimm", "Contacts Orange Gestimm")

	orange_pressyndic.fermeture("Contacts Orange Gestimm")


	projet = exportExcel.ajoutClasseur("Contacts projet")
	projet.ajoutOnglet_contact_projet("Contacts projet", "Contacts projet")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	projet.ecrireVal("B1", p, "Contacts projet", "Contacts projet")
	ETD.rechercher("select Centre from IXI limit 1")
	p = ETD.resultat()[0][0]
	projet.ecrireVal("F1", p, "Contacts projet", "Contacts projet")
	ETD.rechercher(contact_projet)
	v = ETD.resultat()

	projet.ecrireResultat(v, "Contacts projet", "Contacts projet")

	projet.fermeture("Contacts projet")


	etudes = exportExcel.ajoutClasseur("infos études")
	etudes.ajoutOnglet_infos("infos études", "infos études")
	ETD.rechercher("select plaque from IXI limit 1")
	p = ETD.resultat()[0][0]
	etudes.ecrireVal("B1", p, "infos études", "infos études")

	ETD.rechercher(Etudes)
	v = ETD.resultat()

	etudes.ecrireResultat(v, "infos études", "infos études")

	etudes.fermeture("infos études")


	ETD.fermerBdd()
	
else:
	
	ETD = bddSqlite.BaseSqlite()
	ETD.rechercher("Update OPT set AG_Non_pertinente = '0 AG Pertinente'  where AG_Non_pertinente = 'Oui'")
	
os.system("pause")