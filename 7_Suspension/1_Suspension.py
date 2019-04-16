# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcel
from mesRequetes import *
import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Insertion données DPL            #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")



if os.path.exists("Fichiers_Suspension"):

	try :
		os.system("rmdir Fichiers_Suspension /S /Q")
		os.mkdir("Fichiers_Suspension")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:

	try :
		os.mkdir("Fichiers_Suspension")

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


	print("Insertion des données dans les table")
		
	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("NEG", OPT_requestn)
	ETD.insererDonnees("NEG", 0, tupn)
	ETD.fermerBdd()
			
	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("ETD", OPT_requeste)
	ETD.insererDonnees("ETD", 2, tupe)
	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("ETD2", OPT_requeste2)
	ETD.insererDonnees("ETD2", 3, tupe2)
	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()	
	ETD.creerTable("SPI", OPT_requests)
	ETD.insererDonnees("SPI", 4, tups)
	ETD.fermerBdd()

print("Les fichiers ont été chargés")
print("----------------------------")

ETD = bddSqlite.BaseSqlite()
ETD.rechercher(Err_neg)
ETD.rechercher(Err_etd)
ETD.rechercher(tbl)
ETD.rechercher(qn0)
ETD.rechercher(qn1)
ETD.rechercher(prob0)
ETD.rechercher(prob1)
ETD.rechercher(bpt0)
ETD.rechercher(bpt1)
ETD.rechercher(etd0)
ETD.rechercher(etd1)
ETD.rechercher(tl0)
ETD.rechercher(tl1)
ETD.rechercher(doe0)
ETD.rechercher(doe1)
ETD.rechercher(tfx0)
ETD.rechercher(tfx1)
ETD.rechercher(etd_ze)
ETD.rechercher(spi_ze)



ETD.fermerBdd()

ETD = bddSqlite.BaseSqlite()

fichierNEG = exportExcel.ajoutClasseur("Suspension ZN")
fichierNEG.ajoutOnglet_SuspendreN("Suspension ZN", "Suspension ZN")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
fichierNEG.ecrireVal("B1", p, "fichierNEG", "fichierNEG")
ETD.rechercher(suspension_nego)
v = ETD.resultat()

fichierNEG.ecrireResultat(v, "fichierNEG", "fichierNEG")

fichierNEG.fermeture("fichierNEG")


fichierCom = exportExcel.ajoutClasseur("Commentaire Suspension ZN")
fichierCom.ajoutOnglet_Commentaire("Commentaire ZN", "Commentaire ZN")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
fichierCom.ecrireVal("B1", p, "fichierCom", "fichierCom")
ETD.rechercher(commentaire_nego)
v = ETD.resultat()
fichierCom.ecrireResultat(v, "fichierCom", "fichierCom")

fichierCom.fermeture("fichierCom")


fichierETD = exportExcel.ajoutClasseur("Suspension RGT")
fichierETD.ajoutOnglet_SuspendreE("Suspension RGT", "Suspension RGT")

ETD.rechercher(suspension_etude)
v = ETD.resultat()
fichierETD.ecrireResultat(v, "fichierETD", "fichierETD")

fichierETD.fermeture("fichierETD")

# -------------------------------------------------------------------------------------------------------------------------------------------
a1_neg = exportExcel.ajoutClasseur("Liste_A1_Nego")
a1_neg.ajoutOnglet_action_neg("Liste_A1_Nego", "Liste A1 NEGO")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a1_neg.ecrireVal("B1", p, "Liste_A1_Nego", "Liste_A1_Nego")
a1_neg.ecrireVal("D1", 'Liste A1 NEGO', "Liste_A1_Nego", "Liste_A1_Nego")
ETD.rechercher(liste_a1_neg)
v = ETD.resultat()
a1_neg.ecrireResultat(v, "Liste_A1_Nego", "Liste_A1_Nego")

a1_neg.fermeture("Liste_A1_Nego")

# -------------------------------------------------------------------------------------------------------------------------------------------
a1_etd = exportExcel.ajoutClasseur("Liste_A1_Etudes")
a1_etd.ajoutOnglet_action_etd("Liste_A1_Etudes", "Liste A1 Etudes")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a1_etd.ecrireVal("B1", p, "Liste_A1_Etudes", "Liste_A1_Etudes")
a1_etd.ecrireVal("D1", 'Liste A1 Etudes', "Liste_A1_Etudes", "Liste_A1_Etudes")
ETD.rechercher(liste_a1_etd)
v = ETD.resultat()
a1_etd.ecrireResultat(v, "Liste_A1_Etudes", "Liste_A1_Etudes")

a1_etd.fermeture("Liste_A1_Etudes")

# -------------------------------------------------------------------------------------------------------------------------------------------
a2 = exportExcel.ajoutClasseur("Liste_A2")
a2.ajoutOnglet_action("Liste_A2", "Liste A2")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a2.ecrireVal("B1", p, "Liste_A2", "Liste_A2")
a2.ecrireVal("D1", 'Liste A2', "Liste_A2", "Liste_A2")
ETD.rechercher(liste_a2)
v = ETD.resultat()
a2.ecrireResultat(v, "Liste_A2", "Liste_A2")

a2.fermeture("Liste_A2")

# -------------------------------------------------------------------------------------------------------------------------------------------
a3 = exportExcel.ajoutClasseur("Liste_A3")
a3.ajoutOnglet_action("Liste_A3", "Liste A3")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a3.ecrireVal("B1", p, "Liste_A3", "Liste_A3")
a3.ecrireVal("D1", 'Liste A3', "Liste_A3", "Liste_A3")
ETD.rechercher(liste_a3)
v = ETD.resultat()
a3.ecrireResultat(v, "Liste_A3", "Liste_A3")

a3.fermeture("Liste_A3")

# -------------------------------------------------------------------------------------------------------------------------------------------
a4 = exportExcel.ajoutClasseur("Liste_A4")
a4.ajoutOnglet_action("Liste_A4", "Liste A4")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a4.ecrireVal("B1", p, "Liste_A4", "Liste_A4")
a2.ecrireVal("D1", 'Liste A4', "Liste_A4", "Liste_A4")
ETD.rechercher(liste_a4)
v = ETD.resultat()
a4.ecrireResultat(v, "Liste_A4", "Liste_A4")

a4.fermeture("Liste_A4")


# -------------------------------------------------------------------------------------------------------------------------------------------
a5 = exportExcel.ajoutClasseur("Liste_A5")
a5.ajoutOnglet_action("Liste_A5", "Liste A5")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a5.ecrireVal("B1", p, "Liste_A5", "Liste_A5")
a5.ecrireVal("D1", 'Liste A5', "Liste_A5", "Liste_A5")
ETD.rechercher(liste_a5)
v = ETD.resultat()
a5.ecrireResultat(v, "Liste_A5", "Liste_A5")

a5.fermeture("Liste_A5")


# -------------------------------------------------------------------------------------------------------------------------------------------
a6 = exportExcel.ajoutClasseur("Liste_A6")
a6.ajoutOnglet_action("Liste_A6", "Liste A6")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a6.ecrireVal("B1", p, "Liste_A6", "Liste_A6")
a6.ecrireVal("D1", 'Liste A6', "Liste_A6", "Liste_A6")
ETD.rechercher(liste_a6)
v = ETD.resultat()
a6.ecrireResultat(v, "Liste_A6", "Liste_A6")

a6.fermeture("Liste_A6")


# -------------------------------------------------------------------------------------------------------------------------------------------
a7 = exportExcel.ajoutClasseur("Liste_A7")
a7.ajoutOnglet_action("Liste_A7", "Liste A7")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a7.ecrireVal("B1", p, "Liste_A7", "Liste_A7")
a7.ecrireVal("D1", 'Liste A7', "Liste_A7", "Liste_A7")
ETD.rechercher(liste_a7)
v = ETD.resultat()
a7.ecrireResultat(v, "Liste_A7", "Liste_A7")

a7.fermeture("Liste_A7")


# -------------------------------------------------------------------------------------------------------------------------------------------
a8 = exportExcel.ajoutClasseur("Liste_A8")
a8.ajoutOnglet_action("Liste_A8", "Liste A8")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a8.ecrireVal("B1", p, "Liste_A8", "Liste_A8")
a8.ecrireVal("D1", 'Liste A8', "Liste_A8", "Liste_A8")
ETD.rechercher(liste_a8)
v = ETD.resultat()
a8.ecrireResultat(v, "Liste_A8", "Liste_A8")

a8.fermeture("Liste_A8")


# -------------------------------------------------------------------------------------------------------------------------------------------
a9 = exportExcel.ajoutClasseur("Liste_A9")
a9.ajoutOnglet_action("Liste_A9", "Liste A9")
ETD.rechercher("select plaque from SPI limit 1")
p = ETD.resultat()[0][0]
a9.ecrireVal("B1", p, "Liste_A9", "Liste_A9")
a9.ecrireVal("D1", 'Liste A9', "Liste_A9", "Liste_A9")
ETD.rechercher(liste_a9)
v = ETD.resultat()
a9.ecrireResultat(v, "Liste_A9", "Liste_A9")

a9.fermeture("Liste_A9")


# -------------------------------------------------------------------------------------------------------------------------------------------
cem_n = exportExcel.ajoutClasseur("Type processus ZN")
cem_n.ajoutOnglet_CEM_OPT_ZN("Type processus ZN", "Type processus ZN")
ETD.rechercher(cem_opt_zn)
v = ETD.resultat()
cem_n.ecrireResultat(v, "Type processus ZN", "Type processus ZN")

cem_n.fermeture("Type processus ZN")

# -------------------------------------------------------------------------------------------------------------------------------------------
cem_e = exportExcel.ajoutClasseur("Type processus ZE")
cem_e.ajoutOnglet_CEM_OPT_ZE("Type processus ZE", "Type processus ZE")
ETD.rechercher(cem_opt_ze)
v = ETD.resultat()
cem_e.ecrireResultat(v, "Type processus ZE", "Type processus ZE")

cem_e.fermeture("Type processus ZE")

# -------------------------------------------------------------------------------------------------------------------------------------------
cem_t = exportExcel.ajoutClasseur("Type processus RGT")
cem_t.ajoutOnglet_CEM_OPT_RGT("Type processus RGT", "Type processus RGT")
ETD.rechercher(cem_opt_rgt)
v = ETD.resultat()
cem_t.ecrireResultat(v, "Type processus RGT", "Type processus RGT")

cem_t.fermeture("Type processus RGT")

ETD.fermerBdd()


os.system("pause")