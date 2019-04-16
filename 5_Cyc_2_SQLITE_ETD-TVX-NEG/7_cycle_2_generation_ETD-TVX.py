# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcelETD
from mesRequetesETD import *
import datetime
import os


os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#   Génération des fichiers cycle 2 ETD-TVX   #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")

if os.path.exists("Fichiers_Chargement_cycle_2_ETD_TVX"):

	try :
		os.system("rmdir Fichiers_Chargement_cycle_2_ETD_TVX /S /Q")
		os.mkdir("Fichiers_Chargement_cycle_2_ETD_TVX")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:

	try :
		os.mkdir("Fichiers_Chargement_cycle_2_ETD_TVX")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()

centre_etd = input("Entrer le nom du centre d'étude tel qu'il apparait dans IXIEtude : ")

print("Le centre d'étude est :" , centre_etd)

centre_tvx = input("Entrer le nom du centre de travaux tel qu'il apparait dans IXITravaux : ")

print("Le centre de travaux est :" , centre_tvx)


if DPL <= Dep:

# -----------------------------PARTIE ETUDES-------------------------------------------------------------


	ETD = bddSqlite.BaseSqlite()

	ETD.rechercher("Update C set centre_etude = " + "'" + centre_etd + "'")
	ETD.rechercher("Update B1 set centre_etude = " + "'" + centre_etd + "'")
	ETD.rechercher("Update C set centre_travaux = " + "'" + centre_tvx + "'")
	ETD.rechercher("Update B1 set centre_travaux = " + "'" + centre_tvx + "'")
	ETD.rechercher("Update C set Jalon = 'Piquetage non réalisé' where Jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT') and PA in (select PA from C where Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT'))")
	
	# ETD.rechercher(Type_site_Vide_1)
	# ETD.rechercher(Type_site_Vide_2)

	ETD.fermerBdd()

	ETD = bddSqlite.BaseSqlite()
	
	ETD.creerTableInt("C_OK", C_OK)
	ETD.creerTableInt("C_KO", C_KO)

	ETD.creerTableInt("PA_Sans_FIS", PA_Sans_FIS)
	ETD.creerTableInt("PA_avec_FIS", PA_avec_FIS)
	ETD.creerTableInt("PA_SUP_3", PA_SUP_3)

	ETD.creerTableInt("C1", C1)
	ETD.creerTableInt("C2", C2)
	ETD.creerTableInt("C4", C4)
	ETD.creerTableInt("C3", C3)
	ETD.creerTableInt("C3_1", C3_1)

	ETD.creerTableInt("C_1_A", C_1_A)
	ETD.creerTableInt("C_2_A", C_2_A)
	ETD.creerTableInt("C_2_B", C_2_B)
	ETD.creerTableInt("C_3_A", C_3_A)
	ETD.creerTableInt("C_3_B", C_3_B)
	ETD.creerTableInt("C_3_1_A", C_3_1_A)
	ETD.creerTableInt("C_3_1_B", C_3_1_B)
	ETD.creerTableInt("C_4_A", C_4_A)
	ETD.creerTableInt("C_4_B", C_4_B)
	

	ETD = bddSqlite.BaseSqlite()
	ETD.rechercher("Update C_1_A set table_etd = 'RGT_PAV_CTRL'")
	ETD.rechercher("Update C_2_A set table_etd = 'RGT_PAV_CTRL'")
	ETD.rechercher("Update C_3_A set table_etd = 'RGT_PAV_CTRL'")

	ETD.rechercher("Update C_2_B set table_etd = 'RGT_NEG_PAV'")
	ETD.rechercher("Update C_3_1_B set table_etd = 'RGT_NEG_PAV'")

	ETD.rechercher("Update C_3_1_A set table_etd = 'RGT_IMM_CTRL'")
	ETD.rechercher("Update C_4_B set table_etd = 'RGT_IMM_CTRL'")
	
	ETD.rechercher("Update C_4_A set table_etd = 'RGT_IMM_ZN'")
	ETD.rechercher("Update C_3_B set table_etd = 'RGT_IMM_ZN'")

	ETD.rechercher("Update B1 set table_etd = 'RGT_IMM_ZN'")


	ETD.fermerBdd()

# -------------------

	ETD = bddSqlite.BaseSqlite()
	
	ETD.creerTableInt("C_OK_RGT", C_OK_RGT)


	ETD.creerTableInt("CPL_AVC_ETD", completude_avc_etd_tvx)

	ETD = bddSqlite.BaseSqlite()
	

	ETD.rechercher("Update C_OK_RGT set Jalon = 'Piquetage non réalisé' where jalon = 'BPT'")

	ETD.rechercher(QualifND)
	ETD.rechercher(ImportZNND)
	ETD.rechercher(PrevAccordND)
	ETD.rechercher(AccordND)
	ETD.rechercher(DTAND)
	ETD.rechercher(BPEND)
	ETD.rechercher(PINND)
	ETD.rechercher(PIQND)
	ETD.rechercher(ETDND)
	ETD.rechercher(BPTND)
	ETD.rechercher(TLAND)

	ETD.rechercher(QualifNA)
	ETD.rechercher(ImportZNNA)
	ETD.rechercher(Nonqualif)
	
	ETD.fermerBdd()


	ETD = bddSqlite.BaseSqlite()

# 1 ------------------------------------------------Lites des codes IMB en RGT PAV ctrl-------------------------
# RGT PAV Crtl	Contrôle état avancement	1 RGT par avancement

	ETD.rechercher("select distinct PA, count(distinct Jalon), jalon from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_PAV_CTRL'  group by PA having count(distinct Jalon) = 1")
	a = ETD.resultat()

	for i, j, k in a :

		ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + i + "_" + '01' +  "'" + " where PA = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_PAV_CTRL'")
		

	ETD.rechercher("select distinct PA, count(distinct Jalon) from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_PAV_CTRL'  group by PA having count(distinct Jalon) > 1")
	a = ETD.resultat()

	for i, j in a :

		ETD.rechercher("SELECT  distinct Jalon from C_OK_RGT where PA = " + "'" + i + "'" + " and " + " Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE')")
		b = ETD.resultat()
		cpt = 1
		for k in b:

			ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + i + "_" + '0' + str(cpt) +  "'" + " where PA = " "'" + i + "'" + " and Jalon = " "'" + k[0] + "'" + " and table_etd = 'RGT_PAV_CTRL'")
			cpt = cpt + 1



# 2 ------------------------------------------------Lites des codes IMB en RGT NEGO sur PAV-------------------------
# RGT Négo sur PAV	Contrôle état avancement	1 RGT par FIS par avancement

	ETD.rechercher("select distinct CRS, count(distinct Jalon), jalon from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd =  'RGT_NEG_PAV'group by CRS having count(distinct jalon) = 1 order by 1")
	a = ETD.resultat()

	for i, j, k in a :

		ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + i + "_" + '01' +  "'" + " where CRS = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_NEG_PAV'")
		

	ETD.rechercher("select distinct CRS, count(distinct Jalon) from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_NEG_PAV'  group by CRS having count(distinct Jalon) > 1")
	a = ETD.resultat()

	for i, j in a :

		ETD.rechercher("SELECT  distinct Jalon from C_OK_RGT where CRS = " + "'" + i + "'" + " and " + " Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE')")
		b = ETD.resultat()
		cpt = 1
		for k in b:

			ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + i + "_" + '0' + str(cpt) +  "'" + " where CRS = " "'" + i + "'" + " and Jalon = " "'" + k[0] + "'" + " and table_etd = 'RGT_NEG_PAV'")
			cpt = cpt + 1


# 3 ------------------------------------------------Lites des codes IMB en RGT IMM ctrl-------------------------
# RGT IMM Crtl	Contrôle état avancement	1 RGT par avancement

	ETD.rechercher("select distinct PA, count(distinct Jalon), IMM_PAV, count(distinct IMM_PAV), CRS, jalon from C_OK_RGT where Jalon in ('BPT', 'Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_IMM_CTRL' group by PA, CRS having count(distinct jalon) = 1")
	a = ETD.resultat()

	for i, j, u, v, w, x in a :


		if v == 1:

			if u == 'IMM':

				ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-IMM-" + w + "_" + '01' +  "'" + " where CRS = " "'" + w + "'" + " and Jalon = " "'" + x + "'" + " and table_etd = 'RGT_IMM_CTRL'")

			elif u == 'PAV':

				ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + w + "_" + '01' +  "'" + " where CRS = " "'" + w + "'" + " and Jalon = " "'" + x + "'" + " and table_etd = 'RGT_IMM_CTRL'")


		elif v > 1 :

			ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-MIX-" + w + "_" + '01' +  "'" + " where CRS = " "'" + w + "'" + " and Jalon = " "'" + x + "'" + " and table_etd = 'RGT_IMM_CTRL'")


	ETD.rechercher("select distinct CRS, count(distinct Jalon), IMM_PAV, count(distinct IMM_PAV) from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_IMM_CTRL'  group by CRS having count(distinct Jalon) > 1")
	a = ETD.resultat()

	for i, j, v, w in a :

		ETD.rechercher("SELECT  distinct Jalon, count(DISTINCT IMM_PAV), IMM_PAV from C_OK_RGT where CRS = " + "'" + i + "'" + " and " + " Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE')  and table_etd = 'RGT_IMM_CTRL' group by jalon")
		b = ETD.resultat()

		cpt = 1

		for k, l, m in b:

			if l == 1:

				if m == 'IMM':

					ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-IMM-" + i + "_" + '0' + str(cpt) +  "'" + " where CRS = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_IMM_CTRL'")

					cpt = cpt + 1

				elif m == 'PAV':

					ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-PAV-" + i + "_" + '0' + str(cpt) +  "'" + " where CRS = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_IMM_CTRL'")
				
					cpt = cpt + 1


			elif l > 1 :

				ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-MIX-" + i + "_" + '0' + str(cpt) +  "'" + " where CRS = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_IMM_CTRL'")

				cpt = cpt + 1

# 4 ------------------------------------------------Lites des codes IMB en RGT IMM sans ZN-------------------------

# RGT IMM sans ZN	Contrôle état avancement	Pour tout IMM avec BPT OK, 1 RGT par code IMB
# 		Pour les autres, 1 RGT par avancement Négo sans ZN


	ETD.rechercher("select distinct PA, count(distinct Jalon), jalon from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_IMM_ZN'  group by PA having count(distinct Jalon) = 1")
	a = ETD.resultat()

	for i, j, k in a :

		ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-IMM-" + i + "_" + '01' +  "'" + " where PA = " "'" + i + "'" + " and Jalon = " "'" + k + "'" + " and table_etd = 'RGT_IMM_ZN'")
		

	ETD.rechercher("select distinct PA, count(distinct Jalon) from C_OK_RGT where Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE') and table_etd = 'RGT_IMM_ZN'  group by PA having count(distinct Jalon) > 1")
	a = ETD.resultat()

	for i, j in a :

		ETD.rechercher("SELECT  distinct Jalon from C_OK_RGT where PA = " + "'" + i + "'" + " and " + " Jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE')")
		b = ETD.resultat()
		cpt = 1
		for k in b:

			ETD.rechercher("update C_OK_RGT set Regroupement = " + "'" + "RGT-IMM-" + i + "_" + '0' + str(cpt) +  "'" + " where PA = " "'" + i + "'" + " and Jalon = " "'" + k[0] + "'" + " and table_etd = 'RGT_IMM_ZN'")
			cpt = cpt + 1

	ETD.fermerBdd()


# ---------------------------------------

	ETD = bddSqlite.BaseSqlite()

	CA = exportExcelETD.ajoutClasseur("Avec incohérence avancement - liste F")
	CA.ajoutOngletCA("CA", "CA")

	ETD.rechercher(CPL_AVC_ETD_OUT)
	v = ETD.resultat()
	CA.ecrireResultatCPLC(v, "CA", "CA")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	CA.ecrireVal("B1", v, "CA", "CA")
	CA.ecrireVal("D1", "IMB avec avancement NOK", "CA", "CA")
	CA.ecrireVal("F1", centre_etd, "CA", "CA")
	CA.ajoutOngletAVCT("Avancement", "Avancement")
	
	CA.fermeture("CA")


	Sans_PA = exportExcelETD.ajoutClasseur("Avec PA non renseigné  avec avancement - liste E")
	Sans_PA.ajoutOngletSansPA("Prod_Sans_PA", "Prod_Sans_PA")

	ETD.rechercher("select CODE_IMB, Adresse, NB_EL, Jalon, Etat_nego, Code_Syndic, CRS, PA from D")
	v = ETD.resultat()
	Sans_PA.ecrireResultatCPL(v, "Sans_PA", "Sans_PA")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	Sans_PA.ecrireVal("B1", v, "Sans_PA", "Sans_PA")
	Sans_PA.ecrireVal("D1", "Production en cours sans PA", "Sans_PA", "Sans_PA")
	Sans_PA.ecrireVal("F1", centre_etd, "Sans_PA", "Sans_PA")
	Sans_PA.fermeture("Sans_PA")


	non_deb = exportExcelETD.ajoutClasseur("Avec PA renseigné sans avancement - Non débuté - liste A")
	non_deb.ajoutOngletlistes("Non débuté", "Non débuté")

	ETD.rechercher(liste_a)
	v = ETD.resultat()
	non_deb.ecrireResultat(v, "non_deb", "non_deb")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	non_deb.ecrireVal("B1", v, "non_deb", "non_deb")
	non_deb.ecrireVal("D1", "Avec PA renseigné sans avancement - Non débuté", "non_deb", "non_deb")
	non_deb.ecrireVal("F1", centre_etd, "non_deb", "non_deb")
	non_deb.fermeture("non_deb")


	deb = exportExcelETD.ajoutClasseur("Avec PA renseigné  avec avancement - liste B")
	deb.ajoutOngletlistes("PA renseigné avec avancement", "PA renseigné avec avancement")

	ETD.rechercher(liste_b)
	v = ETD.resultat()
	deb.ecrireResultat(v, "deb", "deb")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	deb.ecrireVal("B1", v, "deb", "deb")
	deb.ecrireVal("D1", "Avec PA renseigné  avec avancement ", "deb", "deb")
	deb.ecrireVal("F1", centre_etd, "deb", "deb")
	deb.fermeture("deb")


	non_pa = exportExcelETD.ajoutClasseur("Avec PA non renseigné  sans avancement - liste D")
	non_pa.ajoutOngletlistes("PA non renseigné sans avt", "PA non renseigné sans avt")

	ETD.rechercher(liste_c)
	v = ETD.resultat()
	non_pa.ecrireResultat(v, "non_pa", "non_pa")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	non_pa.ecrireVal("B1", v, "non_pa", "non_pa")
	non_pa.ecrireVal("D1", "PA non renseigné sans avancement", "non_pa", "non_pa")
	non_pa.ecrireVal("F1", centre_etd, "non_pa", "non_pa")
	non_pa.fermeture("non_pa")


	sans_etd = exportExcelETD.ajoutClasseur("Sans avancement étude - liste G")
	sans_etd.ajoutOngletlistes("Sans avancement étude ", "Sans avancement étude ")

	ETD.rechercher(liste_f)
	v = ETD.resultat()
	sans_etd.ecrireResultat(v, "sans_etd", "sans_etd")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	sans_etd.ecrireVal("B1", v, "sans_etd", "sans_etd")
	sans_etd.ecrireVal("D1", "Sans avancement étude ", "sans_etd", "sans_etd")
	sans_etd.ecrireVal("F1", centre_etd, "sans_etd", "sans_etd")
	sans_etd.fermeture("sans_etd")


	fin = exportExcelETD.ajoutClasseur("Avec PA renseigné et Prod terminée - liste C")
	fin.ajoutOngletlistes("Production terminée", "Production terminée")

	ETD.rechercher(prod_terminee)
	v = ETD.resultat()
	fin.ecrireResultat(v, "fin", "fin")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	fin.ecrireVal("B1", v, "fin", "fin")
	fin.ecrireVal("D1", "Production terminée", "fin", "fin")
	fin.ecrireVal("F1", centre_etd, "fin", "fin")
	fin.fermeture("fin")

	eza = exportExcelETD.ajoutClasseur("Etude réalisée - Date Fin EZA à saisir - Liste I1")
	eza.ajoutOngletEZA("Date Fin EZA à saisir", "Date Fin EZA à saisir")

	ETD.rechercher(fin_eza)
	v = ETD.resultat()
	eza.ecrireResultat(v, "eza", "eza")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	eza.ecrireVal("B1", v, "eza", "eza")
	eza.ecrireVal("D1", "Date Fin EZA à saisir", "eza", "eza")
	eza.ecrireVal("F1", centre_etd, "eza", "eza")
	eza.fermeture("eza")

	fci_a = exportExcelETD.ajoutClasseur("Numéro FCI commande d'accès à saisir - Liste I2")
	fci_a.ajoutOngletFCIA("FCI commande d'accès à saisir", "FCI commande d'accès à saisir")

	ETD.rechercher(num_fci_a)
	v = ETD.resultat()
	fci_a.ecrireResultat(v, "fci_a", "fci_a")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	fci_a.ecrireVal("B1", v, "fci_a", "fci_a")
	fci_a.ecrireVal("D1", "Numéro FCI commande d'accès à saisir", "fci_a", "fci_a")
	fci_a.ecrireVal("F1", centre_etd, "fci_a", "fci_a")
	fci_a.fermeture("fci_a")

	fci_f = exportExcelETD.ajoutClasseur("Numéro FCI Fin travaux à saisir - Liste I3")
	fci_f.ajoutOngletFCIF("FCI Fin travaux à saisir", "FCI Fin travaux à saisir")

	ETD.rechercher(num_fci_f)
	v = ETD.resultat()
	fci_f.ecrireResultat(v, "fci_f", "fci_f")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	fci_f.ecrireVal("B1", v, "fci_f", "fci_f")
	fci_f.ecrireVal("D1", "Numéro FCI Fin travaux à saisir", "fci_f", "fci_f")
	fci_f.ecrireVal("F1", centre_etd, "fci_f", "fci_f")
	fci_f.fermeture("fci_f")


	fichierEZA = exportExcelETD.ajoutClasseur("Date EZA")
	fichierEZA.ajoutOnglet_EZA("Date EZA", "Date EZA")

	ETD.rechercher(EZA)
	v = ETD.resultat()
	fichierEZA.ecrireResultat(v, "fichierEZA", "fichierEZA")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	fichierEZA.ecrireVal("B1", v, "fichierEZA", "fichierEZA")
	fichierEZA.ecrireVal("D1", "Date EZA", "fichierEZA", "fichierEZA")
	fichierEZA.ecrireVal("F1", centre_etd, "fichierEZA", "fichierEZA")
	fichierEZA.fermeture("fichierEZA")

	fichierCMD = exportExcelETD.ajoutClasseur("commande FCI")
	fichierCMD.ajoutOnglet_FCI("commande FCI", "commande FCI")

	ETD.rechercher(CMD)
	v = ETD.resultat()
	fichierCMD.ecrireResultat(v, "fichierCMD", "fichierCMD")
	ETD.rechercher(Requete_Plaque)
	v = ETD.resultat()[0][0]
	fichierCMD.ecrireVal("B1", v, "fichierCMD", "fichierCMD")
	fichierCMD.ecrireVal("D1", "commande FCI", "fichierCMD", "fichierCMD")
	fichierCMD.ecrireVal("F1", centre_etd, "fichierCMD", "fichierCMD")
	fichierCMD.fermeture("fichierCMD")

	#-------------------------------------------------------------------------PN------------------------------------------------------------

	ETD = bddSqlite.BaseSqlite()

	PN = exportExcelETD.ajoutClasseur(tables_etd_tvx[1][1])
	PN.ajoutOnglet_Q("PN", "PN")

	ETD.rechercher(select_PN_U)
	v = ETD.resultat()
	PN.ecrireResultat(v, "PN", "PN")

	PN.fermeture("PN")

	#-------------------------------------------------------------------------PR------------------------------------------------------------

	ETD = bddSqlite.BaseSqlite()

	PR = exportExcelETD.ajoutClasseur(tables_etd_tvx[2][1])
	PR.ajoutOnglet_Q("PR", "PR")

	ETD.rechercher(select_PR_U)
	v = ETD.resultat()
	PR.ecrireResultat(v, "PR", "PR")

	PR.fermeture("PR")

	#-------------------------------------------------------------------------Qualif tjs vide------------------------------------------------------------

	ETD = bddSqlite.BaseSqlite()

	PE = exportExcelETD.ajoutClasseur(tables_etd_tvx[3][1])
	PE.ajoutOnglet_Q("PE", "PE")

	ETD.rechercher("select * from C where code_imb = 'code fictif pour retourner un vide' ")
	v = ETD.resultat()
	PE.ecrireResultat(v, "PE", "PE")

	PE.fermeture("PE")

	#-------------------------------------------------------------------------ET---

	ET = exportExcelETD.ajoutClasseur(tables_etd_tvx[4][1])
	ET.ajoutOnglet_E("ET", "ET")

	ETD.rechercher(select_ET_U)
	v = ETD.resultat()
	ET.ecrireResultat(v, "ET", "ET")

	ET.fermeture("ET")

	#-------------------------------------------------------------------------TL---

	TL = exportExcelETD.ajoutClasseur(tables_etd_tvx[5][1])
	TL.ajoutOnglet_TL("TL", "TL")

	ETD.rechercher(select_TL_U)
	v = ETD.resultat()
	TL.ecrireResultat(v, "TL", "TL")

	TL.fermeture("TL")

	#-------------------------------------------------------------------------DO---

	DO = exportExcelETD.ajoutClasseur(tables_etd_tvx[6][1])
	DO.ajoutOnglet_D("DO", "DO")

	ETD.rechercher(select_DO_U)
	v = ETD.resultat()
	DO.ecrireResultat(v, "DO", "DO")

	DO.fermeture("DO")

	#-------------------------------------------------------------------------Suspendu tjs vide------------------------------------------------------------

	ETD = bddSqlite.BaseSqlite()

	SU = exportExcelETD.ajoutClasseur(tables_etd_tvx[7][1])
	SU.ajoutOnglet_S("SU", "SU")

	ETD.rechercher("select * from C where code_imb = 'code fictif pour retourner un vide' ")
	v = ETD.resultat()
	SU.ecrireResultat(v, "SU", "SU")

	SU.fermeture("SU")


	#-------------------------------------------------------------------------TD---

	TD = exportExcelETD.ajoutClasseur(tables_etd_tvx[8][1])
	TD.ajoutOngletTvx_D("TD", "TD")

	ETD.rechercher("select * from C where code_imb = 'code fictif pour retourner un vide' ")
	v = ETD.resultat()
	TD.ecrireResultat(v, "TD", "TD")

	TD.fermeture("TD")

	#-------------------------------------------------------------------------TR---

	TR = exportExcelETD.ajoutClasseur(tables_etd_tvx[9][1])
	TR.ajoutOngletTvx_R("TR", "TR")

	ETD.rechercher(select_TR_U)
	v = ETD.resultat()
	TR.ecrireResultat(v, "TR", "TR")

	TR.fermeture("TR")

	#-------------------------------------------------------------------------TB---

	TB = exportExcelETD.ajoutClasseur(tables_etd_tvx[10][1])
	TB.ajoutOngletTvx_B("TB", "TB")

	ETD.rechercher("select * from C where code_imb = 'code fictif pour retourner un vide' ")
	v = ETD.resultat()
	TB.ecrireResultat(v, "TB", "TB")

	TB.fermeture("TB")


	Suspendre = exportExcelETD.ajoutClasseur("A suspendre")
	Suspendre.ajoutOnglet_Suspendre("Suspendre", "Suspendre")

	ETD.rechercher("select Code_IMB, Adresse, plaque, Centre_etude, Type_suspension from C where Suspendre = 'Oui' ")
	v = ETD.resultat()
	Suspendre.ecrireResultat(v, "Suspendre", "Suspendre")

	Suspendre.fermeture("Suspendre")

	ETD.rechercher("select table_etd from OPT limit 1")
	v = ETD.resultat()[0][0]


	if v == 'ZTD HD' :

		admin_pa = exportExcelETD.ajoutClasseur("Admin PA")
		admin_pa.ajoutOnglet_Admin_PA("Admin PA", "Admin PA")

		ETD.rechercher(Admin_PA)
		v = ETD.resultat()
		admin_pa.ecrireResultat(v, "Admin PA", "Admin PA")

		admin_pa.fermeture("Admin PA")



	SY = exportExcelETD.ajoutClasseur("Rapport ETD-TVX")
	SY.ajoutOngletSyntheseETD("SY", "Tableaux ETD-TVX")

	ETD.rechercher(qte_plaque)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C4", v, "SY", "SY")

	ETD.rechercher(qte_plaque_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D4", v, "SY", "SY")

	ETD.rechercher(nego_declare)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C5", v, "SY", "SY")

	ETD.rechercher(nego_declare_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D5", v, "SY", "SY")

	ETD.rechercher(nego_pds)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C6", v, "SY", "SY")

	ETD.rechercher(nego_pds_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D6", v, "SY", "SY")

	ETD.rechercher(non_debute)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C10", v, "SY", "SY")

	ETD.rechercher(non_debute_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D10", v, "SY", "SY")

	ETD.rechercher(pa_encours)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C11", v, "SY", "SY")

	ETD.rechercher(pa_encours_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D11", v, "SY", "SY")

	ETD.rechercher(pa_fin)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C12", v, "SY", "SY")

	ETD.rechercher(pa_fin_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D12", v, "SY", "SY")

	ETD.rechercher(sans_pa)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C13", v, "SY", "SY")

	ETD.rechercher(sans_pa_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D13", v, "SY", "SY")

	ETD.rechercher(sans_pa_avc)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C14", v, "SY", "SY")

	ETD.rechercher(sans_pa_avc_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D14", v, "SY", "SY")


	ETD.rechercher(valide)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C18", v, "SY", "SY")

	ETD.rechercher(valide_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D18", v, "SY", "SY")


	ETD.rechercher(avc_nok)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C19", v, "SY", "SY")

	ETD.rechercher(avc_nok_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D19", v, "SY", "SY")

	ETD.rechercher(nok_etd_neg_nb)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C20", v, "SY", "SY")

	ETD.rechercher(nok_etd_neg_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D20", v, "SY", "SY")

	ETD.rechercher(sans_etude)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C21", v, "SY", "SY")

	ETD.rechercher(sans_etude_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D21", v, "SY", "SY")


	ETD.rechercher(pn)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C24", v, "SY", "SY")

	ETD.rechercher(pn_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D24", v, "SY", "SY")

	ETD.rechercher(pr)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C25", v, "SY", "SY")

	ETD.rechercher(pr_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D25", v, "SY", "SY")

	SY.ecrireValSyn("C26", 0, "SY", "SY")
	SY.ecrireValSyn("D26", 0, "SY", "SY")

	ETD.rechercher(et)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C27", v, "SY", "SY")

	ETD.rechercher(et_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D27", v, "SY", "SY")

	ETD.rechercher(tl)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C28", v, "SY", "SY")

	ETD.rechercher(tl_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D28", v, "SY", "SY")

	SY.ecrireValSyn("C29", 0, "SY", "SY")
	SY.ecrireValSyn("D29", 0, "SY", "SY")

	ETD.rechercher(do)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C30", v, "SY", "SY")

	ETD.rechercher(do_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D30", v, "SY", "SY")

	ETD.rechercher(do)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C31", v, "SY", "SY")

	ETD.rechercher(do_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D31", v, "SY", "SY")


	ETD.rechercher(fin_eza_nb)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C34", v, "SY", "SY")

	ETD.rechercher(fin_eza_nb_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D34", v, "SY", "SY")

	ETD.rechercher(num_fci_a_nb)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C35", v, "SY", "SY")

	ETD.rechercher(num_fci_a_nb_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D35", v, "SY", "SY")

	ETD.rechercher(num_fci_f_nb)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("C36", v, "SY", "SY")

	ETD.rechercher(num_fci_f_nb_el)
	v = ETD.resultat()[0][0]
	SY.ecrireValSyn("D36", v, "SY", "SY")


	SY.fermeture("SY")



	ETD.fermerBdd()

os.system("pause")