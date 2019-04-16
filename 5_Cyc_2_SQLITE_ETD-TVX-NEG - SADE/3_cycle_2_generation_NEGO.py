# -*- coding: utf-8 -*-
import bddSqlite
from req import *
import exportExcelNEG
from mesRequetesNEG import *
import datetime
import os

os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#     Génération des fichiers cycle 2 NEG     #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")


if os.path.exists("Fichiers_Chargement_cycle_2_NEGO"):
	try :
		os.system("rmdir Fichiers_Chargement_cycle_2_NEGO /S /Q")
		os.mkdir("Fichiers_Chargement_cycle_2_NEGO")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()
else:
	
	try :
		os.mkdir("Fichiers_Chargement_cycle_2_NEGO")

	except FileExistsError:
		print("------------------------------------------------------------------------------------")
		print("")
		print("\x1b[5;41;80m Un fichier est ouvert, veuillez fermer les fichiers qui vont être générés de nouveau \x1b[0m")	
		print("")
		print("------------------------------------------------------------------------------------")
		exit()

centre_nego = input("Entrer le nom du centre de négo tel qu'il apparait dans IXINEGO : ")

print("Le centre de NEGO est :" , centre_nego)



if DPL <= Dep:

	NEG = bddSqlite.BaseSqlite()


	NEG.update('Escalier', '')
	NEG.update('CRS', '')
	NEG.update('Code_Syndic', '')
	NEG.update('Civ_PDT_SYN', '')
	NEG.update('Mail_PDT_SYN', '')
	NEG.update('Fax_PDT_SYN', '')
	NEG.update('Nom_PDT_SYN', '')
	NEG.update('Prenom_PDT_SYN', '')
	NEG.update('Tel_PDT_SYN', '')
	NEG.update('Mob_PDT_SYN', '')
	NEG.update('Nom_contact_syndic', '')
	NEG.update('Prenom_Contact_Syndic', '')
	NEG.update('Tel_contact_syndic', '')
	NEG.update('Mob_contact_syndic', '')
	NEG.update('Resultat_etude_site', '')
	NEG.update('Nom_contact_syndic', '')
	NEG.update('AG_non_pertinente', '')
	NEG.update('Attente_Probation', '')

	# NEG.rechercher(Update_01)
	# NEG.rechercher(Update_02)
	# NEG.rechercher(Update_03)
	# NEG.rechercher(Update_04)
	# NEG.rechercher(Update_05)

	NEG.rechercher(updateIMM)

	NEG.rechercher(Update_Mob_contact_syndic)
	NEG.rechercher(Update_Tel_contact_syndic)
	NEG.rechercher(Update_Tel_PDT_SYN)
	NEG.rechercher(Update_Fax_PDT_SYN)
	NEG.rechercher(Update_Mob_PDT_SYN)

	NEG.rechercher("Update OPT set Nego_non_deb = 'Non', Qualifie = 'Non',  ImportZN = 'Non', PrevAccord = 'Non', Accord = 'Non', DTA = 'Non', BPE = 'Non', BPT = 'Non'")


	NEG.rechercher(updateNondebute)
	NEG.rechercher(updateQualifie)
	NEG.rechercher(updateImportZN)
	NEG.rechercher(updatePrevAccord)
	NEG.rechercher(updateAccord)
	NEG.rechercher(updateDTA)
	NEG.rechercher(updateBPE)
	NEG.rechercher(updatePiquetage)
	NEG.rechercher(updateBPT)
	NEG.rechercher(updateEtudes)
	NEG.rechercher(updateTravauxLances)
	NEG.rechercher(updateTravauxRealises)
	NEG.rechercher(updateDOE)


	
	NEG.fermerBdd()

	NEG = bddSqlite.BaseSqlite()
	
	NEG.rechercher(JalonSansProd)
	NEG.rechercher(JalonQualif)
	NEG.rechercher(JalonImportZN)
	NEG.rechercher(JalonPrevAccord)
	NEG.rechercher(JalonAccord)
	# NEG.rechercher(JalonDTA)
	NEG.rechercher(JalonBPE)
	NEG.rechercher(JalonPiquetage)
	NEG.rechercher(JalonBPT)
	NEG.rechercher(JalonEtude)
	NEG.rechercher(JalonTL)
	NEG.rechercher(JalonDOE)

	NEG.fermerBdd()

	NEG = bddSqlite.BaseSqlite()

	NEG.creerTableInt("Prod_Fin", Prod_Fin)
	NEG.creerTableInt("Sans_Prod", Sans_Prod)
	NEG.creerTableInt("B1", B1)
	NEG.creerTableInt("B2", B2)
	NEG.creerTableInt("B3", B3)
	NEG.creerTableInt("C", C)
	NEG.creerTableInt("D", D)
	NEG.creerTableInt("CD", CD)


	NEG = bddSqlite.BaseSqlite()

	NEG.creerTableInt("C1_0_KO", C1_0_KO)
	NEG.creerTableInt("C1_0_OK", C1_0_OK)
	NEG.creerTableInt("C1_1_OK", C1_1_OK)
	NEG.creerTableInt("C1_1_KO", C1_1_KO)
	NEG.creerTableInt("C2_1_OK", C2_1_OK)
	NEG.creerTableInt("C2_1_KO", C2_1_KO)
	NEG.creerTableInt("C1_C2_OK", C1_C2_OK)
	NEG.creerTableInt("CPL_AVC", completude_avc)

	NEG = bddSqlite.BaseSqlite()

	# NEG.rechercher("Update CPL_AVC set Jalon = 'BPE' where Jalon in ('Piquetage', 'Etudes')")
	# NEG.rechercher("Update CPL_AVC set Jalon = 'BPT' where Jalon in ('Travaux lancés', 'DOE')")

	NEG.rechercher(QualifND)
	NEG.rechercher(ImportZNND)
	NEG.rechercher(PrevAccordND)
	NEG.rechercher(AccordND)
	# NEG.rechercher(DTAND)
	NEG.rechercher(BPEND)
	# NEG.rechercher(PIQND)
	# NEG.rechercher(ETDND)
	# NEG.rechercher(BPTND)
	# NEG.rechercher(TLAND)

	NEG.rechercher(QualifNA)
	NEG.rechercher(ImportZNNA)
	
	NEG.fermerBdd()

	NEG_Q = bddSqlite.BaseSqlite()


	NonQualif = exportExcelNEG.ajoutClasseur(tables_nego[1][1])
	NonQualif.ajoutOngletQ("NonQualif", "NonQualif")

	NEG_Q.rechercher(fichierNonQualif)
	v = NEG_Q.resultat()
	NonQualif.ecrireResultat(v, "NonQualif", "NonQualif")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	NonQualif.ecrireVal("B1", v, "NonQualif", "NonQualif")
	NonQualif.ecrireVal("D1", "Non Qualif", "NonQualif", "NonQualif")
	NonQualif.ecrireVal("F1", centre_nego, "NonQualif", "NonQualif")
	NonQualif.fermeture("NonQualif")


	Qualif = exportExcelNEG.ajoutClasseur(tables_nego[2][1])
	Qualif.ajoutOngletQ("Qualif", "Qualif")

	NEG_Q.rechercher(fichierQualif)
	v = NEG_Q.resultat()
	Qualif.ecrireResultat(v, "Qualif", "Qualif")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	Qualif.ecrireVal("B1", v, "Qualif", "Qualif")
	Qualif.ecrireVal("D1", "Qualif", "Qualif", "Qualif")
	Qualif.ecrireVal("F1", centre_nego, "Qualif", "Qualif")
	Qualif.fermeture("Qualif")


	ImportZN = exportExcelNEG.ajoutClasseur(tables_nego[3][1])
	ImportZN.ajoutOngletI("ImportZN", "ImportZN")

	NEG_Q.rechercher(fichierImportZN)
	v = NEG_Q.resultat()
	ImportZN.ecrireResultat(v, "ImportZN", "ImportZN")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	ImportZN.ecrireVal("B1", v, "ImportZN", "ImportZN")
	ImportZN.ecrireVal("D1", "ImportZN", "ImportZN", "ImportZN")
	ImportZN.ecrireVal("F1", centre_nego, "ImportZN", "ImportZN")
	ImportZN.fermeture("ImportZN")


	PrevAccord = exportExcelNEG.ajoutClasseur(tables_nego[4][1])
	PrevAccord.ajoutOngletPA("PrevAccord", "PrevAccord")

	NEG_Q.rechercher(fichierPrevAccord)
	v = NEG_Q.resultat()
	PrevAccord.ecrireResultat(v, "PrevAccord", "PrevAccord")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	PrevAccord.ecrireVal("B1", v, "PrevAccord", "PrevAccord")
	PrevAccord.ecrireVal("D1", "PrevAccord", "PrevAccord", "PrevAccord")
	PrevAccord.ecrireVal("F1", centre_nego, "PrevAccord", "PrevAccord")
	PrevAccord.fermeture("PrevAccord")


	Accord = exportExcelNEG.ajoutClasseur(tables_nego[5][1])
	Accord.ajoutOngletPA("Accord", "Accord")

	NEG_Q.rechercher(fichierAccord)
	v = NEG_Q.resultat()
	Accord.ecrireResultat(v, "Accord", "Accord")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	Accord.ecrireVal("B1", v, "Accord", "Accord")
	Accord.ecrireVal("D1", "Accord", "Accord", "Accord")
	Accord.ecrireVal("F1", centre_nego, "Accord", "Accord")
	Accord.fermeture("Accord")


	BPE = exportExcelNEG.ajoutClasseur(tables_nego[6][1])
	BPE.ajoutOngletE("BPE", "BPE")

	NEG_Q.rechercher(fichierBPE)
	v = NEG_Q.resultat()
	BPE.ecrireResultat(v, "BPE", "BPE")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	BPE.ecrireVal("B1", v, "BPE", "BPE")
	BPE.ecrireVal("D1", "BPE", "BPE", "BPE")
	BPE.ecrireVal("F1", centre_nego, "BPE", "BPE")
	BPE.fermeture("BPE")


	BPT = exportExcelNEG.ajoutClasseur(tables_nego[7][1])
	BPT.ajoutOngletT("BPT", "BPT")

	NEG_Q.rechercher(fichierBPT)
	v = NEG_Q.resultat()
	BPT.ecrireResultat(v, "BPT", "BPT")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	BPT.ecrireVal("B1", v, "BPT", "BPT")
	BPT.ecrireVal("D1", "BPT", "BPT", "BPT")
	BPT.ecrireVal("F1", centre_nego, "BPT", "BPT")
	BPT.fermeture("BPT")

	DTA = exportExcelNEG.ajoutClasseur(tables_nego[8][1])
	DTA.ajoutOngletD("DTA", "DTA")

	NEG_Q.rechercher(fichierDTA)
	v = NEG_Q.resultat()
	DTA.ecrireResultat(v, "DTA", "DTA")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	DTA.ecrireVal("B1", v, "DTA", "DTA")
	DTA.ecrireVal("D1", "DTA", "DTA", "DTA")
	DTA.ecrireVal("F1", centre_nego, "DTA", "DTA")
	DTA.fermeture("DTA")


	InfosCompl = exportExcelNEG.ajoutClasseur(tables_nego[9][1])
	InfosCompl.ajoutOngletInfosCompl("Infos compl", "Infos compl")

	NEG_Q.rechercher(fichierInfosCpl)
	v = NEG_Q.resultat()
	InfosCompl.ecrireResultat(v, "Infos compl", "Infos compl")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	InfosCompl.ecrireVal("B1", v, "Infos compl", "Infos compl")
	InfosCompl.ecrireVal("D1", "Infos compl", "Infos compl", "Infos compl")
	InfosCompl.ecrireVal("F1", centre_nego, "Infos compl", "Infos compl")
	InfosCompl.fermeture("Infos compl")
	

	CI = exportExcelNEG.ajoutClasseur("Complétude Infos FIS non conforme - liste B")
	CI.ajoutOngletCI("CI", "CI")

	NEG_Q.rechercher(completude_infos)
	v = NEG_Q.resultat()
	CI.ecrireResultatCPL(v, "CI", "CI")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	CI.ecrireVal("B1", v, "CI", "CI")
	CI.ecrireVal("D1", "FIS avec infos différentes", "CI", "CI")
	CI.ecrireVal("F1", centre_nego, "CI", "CI")
	CI.fermeture("CI")


	CA = exportExcelNEG.ajoutClasseur("Avec incohérence avancement - liste A")
	CA.ajoutOngletCA("CA", "CA")
	NEG_Q.rechercher(CPL_AVC_OUT)
	v = NEG_Q.resultat()
	CA.ecrireResultatCPLC(v, "CA", "CA")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	CA.ecrireVal("B1", v, "CA", "CA")
	CA.ecrireVal("D1", "IMB avec avancement NOK", "CA", "CA")
	CA.ecrireVal("F1", centre_nego, "CA", "CA")
	CA.ajoutOngletAVCT("Avancement", "Avancement")
	
	CA.fermeture("CA")

	





	SF = exportExcelNEG.ajoutClasseur("Codes IMB Pavillon sans FIS - liste C")
	SF.ajoutOngletPavSansFIS("SF", "PAV sans FIS")

	NEG_Q.rechercher(pav_sans_fis_fic)
	v = NEG_Q.resultat()
	SF.ecrireResultat(v, "SF", "SF")
	NEG_Q.rechercher(Requete_Plaque)
	v = NEG_Q.resultat()[0][0]
	SF.ecrireVal("B1", v, "SF", "SF")
	SF.ecrireVal("D1", "IMB avec avancement NOK", "SF", "SF")
	SF.ecrireVal("F1", centre_nego, "SF", "SF")
	SF.fermeture("SF")





	SY = exportExcelNEG.ajoutClasseur("Rapport Nego")
	SY.ajoutOngletSynthese("SY", "Tableaux NEGO")

	NEG_Q.rechercher(qte_plaque)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C4", v, "SY", "SY")

	NEG_Q.rechercher(qte_plaque_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D4", v, "SY", "SY")

	NEG_Q.rechercher(nego_declare)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C5", v, "SY", "SY")

	NEG_Q.rechercher(nego_declare_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D5", v, "SY", "SY")

	NEG_Q.rechercher(nego_pds)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C6", v, "SY", "SY")

	NEG_Q.rechercher(nego_pds_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D6", v, "SY", "SY")

	NEG_Q.rechercher(valide)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C10", v, "SY", "SY")

	NEG_Q.rechercher(valide_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D10", v, "SY", "SY")	

	NEG_Q.rechercher(stock_non_deb)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C11", v, "SY", "SY")

	NEG_Q.rechercher(stock_non_deb_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D11", v, "SY", "SY")	


	NEG_Q.rechercher(avc_nok)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C12", v, "SY", "SY")

	NEG_Q.rechercher(avc_nok_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D12", v, "SY", "SY")	

	NEG_Q.rechercher(infos_nok)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C13", v, "SY", "SY")

	NEG_Q.rechercher(infos_nok_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D13", v, "SY", "SY")

	NEG_Q.rechercher(infos_nok_pds)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C14", v, "SY", "SY")

	NEG_Q.rechercher(infos_nok_pds_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D14", v, "SY", "SY")

	NEG_Q.rechercher(pav_sans_fis)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C15", v, "SY", "SY")

	NEG_Q.rechercher(pav_sans_fis_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D15", v, "SY", "SY")






	NEG_Q.rechercher(nq)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C18", v, "SY", "SY")

	NEG_Q.rechercher(nq_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D18", v, "SY", "SY")

	NEG_Q.rechercher(q)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C19", v, "SY", "SY")

	NEG_Q.rechercher(q_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D19", v, "SY", "SY")	

	NEG_Q.rechercher(im)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C20", v, "SY", "SY")

	NEG_Q.rechercher(im_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D20", v, "SY", "SY")	

	NEG_Q.rechercher(pa)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C21", v, "SY", "SY")

	NEG_Q.rechercher(pa_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D21", v, "SY", "SY")	

	NEG_Q.rechercher(ac)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C22", v, "SY", "SY")

	NEG_Q.rechercher(ac_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D22", v, "SY", "SY")	

	NEG_Q.rechercher(dta)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C23", v, "SY", "SY")

	NEG_Q.rechercher(dta_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D23", v, "SY", "SY")	

	NEG_Q.rechercher(bpe)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C24", v, "SY", "SY")

	NEG_Q.rechercher(bpe_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D24", v, "SY", "SY")	

	NEG_Q.rechercher(bpt)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("C25", v, "SY", "SY")

	NEG_Q.rechercher(bpt_el)
	v = NEG_Q.resultat()[0][0]
	SY.ecrireValSyn("D25", v, "SY", "SY")	

	SY.fermeture("SY")

	NEG_Q.fermerBdd()



os.system("pause")