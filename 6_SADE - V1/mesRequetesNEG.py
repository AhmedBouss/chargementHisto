# -*- coding: utf-8 -*-
from datetime import datetime
from urllib.request import urlopen


if 1 :
	
	Dep = 43539
	debut = datetime(1900, 1, 1, 0, 0, 0)

	res = urlopen('http://www.google.com')
	headers = res.info()
	date = headers['date'].split(' ')
	date = date[1:4]
	date = " ".join(date)

	newdate = datetime.date(datetime.strptime(str(date), '%d %b %Y'))

	datefin = datetime(newdate.year, newdate.month, newdate.day, 0, 0, 0)

	DPL = (datefin - debut).days + 2


table_OPT = {
				1 : ['Code_IMB', 'Code_IMB',  'CHARINT', 'liste[1]', '?'],  
				2 : ['Adresse', 'Adresse',  'CHARINT', 'liste[1]', '?'],  
				3 : ['Localite', 'Localite',  'CHARINT', 'liste[1]', '?'],  
				4 : ['CP', 'CP',  'CHARINT', 'liste[1]', '?'],  
				5 : ['NB_EL', 'NB_EL',  'CHARINT', 'liste[1]', '?'],  
				6 : ['NB_EL_R', 'NB_EL_R',  'CHARINT', 'liste[1]', '?'],  
				7 : ['NB_EL_P', 'NB_EL_P',  'CHARINT', 'liste[1]', '?'],  
				8 : ['PM', 'PM',  'CHARINT', 'liste[1]', '?'],  
				9 : ['PA', 'PA',  'CHARINT', 'liste[1]', '?'],  
				10 : ['Dernier_jalon', 'Dernier_jalon',  'CHARINT', 'liste[1]', '?'],  
				11 : ['Jalon_dta_ok', 'Jalon_dta_ok',  'CHARINT', 'liste[1]', '?'],  
				12 : ['Regroupemen_ZN', 'Regroupemen_ZN',  'CHARINT', 'liste[1]', '?'],  
				13 : ['ID_SYNDIC', 'ID_SYNDIC',  'CHARINT', 'liste[1]', '?'],  
				14 : ['Nom_Syndic', 'Nom_Syndic',  'CHARINT', 'liste[1]', '?'],  
				15 : ['Type_syndic', 'Type_syndic',  'CHARINT', 'liste[1]', '?'],  
				16 : ['Code_regroupement_syndic', 'Code_regroupement_syndic',  'CHARINT', 'liste[1]', '?'],  
				17 : ['Date_prev_accord', 'Date_prev_accord',  'CHARINT', 'liste[1]', '?'],  
				18 : ['AG_pertinente', 'AG_pertinente',  'CHARINT', 'liste[1]', '?'],  
				19 : ['Date_AG', 'Date_AG',  'CHARINT', 'liste[1]', '?'],  
				20 : ['Date_Envoi_Accord_Type', 'Date_Envoi_Accord_Type',  'CHARINT', 'liste[1]', '?'],  
				21 : ['Date_d_accord_syndic', 'Date_d_accord_syndic',  'CHARINT', 'liste[1]', '?'],  
				22 : ['Attente_Probation', 'Attente_Probation',  'CHARINT', 'liste[1]', '?'],  
				23 : ['Date_retour_convention', 'Date_retour_convention',  'CHARINT', 'liste[1]', '?'],  
				24 : ['Presence_DTA', 'Presence_DTA',  'CHARINT', 'liste[1]', '?'],  
				25 : ['Date_envoi_PE_au_syndic', 'Date_envoi_PE_au_syndic',  'CHARINT', 'liste[1]', '?'],  
				26 : ['Date_de_validation_PE', 'Date_de_validation_PE',  'CHARINT', 'liste[1]', '?'],  
				27 : ['Dernier_jalon_ETD', 'Dernier_jalon_ETD',  'CHARINT', 'liste[1]', '?'],  
				28 : ['Regroupement_ZE', 'Regroupement_ZE',  'CHARINT', 'liste[1]', '?'],  
				29 : ['Regroupement_RGT', 'Regroupement_RGT',  'CHARINT', 'liste[1]', '?'],  
				30 : ['NB_escaliers', 'NB_escaliers',  'CHARINT', 'liste[1]', '?'],  
				31 : ['NB_etages', 'NB_etages',  'CHARINT', 'liste[1]', '?'],  
				32 : ['Type_de_site', 'Type_de_site',  'CHARINT', 'liste[1]', '?'],  
				33 : ['Date_fin_EZA', 'Date_fin_EZA',  'CHARINT', 'liste[1]', '?'],  
				34 : ['Resultat_etude_d2', 'Resultat_etude_d2',  'CHARINT', 'liste[1]', '?'],  
				35 : ['Numero_fci_acces', 'Numero_fci_acces',  'CHARINT', 'liste[1]', '?'],  
				36 : ['Date_pose_PB', 'Date_pose_PB',  'CHARINT', 'liste[1]', '?'],  
				37 : ['Numero_FCI_fin', 'Numero_FCI_fin',  'CHARINT', 'liste[1]', '?'],  
				38 : ['Qualificateur', 'Qualificateur',  'CHARINT', 'liste[1]', '?'],  
				39 : ['Nom_contact', 'Nom_contact',  'CHARINT', 'liste[1]', '?'],  
				40 : ['Type_contact', 'Type_contact',  'CHARINT', 'liste[1]', '?'],  
				41 : ['Adresse_contact', 'Adresse_contact',  'CHARINT', 'liste[1]', '?'],  
				42 : ['Tel_contact', 'Tel_contact',  'CHARINT', 'liste[1]', '?'],  
				43 : ['Mail_contact', 'Mail_contact',  'CHARINT', 'liste[1]', '?'],  
				44 : ['Commentaire_Qualif', 'Commentaire_Qualif',  'CHARINT', 'liste[1]', '?'],  
				45 : ['Civilite_President', 'Civilite_President',  'CHARINT', 'liste[1]', '?'],  
				46 : ['Email_PdtConseilSyndical', 'Email_PdtConseilSyndical',  'CHARINT', 'liste[1]', '?'],  
				47 : ['N_fax_pdtconseilsyndical', 'N_fax_pdtconseilsyndical',  'CHARINT', 'liste[1]', '?'],  
				48 : ['Nom_PdtConseilSyndical', 'Nom_PdtConseilSyndical',  'CHARINT', 'liste[1]', '?'],  
				49 : ['Prenom_PdtConseilSyndical', 'Prenom_PdtConseilSyndical',  'CHARINT', 'liste[1]', '?'],  
				50 : ['N_telephone_pdtconseilsyndical', 'N_telephone_pdtconseilsyndical',  'CHARINT', 'liste[1]', '?'],  
				51 : ['N_mobile_pdtconseilsyndical', 'N_mobile_pdtconseilsyndical',  'CHARINT', 'liste[1]', '?'],  
				52 : ['Civilite_Responsable', 'Civilite_Responsable',  'CHARINT', 'liste[1]', '?'],  
				53 : ['Email_Syndic', 'Email_Syndic',  'CHARINT', 'liste[1]', '?'],  
				54 : ['N_Fax_syndic', 'N_Fax_syndic',  'CHARINT', 'liste[1]', '?'],  
				55 : ['Nom_Responsable', 'Nom_Responsable',  'CHARINT', 'liste[1]', '?'],  
				56 : ['Prenom_Responsable', 'Prenom_Responsable',  'CHARINT', 'liste[1]', '?'],  
				57 : ['N_Tel_syndic', 'N_Tel_syndic',  'CHARINT', 'liste[1]', '?'],  
				58 : ['N_Mobile_Syndic', 'N_Mobile_Syndic',  'CHARINT', 'liste[1]', '?'],  
				59 : ['Commentaire_ZN', 'Commentaire_ZN',  'CHARINT', 'liste[1]', '?'],  
				60 : ['Negociateur', 'Negociateur',  'CHARINT', 'liste[1]', '?'],  
				61 : ['PB_Geof_Geofibre', 'PB_Geof_Geofibre',  'CHARINT', 'liste[1]', '?'],  
				62 : ['Adduction_pb_vers_PA', 'Adduction_pb_vers_PA',  'CHARINT', 'liste[1]', '?'],  
				63 : ['Type_qualif_ETD', 'Type_qualif_ETD',  'CHARINT', 'liste[1]', '?'],  
				64 : ['Numero_support_site', 'Numero_support_site',  'CHARINT', 'liste[1]', '?'],  
				65 : ['Adduction_pb_vers_EL', 'Adduction_pb_vers_EL',  'CHARINT', 'liste[1]', '?'],  
				66 : ['PT_IPON', 'PT_IPON',  'CHARINT', 'liste[1]', '?'],  
				67 : ['Commentaire_RGT', 'Commentaire_RGT',  'CHARINT', 'liste[1]', '?'],  
				68 : ['Conducteur_de_travaux', 'Conducteur_de_travaux',  'CHARINT', 'liste[1]', '?'],  
				69 : ['Prod_int_ext', 'Prod_int_ext',  'CHARINT', 'liste[1]', '?'],  
				70 : ['Entreprise', 'Entreprise',  'CHARINT', 'liste[1]', '?'],
				71 : ['Etat_NEGO', 'Etat_NEGO',  'CHARINT', 'liste[1]', '?'],
				72 : ['Date_reception_BPE', 'Date_reception_BPE',  'CHARINT', 'liste[1]', '?'],
				73 : ['Date_prog_racco', 'Date_prog_racco',  'CHARINT', 'liste[1]', '?']

				}



DPL_OPT = "select Code_IMB, Adresse, CP, Localite, '' plaque, ID_SYNDIC, NB_EL, '' PB, PA, PM, Date_envoi_PE_au_syndic, NB_escaliers, NB_EL_R, NB_EL_P, Type_de_site, Resultat_etude_d2, Date_pose_PB, '' Date_MaD_PE, Etat_NEGO, Date_prog_racco, Nom_Syndic, Code_regroupement_syndic, Nom_contact, '' Type_de_regroupement, '' Etat_reseau_TD, '' Date_realisation_TD, AG_pertinente, '' Date_AG, Date_Envoi_Accord_Type, '' Date_retour_accord, Date_d_accord_syndic, Date_retour_convention, Attente_Probation, Date_de_validation_PE, Presence_DTA, Civilite_President, Nom_PdtConseilSyndical, Prenom_PdtConseilSyndical, N_telephone_pdtconseilsyndical, N_mobile_pdtconseilsyndical, N_fax_pdtconseilsyndical, '' mail_president, Prenom_Responsable, N_Tel_syndic, N_Mobile_Syndic, Date_reception_BPE, '' Date_souhaitee_PE, Date_fin_EZA from OPT where code_imb <> ''"

# --------------------------------Itération pour création de la tables OPT à partir du dictionnaire----------------------------------------------
request = []
tup = []
ma_liste = []

for element in table_OPT:

	champ = table_OPT[element][0]
	typechamp = table_OPT[element][2]
	ins = table_OPT[element][3]
	sep = table_OPT[element][4]

	item = champ + " " + typechamp

	request.append(item)
	ma_liste.append(ins)
	tup.append(sep)

OPT_request = ", ".join(request)
tup = ", ".join(tup)
ma_liste = ", ".join(ma_liste)

