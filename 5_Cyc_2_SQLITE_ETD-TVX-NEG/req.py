# -*- coding: utf-8 -*-
from datetime import datetime



tables_nego = {
         1 : ['NQU', '1_Format Fichier Cycle2-Non Qualif', 'Non_Qualif', 1],
               2 : ['QUA', '2_Format Fichier Cycle2-Qualif', 'Qualif', 2],
               3 : ['IMP', '3_Format Fichier Cycle2-Import ZN', 'ImportZN', 3],
               4 : ['PAC', '4_Format Fichier Cycle2-Prev Accord', 'PrevAccord', 4],
               5 : ['ACC', '5_Format Fichier Cycle2-Accord', 'Accord', 5],
               6 : ['BPE', '6_Format Fichier Cycle2-BPE', 'BPE', 6],
               7 : ['BPT', '7_Format Fichier Cycle2-BPT', 'BPT', 7],
               8 : ['DTA', '8_Format Fichier Cycle2-DTA', 'DTA', 8],
               9 : ['infosCompl', '9_Format Fichier Cycle2-infos Compl', 'infosCompl', 9]
               }



defaultValue = {
        1 : ['Type_syndic', '12-Autre Syndic'],
        2 : ['Type_contact', 'Gestimm'], 
        3 : ['Statut_affectation', '1-Affecté'], 
        4 : ['Flag_immeuble', '0-Immeuble ancien'], 
        5 : ['Site_copro', 'Non']
        }

table_OPT = {
        1 : ['Code_IMB', 'Dossier - Code IMB', 'CHARINT', 'liste[1]', '?'],
        2 : ['Adresse', 'Adresse code IMB', 'CHARINT', 'liste[2]', '?'],
        3 : ['CP', 'Code postal - OPT', 'CHARINT', 'liste[3]', '?'],
        4 : ['Localite', 'Localité', 'CHARINT', 'liste[4]', '?'],
        5 : ['Plaque', 'Identifiant Processus - OPT', 'CHARINT', 'liste[5]', '?'],
        6 : ['Code_Syndic', 'Code Syndic - OPT', 'CHARINT', 'liste[6]', '?'],
        7 : ['NB_EL', 'Nb logements - OPT', 'CHARINT', 'liste[7]', '?'],
        8 : ['PB', 'Identifiant PB - OPT', 'CHARINT', 'liste[8]', '?'],
        9 : ['PA', 'Identifiant PA', 'CHARINT', 'liste[9]', '?'],
        10 : ['PM', 'Identifiant PM - OPT', 'CHARINT', 'liste[10]', '?'],
        11 : ['Date_env_pe_syndic', 'Date envoi pré-étude au syndic - OPT', 'CHARINT', 'liste[11]', '?'],
        12 : ['Escalier', 'Escalier - OPT', 'CHARINT', 'liste[12]', '?'],
        13 : ['NB_EL_P', 'Nb logements P - OPT', 'CHARINT', 'liste[13]', '?'],
        14 : ['NB_EL_R', 'Nb logements R - OPT', 'CHARINT', 'liste[14]', '?'],
        15 : ['Type_Site', 'Type site - OPT', 'CHARINT', 'liste[15]', '?'],
        16 : ['Resultat_etude_site', 'Résultat de l’étude site - OPT', 'CHARINT', 'liste[16]', '?'],
        17 : ['Date_pose_PB', 'Date pose du PB - OPT', 'CHARINT', 'liste[17]', '?'],
        18 : ['Date_mad_PE', 'Date mise à disposition pré-étude - OPT', 'CHARINT', 'liste[18]', '?'],
        19 : ['Etat_Nego', 'Etat Négociation - OPT', 'CHARINT', 'liste[19]', '?'],
        20 : ['Date_prog_racco', 'Date programmée raccordement - OPT', 'CHARINT', 'liste[20]', '?'],
        21 : ['Nom_Syndic', 'Nom Syndic - OPT', 'CHARINT', 'liste[21]', '?'],
        22 : ['CRS', 'Code regroup. Syndic - OPT', 'CHARINT', 'liste[22]', '?'],
        23 : ['Nom_contact_syndic', 'Nom contact Syndic - OPT', 'CHARINT', 'liste[42]', '?'],
        24 : ['Type_regroupement', 'Type de regroupement', 'CHARINT', 'liste[23]', '?'],
        25 : ['Etat_reseau_TD', 'Etat réseau T/D - OPT', 'CHARINT', 'liste[24]', '?'],
        26 : ['Date_reseau_TD', 'Date réalisation réseau T/D - OPT', 'CHARINT', 'liste[25]', '?'],
        27 : ['AG_non_pertinente', 'AG non pertinente - OPT', 'CHARINT', 'liste[26]', '?'],
        28 : ['Date_AG', 'Date AG - OPT', 'CHARINT', 'liste[27]', '?'],
        29 : ['Date_envoi_AS', 'Date envoi accord syndic - OPT', 'CHARINT', 'liste[28]', '?'],
        30 : ['Date_retour_AS', 'Date retour accord syndic - OPT', 'CHARINT', 'liste[29]', '?'],
        31 : ['Date_saisie_AS', 'Date de saisie de l’accord syndic - OPT', 'CHARINT', 'liste[30]', '?'],
        32 : ['Date_retour_convention', 'Date Retour Convention - OPT', 'CHARINT', 'liste[31]', '?'],
        33 : ['Attente_probation', 'Attente Probation - OPT', 'CHARINT', 'liste[32]', '?'],
        34 : ['Date_valid_PE', 'Date validation pré-étude par syndic - OPT', 'CHARINT', 'liste[33]', '?'],
        35 : ['Presence_DTA', 'Présence DTA - OPT', 'CHARINT', 'liste[34]', '?'],
        36 : ['Civ_PDT_SYN', 'Civilité Pdt conseil syndical - OPT', 'CHARINT', 'liste[35]', '?'],
        37 : ['Nom_PDT_SYN', 'Nom Pdt conseil syndical - OPT', 'CHARINT', 'liste[36]', '?'],
        38 : ['Prenom_PDT_SYN', 'Prénom Pdt conseil syndical - OPT', 'CHARINT', 'liste[37]', '?'],
        39 : ['Tel_PDT_SYN', 'N° Tel. Pdt conseil syndical - OPT', 'CHAR', 'liste[38]', '?'],
        40 : ['Mob_PDT_SYN', 'N° Mobile Pdt conseil syndical - OPT', 'CHAR', 'liste[39]', '?'],
        41 : ['Fax_PDT_SYN', 'Fax Pdt conseil syndical - OPT', 'CHAR', 'liste[40]', '?'],
        42 : ['Mail_PDT_SYN', 'Email Pdt conseil syndical - OPT', 'CHAR', 'liste[41]', '?'],
        43 : ['Prenom_contact_syndic', 'Prénom Contact Syndic - OPT', 'CHARINT', 'liste[43]', '?'],
        44 : ['Tel_contact_syndic', 'N° Tel. Contact Syndic - OPT', 'CHAR', 'liste[44]', '?'],
        45 : ['Mob_contact_syndic', 'N° Mobile Contact Syndic - OPT', 'CHAR', 'liste[45]', '?'],
        46 : ['Date_reception_bon_PE', 'Date récept. bon pré-étude - OPT', 'CHARINT', 'liste[46]', '?'],
        47 : ['Date_souhaitee_PE', 'Date souhaitée pré-étude - OPT', 'CHARINT', 'liste[47]', '?'],
        48 : ['Date_fin_EZA', 'Date fin etd ZA PA-OPT', 'CHARINT', 'liste[48]', '?'],
        49 : ['Statut_aff_Syndic', 'Statut affectation syndic', 'CHARINT', 'liste[49]', '?'],
        50 : ['Type_DTA', 'Type DTA', 'CHARINT', 'liste[50]', '?'],
        51 : ['Nego_non_deb', 'Nego_non_deb', 'CHARINT', 'liste[51]', '?'],
        52 : ['Qualifie', 'Qualifie', 'CHARINT', 'liste[52]', '?'],
        53 : ['ImportZN', 'ImportZN', 'CHARINT', 'liste[53]', '?'],
        54 : ['PrevAccord', 'PrevAccord', 'CHARINT', 'liste[54]', '?'],
        55 : ['Accord', 'Accord', 'CHARINT', 'liste[55]', '?'],
        56 : ['DTA', 'DTA', 'CHARINT', 'liste[56]', '?'],
        57 : ['BPE', 'BPE', 'CHARINT', 'liste[57]', '?'],
        58 : ['Piquetage', 'Piquetage', 'CHARINT', 'liste[58]', '?'],
        59 : ['BPT', 'BPT', 'CHARINT', 'liste[59]', '?'],
        60 : ['Etudes', 'Etudes', 'CHARINT', 'liste[60]', '?'],
        61 : ['Travaux_lances', 'Travaux_lances', 'CHARINT', 'liste[61]', '?'],
        62 : ['Travaux_realises', 'Travaux_realises', 'CHARINT', 'liste[62]', '?'],
        63 : ['DOE', 'DOE', 'CHARINT', 'liste[63]', '?'],
        64 : ['Jalon', 'Jalon', 'CHARINT', 'liste[64]', '?'],
        65 : ['IMM_PAV', 'IMM_PAV', 'CHARINT', 'liste[64]', '?'],
        66 : ['Regroupement', 'Regroupement', 'CHAR', 'liste[66]', '?'],
        67 : ['centre_etude', 'centre_etude', 'CHAR', 'liste[67]', '?'],
        68 : ['centre_travaux', 'centre_travaux', 'CHAR', 'liste[68]', '?'],
        69 : ['fci_acces', 'Numéro FCI d’accès', 'CHAR', 'liste[69]', '?'],
        70 : ['fci_fin', 'Numéro FCI fin de travaux', 'CHAR', 'liste[70]', '?'],
        71 : ['date_fci', 'date FCI', 'INT', 'liste[71]', '?'],
        72 : ['table_etd', 'table etude', 'CHAR', 'liste[72]', '?'],
        73 : ['Suspendre', 'Suspendre', 'CHAR', 'liste[73]', '?'],
        74 : ['Type_suspension', 'Type suspension', 'CHAR', 'liste[74]', '?'],
        75 : ['SPI', 'Migré SPI', 'CHAR', 'liste[75]', '?']
        }

updateIMM = "update OPT set IMM_PAV = 'IMM' where NB_EL > 3"

updateNondebute = "update OPT set Nego_non_deb = 'Oui' where Etat_Nego <> 'PAS DE SYNDIC' and Code_Syndic = ''"

updateQualifie = "update OPT set Qualifie = 'Oui' where Code_Syndic <> '' and CRS = ''"

updateImportZN = "update OPT set ImportZN = 'Oui' where Code_Syndic <> '' and CRS <> ''"

updatePrevAccord = "update OPT set PrevAccord = 'Oui' where Date_envoi_AS <> '' and (AG_non_pertinente = '1-AG non pertinente' OR AG_non_pertinente = '0-AG pertinente' )"

updateAccord = "update OPT set Accord = 'Oui' where Date_retour_convention <> '' OR Date_saisie_AS <> ''" 

updateDTA = "update OPT set DTA = 'Oui' where Presence_DTA in ('1-DTA présent', '2-DTA inutile')"

updateBPE = "update OPT set BPE = 'Oui' where Date_reception_bon_PE <> ''" 

updatePiquetage = "update OPT set Piquetage = 'Oui' where (Etat_Nego <> 'PAS DE SYNDIC' and Date_mad_PE <> '') OR (Etat_Nego = 'PAS DE SYNDIC' and Resultat_etude_site <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '')"

updateBPT = "update OPT set BPT = 'Oui' where Date_valid_PE <> '' and Date_saisie_AS <> '' and Date_env_pe_syndic <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> ''  and Presence_DTA in ('1-DTA présent', '2-DTA inutile')"

updateEtudes = "update OPT set Etudes = 'Oui' where (Etat_Nego <> 'PAS DE SYNDIC' and Resultat_etude_site <> ''  and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '') OR (Etat_Nego = 'PAS DE SYNDIC' and Resultat_etude_site <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '')" 

updateTravauxLances = "update OPT set Travaux_lances = 'Oui' where Date_prog_racco <> ''"

updateTravauxRealises = "update OPT set Travaux_realises = 'Oui' where Date_pose_PB <> '' "

updateDOE = "update OPT set DOE = 'Oui' where Date_pose_PB <> '' "


Prod_Fin = "select PA, count(PA) nb_prod_fin from OPT where BPT  = 'Oui' and DOE = 'Oui' group by PA order by 2"

Sans_Prod = "select PA, count(PA) nb_non_prod from OPT \
      where PA <> '' and Qualifie = 'Non' and ImportZN = 'Non' and PrevAccord = 'Non' and Accord = 'Non' and DTA = 'Non' \
      and BPE = 'Non' and Piquetage = 'Non' and Etudes = 'Non' and BPT = 'Non' and Travaux_lances = 'Non' and DOE = 'Non' \
      group by PA order by 2"

All = "select * from OPT"

# Table B1 = Codes IMB d'un même PA terminés = Prod terminée
B1 = "select * from OPT  where PA in (select OPT.PA from OPT, prod_fin where OPT.PA = prod_fin.PA  and OPT.PA <> ''group by OPT.PA having count(OPT.PA) = prod_fin.nb_prod_fin)"

# Table B2 = Codes IMB avec PA renseigné et aucun avancement déclaré = en stock
B2 = "select * from OPT  where PA in (select OPT.PA from OPT, Sans_Prod where OPT.PA = Sans_Prod.PA group by OPT.PA having count(OPT.PA) = Sans_Prod.nb_non_prod)" 

# Table B3 = Codes IMB sans PA renseigné et aucun avancement déclaré = en stock
B3 = "select * from OPT where PA = '' and Qualifie = 'Non' and ImportZN = 'Non' and PrevAccord = 'Non' and Accord = 'Non' and DTA = 'Non' and BPE = 'Non' and Piquetage = 'Non' and Etudes = 'Non' and BPT = 'Non' and Travaux_lances = 'Non' and DOE = 'Non'"

# Table C = Codes IMB avec PA renseigné et avancement déclaré = en cours
C = "select * from OPT where PA <> '' and code_IMB not in (select code_IMB from B1) and code_IMB not in (select code_IMB from B2)"

# Table D = Codes IMB sans PA renseigné mais avec avancement déclaré
D = "select * from OPT where PA = '' and code_IMB not in (select code_IMB from B3)"

# Table CD = Codes IMB des tables C et D car pour la négo pas besoin du PA
CD = "Select * from C UNION select * from D"

# Table C1_0_KO = Combinaison d'information NEGO différente pour un même FIS = Fichier de complétude infos
C1_0_KO = "select * from CD where CRS in (select CRS from CD where CRS <> '' group by CRS having count(distinct(Code_Syndic || Etat_Nego || CRS || Date_envoi_AS || AG_non_pertinente || Date_retour_convention || Date_saisie_AS || Date_reception_bon_PE)) > 1)"

# Table C1_0_OK = Combinaison d'information NEGO OK pour un même FIS 
C1_0_OK = "select * from CD where CRS in (select CRS from CD where CRS <> '' group by CRS having count(distinct(Code_Syndic || Etat_Nego || CRS || Date_envoi_AS || AG_non_pertinente || Date_retour_convention || Date_saisie_AS || Date_reception_bon_PE)) = 1)"


# Table C1_1_OK = Vérification de la cohérence de l'avancement NEGO OK pour FIS existant
C1_1_OK = "select * from C1_0_OK where (CRS  <> '' and Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in (\
'NonNonNonNonNonNon',\
'NonOuiNonNonNonNon',\
'NonOuiOuiNonNonNon',\
'NonOuiOuiOuiNonNon',\
'NonOuiOuiOuiNonNon',\
'NonOuiOuiOuiOuiNon',\
'NonOuiOuiOuiOuiOui'))"

# Table C1_1_OK = Vérification de la cohérence de l'avancement NEGO NOK pour FIS existant
C1_1_KO = "select * from C1_0_OK where (CRS  <> '' and Qualifie || ImportZN || PrevAccord || Accord || BPE  || BPT not in (\
'NonNonNonNonNonNon', \
'NonOuiNonNonNonNon', \
'NonOuiOuiNonNonNon', \
'NonOuiOuiOuiNonNon', \
'NonOuiOuiOuiNonNon', \
'NonOuiOuiOuiOuiNon',\
'NonOuiOuiOuiOuiOui'))"

# Table C2_1_OK = Vérification de la cohérence de l'avancement NEGO OK pour FIS non existant
C2_1_OK = "select * from CD where (CRS  = '' and NB_EL > 3 and Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in (\
'NonNonNonNonNonNon',\
'OuiNonNonNonNonNon',\
'OuiNonOuiNonNonNon',\
'OuiNonOuiOuiNonNon',\
'OuiNonOuiOuiNonNon',\
'OuiNonOuiOuiOuiNon',\
'OuiNonOuiOuiOuiOui'))"

# Table C2_1_OK = Vérification de la cohérence de l'avancement NEGO NOK pour FIS non existant
C2_1_KO = "select * from CD where (CRS  = '' and NB_EL > 3 and Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT not in (\
'NonNonNonNonNonNon',\
'OuiNonNonNonNonNon', \
'OuiNonOuiNonNonNon',\
'OuiNonOuiOuiNonNon',\
'OuiNonOuiOuiNonNon',\
'OuiNonOuiOuiOuiNon',\
'OuiNonOuiOuiOuiOui'))"


# Table C1_C2_OK concaténation des tables OK des avancement OK
C1_C2_OK =  "Select * from C1_1_OK UNION select * from C2_1_OK"


# Complétudes infos issues de la table C1_0_KO - infos différentes pour un même FIS
completude_infos = "select Code_IMB, Adresse, NB_EL, CRS, Code_Syndic, Etat_Nego, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, \
Date_reception_bon_PE from C1_0_KO order by CRS"

# completude_avc = "select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, DTA, BPE,  BPT, Etat_Nego, Code_Syndic, CRS, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, Presence_DTA, Date_mad_PE, Date_fin_EZA, Resultat_etude_site, Date_env_pe_syndic, Type_Site, NB_EL_P, NB_EL_R, Date_valid_PE, Date_theorique_racco, Date_pose_PB from  C1_1_KO where Etat_Nego <>'PAS DE SYNDIC' \
# Union select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, DTA, BPE, BPT, Etat_Nego, Code_Syndic, CRS, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, Presence_DTA, Date_mad_PE, Date_fin_EZA, Resultat_etude_site, Date_env_pe_syndic, Type_Site, NB_EL_P, NB_EL_R, Date_valid_PE, Date_theorique_racco, Date_pose_PB from  C2_1_KO where Etat_Nego <>'PAS DE SYNDIC' order by CRS"

# Complétudes avancement issues des tables C1_1_KO - C2_1_KO
completude_avc = "select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, BPE,  BPT, Etat_Nego, \
Code_Syndic, CRS, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, Date_reception_bon_PE, Presence_DTA, Date_env_pe_syndic, Date_valid_PE, NB_EL_P, NB_EL_R, Type_Site, '' Suspendre, '' Type_Suspension \
from  C1_1_KO where Etat_Nego <> 'PAS DE SYNDIC' \
Union \
select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, BPE, BPT, Etat_Nego, Code_Syndic, CRS, \
Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, Date_reception_bon_PE, Presence_DTA, Date_env_pe_syndic, Date_valid_PE, NB_EL_P, NB_EL_R, Type_Site, '' Suspendre, '' Type_Suspension \
from  C2_1_KO where Etat_Nego <>'PAS DE SYNDIC'  order by CRS"


fichierNonQualif = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, '' Type_Syndic, Nom_contact_syndic, '' Type_contact, '' Adresse_Contact, Tel_contact_syndic, '' mail_contact_syndic, Etat_reseau_TD, Date_reseau_TD, '' Qualificateur  from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

fichierQualif = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, '' Type_Syndic, Nom_contact_syndic, '' Type_contact, '' Adresse_Contact, Tel_contact_syndic, '' mail_contact_syndic, Etat_reseau_TD, Date_reseau_TD, '' Qualificateur  from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'OuiNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

fichierImportZN = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonOuiNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

fichierPrevAccord = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiNonNonNon', 'OuiNonOuiNonNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"

fichierAccord = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiNonNon', 'OuiNonOuiOuiNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"

fichierDTA = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Presence_DTA, Type_DTA from C1_C2_OK where (CRS in (select distinct CRS from C1_C2_OK group by CRS having count(Distinct DTA) = 1) and DTA = 'Oui') \
OR ( DTA = 'Oui' and CRS = '')"

fichierBPE = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf, Attente_probation, Date_retour_convention, Presence_DTA from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiNon', 'OuiNonOuiOuiOuiNon') and Etat_Nego <> 'PAS DE SYNDIC'"

fichierBPT = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf, Attente_probation, Date_retour_convention, Presence_DTA, Date_retour_AS, Date_valid_PE, Civ_PDT_SYN, Mail_PDT_SYN, Fax_PDT_SYN, Nom_PDT_SYN, Prenom_PDT_SYN, Tel_PDT_SYN, Mob_PDT_SYN, '' Civilite_resp, '' Email_syndic, '' Fax_syndic, Nom_contact_syndic, Prenom_Contact_Syndic, Tel_contact_syndic, Mob_contact_syndic, Date_env_pe_syndic from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiOui', 'OuiNonOuiOuiOuiOui') and Etat_Nego <> 'PAS DE SYNDIC' \
union \
select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf, Attente_probation, Date_retour_convention, Presence_DTA, Date_retour_AS, Date_valid_PE, Civ_PDT_SYN, Mail_PDT_SYN, Fax_PDT_SYN, Nom_PDT_SYN, Prenom_PDT_SYN, Tel_PDT_SYN, Mob_PDT_SYN, '' Civilite_resp, '' Email_syndic, '' Fax_syndic, Nom_contact_syndic, Prenom_Contact_Syndic, Tel_contact_syndic, Mob_contact_syndic, Date_env_pe_syndic  \
From B1 where ETAT_NEGO <> 'PAS DE SYNDIC'" 

fichierInfosCpl = "select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf, Attente_probation, Date_retour_convention, Presence_DTA, '' Type_DTA, Date_retour_AS, Date_valid_PE, Civ_PDT_SYN, Mail_PDT_SYN, Fax_PDT_SYN, Nom_PDT_SYN, Prenom_PDT_SYN, Tel_PDT_SYN, Mob_PDT_SYN, '' Civilite_resp, '' Email_syndic, '' Fax_syndic, Nom_contact_syndic, Prenom_Contact_Syndic, Tel_contact_syndic, Mob_contact_syndic, Date_env_pe_syndic  from C1_C2_OK where Etat_Nego <> 'PAS DE SYNDIC' \
UNION  \
select Code_IMB, Adresse, Localite, CP, Code_Syndic, Nom_Syndic, NB_EL, CRS, Type_regroupement, '' Date_Realisation, Statut_aff_Syndic, AG_non_pertinente, Date_AG, Date_envoi_AS, '' site_en_copro, '' flag_imm_neuf, Attente_probation, Date_retour_convention, Presence_DTA, '' Type_DTA, Date_retour_AS, Date_valid_PE, Civ_PDT_SYN, Mail_PDT_SYN, Fax_PDT_SYN, Nom_PDT_SYN, Prenom_PDT_SYN, Tel_PDT_SYN, Mob_PDT_SYN, '' Civilite_resp, '' Email_syndic, '' Fax_syndic, Nom_contact_syndic, Prenom_Contact_Syndic, Tel_contact_syndic, Mob_contact_syndic, Date_env_pe_syndic  from B1 where Etat_Nego <> 'PAS DE SYNDIC'"

Requete_Plaque ="select plaque from C1_C2_OK limit 1"


Update_01 = "Update OPT set Mob_contact_syndic = '0'|| Mob_contact_syndic where Mob_contact_syndic not in ('', 'Optionnelle')"

Update_02 = "Update OPT set Tel_contact_syndic = '0'|| Tel_contact_syndic where Tel_contact_syndic not in ('', 'Optionnelle')"

Update_03 = "Update OPT set Tel_PDT_SYN = '0'|| Tel_PDT_SYN where Tel_PDT_SYN not in ('', 'Optionnelle')"

Update_04 = "Update OPT set Fax_PDT_SYN = '0'|| Fax_PDT_SYN where Fax_PDT_SYN not in ('', 'Optionnelle')"

Update_05 = "Update OPT set Mob_PDT_SYN = '0'|| Mob_PDT_SYN where Mob_PDT_SYN not in ('', 'Optionnelle')"


Update_Mob_contact_syndic = "update OPT set Mob_contact_syndic = replace(Mob_contact_syndic, '.0','')"

Update_Tel_contact_syndic = "update OPT set Tel_contact_syndic = replace(Tel_contact_syndic, '.0','')"

Update_Tel_PDT_SYN = "update OPT set Tel_PDT_SYN = replace(Tel_PDT_SYN, '.0','')"

Update_Fax_PDT_SYN = "update OPT set  Fax_PDT_SYN = replace( Fax_PDT_SYN, '.0','')"

Update_Mob_PDT_SYN = "update OPT set Mob_PDT_SYN = replace(Mob_PDT_SYN, '.0','')"


Update_Presence_DTA_1 = "Update OPT set Presence_DTA = '1-DTA présent' where Presence_DTA = 'Présent'"

Update_Presence_DTA_2 =  "Update OPT set Presence_DTA = '0-DTA non présent et nécessaire' where Presence_DTA = 'Pas présent'"

Update_Presence_DTA_3 = "Update OPT set Presence_DTA = '2-DTA inutile' where Presence_DTA = 'Pas nécessaire'"

Update_AG_non_pertinente_1 = "Update OPT set AG_non_pertinente = '1-AG non pertinente' where AG_non_pertinente = 'Oui'"

Update_AG_non_pertinente_2 = "Update OPT set AG_non_pertinente = '0-AG pertinente' where AG_non_pertinente = 'Non'"

Update_statut_affectation = "Update OPT set Statut_aff_Syndic = '1-Affecté' where CODE_SYNDIC <> ''"

Update_Type_DTA_1 = "Update OPT set Type_DTA = 'DTA' where Presence_DTA = '1-DTA présent'"
Update_Type_DTA_2 = "Update OPT set Type_DTA = 'Attestation' where Presence_DTA = '2-DTA inutile'"



JalonSansProd = "Update OPT set Jalon = 'Sans_Prod' where Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'NonNonNonNonNonNonNonNonNonNon'"

JalonQualif = "Update OPT set Jalon = 'Qualif' where Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNonNonNon'"

JalonImportZN = "Update OPT set Jalon = 'ImportZN' where ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNonNon'"

JalonPrevAccord = "Update OPT set Jalon = 'PrevAccord' where PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNon' "

JalonAccord = "Update OPT set Jalon = 'Accord' where Accord || BPE || Piquetage || BPT || Etudes ||  Travaux_Lances || DOE = 'OuiNonNonNonNonNonNon'"

# JalonDTA = "Update OPT set Jalon = 'DTA' where DTA || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNon'"

JalonBPE = "Update OPT set Jalon = 'BPE' where BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNon'"

JalonPiquetage = "Update OPT set Jalon = 'Piquetage' where  Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNon'"

JalonBPT = "Update OPT set Jalon = 'BPT' where  BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNon'"

JalonEtude = "Update OPT set Jalon = 'Etude' where Etudes || Travaux_Lances || DOE = 'OuiNonNon'"

JalonTL = "Update OPT set Jalon = 'Travaux lancés' where Travaux_Lances || DOE ='OuiNon' "

JalonDOE = "Update OPT set Jalon = 'DOE' where  DOE = 'Oui'"



# QualifND = "Update CPL_AVC set ImportZN = 'N.A.', PrevAccord = 'N.D.',  Accord = 'N.D.',  DTA = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Qualif' "
# ImportZNND = "Update CPL_AVC set  PrevAccord = 'N.D.',  Accord = 'N.D.',  DTA = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'ImportZN' "
# PrevAccordND = "Update CPL_AVC set  Accord = 'N.D.',  DTA = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'PrevAccord' "
# AccordND = "Update CPL_AVC set  DTA = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Accord' "
# DTAND = "Update CPL_AVC set BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'DTA' "
# BPEND = "Update CPL_AVC set Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'BPE' "
# PIQND = "Update CPL_AVC set  Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Piquetage' "
# ETDND = "Update CPL_AVC set  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Etudes' "
# BPTND = "Update CPL_AVC set  Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'BPT' "
# TLAND = "Update CPL_AVC set  DOE = 'N.D.' where jalon = 'Travaux lancés' "
# QualifNA = "Update CPL_AVC set Qualifie = 'N.A.' where CRS <> ''"
# ImportZNNA = "Update CPL_AVC set ImportZN = 'N.A.' where CRS = ''"

QualifND = "Update CPL_AVC set ImportZN = 'N.A.', PrevAccord = 'N.D.',  Accord = 'N.D.', BPE = 'N.D.', BPT = 'N.D.' where jalon = 'Qualif' "
ImportZNND = "Update CPL_AVC set  PrevAccord = 'N.D.',  Accord = 'N.D.', BPE = 'N.D.', BPT = 'N.D.' where jalon = 'ImportZN' "
PrevAccordND = "Update CPL_AVC set Accord = 'N.D.', BPE = 'N.D.', BPT = 'N.D.' where jalon = 'PrevAccord' "
AccordND = "Update CPL_AVC set BPE = 'N.D.', BPT = 'N.D.' where jalon = 'Accord' "
# DTAND = "Update CPL_AVC set BPE = 'N.D.', BPT = 'N.D.' where jalon = 'DTA' "
BPEND = "Update CPL_AVC set BPT = 'N.D.' where jalon in ('BPE', 'Piquetage')"
QualifNA = "Update CPL_AVC set Qualifie = 'N.A.' where CRS <> ''"
ImportZNNA = "Update CPL_AVC set ImportZN = 'N.A.' where CRS = ''"

CPL_AVC_OUT = "select * from CPL_AVC"

pav_sans_fis_fic = "select code_IMB, CRS, NB_EL from CD where NB_EL < 4 and CRS = '' and ETAT_NEGO <> 'PAS DE SYNDIC'"



qte_plaque = "select count(code_imb) from OPT"
qte_plaque_el = "select ifnuLL(sum(NB_EL),0) from OPT"

nego_declare = "select count(code_imb) from OPT where Etat_Nego <> 'PAS DE SYNDIC'"
nego_declare_el = "select ifnuLL(sum(NB_EL),0) from OPT where Etat_Nego <> 'PAS DE SYNDIC'"

nego_pds = "select count(code_imb) from OPT where Etat_Nego = 'PAS DE SYNDIC'"
nego_pds_el = "select ifnuLL(sum(NB_EL),0) from OPT where Etat_Nego = 'PAS DE SYNDIC'"


valide = "select count(code_imb) from (select * from C1_C2_OK UNION select * from B1) where Etat_Nego <> 'PAS DE SYNDIC'"
valide_el = "select ifnuLL(sum(NB_EL),0) from (select * from C1_C2_OK UNION select * from B1) where Etat_Nego <> 'PAS DE SYNDIC'"

avc_nok = "select count(code_imb) from CPL_AVC"
avc_nok_el = "select ifnuLL(sum(NB_EL),0) from CPL_AVC"

infos_nok = "select count(code_imb) from C1_0_KO where ETAT_NEGO <> 'PAS DE SYNDIC'"
infos_nok_el = "select ifnuLL(sum(NB_EL),0) from C1_0_KO where ETAT_NEGO <> 'PAS DE SYNDIC'"

infos_nok_pds = "select count(code_imb) from C1_0_KO where ETAT_NEGO = 'PAS DE SYNDIC'"
infos_nok_pds_el = "select ifnuLL(sum(NB_EL),0) from C1_0_KO where ETAT_NEGO = 'PAS DE SYNDIC'"


nq = "select count(code_imb) from C1_C2_OK where  Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"
nq_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

q = "select count(code_imb) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'OuiNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"
q_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'OuiNonNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

im = "select count(code_imb) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonOuiNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"
im_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT = 'NonOuiNonNonNonNon' and Etat_Nego <> 'PAS DE SYNDIC'"

pa = "select count(code_imb) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiNonNonNon', 'OuiNonOuiNonNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"
pa_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiNonNonNon', 'OuiNonOuiNonNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"

ac = "select count(code_imb) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiNonNon', 'OuiNonOuiOuiNonNon','NonOuiOuiOuiNonNon', 'OuiNonOuiOuiNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"
ac_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiNonNon', 'OuiNonOuiOuiNonNon','NonOuiOuiOuiNonNon', 'OuiNonOuiOuiNonNon') and Etat_Nego <> 'PAS DE SYNDIC'"

dta = "select count(code_imb) from C1_C2_OK where  DTA = 'Oui' and  Qualifie || ImportZN || PrevAccord in ('OuiNonOui', 'NonOuiOui', 'NonOuiNon') and Etat_Nego <> 'PAS DE SYNDIC'"
dta_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where DTA = 'Oui' and  Qualifie || ImportZN || PrevAccord in ('OuiNonOui', 'NonOuiOui', 'NonOuiNon') and Etat_Nego <> 'PAS DE SYNDIC'"

bpe = "select count(code_imb) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiNon', 'OuiNonOuiOuiOuiNon') and Etat_Nego <> 'PAS DE SYNDIC'"
bpe_el = "select ifnuLL(sum(NB_EL),0) from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiNon', 'OuiNonOuiOuiOuiNon') and Etat_Nego <> 'PAS DE SYNDIC'"

bpt = "select count(code_imb) from (select * from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiOui', 'OuiNonOuiOuiOuiOui') and Etat_Nego <> 'PAS DE SYNDIC' union select * from B1 where Etat_Nego <> 'PAS DE SYNDIC')"
bpt_el = "select ifnuLL(sum(NB_EL),0) from (select * from C1_C2_OK where Qualifie || ImportZN || PrevAccord || Accord || BPE || BPT in ('NonOuiOuiOuiOuiOui', 'OuiNonOuiOuiOuiOui') and Etat_Nego <> 'PAS DE SYNDIC' union select * from B1 where Etat_Nego <> 'PAS DE SYNDIC')"



pav_sans_fis = "select count(code_imb) from CD where (CRS  = '' and NB_EL < 4  and ETAT_NEGO <> 'PAS DE SYNDIC')"
pav_sans_fis_el = "select ifnuLL(sum(NB_EL),0) from CD where (CRS  = '' and NB_EL < 4  and ETAT_NEGO <> 'PAS DE SYNDIC')"

stock_non_deb = "select count(code_imb) from (select * from B2 UNION select * from B3) where Etat_Nego <> 'PAS DE SYNDIC'"
stock_non_deb_el = "select ifnuLL(sum(NB_EL),0) from (select * from B2 UNION select * from B3) where Etat_Nego <> 'PAS DE SYNDIC'"



if 1 :
  from urllib.request import urlopen
  Dep = 43588
  debut = datetime(1900, 1, 1, 0, 0, 0)

  res = urlopen('http://www.python.org')
  headers = res.info()
  date = headers['date'].split(' ')
  date = date[1:4]
  date = " ".join(date)

  newdate = datetime.date(datetime.strptime(str(date), '%d %b %Y'))

  datefin = datetime(newdate.year, newdate.month, newdate.day, 0, 0, 0)

  DPL = (datefin - debut).days + 2


#  ceci est un test git 