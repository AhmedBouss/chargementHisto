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


tables_etd_tvx = {
                 1 : ['PN','1_Trame IXIETUDES_piquetage non realise', 'table_pn', 1],
                 2 : ['PR','2_Trame IXIETUDES_piquetage realise', 'table_pr', 2],
                 3 : ['PE','3_Trame IXIETUDES_pre-etude realisee', 'table_pe', 3], 
                 4 : ['ET','4_Trame IXIETUDES_etude realisee', 'table_et', 4], 
                 5 : ['TL','5_Trame IXIETUDES_ETD OK-lancées en TVX', 'table_tl', 5], 
                 6 : ['DO','6_Trame IXIETUDES_DOE realisee', 'table_do', 6], 
                 7 : ['SU','7_Trame IXIETUDES_ZE suspendue', 'table_su', 7],
                 8 : ['TD', '8_Trame IXITRAVAUX_ETD OK-TVX débutés', 'table_td', 8],
                 9 : ['TR', '9_Trame IXITRAVAUX_ETD OK-TVX réalisés', 'table_tr', 9],
                 10 : ['TB', '10_Trame IXITRAVAUX_TVX bloques', 'table_tb', 10]
                 }


defaultValue = {
				1 : ['Type_syndic', '12-Autre Syndic'],
				2 : ['Type_contact', 'Gestimm'], 
				3 : ['Statut_affectation', '1-Affecté'], 
				4 : ['Flag_immeuble', '0-Immeuble ancien'], 
				5 : ['Site_copro', 'Non']
				}



C_OK = "select * from C where ( ETAT_NEGO <> 'PAS DE SYNDIC' and Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage ||  BPT || Etudes || Travaux_Lances || DOE in( \
'NonNonNonNonNonNonNonNonNonNon', \
'NonOuiNonNonNonNonNonNonNonNon', \
'NonOuiOuiNonNonNonNonNonNonNon', \
'NonOuiOuiOuiNonNonNonNonNonNon', \
'NonOuiOuiOuiNonNonNonNonNonNon', \
'NonOuiOuiOuiOuiNonNonNonNonNon', \
'NonOuiOuiOuiOuiOuiNonNonNonNon', \
'NonOuiOuiOuiOuiOuiOuiNonNonNon', \
'NonOuiOuiOuiOuiOuiOuiOuiNonNon', \
'NonOuiOuiOuiOuiOuiOuiOuiOuiNon', \
'NonOuiOuiOuiOuiOuiOuiOuiOuiOui', \
'NonNonNonNonNonNonNonNonNonNon', \
'OuiNonNonNonNonNonNonNonNonNon', \
'OuiNonOuiNonNonNonNonNonNonNon', \
'OuiNonOuiOuiNonNonNonNonNonNon', \
'OuiNonOuiOuiNonNonNonNonNonNon', \
'OuiNonOuiOuiOuiNonNonNonNonNon', \
'OuiNonOuiOuiOuiOuiNonNonNonNon', \
'OuiNonOuiOuiOuiOuiOuiNonNonNon', \
'OuiNonOuiOuiOuiOuiOuiOuiNonNon', \
'OuiNonOuiOuiOuiOuiOuiOuiOuiNon', \
'OuiNonOuiOuiOuiOuiOuiOuiOuiOui', \
'OuiOuiNonNonNonNonNonNonNonNon', \
'OuiOuiOuiNonNonNonNonNonNonNon', \
'OuiOuiOuiOuiNonNonNonNonNonNon', \
'OuiOuiOuiOuiOuiNonNonNonNonNon', \
'OuiOuiOuiOuiOuiOuiNonNonNonNon', \
'OuiOuiOuiOuiOuiOuiOuiNonNonNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiNonNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiOuiNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiOuiOui'\
)) \
OR \
(ETAT_NEGO = 'PAS DE SYNDIC' and Piquetage || Etudes || Travaux_Lances || DOE in (\
'NonNonNonNon', 'OuiNonNonNon', 'OuiOuiNonNon', 'OuiOuiOuiNon','OuiOuiOuiOui')) \
OR \
(ETAT_NEGO <> 'PAS DE SYNDIC' and CRS = '' and NB_EL < 4 and Piquetage || Etudes || Travaux_Lances || DOE in (\
'NonNonNonNon', 'OuiNonNonNon', 'OuiOuiNonNon', 'OuiOuiOuiNon','OuiOuiOuiOui') \
OR \
Jalon = 'Piquetage non réalisé')"




C_KO = "select * from C where (( ETAT_NEGO <> 'PAS DE SYNDIC' and Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE not in( \
'NonNonNonNonNonNonNonNonNonNon', \
'NonOuiNonNonNonNonNonNonNonNon', \
'NonOuiOuiNonNonNonNonNonNonNon', \
'NonOuiOuiOuiNonNonNonNonNonNon', \
'NonOuiOuiOuiNonNonNonNonNonNon', \
'NonOuiOuiOuiOuiNonNonNonNonNon', \
'NonOuiOuiOuiOuiOuiNonNonNonNon', \
'NonOuiOuiOuiOuiOuiOuiNonNonNon', \
'NonOuiOuiOuiOuiOuiOuiOuiNonNon', \
'NonOuiOuiOuiOuiOuiOuiOuiOuiNon', \
'NonOuiOuiOuiOuiOuiOuiOuiOuiOui', \
'NonNonNonNonNonNonNonNonNonNon', \
'OuiNonNonNonNonNonNonNonNonNon', \
'OuiNonOuiNonNonNonNonNonNonNon', \
'OuiNonOuiOuiNonNonNonNonNonNon', \
'OuiNonOuiOuiNonNonNonNonNonNon', \
'OuiNonOuiOuiOuiNonNonNonNonNon', \
'OuiNonOuiOuiOuiOuiNonNonNonNon', \
'OuiNonOuiOuiOuiOuiOuiNonNonNon', \
'OuiNonOuiOuiOuiOuiOuiOuiNonNon', \
'OuiNonOuiOuiOuiOuiOuiOuiOuiNon', \
'OuiNonOuiOuiOuiOuiOuiOuiOuiOui', \
'OuiOuiNonNonNonNonNonNonNonNon', \
'OuiOuiOuiNonNonNonNonNonNonNon', \
'OuiOuiOuiOuiNonNonNonNonNonNon', \
'OuiOuiOuiOuiOuiNonNonNonNonNon', \
'OuiOuiOuiOuiOuiOuiNonNonNonNon', \
'OuiOuiOuiOuiOuiOuiOuiNonNonNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiNonNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiOuiNon', \
'OuiOuiOuiOuiOuiOuiOuiOuiOuiOui') and Jalon <> 'Piquetage non réalisé') \
and code_imb not in (select code_IMB from C where NB_EL < 4 and CRS = '' and ETAT_NEGO <> 'PAS DE SYNDIC'and Jalon <> 'Piquetage non réalisé' ))\
OR \
(ETAT_NEGO = 'PAS DE SYNDIC' and Jalon <> 'Piquetage non réalisé' and Piquetage || Etudes || Travaux_Lances || DOE not in (\
'NonNonNonNon', 'OuiNonNonNon', 'OuiOuiNonNon', 'OuiOuiOuiNon','OuiOuiOuiOui')) \
OR \
(ETAT_NEGO <> 'PAS DE SYNDIC' and CRS = '' and NB_EL < 4 and Jalon <> 'Piquetage non réalisé' and Piquetage || Etudes || Travaux_Lances || DOE not in (\
'NonNonNonNon', 'OuiNonNonNon', 'OuiOuiNonNon', 'OuiOuiOuiNon','OuiOuiOuiOui') \
OR \
Jalon <> 'Piquetage non réalisé' and Piquetage || Etudes || Travaux_Lances || DOE not in (\
'NonNonNonNon', 'OuiNonNonNon', 'OuiOuiNonNon', 'OuiOuiOuiNon','OuiOuiOuiOui' ))"



PA_Sans_FIS = "select PA, count(PA) PA_PAV_0_FIS from C_OK where NB_EL < 4 and CRS = '' group by PA"

PA_avec_FIS = "select PA, count(PA) PA_PAV_1_FIS from C_OK where NB_EL < 4 and CRS <> '' group by PA"

PA_SUP_3 = "select PA, count(*) Nb_SUP_3 from C_OK where NB_EL > 3 group by PA"

# PA avec seulement des Codes IMB < 4 sans FIS
C1 = "select * from C_OK where PA in (select  C_OK.PA from C_OK, PA_sans_FIS where  C_OK.PA = PA_sans_FIS.PA group by  C_OK.PA having count( C_OK.PA) = PA_sans_FIS.PA_PAV_0_FIS) "

# PA avec seulement des Codes IMB < 4 et au moins un Code IMB avec FIS
C2 = "select * from C_OK where PA in (select  C_OK.PA from C_OK, PA_avec_FIS where  C_OK.PA = PA_avec_FIS.PA group by  C_OK.PA having count( C_OK.PA) = PA_avec_FIS.PA_PAV_1_FIS)"

# PA avec codes IMB < 4 ou > 3 avec ou sans FIS
C3 = "select * from C_OK where code_IMB not in (select code_IMB from C1 UNION select code_IMB from C2 UNION select code_IMB from C4)"

# Codes IMB avec FIS
C3_1 = "select * from C3 where CRS <> ''"

# PA avec code IMB > 3
C4 = "select * from C_OK where PA in (select  C_OK.PA from C_OK, PA_SUP_3 where  C_OK.PA = PA_SUP_3.PA group by  C_OK.PA having count( C_OK.PA) = PA_SUP_3.NB_SUP_3)"


C_1_A = "select * from C_OK where PA in (select  C_OK.PA from C_OK, PA_sans_FIS where  C_OK.PA = PA_sans_FIS.PA group by  C_OK.PA having count( C_OK.PA) = PA_sans_FIS.PA_PAV_0_FIS) "

C_2_A = "select * from C2 where CRS = ''"

C_2_B = "select * from C2 where CRS <> ''"

C_3_A = "select * from C3 where CRS = '' and NB_EL < 4 "

C_3_B = "select * from C3 where CRS = '' and NB_EL > 3 "

C_3_1_A = "select * from C3_1 where CRS in (select CRS from C3_1 where CRS <> ''  group by CRS having count(distinct IMM_PAV) > 1) OR IMM_PAV = 'IMM' order by 2"

C_3_1_B = "select * from C3_1 where NB_EL < 4 and CRS in (select CRS from C3_1 where CRS <> ''  group by CRS having count(distinct IMM_PAV) = 1) order by 2"

C_4_A = "select * from C4 where CRS = '' "

C_4_B = "select * from C4 where CRS <> '' "






RGT_PAV_CTRL = "select * from C1 UNION select * from C2 where CRS = '' UNION select * from C3 where CRS = '' and NB_EL < 4"

# RGT_NEGO_PAV = ""

# RGT_IMM_CTRL = ""

# RGT_IMM_SANS_ZN = ""


completude_infos = "select Code_IMB, Adresse, NB_EL, CRS, Code_Syndic, Etat_Nego, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, Presence_DTA, Date_reception_bon_PE from C1_0_KO where Etat_Nego <>'PAS DE SYNDIC' order by CRS"

completude_avc_etd_tvx = "select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, BPE, Piquetage, BPT, Etudes,  \
Travaux_lances, DOE, Etat_Nego, Code_Syndic, CRS, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS, \
Date_reception_bon_PE, Date_mad_PE, Date_fin_EZA, Presence_DTA, Date_env_pe_syndic, Type_Site, NB_EL_P, NB_EL_R, Date_valid_PE, \
Resultat_etude_site, Date_prog_racco, Date_pose_PB, '' Suspendre, '' Type_Suspension from  C_KO where code_imb not in (select code_imb from C1_1_KO) and code_imb not in (select code_imb from C2_1_KO) and code_imb not in (select code_imb from C1_0_KO)"

nok_etd_neg = "select Code_IMB, Adresse, NB_EL, Jalon, Qualifie, ImportZN, PrevAccord, Accord, BPE, Piquetage, BPT, Etudes,\
Travaux_lances, DOE, Etat_Nego, Code_Syndic, CRS, Date_envoi_AS, AG_non_pertinente, Date_retour_convention, Date_saisie_AS,\
Date_reception_bon_PE, Date_mad_PE, Date_fin_EZA, Presence_DTA, Date_env_pe_syndic, Type_Site, NB_EL_P, NB_EL_R, Date_valid_PE,\
Resultat_etude_site, Date_prog_racco, Date_pose_PB, '' Suspendre, '' Type_Suspension from  C_KO \
where (code_imb in (select code_imb from C1_1_KO) OR code_imb in (select code_imb from C2_1_KO) OR code_imb in (select code_imb from C1_0_KO))"

Requete_Plaque ="select plaque from C1_C2_OK limit 1"


QualifND = "Update CPL_AVC_ETD set ImportZN = 'N.A.', PrevAccord = 'N.D.',  Accord = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Qualif' "
ImportZNND = "Update CPL_AVC_ETD set  PrevAccord = 'N.D.',  Accord = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'ImportZN' "
PrevAccordND = "Update CPL_AVC_ETD set  Accord = 'N.D.', BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'PrevAccord' "
AccordND = "Update CPL_AVC_ETD set  BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Accord' "
DTAND = "Update CPL_AVC_ETD set BPE = 'N.D.', Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'DTA' "
BPEND = "Update CPL_AVC_ETD set Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'BPE' "
PINND = "Update CPL_AVC_ETD set  Piquetage = 'N.D.', Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Piquetage non réalisé' "
PIQND = "Update CPL_AVC_ETD set  Etudes = 'N.D.',  BPT = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'Piquetage' "
BPTND = "Update CPL_AVC_ETD set  Etudes = 'N.D.', Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon = 'BPT' "
ETDND = "Update CPL_AVC_ETD set  Travaux_Lances = 'N.D.', DOE = 'N.D.' where jalon in ('Etudes', 'Etude') "
TLAND = "Update CPL_AVC_ETD set  DOE = 'N.D.' where jalon = 'Travaux lancés' "

QualifNA = "Update CPL_AVC_ETD set Qualifie = 'N.A.' where CRS <> ''"
ImportZNNA = "Update CPL_AVC_ETD set ImportZN = 'N.A.' where CRS = ''"
Nonqualif = "Update CPL_AVC_ETD set Qualifie = 'N.A.', ImportZN = 'N.A.', PrevAccord = 'N.A.',  Accord = 'N.A.', BPE = 'N.A.', BPT = 'N.A.' where ETAT_NEGO = 'PAS DE SYNDIC'"


# Type_site_Vide_1 = "Update C set  Type_site = 'C-Pour un immeuble' where Type_site = 'Immeuble'"

# Type_site_Vide_2 = "Update C set  Type_site = 'S-Pour une maison individuelle' where Type_site = 'Maison'"



CPL_AVC_ETD_OUT = "select * from CPL_AVC_ETD"


C_OK_RGT = "SELECT * FROM C_1_A UNION \
SELECT * FROM C_2_A UNION \
SELECT * FROM C_2_B UNION \
SELECT * FROM C_3_A UNION \
SELECT * FROM C_3_B UNION \
SELECT * FROM C_3_1_A UNION \
SELECT * FROM C_3_1_B UNION \
SELECT * FROM C_4_A UNION \
SELECT * FROM C_4_B UNION \
SELECT * FROM B1 "


select_PN_U = "select code_imb, Adresse, plaque, centre_etude, localite, CODE_SYNDIC, NB_EL, PB, '' FI, PA, PM, '' Typo_PB, '' Date_realisation, 'T' Nature_Code_IMB, '' Code_Pere, '' Nature_Facade, Regroupement, Escalier, 0 NB_Etages, NB_EL_P, NB_EL_R, Type_Site  from C_OK_RGT \
where Jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT') \
and PA in (select PA from C_OK_RGT where Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'Piquetage non réalisé'))"

select_PR_U = "select code_imb, Adresse, plaque, centre_etude, localite, CODE_SYNDIC, NB_EL, PB, '' FI, PA, PM, '' Typo_PB, '' Date_realisation, 'T' Nature_Code_IMB, '' Code_Pere, '' Nature_Facade, Regroupement, Escalier, 0 NB_Etages, NB_EL_P, NB_EL_R, Type_Site from C_OK_RGT where Jalon in ('Piquetage', 'BPT')"

select_ET_U = "select code_imb, Adresse, plaque, centre_etude, localite, CODE_SYNDIC, NB_EL, PB, '' FI, PA, PM, '' Typo_PB, '' Date_realisation, 'T' Nature_Code_IMB, '' Code_Pere, '' Nature_Facade, Regroupement, Escalier, 0 NB_Etages, NB_EL_P, NB_EL_R, Type_Site, Date_fin_EZA, Resultat_etude_site from C_OK_RGT where Jalon in ('Etude')"

select_TL_U = "select code_imb, Adresse, plaque, centre_etude, localite, CODE_SYNDIC, NB_EL, PB, '' FI, PA, PM, '' Typo_PB, '' Date_realisation, 'T' Nature_Code_IMB, '' Code_Pere, '' Nature_Facade, Regroupement, Escalier, 0 NB_Etages, NB_EL_P, NB_EL_R, Type_Site, Date_fin_EZA, Resultat_etude_site, FCI_ACCES from C_OK_RGT where Jalon in ('Travaux lancés') and Date_prog_racco <> ''"

select_TR_U = "select code_imb, Adresse, plaque, centre_travaux, '' CDT, localite, NB_EL, PB, PA, PM, '' Date_realisation, 'Interne' Prod_int_presta, '' Entreprise, FCI_ACCES, date_pose_pb from C_OK_RGT where Jalon in ('DOE') and date_pose_pb <> ''"

select_DO_U = "select code_imb, Adresse, plaque, centre_etude, localite, CODE_SYNDIC, NB_EL, PB, '' FI, PA, PM, \
     '' Typo_PB, '' Date_realisation, 'T' Nature_Code_IMB, '' Code_Pere, '' Nature_Facade, Regroupement, \
     Escalier, 0 NB_Etages, NB_EL_P, NB_EL_R, Type_Site, Date_fin_EZA, Resultat_etude_site, FCI_ACCES, \
     date_pose_pb, FCI_FIN from C_OK_RGT where Jalon in ('DOE') and date_pose_pb <> ''"


qte_plaque = "select count(code_imb) from OPT"
qte_plaque_el = "select ifnuLL(sum(NB_EL),0) from OPT"

nego_declare = "select count(code_imb) from OPT where Etat_Nego <> 'PAS DE SYNDIC'"
nego_declare_el = "select ifnuLL(sum(NB_EL),0) from OPT where Etat_Nego <> 'PAS DE SYNDIC'"

nego_pds = "select count(code_imb) from OPT where Etat_Nego = 'PAS DE SYNDIC'"
nego_pds_el = "select ifnuLL(sum(NB_EL),0) from OPT where Etat_Nego = 'PAS DE SYNDIC'"

valide = "select count(code_imb) from ( select code_imb from C_OK_RGT  where jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT') \
and PA in (select PA from C_OK_RGT where Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'Piquetage non réalisé')) OR Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT'))" 

valide_el = "select ifnuLL(sum(NB_EL),0) from ( select * from C_OK_RGT  where jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT') \
and PA in (select PA from C_OK_RGT where Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'Piquetage non réalisé')) OR Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'BPT'))"  

non_debute = " select count(code_imb) from B2 "
non_debute_el = " select ifnuLL(sum(NB_EL),0) from B2 "

avc_nok = "select count(code_imb) from CPL_AVC_ETD"
avc_nok_el = "select ifnuLL(sum(NB_EL),0) from CPL_AVC_ETD"

sans_pa = " select count(code_imb) from B3 "
sans_pa_el = " select ifnuLL(sum(NB_EL),0) from B3 "

sans_pa_avc = " select count(code_imb) from D "
sans_pa_avc_el = " select ifnuLL(sum(NB_EL),0) from D "

pa_encours = " select count(code_imb) from C "
pa_encours_el = " select ifnuLL(sum(NB_EL),0) from C "

sans_etude = "select count(code_imb) from ( select * from C_OK_RGT  where jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'Piquetage non réalisé', 'BPT')) "

sans_etude_el = "select ifnuLL(sum(NB_EL),0) from ( select * from C_OK_RGT  where jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE', 'Piquetage non réalisé', 'BPT'))"

pa_fin = " select count(code_imb) from B1 "
pa_fin_el = " select ifnuLL(sum(NB_EL),0) from B1 "

pn = "select count(code_imb) from (" + select_PN_U + ")"
pn_el = "select ifnuLL(sum(NB_EL),0) from (" + select_PN_U + ")"

pr = "select count(code_imb) from (" + select_PR_U + ")"
pr_el = "select ifnuLL(sum(NB_EL),0) from (" + select_PR_U + ")"

et = "select count(code_imb) from (" + select_ET_U + ")"
et_el = "select ifnuLL(sum(NB_EL),0) from (" + select_ET_U + ")"

tl = "select count(code_imb) from (" + select_TL_U + ")"
tl_el = "select ifnuLL(sum(NB_EL),0) from (" + select_TL_U + ")"

do = "select count(code_imb) from (" + select_DO_U + ")"
do_el = "select ifnuLL(sum(NB_EL),0) from (" + select_DO_U + ")"


liste_a = "select code_IMB, Adresse, NB_EL, PA from B2"

liste_b = "select code_IMB, Adresse, NB_EL, PA from C"

liste_c = "select code_IMB, Adresse, NB_EL, PA from B3"

liste_f = "select code_IMB, Adresse, NB_EL, PA from ( select * from C_OK_RGT  where jalon not in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE') \
and PA  not in (select PA from C_OK_RGT where Jalon in ('Piquetage', 'Etude', 'Travaux lancés', 'DOE')))"

prod_terminee = "select code_IMB, Adresse, NB_EL, PA from B1 "

fin_eza = "select code_IMB, Adresse, NB_EL, Date_fin_EZA from C where Resultat_etude_site in ('OK', 'ok', 'Ok') and Date_fin_EZA = '' and Jalon in ('Etude', 'DOE', 'Travaux lancés') \
Union select code_IMB, Adresse, NB_EL, Date_fin_EZA from B1 where Resultat_etude_site in ('OK', 'ok', 'Ok') and Date_fin_EZA = '' and Jalon in ('Etude', 'DOE', 'Travaux lancés')"

num_fci_a = "select code_IMB, Adresse, NB_EL, '' from C where Jalon in ('DOE', 'Travaux lancés') and fci_acces = '' \
Union select code_IMB, Adresse, NB_EL, '' from B1 where Jalon in ('DOE', 'Travaux lancés') and fci_acces = '' and Date_prog_racco <> ''"

num_fci_f = "select code_IMB, Adresse, NB_EL, '' from C where Jalon in ('DOE') and fci_fin = '' \
Union select code_IMB, Adresse, NB_EL, '' from B1 where Jalon in ('DOE') and fci_fin = '' and Date_pose_PB <> ''"


fin_eza_nb = "select count(*)  from (select * from C where Resultat_etude_site in ('OK', 'ok', 'Ok') and Date_fin_EZA = '' and Jalon in ('Etude', 'DOE', 'Travaux lancés') Union select * from B1 where Date_fin_EZA = '' )"
fin_eza_nb_el = "select ifnuLL(sum(NB_EL),0)  from (select * from C where Resultat_etude_site in ('OK', 'ok', 'Ok') and Date_fin_EZA = '' and Jalon in ('Etude', 'DOE', 'Travaux lancés') Union select * from B1 where Date_fin_EZA = '' )"

num_fci_a_nb = "select count(*)  from (select * from C where Jalon in ('DOE', 'Travaux lancés') Union select * from B1) where fci_acces = '' and Date_prog_racco <> ''"
num_fci_a_nb_el = "select ifnuLL(sum(NB_EL),0)  from (select * from C where Jalon in ('DOE', 'Travaux lancés') Union select * from B1) where fci_acces = ''and Date_prog_racco <> ''"

num_fci_f_nb = "select count(*)  from (select * from C where Jalon in ('DOE')  Union select * from B1) where fci_fin = ''  and Date_pose_PB <> ''"
num_fci_f_nb_el = "select ifnuLL(sum(NB_EL),0)  from (select * from C where Jalon in ('DOE')  Union select * from B1) where fci_fin = ''  and Date_pose_PB <> ''"

nok_etd_neg_nb = "select count(code_imb) from (" + nok_etd_neg + ")"
nok_etd_neg_el = "select ifnuLL(sum(NB_EL),0) from (" + nok_etd_neg + ")"


updateNondebute = "update C set Nego_non_deb = 'Oui' where Etat_Nego <> 'PAS DE SYNDIC' and Code_Syndic = ''"

updateQualifie = "update C set Qualifie = 'Oui' where Code_Syndic <> '' and CRS = ''"

updateImportZN = "update C set ImportZN = 'Oui' where Code_Syndic <> '' and CRS <> ''"

updatePrevAccord = "update C set PrevAccord = 'Oui' where Date_envoi_AS <> '' and (AG_non_pertinente = '1-AG non pertinente' OR AG_non_pertinente = '0-AG pertinente' )"

updateAccord = "update C set Accord = 'Oui' where Date_retour_convention <> '' OR Date_saisie_AS <> ''" 

updateDTA = "update C set DTA = 'Oui' where Presence_DTA in ('1-DTA présent', '2-DTA inutile')"

updateBPE = "update C set BPE = 'Oui' where Date_reception_bon_PE <> ''" 

updatePiquetage = "update C set Piquetage = 'Oui' where (Etat_Nego <> 'PAS DE SYNDIC' and Date_mad_PE <> '') OR (Etat_Nego = 'PAS DE SYNDIC' and Resultat_etude_site <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '')"

updateBPT = "update C set BPT = 'Oui' where Date_valid_PE <> '' and Date_saisie_AS <> '' and Date_env_pe_syndic <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> ''  and Presence_DTA in ('1-DTA présent', '2-DTA inutile')"

updateEtudes = "update C set Etudes = 'Oui' where (Etat_Nego <> 'PAS DE SYNDIC' and Resultat_etude_site <> ''  and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '') OR (Etat_Nego = 'PAS DE SYNDIC' and Resultat_etude_site <> '' and Type_Site <> '' and NB_EL_P <> '' and NB_EL_R <> '')" 

updateTravauxLances = "update C set Travaux_lances = 'Oui' where Date_prog_racco <> ''"

updateTravauxRealises = "update C set Travaux_realises = 'Oui' where Date_pose_PB <> '' "

updateDOE = "update C set DOE = 'Oui' where Date_pose_PB <> '' "



JalonSansProd = "Update C set Jalon = 'Sans_Prod' where Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'NonNonNonNonNonNonNonNonNonNon'"

JalonQualif = "Update C set Jalon = 'Qualif' where Qualifie || ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNonNonNon'"

JalonImportZN = "Update C set Jalon = 'ImportZN' where ImportZN || PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNonNon'"

JalonPrevAccord = "Update C set Jalon = 'PrevAccord' where PrevAccord || Accord || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNonNon' "

JalonAccord = "Update C set Jalon = 'Accord' where Accord || BPE || Piquetage || BPT || Etudes ||  Travaux_Lances || DOE = 'OuiNonNonNonNonNonNon'"

# JalonDTA = "Update C set Jalon = 'DTA' where DTA || BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNonNon'"

JalonBPE = "Update C set Jalon = 'BPE' where BPE || Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNonNon'"

JalonPiquetage = "Update C set Jalon = 'Piquetage' where  Piquetage || BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNonNon'"

JalonBPT = "Update C set Jalon = 'BPT' where  BPT || Etudes || Travaux_Lances || DOE = 'OuiNonNonNon'"

JalonEtude = "Update C set Jalon = 'Etude' where Etudes || Travaux_Lances || DOE = 'OuiNonNon'"

JalonTL = "Update C set Jalon = 'Travaux lancés' where Travaux_Lances || DOE ='OuiNon' "

JalonDOE = "Update C set Jalon = 'DOE' where  DOE = 'Oui'"

EZA = "select Code_IMB, Adresse, plaque, centre_etude, localite, Date_fin_EZA from C_OK_RGT where Resultat_etude_site in ('OK', 'ok', 'Ok') and Date_fin_EZA <> ''"

CMD = "select Code_IMB, Adresse, plaque, centre_etude, localite, FCI_ACCES, 'Simple', FCI_ACCES, Date_FCI from C_OK_RGT where FCI_ACCES <> '' and Date_FCI <> '' and jalon in ('Travaux lancés', 'DOE')"

centre_etd = "select centre_etude from C limit 1"


Admin_PA = "select code_IMB, PA_INI from C_OK_RGT where jalon in ('Piquetage non réalisé', 'Piquetage', 'Etude', 'Travaux lancés', 'DOE')"