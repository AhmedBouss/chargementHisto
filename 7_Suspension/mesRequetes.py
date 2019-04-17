# -*- coding: utf-8 -*-
from datetime import datetime


table_NEGO = {
				1 : ['Code_IMB', 'Dossier - Code IMB', 'CHARINT', 'liste[1]', '?'],
				2 : ['Plaque', 'Plaque', 'CHARINT', 'liste[2]', '?'],
				3 : ['Existant', 'Existant SPI ', 'CHARINT', 'liste[3]', '?'],
				4 : ['Statut', 'Statut', 'CHARINT', 'liste[4]', '?'],
				5 : ['ID_ZN', 'ID_ZN', 'CHARINT', 'liste[5]', '?'],
				6 : ['Erreur', 'Erreur', 'CHARINT', 'liste[6]', '?']
				}


table_ETUDE = {
				1 : ['Code_IMB', 'Dossier - Code IMB', 'CHARINT', 'liste[1]', '?'],
				2 : ['Plaque', 'Plaque', 'CHARINT', 'liste[2]', '?'],
				3 : ['Existant', 'Existant SPI', 'CHARINT', 'liste[3]', '?'],
				4 : ['Statut', 'Statut', 'CHARINT', 'liste[4]', '?'],
				5 : ['ID_ZE', 'ID_ZN', 'CHARINT', 'liste[5]', '?'],
				6 : ['ID_RGT', 'ID_RGT', 'CHARINT', 'liste[6]', '?'],
				7 : ['Erreur', 'Erreur', 'CHARINT', 'liste[7]', '?']
				}


table_ETUDE_2 = {
				1 : ['Code_IMB', 'Dossier - Code IMB', 'CHARINT', 'liste[1]', '?'],
				2 : ['Plaque', 'Plaque', 'CHARINT', 'liste[2]', '?'],
				3 : ['Existant', 'Existant SPI', 'CHARINT', 'liste[3]', '?'],
				4 : ['Statut', 'Statut', 'CHARINT', 'liste[4]', '?'],
				5 : ['ID_ZE', 'ID_ZN', 'CHARINT', 'liste[5]', '?'],
				6 : ['ID_RGT', 'ID_RGT', 'CHARINT', 'liste[6]', '?'],
				7 : ['CEM', 'CEM', 'CHARINT', 'liste[7]', '?'],
				8 : ['Importe', 'Importe', 'CHARINT', 'liste[8]', '?'],
				9 : ['Erreur', 'Erreur', 'CHARINT', 'liste[9]', '?']
				}


table_SPI = {
				1 : ['Code_IMB', 'Dossier - Code IMB', 'CHARINT', 'liste[1]', '?'],
				2 : ['Plaque', 'Plaque', 'CHARINT', 'liste[2]', '?'],
				3 : ['Action', 'Action', 'CHARINT', 'liste[3]', '?'],
				4 : ['ID_ZN', 'ID_ZN', 'CHARINT', 'liste[4]', '?'],
				5 : ['ID_ZE', 'ID_ZE', 'CHARINT', 'liste[5]', '?'],
				6 : ['ID_RGT', 'ID_RGT', 'CHARINT', 'liste[6]', '?'],
				7 : ['Statut', 'Statut', 'CHARINT', 'liste[7]', '?'],
				8 : ['Erreur', 'Erreur', 'CHARINT', 'liste[8]', '?']
				}


Neg = "select * from NEG where Existant = 'Non existant SPI-Existant NEG' and Statut <> 'Existant SPI-Non existant NEG'  and statut <> 'ZN terminée'"
Err_neg = "Update NEG set Erreur  = 'Erreur 1-0-0-0-2' where Existant = 'Non existant SPI-Existant NEG' and Statut <> 'Existant SPI-Non existant NEG'  and statut <> 'ZN terminée'"

Etd = "select * from ETD where Existant = 'Non existant SPI-Existant ETD' and statut <> 'DOE réalisée'"
Err_etd = "Update ETD set Erreur  = 'Erreur 2-0-0-0-0-2' where Existant = 'Non existant SPI-Existant ETD' and statut <> 'DOE réalisée'"

Spi_neg = "select * from SPI where action = 'Qualif Négo' and statut in('Qualifié', 'En attente prev accord') \
OR action = 'Probation' and statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE') \
OR action = 'Obtention BPT' and statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE', 'En attente BPT')"


Spi_etd = "select * from SPI where action = 'Réalisation étude' and statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours' ) \
OR action = 'Lancement travaux' and statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours') \
OR action = 'Date Pose PB' and statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours') \
OR action = 'TFX' and statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours')"


tbl = "create table tbl as select distinct Code_IMB, Plaque, '' qualif_nego, '' probation, '' BPT, '' Etude, '' TL, '' DOE, '' TFX, statut, ID_ZN, ID_ZE, ID_RGT \
from SPI where statut not in ('DOE réalisée', 'ZN terminée')"
# '' , non_neg, '' quafie, '' attente_pa, '' attente_a, '' attente_bpe, '' attente_bpt, '' attente_dta, \
# '' non_etd, '' piquetage, '' qualif_en_cours, '' etude_en_cours, '' lct_tvx, '' tvx_eff, '' tvx_en_cours, '' doe_en_cours  \

qn0 = "update tbl set qualif_NEGO = '1' where Code_IMB in (select code_imb from SPI where action = 'Qualif Négo' and statut not in('En attente accord Syndic', 'En attente BPE', 'En attente BPT', 'En attente DTA', 'ZN terminée'))"
qn1 = "update tbl set qualif_NEGO = '0' where Code_IMB not in  (select code_imb from SPI where action = 'Qualif Négo' and statut not in('En attente accord Syndic', 'En attente BPE', 'En attente BPT', 'En attente DTA', 'ZN terminée'))"

prob0 = "update tbl set Probation = '1' where Code_IMB in (select code_imb from SPI where action = 'Probation' and statut not in('En attente BPE', 'En attente BPT', 'En attente DTA', 'ZN terminée'))"
prob1 = "update tbl set Probation = '0' where Code_IMB not in  (select code_imb from SPI where action = 'Probation' and statut not in('En attente BPE', 'En attente BPT', 'En attente DTA', 'ZN terminée'))"

bpt0 = "update tbl set BPT = '1' where Code_IMB in (select code_imb from SPI where action = 'Obtention BPT' and statut not in('En attente DTA', 'ZN terminée'))"
bpt1 = "update tbl set BPT = '0' where Code_IMB not in  (select code_imb from SPI where action = 'Obtention BPT' and statut not in('En attente DTA', 'ZN terminée'))"

etd0 = "update tbl set Etude = '1' where Code_IMB in (select code_imb from SPI where action = 'Réalisation étude' and statut not in('Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours'))"
etd1 = "update tbl set Etude = '0' where Code_IMB not in (select code_imb from SPI where action = 'Réalisation étude' and statut not in('Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours'))"

tl0 = "update tbl set TL = '1' where Code_IMB in (select code_imb from SPI where action = 'Lancement travaux' and statut not in('Travaux en cours', 'DOE en cours', 'RET CH en cours'))"
tl1 = "update tbl set TL = '0' where Code_IMB not in  (select code_imb from SPI where action = 'Lancement travaux' and statut not in('Travaux en cours', 'DOE en cours', 'RET CH en cours'))"

doe0 = "update tbl set DOE = '1' where Code_IMB in (select code_imb from SPI where action = 'Date Pose PB' and statut not in('DOE réalisée'))"
doe1 = "update tbl set DOE = '0' where Code_IMB not in  (select code_imb from SPI where action = 'Date Pose PB' and statut not in('DOE réalisée'))"

tfx0 = "update tbl set TFX = '1' where Code_IMB in (select code_imb from SPI where action = 'TFX' and statut not in('DOE réalisée'))"
tfx1 = "update tbl set TFX = '0' where Code_IMB not in  (select code_imb from SPI where action = 'TFX' and statut not in('DOE réalisée'))"



suspension_nego = "select SPI.code_imb, 'MIG pb SI Traitement asynchrone Admin SI', 'Erreur 1-' || tbl.qualif_nego || '-' || tbl.probation || '-' || tbl.bpt || '-0'  from SPI, tbl where \
(action = 'Qualif Négo' and SPI.statut in('Qualifié', 'En attente prev accord') \
OR action = 'Probation' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE') \
OR action = 'Obtention BPT' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE', 'En attente BPT')) \
and spi.code_imb = tbl.code_imb and SPI.ID_ZN <> 'SO' \
UNION \
select code_imb, 'MIG pb SI Traitement asynchrone Admin SI', Erreur from NEG where Erreur <> '' and ID_ZN <> 'SO'"


commentaire_nego = "select 'CODE IMB', SPI.code_imb, SPI.ID_ZN, '' Date_commentaire, 'SUPNEG_CLA_PIL_02', 'Erreur 1-' || tbl.qualif_nego || '-' || tbl.probation || '-' || tbl.bpt || '-0'  from SPI, tbl where \
(action = 'Qualif Négo' and SPI.statut in('Qualifié', 'En attente prev accord') \
OR action = 'Probation' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE') \
OR action = 'Obtention BPT' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE', 'En attente BPT')) \
and spi.code_imb = tbl.code_imb and SPI.ID_ZN = 'SO' \
UNION \
select 'CODE IMB', code_imb, ID_ZN, '' Date_commentaire, 'ADMIN_IXIN', Erreur from NEG where Erreur <> '' and ID_ZN = 'SO'"



suspension_etude = "select distinct SPI_ZE.code_imb, SPI_ZE.Plaque, 'MIG pb SI Traitement asynchrone Admin SI', \
'Erreur 2-' || tbl.etude || '-' || tbl.TL || '-' || tbl.DOE || '-' || tbl.TFX || '-0'  from SPI_ZE, tbl \
where SPI_ZE.code_imb = tbl.code_imb and SPI_ZE.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select SPI.code_imb, SPI.Plaque, 'MIG pb SI Traitement asynchrone Admin SI', 'Erreur 120'  from SPI,  tbl \
where (SPI.ID_ZE in (select ID_ZE from spi_ze) and SPI.code_imb not in (select SPI_ZE.code_imb from spi_ze)) \
and SPI.code_imb = tbl.code_imb and SPI.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select code_imb, Plaque, 'MIG pb SI Traitement asynchrone Admin SI', Erreur from ETD where Erreur <> '' \
UNION \
select code_imb, Plaque, 'MIG pb SI Traitement asynchrone Admin SI', 'Erreur 120' from ETD2 where ID_ZE in (select ID_ZE from ETD where erreur <> '') \
UNION \
select code_imb, Plaque, 'MIG pb SI Traitement asynchrone Admin SI', 'Erreur 120' from ETD2 where ID_ZE in (select ID_ZE from SPI_ZE)"


etd_ze = "update ETD set erreur = 'Erreur 120' where ID_ZE in (select ID_ZE from ETD where Erreur <> '') and Erreur = '' and statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') "

spi_ze = "create table spi_ze as select SPI.code_imb, '' Adresse, SPI.Plaque, '' centre_etd, '' Localite, 'MIG pb SI Traitement asynchrone Admin SI', 'Erreur 2-' || tbl.etude || '-' || tbl.TL || '-' || tbl.DOE || '-' || tbl.TFX || '-0' Erreur, SPI.ID_ZE, SPI.Statut, SPI.ID_RGT  from SPI, tbl where \
(action = 'Réalisation étude' and SPI.statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours' ) \
OR action = 'Lancement travaux' and SPI.statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours') \
OR action = 'Date Pose PB' and SPI.statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours') \
OR action = 'TFX' and SPI.statut in('Piquetage en cours', 'Qualif en cours', 'Etude en cours', 'Lct Tvx RGT en cours', 'Travaux à effectuer', 'Travaux en cours', 'DOE en cours', 'RET CH en cours'))\
and spi.code_imb = tbl.code_imb and SPI.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée')"

# Suspension pour MaJ RGT en retard - Erreur 120


cem_opt_zn = "select distinct SPI.Code_IMB, SPI.ID_ZN, 'CEM OPT'  from SPI, tbl where \
(action = 'Qualif Négo' and SPI.statut in('Qualifié', 'En attente prev accord')  \
OR action = 'Probation' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE')  \
OR action = 'Obtention BPT' and SPI.statut in('Qualifié', 'En attente prev accord', 'En attente accord Syndic', 'En attente BPE', 'En attente BPT'))  \
and spi.code_imb = tbl.code_imb and SPI.ID_ZN <> 'SO'  \
UNION  \
select distinct Code_IMB, ID_ZN, 'CEM OPT' from NEG where Erreur <> '' and ID_ZN <> 'SO' "

cem_opt_ze = "select distinct SPI_ZE.ID_ZE, 'CEM OPT' from SPI_ZE, tbl \
where SPI_ZE.code_imb = tbl.code_imb and SPI_ZE.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select distinct SPI.ID_ZE, 'CEM OPT'  from SPI,  tbl \
where (SPI.ID_ZE in (select ID_ZE from spi_ze) and SPI.code_imb not in (select SPI_ZE.code_imb from spi_ze)) \
and SPI.code_imb = tbl.code_imb and SPI.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select distinct ID_ZE, 'CEM OPT' from ETD where Erreur <> '' \
UNION \
select distinct ID_ZE, 'CEM OPT' from ETD2 where ID_ZE in (select ID_ZE from ETD where erreur <> '')\
UNION \
select distinct ID_ZE, 'CEM OPT' from ETD2 where ID_ZE in (select ID_ZE from SPI_ZE)"

cem_opt_rgt = "select distinct SPI_ZE.ID_RGT, 'CEM OPT' from SPI_ZE, tbl \
where SPI_ZE.code_imb = tbl.code_imb and SPI_ZE.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select distinct SPI.ID_RGT, 'CEM OPT'  from SPI,  tbl \
where (SPI.ID_RGT in (select ID_RGT from spi_ze) and SPI.code_imb not in (select SPI_ZE.code_imb from spi_ze)) \
and SPI.code_imb = tbl.code_imb and SPI.statut not in ('DOE réalisée', 'Non existant IXIETUDES', 'Non existant IXINEGO', 'ZN terminée') \
UNION \
select distinct ID_RGT, 'CEM OPT' from ETD where Erreur <> '' \
UNION \
select distinct ID_RGT, 'CEM OPT' from ETD2 where ID_RGT in (select ID_RGT from ETD where erreur <> '')\
UNION \
select distinct ID_RGT, 'CEM OPT' from ETD2 where ID_RGT in (select ID_RGT from SPI_ZE)"


liste_a1_neg = "select code_IMB, Plaque, Existant, Statut, ID_ZN, Erreur from NEG where Existant = 'Non existant SPI-Existant NEG' and Statut <> 'Existant SPI-Non existant NEG'  and statut <> 'ZN terminée'"

liste_a1_etd= "select code_IMB, Plaque, Existant, Statut, ID_ZE, ID_RGT, Erreur from ETD where Existant = 'Non existant SPI-Existant ETD' and Statut <> 'Existant SPI-Non existant ETD'  and statut <> 'DOE réalisée'"

liste_a2 = "select * from SPI where code_imb in (select Code_IMB from tbl where statut in ('Non existant IXINEGO', 'Non existant IXIETUDES'))"


# qualif_nego || probation || BPT || Etude || TL || DOE || TFX

liste_a3 = "select Code_IMB, Plaque, 'Qualif Négo' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where qualif_nego  = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a4 = "select Code_IMB, Plaque, 'Probation' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where probation = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a5 = "select Code_IMB, Plaque, 'Obtention BPT' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where BPT = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a6 = "select Code_IMB, Plaque, 'Réalisation étude' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where Etude = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a7 = "select Code_IMB, Plaque, 'Lancement travaux' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where  TL = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a8 = "select Code_IMB, Plaque, 'Date Pose PB' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where  DOE = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"

liste_a9 = "select Code_IMB, Plaque, 'TFX' Action, ID_ZN, ID_ZE, ID_RGT, Statut from tbl where TFX = '1' and statut not in ('Non existant IXINEGO', 'Non existant IXIETUDES')"


# --------------------------------Itération pour création de la tables NEGO à partir du dictionnaire----------------------------------------------
requestn = []
tupn = []
ma_listen = []

for element in table_NEGO:

	champn = table_NEGO[element][0]
	typechampn = table_NEGO[element][2]
	insn = table_NEGO[element][3]
	sepn = table_NEGO[element][4]

	itemn = champn + " " + typechampn

	requestn.append(itemn)
	ma_listen.append(insn)
	tupn.append(sepn)

OPT_requestn = ", ".join(requestn)
tupn = ", ".join(tupn)
ma_listen = ", ".join(ma_listen)

# --------------------------------Itération pour création de la tables NEGO à partir du dictionnaire----------------------------------------------
requeste = []
tupe = []
ma_listee = []

for element in table_ETUDE:

	champe = table_ETUDE[element][0]
	typechampe = table_ETUDE[element][2]
	inse = table_ETUDE[element][3]
	sepe = table_ETUDE[element][4]

	iteme = champe + " " + typechampe

	requeste.append(iteme)
	ma_listee.append(inse)
	tupe.append(sepe)

OPT_requeste = ", ".join(requeste)
tupe = ", ".join(tupe)
ma_listee = ", ".join(ma_listee)


# --------------------------------Itération pour création de la tables NEGO à partir du dictionnaire----------------------------------------------
requeste2 = []
tupe2 = []
ma_listee2 = []

for element in table_ETUDE_2:

	champe2 = table_ETUDE_2[element][0]
	typechampe2 = table_ETUDE_2[element][2]
	inse2 = table_ETUDE_2[element][3]
	sepe2 = table_ETUDE_2[element][4]

	iteme2 = champe2 + " " + typechampe2

	requeste2.append(iteme2)
	ma_listee2.append(inse2)
	tupe2.append(sepe2)

OPT_requeste2 = ", ".join(requeste2)
tupe2 = ", ".join(tupe2)
ma_listee2 = ", ".join(ma_listee2)

# --------------------------------Itération pour création de la tables NEGO à partir du dictionnaire----------------------------------------------
requests = []
tups = []
ma_listes = []

for element in table_SPI:

	champs = table_SPI[element][0]
	typechamps = table_SPI[element][2]
	inss = table_SPI[element][3]
	seps = table_SPI[element][4]

	items = champs + " " + typechamps

	requests.append(items)
	ma_listes.append(inss)
	tups.append(seps)

OPT_requests = ", ".join(requests)
tups = ", ".join(tups)
ma_listes = ", ".join(ma_listes)
