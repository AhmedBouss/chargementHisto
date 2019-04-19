# -*- coding: utf-8 -*-
import Db_SPI
import Export_Excel
from mesRequetes import *
import os

# --------------------------------------------------------------------------------------------------------------------------------

maTable_SPI = Db_SPI.BaseSqlite()
maTable_SPI.creerTableSPI("Qualif")
maTable_SPI.creerTableSPI("Probation")
maTable_SPI.creerTableSPI("Etude")
maTable_SPI.creerTableSPI("BPT")
maTable_SPI.creerTableSPI("Lancement")
maTable_SPI.creerTableSPI("Realisation")
maTable_SPI.creerTableSPI("TFX")


for element in os.listdir('data'):

	if element[0:15] == 'IXI_Qualif Négo':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("Qualif")
		maTable_SPI.insererDonneesSPI(element, "Qualif")


	if element[0:13] == 'IXI_Probation':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("Probation")
		maTable_SPI.insererDonneesSPI(element, "Probation")


	if element[0:21] == 'IXI_Réalisation étude':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("Etude")
		maTable_SPI.insererDonneesSPI(element, "Etude")


	if element[0:17] == 'IXI_Obtention BPT':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("BPT")
		maTable_SPI.insererDonneesSPI(element, "BPT")


	if element[0:21] == 'IXI_Lancement travaux':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("Lancement")
		maTable_SPI.insererDonneesSPI(element, "Lancement")


	if element[0:16] == 'IXI_Date Pose PB':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("Realisation")
		maTable_SPI.insererDonneesSPI(element, "Realisation")


	if element[0:7] == 'IXI_TFX':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableSPI("TFX")
		maTable_SPI.insererDonneesSPI(element, "TFX")

	if element[0:17] == 'IXI_Etat confirmé':
		maTable_SPI = Db_SPI.BaseSqlite()
		maTable_SPI.creerTableACT("ACT")
		maTable_SPI.insererDonneesACT(element, "ACT")


maTable_SPI = Db_SPI.BaseSqlite()

maTable_SPI.rechercher(TFX_1)
maTable_SPI.rechercher(TFX_2)
maTable_SPI.rechercher(TFX_3)
maTable_SPI.rechercher(TFX_4)
maTable_SPI.rechercher(TFX_5)
maTable_SPI.rechercher(TFX_6)

maTable_SPI.rechercher(Real_1)
maTable_SPI.rechercher(Real_2)
maTable_SPI.rechercher(Real_3)
maTable_SPI.rechercher(Real_4)
maTable_SPI.rechercher(Real_5)

maTable_SPI.rechercher(Lanc_1)
maTable_SPI.rechercher(Lanc_2)
maTable_SPI.rechercher(Lanc_3)
maTable_SPI.rechercher(Lanc_4)

maTable_SPI.rechercher(BPT_1)
maTable_SPI.rechercher(BPT_2)
maTable_SPI.rechercher(BPT_3)

maTable_SPI.rechercher(Etd_1)
maTable_SPI.rechercher(Etd_2)

maTable_SPI.rechercher(Prb_1)

maTable_SPI.rechercher(lance)
maTable_SPI.rechercher(BPT)
maTable_SPI.rechercher(etude)
maTable_SPI.rechercher(probation)
maTable_SPI.rechercher(qualif)

maTable_SPI.fermerBdd()
maTable_SPI = Db_SPI.BaseSqlite()

export = Export_Excel.ajoutClasseur("IXISPI_CEM")
maTable_SPI.rechercher(all)
a = maTable_SPI.resultat()
export.ajoutOnglet("IXISPI_CEM", "Cohérence IXISPI_CEM")
export.ecrireResultat(a, "IXISPI_CEM", "Cohérence IXISPI_CEM")

export.ajoutOngletSyn("IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_q)
b = maTable_SPI.resultat()[0][0]
export.ecrireVal("C4", b, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_q)
c = maTable_SPI.resultat()[0][0]
export.ecrireVal("D4", c, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_p)
d = maTable_SPI.resultat()[0][0]
export.ecrireVal("C5", d, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_p)
e = maTable_SPI.resultat()[0][0]
export.ecrireVal("D5", e, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_e)
f = maTable_SPI.resultat()[0][0]
export.ecrireVal("C6", f, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_e)
g = maTable_SPI.resultat()[0][0]
export.ecrireVal("D6", g, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_b)
h = maTable_SPI.resultat()[0][0]
export.ecrireVal("C7", h, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_b)
i = maTable_SPI.resultat()[0][0]
export.ecrireVal("D7", i, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_l)
j = maTable_SPI.resultat()[0][0]
export.ecrireVal("C8", j, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_l)
k = maTable_SPI.resultat()[0][0]
export.ecrireVal("D8", k, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_r)
l = maTable_SPI.resultat()[0][0]
export.ecrireVal("C9", l, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_r)
m = maTable_SPI.resultat()[0][0]
export.ecrireVal("D9", m, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_t)
n = maTable_SPI.resultat()[0][0]
export.ecrireVal("C10", n, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_t)
o = maTable_SPI.resultat()[0][0]
export.ecrireVal("D10", o, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_p_q)
p = maTable_SPI.resultat()[0][0]
export.ecrireVal("E5", p, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_p_q)
q = maTable_SPI.resultat()[0][0]
export.ecrireVal("F5", q, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_p_q)
r = maTable_SPI.resultat()[0][0]
export.ecrireVal("G5", r, "IXISPI_CEM", "Synthèse")

#---------------------------------------------------------

maTable_SPI.rechercher(nb_ano_e_q)
s = maTable_SPI.resultat()[0][0]
export.ecrireVal("G6", s, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_e_p)
t = maTable_SPI.resultat()[0][0]
export.ecrireVal("H6", t, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_e_q)
u = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_e_p)
v = maTable_SPI.resultat()[0][0]


export.ecrireVal("E6", s + t, "IXISPI_CEM", "Synthèse")
export.ecrireVal("F6", u + v, "IXISPI_CEM", "Synthèse")

#---------------------------------------------------------

maTable_SPI.rechercher(nb_ano_b_q)
s = maTable_SPI.resultat()[0][0]
export.ecrireVal("G7", s, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_b_p)
t = maTable_SPI.resultat()[0][0]
export.ecrireVal("H7", t, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_b_e)
u = maTable_SPI.resultat()[0][0]
export.ecrireVal("I7", u, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_b_q)
v = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_b_p)
w = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_b_e)
x = maTable_SPI.resultat()[0][0]


export.ecrireVal("E7", s + t + u, "IXISPI_CEM", "Synthèse")
export.ecrireVal("F7", v + w + x, "IXISPI_CEM", "Synthèse")

#---------------------------------------------------------

maTable_SPI.rechercher(nb_ano_l_q)
s = maTable_SPI.resultat()[0][0]
export.ecrireVal("G8", s, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_l_p)
t = maTable_SPI.resultat()[0][0]
export.ecrireVal("H8", t, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_l_e)
u = maTable_SPI.resultat()[0][0]
export.ecrireVal("I8", u, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_l_b)
v = maTable_SPI.resultat()[0][0]
export.ecrireVal("J8", v, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_l_q)
w = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_l_p)
x = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_l_e)
y = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_l_b)
z = maTable_SPI.resultat()[0][0]

export.ecrireVal("E8", s + t + u + v, "IXISPI_CEM", "Synthèse")
export.ecrireVal("F8", w + x + y + z, "IXISPI_CEM", "Synthèse")

#---------------------------------------------------------

maTable_SPI.rechercher(nb_ano_r_q)
s = maTable_SPI.resultat()[0][0]
export.ecrireVal("G9", s, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_p)
t = maTable_SPI.resultat()[0][0]
export.ecrireVal("H9", t, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_e)
u = maTable_SPI.resultat()[0][0]
export.ecrireVal("I9", u, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_b)
v = maTable_SPI.resultat()[0][0]
export.ecrireVal("J9", v, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_l)
w = maTable_SPI.resultat()[0][0]
export.ecrireVal("K9", w, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_r_q)
x = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_r_p)
y = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_r_e)
z = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_r_b)
aa = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_r_l)
ab = maTable_SPI.resultat()[0][0]


export.ecrireVal("E9", s + t + u + v + w , "IXISPI_CEM", "Synthèse")
export.ecrireVal("F9", x + y + z + aa + ab, "IXISPI_CEM", "Synthèse")

#---------------------------------------------------------

maTable_SPI.rechercher(nb_ano_t_q)
s = maTable_SPI.resultat()[0][0]
export.ecrireVal("G10", s, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_p)
t = maTable_SPI.resultat()[0][0]
export.ecrireVal("H10", t, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_e)
u = maTable_SPI.resultat()[0][0]
export.ecrireVal("I10", u, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_b)
v = maTable_SPI.resultat()[0][0]
export.ecrireVal("J10", v, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_l)
w = maTable_SPI.resultat()[0][0]
export.ecrireVal("K10", w, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_r)
x= maTable_SPI.resultat()[0][0]
export.ecrireVal("L10", x, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_t_q)
y = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_p)
z = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_e)
aa = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_b)
ab = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_l)
ac = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_r)
ad = maTable_SPI.resultat()[0][0]

export.ecrireVal("E10", s + t + u + v + w + x , "IXISPI_CEM", "Synthèse")
export.ecrireVal("F10", y + z + aa + ab + ac + ad, "IXISPI_CEM", "Synthèse")




maTable_SPI.rechercher(nb_e_s)
a = maTable_SPI.resultat()[0][0]
export.ecrireVal("C18", a, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_e_s)
b = maTable_SPI.resultat()[0][0]
export.ecrireVal("D18", b, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_l_s)
c = maTable_SPI.resultat()[0][0]
export.ecrireVal("C20", c, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_l_s)
d = maTable_SPI.resultat()[0][0]
export.ecrireVal("D20", d, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_l_e_s)
e = maTable_SPI.resultat()[0][0]
export.ecrireVal("E20", e, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_l_e_s)
e = maTable_SPI.resultat()[0][0]
export.ecrireVal("F20", e, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_l_e_s)
g = maTable_SPI.resultat()[0][0]
export.ecrireVal("I20", g, "IXISPI_CEM", "Synthèse")


maTable_SPI.rechercher(nb_r_s)
h = maTable_SPI.resultat()[0][0]
export.ecrireVal("C21", h, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_r_s)
i = maTable_SPI.resultat()[0][0]
export.ecrireVal("D21", i, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_e_s)
j = maTable_SPI.resultat()[0][0]
export.ecrireVal("I21", j, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_r_l_s)
k = maTable_SPI.resultat()[0][0]
export.ecrireVal("K21", k, "IXISPI_CEM", "Synthèse")

export.ecrireVal("E21", j + k, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_r_e_s)
l = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_r_l_s)
m = maTable_SPI.resultat()[0][0]

export.ecrireVal("F21", l + m, "IXISPI_CEM", "Synthèse")



maTable_SPI.rechercher(nb_t_s)
n = maTable_SPI.resultat()[0][0]
export.ecrireVal("C22", n, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_el_t_s)
o = maTable_SPI.resultat()[0][0]
export.ecrireVal("D22", o, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_e_s)
p = maTable_SPI.resultat()[0][0]
export.ecrireVal("I22", p, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_l_s)
q = maTable_SPI.resultat()[0][0]
export.ecrireVal("K22", q, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_t_r_s)
r = maTable_SPI.resultat()[0][0]
export.ecrireVal("L22", r, "IXISPI_CEM", "Synthèse")

export.ecrireVal("E22", p + q + r, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(nb_ano_el_t_e_s)
s = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_l_s)
t = maTable_SPI.resultat()[0][0]

maTable_SPI.rechercher(nb_ano_el_t_r_s)
u = maTable_SPI.resultat()[0][0]

export.ecrireVal("F22", s + t + u, "IXISPI_CEM", "Synthèse")

maTable_SPI.rechercher(plaque)
v = maTable_SPI.resultat()[0][0]
export.ecrireVal2("B2", v, "IXISPI_CEM", "Synthèse")
export.ecrireVal2("B14", v, "IXISPI_CEM", "Synthèse")

export.fermeture("IXISPI_CEM")

maTable_SPI.fermerBdd()

os.system("pause")