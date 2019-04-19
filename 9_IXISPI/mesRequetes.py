# -*- coding: utf-8 -*-

lance =     "update Lancement set realise = 'SO'"
BPT =       "update BPT set realise = 'SO', lance = 'SO'"
etude =     "update Etude set realise = 'SO', lance = 'SO', BPT = 'SO'"
probation = "update Probation set realise = 'SO', lance = 'SO', BPT = 'SO', Etude = 'SO'"
qualif =    "update Qualif set realise = 'SO', lance = 'SO', BPT = 'SO', Etude = 'SO', Probation = 'SO'"

TFX_1 = "update TFX set Realise = 'NOK' where code_IMB not in (select code_IMB from Realisation)"
TFX_2 = "update TFX set Lance = 'NOK' where code_IMB not in (select code_IMB from Lancement)"
TFX_3 = "update TFX set BPT = 'NOK' where code_IMB not in (select code_IMB from BPT)"
TFX_4 = "update TFX set Etude = 'NOK' where code_IMB not in (select code_IMB from Etude)"
TFX_5 = "update TFX set Probation = 'NOK' where code_IMB not in (select code_IMB from Probation)"
TFX_6 = "update TFX set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

Real_1 = "update Realisation set Lance = 'NOK' where code_IMB not in (select code_IMB from Lancement)"
Real_2 = "update Realisation set BPT = 'NOK' where code_IMB not in (select code_IMB from BPT)"
Real_3 = "update Realisation set Etude = 'NOK' where code_IMB not in (select code_IMB from Etude)"
Real_4 = "update Realisation set Probation = 'NOK' where code_IMB not in (select code_IMB from Probation)"
Real_5 = "update Realisation set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

Lanc_1 = "update Lancement set BPT = 'NOK' where code_IMB not in (select code_IMB from BPT)"
Lanc_2 = "update Lancement set Etude = 'NOK' where code_IMB not in (select code_IMB from Etude)"
Lanc_3 = "update Lancement set Probation = 'NOK' where code_IMB not in (select code_IMB from Probation)"
Lanc_4 = "update Lancement set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

BPT_1 = "update BPT set Etude = 'NOK' where code_IMB not in (select code_IMB from Etude)"
BPT_2 = "update BPT set Probation = 'NOK' where code_IMB not in (select code_IMB from Probation)"
BPT_3 = "update BPT set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

Etd_1 = "update Etude set Probation = 'NOK' where code_IMB not in (select code_IMB from Probation)"
Etd_2 = "update Etude set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

Prb_1 = "update Probation set Qualif = 'NOK' where code_IMB not in (select code_IMB from Qualif)"

all = "select * from TFX union all \
select * from Realisation where code_IMB not in (select code_IMB from TFX) union all  \
select * from Lancement where code_IMB not in (select code_IMB from Realisation union all select code_IMB from TFX) union all  \
select * from BPT where code_IMB not in (select code_IMB from Lancement union all select code_IMB from Realisation union all select code_IMB from TFX) union all \
select * from Etude where code_IMB not in (select code_IMB from BPT union all select code_IMB from Lancement union all select code_IMB from Realisation union all select code_IMB from TFX) union all  \
select * from Probation where code_IMB not in (select code_IMB from Etude union all select code_IMB from BPT union all select code_IMB from Lancement union all select code_IMB from Realisation union all select code_IMB from TFX) union all  \
select * from Qualif where code_IMB not in (select code_IMB from Probation union all select code_IMB from Etude union all select code_IMB from BPT union all select code_IMB from Lancement union all select code_IMB from Realisation union all select code_IMB from TFX)"


nb_q = "select count(*) from Qualif where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_q = "select ifnull(sum(nb_el), 0) from Qualif where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_p = "select count(*) from Probation where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_p = "select ifnull(sum(nb_el), 0) from Probation where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_p_q = "select count(*) from Probation where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_p_q = "select ifnull(sum(nb_el), 0) from Probation where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_e = "select count(*) from Etude where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_e = "select ifnull(sum(nb_el), 0) from Etude where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_e_q = "select count(*) from Etude where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_e_q =  "select ifnull(sum(nb_el), 0) from Etude where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_e_p = "select count(*) from Etude where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_e_p = "select ifnull(sum(nb_el), 0) from Etude where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_b = "select count(*) from BPT where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_b = "select ifnull(sum(nb_el), 0) from BPT where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_b_q = "select count(*) from BPT where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_b_q = "select ifnull(sum(nb_el), 0) from BPT where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_b_p = "select count(*) from BPT where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_b_p =  "select ifnull(sum(nb_el), 0) from BPT where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_b_e = "select count(*) from BPT where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_b_e = "select ifnull(sum(nb_el), 0) from BPT where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_l = "select count(*) from Lancement where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_l = "select ifnull(sum(nb_el), 0) from Lancement where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_l_q = "select count(*) from Lancement where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_l_q =  "select ifnull(sum(nb_el), 0) from Lancement where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_l_p = "select count(*) from Lancement where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_l_p =  "select ifnull(sum(nb_el), 0) from Lancement where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_l_e = "select count(*) from Lancement where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_l_e =  "select ifnull(sum(nb_el), 0) from Lancement where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_l_b = "select count(*) from Lancement where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_l_b =  "select ifnull(sum(nb_el), 0) from Lancement where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_r = "select count(*) from Realisation where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_r = "select ifnull(sum(nb_el), 0) from Realisation where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_r_q = "select count(*) from Realisation where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_r_q =  "select ifnull(sum(nb_el), 0) from Realisation where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_r_p = "select count(*) from Realisation where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_r_p =  "select ifnull(sum(nb_el), 0) from Realisation where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_r_e = "select count(*) from Realisation where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_r_e =  "select ifnull(sum(nb_el), 0) from Realisation where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_r_b = "select count(*) from Realisation where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_r_b =  "select ifnull(sum(nb_el), 0) from Realisation where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_r_l = "select count(*) from Realisation where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_r_l =  "select ifnull(sum(nb_el), 0) from Realisation where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

nb_t = "select count(*) from TFX where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_el_t = "select ifnull(sum(nb_el), 0) from TFX where code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_q = "select count(*) from TFX where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_q =  "select ifnull(sum(nb_el), 0) from TFX where Qualif = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_p = "select count(*) from TFX where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_p =  "select ifnull(sum(nb_el), 0) from TFX where Probation = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_e = "select count(*) from TFX where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_e =  "select ifnull(sum(nb_el), 0) from TFX where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_b = "select count(*) from TFX where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_b =  "select ifnull(sum(nb_el), 0) from TFX where BPT = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_l = "select count(*) from TFX where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_l =  "select ifnull(sum(nb_el), 0) from TFX where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_t_r = "select count(*) from TFX where Realise = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"
nb_ano_el_t_r =  "select ifnull(sum(nb_el), 0) from TFX where Realise = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 <> 'Sans activité 0020')"

# ---------------------------------------------------------------------------------------------------------------------------------------------------

nb_e_s = "select count(*) from Etude where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_el_e_s = "select ifnull(sum(nb_el), 0) from Etude where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"

nb_l_s = "select count(*) from Lancement where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_el_l_s = "select ifnull(sum(nb_el), 0) from Lancement where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_l_e_s = "select count(*) from Lancement where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_l_e_s =  "select ifnull(sum(nb_el), 0) from Lancement where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"

nb_r_s = "select count(*) from Realisation where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_el_r_s = "select ifnull(sum(nb_el), 0) from Realisation where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_r_e_s = "select count(*) from Realisation where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_r_e_s =  "select ifnull(sum(nb_el), 0) from Realisation where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_r_l_s = "select count(*) from Realisation where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_r_l_s =  "select ifnull(sum(nb_el), 0) from Realisation where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"


nb_t_s = "select count(*) from TFX where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_el_t_s = "select ifnull(sum(nb_el), 0) from TFX where code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_t_e_s = "select count(*) from TFX where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_t_e_s =  "select ifnull(sum(nb_el), 0) from TFX where Etude = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_t_l_s = "select count(*) from TFX where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_t_l_s =  "select ifnull(sum(nb_el), 0) from TFX where Lance = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_t_r_s = "select count(*) from TFX where Realise = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"
nb_ano_el_t_r_s =  "select ifnull(sum(nb_el), 0) from TFX where Realise = 'NOK' and code_imb in (select code_IMB from ACT where Statut_act_20 = 'Sans activité 0020')"


plaque = "select code_plaque from ACT limit 1"