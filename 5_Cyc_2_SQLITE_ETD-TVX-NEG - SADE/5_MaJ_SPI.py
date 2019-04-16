# -*- coding: utf-8 -*-
import sqlite3
import xlrd
import bddSqlite
import os

import os
os.system("@echo off")
os.system("@echo 						###############################################")
os.system("@echo 						#                                             #")
os.system("@echo 						#            Mise à jour données DPL          #")
os.system("@echo 						#                 Migration SPI               #")
os.system("@echo 						#                                             #")
os.system("@echo 						###############################################")

for element in os.listdir('data'):

	if element[0:9] == 'IXI_C_IMB':

		book = xlrd.open_workbook('data/' + element, encoding_override="utf-8")
		sheet = book.sheet_by_index(0)

		ETD = bddSqlite.BaseSqlite()
		print('mise à jour du fichier ' + element)
		
		for row in range (5, sheet.nrows) :
			liste = (sheet.row_values(row))
			
			ETD.rechercher("Update C set SPI = 'Oui'"  + " where code_IMB = " + "'" + liste[1] + "'")

		ETD.fermerBdd()

os.system("pause")
