# -*- coding: utf-8 -*-
import xlrd
import sqlite3

class BaseSqlite :

	def __init__(self):

		self.con = sqlite3.connect("ma_base.db")
		self.cursor = self.con.cursor() 

	def creerTableSPI(self, table):

		self.table = table
		requeteTable = "CREATE TABLE " + table + "(CODE_IMB Varchar2(20), ID_Chantier Number, NB_EL Number, Cle_Site Varchar2(20), Statut_Syndic Varchar2(20), Action Varchar2(20), \
		Realise Varchar2(20), Lance Varchar2(20), BPT Varchar2(20), Etude Varchar2(20), Probation Varchar2(20), Qualif Varchar2(20))"
		
		self.cursor.execute("Drop Table IF EXISTS " + table)
		self.cursor.execute(requeteTable)

	def creerTableACT(self, table):

		self.table = table
		requeteTable = "CREATE TABLE " + table + "(CODE_IMB Varchar2(20), ID_Chantier Number, NB_EL Number, Code_Plaque Varchar2(20), Statut_act_20 Varchar2(20), Qualif Varchar2(20), \
		Probation Varchar2(20), BPT Varchar2(20), Statut_act_40 Varchar2(20), Etude Varchar2(20), Lancement Varchar2(20), Realisation Varchar2(20), TFX Varchar2(20))"
		
		self.cursor.execute("Drop Table IF EXISTS " + table)
		self.cursor.execute(requeteTable)


	def insererDonneesSPI(self, fichier, table) :

		self.fichier = fichier
		self.table = table

		try:                      
			book = xlrd.open_workbook('data/' + fichier , encoding_override="utf-8")
			sheet = book.sheet_by_index(0)
		  
			for row in range (5, sheet.nrows) :
				
				liste = (sheet.row_values(row))
				
				ma_liste = [liste[1], liste[2], liste[5], liste[11], liste[12], "OK", "OK", "OK", "OK", "OK", "OK"]
				ma_liste.append(table)
				ma_liste2 = (ma_liste[0], ma_liste[1], ma_liste[2], ma_liste[3], ma_liste[4], ma_liste[11], "OK", "OK", "OK", "OK", "OK", "OK")

				tup = tuple(ma_liste2)

				self.cursor.execute(("Insert into " +  table + " values (?,?,?,?,?,?,?,?,?,?,?,?)"), tup)

			print("Le fichier : " + fichier + " a été chargé dans la table " + table)
			print("-----------------------------------------------------------------------------------------------------------------------------")

		except FileNotFoundError:

			print("Le fichier : optimum.xlsx n'est pas disponible")
			print("-----------------------------------------------------------------------------------------------------------------------------")

		except sqlite3.ProgrammingError:

			print("Le fichier : optimum.xlsx est incorrect")
			print("-----------------------------------------------------------------------------------------------------------------------------")

		self.con.commit()

		self.fermerBdd()


	def insererDonneesACT(self, fichier, table) :

		self.fichier = fichier
		self.table = table

		try:                      
			book = xlrd.open_workbook('data/' + fichier , encoding_override="utf-8")
			sheet = book.sheet_by_index(0)
		  
			for row in range (5, sheet.nrows) :
				
				liste = (sheet.row_values(row))
				
				ma_liste = [liste[1], liste[2], liste[3], liste[4], liste[5], liste[6], liste[7], liste[8], liste[9], liste[10], liste[11], liste[12], liste[13]]

				tup = tuple(ma_liste)

				self.cursor.execute(("Insert into " +  table + " values (?,?,?,?,?,?,?,?,?,?,?,?,?)"), tup)

			print("Le fichier : " + fichier + " a été chargé dans la table " + table)
			print("-----------------------------------------------------------------------------------------------------------------------------")

		except FileNotFoundError :

			print("Le fichier : optimum.xlsx n'est pas disponible")
			print("-----------------------------------------------------------------------------------------------------------------------------")

		except sqlite3.ProgrammingError :

			print("Le fichier : optimum.xlsx est incorrect")
			print("-----------------------------------------------------------------------------------------------------------------------------")

		self.con.commit()

		self.fermerBdd()


	def rechercher(self, requete):
					   
		self.cursor.execute(requete)
				
	def resultat(self):
		
		return self.cursor.fetchall()	

	def fermerBdd(self):
		self.con.commit()
		self.cursor.close()
		self.con.close()