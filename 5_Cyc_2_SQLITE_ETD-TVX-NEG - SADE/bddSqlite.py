# -*- coding: utf-8 -*-
import sqlite3
import xlrd
from mesRequetesNEG import *
import os


class BaseSqlite :

 def __init__(self):

  self.con = sqlite3.connect("ma_base.db")
  self.cursor = self.con.cursor() 
   
 def creerTable(self, table, request):

  self.request = request
  self.table = table
  requeteTable = "CREATE TABLE " + table + " (" + request +")"
  
  self.cursor.execute("Drop Table IF EXISTS " + table)
  self.cursor.execute(requeteTable)

 def creerTableInt(self, table, request):

  self.request = request
  self.table = table
  requeteTable = "CREATE TABLE " + table + " as " + request 
  
  self.cursor.execute("Drop Table IF EXISTS " + table)
  self.cursor.execute(requeteTable)

 def creerTableIXI(self):

  requeteTable = "CREATE TABLE IXI (Code_IMB CHAR, Adresse CHAR, Localite CHAR, Plaque CHAR, NB_EL INT, PA CHAR, PM CHAR, PB CHAR, CS CHAR, CRS CHAR)" 
  
  self.cursor.execute("Drop Table IF EXISTS IXI")
  self.cursor.execute(requeteTable)

   
 def insererDonneesOPT(self) :

  for fichier in os.listdir('data'):

    if fichier[0:4] == 'DPL_': 

      book = xlrd.open_workbook('data/' + fichier, encoding_override="utf-8")
      sheet = book.sheet_by_index(0)

      for row in range (5, sheet.nrows) :
        liste = (sheet.row_values(row))

        ext = ['', '2-Non Affecté','', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', 'Non', '','PAV', '','','','','', '', '', 'Non', '', 'Oui']  # Colonne supplémentaires pour vérification des conditions
        liste.extend(ext)
        del liste[0]
        liste = tuple(liste)
        
        self.cursor.execute(("Insert into OPT values (" + tup + ")" ), liste)


 def insererDonneesIXI(self) :

  for fichier in os.listdir('data'):

    if fichier[0:10] == 'Chargement': 

      book = xlrd.open_workbook('data/' + fichier, encoding_override="utf-8")
      sheet = book.sheet_by_index(0)

      for row in range (5, sheet.nrows) :
        liste = (sheet.row_values(row))

        del liste[0]
        liste = tuple(liste)
        
        self.cursor.execute(("Insert into IXI values (?,?,?,?,?,?,?,?,?,?)" ), liste)

      
 def rechercher(self, requete) :
        
  self.cursor.execute(requete)
  
  
 def resultat(self) :
  return self.cursor.fetchall()

 def update(self, champ, valeur) :
  self.champ = champ
  self.valeur = valeur

  return self.cursor.execute("update OPT set " + champ +  " = ''" +  valeur + " where " + champ + " = 'Optionnelle'")
  self.con.commit()

 def updateT(self, table, champ, champO, valeur, valeurO) :
  self.champ = champ
  self.champO = champO
  self.valeur = valeur
  self.valeurO = valeurO
  self.table = table


  return self.cursor.execute("update " + table + " set " + champ +  " = " +  "'" + valeur + "'" + " where " + champO +  " = " +  "'" + valeurO + "'")
  self.con.commit()
   
 def fermerBdd(self):

  self.con.commit()
  self.cursor.close()
  self.con.close()
  
  