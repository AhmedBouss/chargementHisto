# -*- coding: utf-8 -*-
import sqlite3
import xlrd
from mesRequetes import *
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






   
 def insererDonnees(self, table, indice, donnees) :

  self.table = table
  self.indice = indice
  self.donnees = donnees

  for fichier in os.listdir('data'):

    if fichier[0:4] == 'Init': 

      book = xlrd.open_workbook('data/' + fichier, encoding_override="utf-8")
      sheet = book.sheet_by_index(indice)

      for row in range (5, sheet.nrows) :
        liste = (sheet.row_values(row))

        ext = ['']  # Colonne supplémentaires pour vérification des conditions
        liste.extend(ext)
        del liste[0]
        liste = tuple(liste)
        
        self.cursor.execute(("Insert into "  + table + " values (" + donnees + ")" ), liste)



      
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
  
  