#Décodage de la trame --- outdated

from datetime import datetime
import xlwt
from transfertDonnees import *

heure = str(datetime.now().time())[0:8]  #Heure tronquée aux secondes

nbTrame = 2
trame = ["$ 1, 26.5, 1020, 41, 2000, 30, 45 * Checksum", "$ 1, 26.5, 1020, 40, 2100, 31, 43"]  #Trame d'essai


#Transfert sur Excel

wb = xlwt.Workbook()
sheet = wb.add_sheet("Mesures")

#Mise en place des colonnes de données
sheet.write(0, 0, "Heure")
sheet.write(0, 1, "N° Ballon")
sheet.write(0, 2, "Température (°C)")
sheet.write(0, 3, "Pression (hPA)")
sheet.write(0, 4, "Humidité (%)")
sheet.write(0, 5, "Altitude (m)")
sheet.write(0, 6, "Latitude (m)")
sheet.write(0, 7, "Longitude (m)")

for i in range(nbTrame):       #Ecriture des données dans les cellules du fichier Excel pour chaque trame
    elem = formatage(trame[i])
    sheet.write(i+1, 0, heure)
    sheet.write(i+1, 1, elem[0])
    sheet.write(i+1, 2, elem[1])
    sheet.write(i+1, 3, elem[2])
    sheet.write(i+1, 4, elem[3])
    sheet.write(i+1, 5, elem[4])
    sheet.write(i+1, 6, elem[5])
    sheet.write(i+1, 7, elem[6])


wb.save(r"C:\Users\adib4\OneDrive\Documents\Formations\Programmation\Projet Ballon Sonde\donnees_ballon_sonde.xls")     #Sauvegarde du fichier
