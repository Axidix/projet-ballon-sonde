"""Fonctions permettant d'écrire un fichier Excel contenant les données reçues par trames"""

from datetime import datetime
import xlwt

def formatage(trame):
    """Fonction transformant la trame en un tableau de données"""

    #Conservation des données utiles
    elem = trame.split(", ")
    elem[0] = elem[0].strip("$ ")
    elem[6] = elem[6].split()[0]
    elem.append(str(datetime.now().time())[0:8])  #Ajout de l'heure tronquée aux secondes

    return(elem)        #Renvoi du tableau de données




def fichierExcel():
    """Fonction créant un fichier Excel adapté pour les données"""

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

    return(sheet, wb)     #Renvoi du fichier créé



def transfert_donnees(elem, sheet, wb):
    """Fonction inscrivant les données sur le ficheir Excel à partir d'un tableau de données"""

    for i in range(nbTrame):        #Ecriture des données dans les cellules du fichier Excel pour chaque trame

        sheet.write(i+1, 0, elem[7])
        sheet.write(i+1, 1, elem[0])
        sheet.write(i+1, 2, elem[1])
        sheet.write(i+1, 3, elem[2])
        sheet.write(i+1, 4, elem[3])
        sheet.write(i+1, 5, elem[4])
        sheet.write(i+1, 6, elem[5])
        sheet.write(i+1, 7, elem[6])

        wb.save(r"C:\Users\adib4\OneDrive\Documents\Formations\Programmation\Projet Ballon Sonde\donnees_ballon_sonde.xls")     #Sauvegarde du fichier