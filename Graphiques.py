import matplotlib.pyplot as plt
import xlrd


donnees = xlrd.open_workbook(r"C:\Users\adib4\OneDrive\Documents\Formations\Programmation\Projet Ballon Sonde\donnees_ballon_sonde.xls")
sheet = donnees.sheet_by_name("Mesures")


mesures_ballons = {}
mesures = [""]*7
nonVide = True
nLigne = 1

while nonVide:

    try:
        mesures[0] = sheet.cell_value(nLigne, 0)
        for i in range(6):
            mesures[i+1] = sheet.cell_value(nLigne, i+2)

        try:
            mesures_ballons[sheet.cell_value(nLigne, 1)] += [list(mesures)]

        except:
            mesures_ballons[sheet.cell_value(nLigne, 1)] = [list(mesures)]

        nLigne += 1

    except:
        nonVide = False



premDate = sheet.cell_value(1,0).split(":")
refHoraire = int(premDate[0])*60 + int(premDate[1]) + int(premDate[2])/60



tabMesures = []     #Tableau contenant les mesures de tous les ballons oraganisées par nature des mesures
dicoMesures = {"Température": [], "Pression": [], "Latitude": [], "Longitude": [], "Altitude": [], "Humidité": [], "Date": [], "Référence": ""}
k=0

for cle, valeur in mesures_ballons.items():
    tabMesures.append(dicoMesures)
    tabMesures[k]["Référence"] = cle

    for i in range(len(valeur)):
        date = valeur[i][0].split(":")
        horaire = int(date[0])*60 + int(date[1]) + int(date[2])/60 - refHoraire

        tabMesures[k]["Date"] += [horaire]
        tabMesures[k]["Température"].append(valeur[i][1])
        tabMesures[k]["Pression"].append(valeur[i][2])
        tabMesures[k]["Longitude"].append(valeur[i][6])
        tabMesures[k]["Altitude"].append(valeur[i][4])
        tabMesures[k]["Humidité"].append(valeur[i][3])
        tabMesures[k]["Latitude"].append(valeur[i][5])
    k += 1



#Affichage des mesures

plt.subplot(321)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Température"])
plt.xlabel("Date")
plt.ylabel("Température")


plt.subplot(322)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Pression"])
plt.xlabel("Date")
plt.ylabel("Pression")

plt.subplot(323)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Humidité"])
plt.xlabel("Date")
plt.ylabel("Humidité")

plt.subplot(324)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Altitude"])
plt.xlabel("Date")
plt.ylabel("Altitude")

plt.subplot(325)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Latitude"])
plt.xlabel("Date")
plt.ylabel("Latitude")

plt.subplot(326)
plt.plot(tabMesures[0]["Date"], tabMesures[0]["Longitude"])
plt.xlabel("Date")
plt.ylabel("Longitude")


plt.show()
