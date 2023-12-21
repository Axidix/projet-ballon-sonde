"""Fichier principal du projet
Etablis la liaison série et écrit le fichier Excel en faisant appel au fichier "transfertDonnees.py" """

import serial
from transfertDonnees import *

sheet, wb = fichierExcel()

with serial.Serial(port="COM2", baudrate=9600, timeout=1, writeTimeout=1) as liaison:
    if liaison.isOpen():
        while True:
            trame = liaison.readline() #.decode("utf-8") ?
            print(trame)
            donnees = formatage(trame, ballons)
            transfert_donnees(donnees, sheet, wb)