# Aide-mémoire des arguments à appeller avec sys.argv[index]
# [0] = Chemin du logiciel PyCall
# [1] = Chemin complet du fichier
# [2] = Chemin complet du dossier
# [3] = Nom du fichier
# [4] = Date du jour
# [5] = Addresse Courriel Technicien

import os
import subprocess
import glob
import shutil
import sys

# Affectation des arguements

fullFile = sys.argv[1]
#fullFile = "C:\\PyCall\\Monitoring\\Appels de service\\20160623131510.pdf"
fullDir = sys.argv[2]
#fullDir = "C:\\PyCall\\Monitoring\\Appels de service"
fileName = sys.argv[3]
#fileName = "20160623131510.pdf"
dateOfDay = sys.argv[4]
#dateOfDay = "2016-09-27"
techEmail = sys.argv[5]
#techEmail = "tomgir@outlook.com"

# Selon le mois, crée une variable qui écrit le mois au complet.

monthList = ["01 - Janvier", "02 - Fevrier", "03 - Mars", "04 - Avril", "05 - Mai", "06 - Juin", "07 - Juillet", "08 - Aout", "09 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Decembre"]
monthName = monthList[int(dateOfDay[5:7]) - 1]
year = dateOfDay[0:4]

# Affectation des dossiers et des stamps

tempFolder = fullDir + "\\.." + "\\" + os.path.splitext(fileName)[0]
dossierArchive = fullDir + "\\Archives" + "\\" + year + "\\" + monthName
dossierTermine = fullDir + "\\Termine"
appelService = tempFolder + "\\appelservice.pdf"
pageDocument = tempFolder + "\\documents.pdf"

appelEnCours = fullDir + "\\.." + "\\.." + "\\Appels en cours"

# Création des fichiers de travails s'ils n'existent pas, à inclure éventuellement

if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
if os.path.isdir(appelEnCours) == False : os.makedirs(appelEnCours)

# Cette fonction permet d'extraire les champs de texte dans un string donné :
# ex : << /V (MEGABURO ALMA) /T (Nom Entreprise)>>
# retournera : MEGABURO ALMA
def getName (givenLine):
        capture = False
        tempName = ""
        for i in range(0, len(givenLine)):
        
            if givenLine[i] == (")"):
                break
            if capture == True:
                tempName += givenLine[i]
            if givenLine[i] == ("("):
                capture = True
        return tempName.upper()

# Lecture du fichier texte associé à l'appel, si existant, puis crée les variables
# de travail à partir des informations obtenues
if os.path.isfile(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt"):
    f = open(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt", "r")
    filedata = f.read()
    f.close()
    myList = [line for line in filedata.split('\n')]
    noAppel = myList[0]
    coName = myList[1]
    coAdresse = myList[2]
    coVille = myList[3]
    googleSearch = myList[4]

# Si le fichier texte associé est inexistant, le logiciel traite l'appel comme si c'était un
# appel fait à la main. Il execute la fonction fdf de pdftk afin d'extraire les champs voulu,
# pour ensuite créer les variables associées.
else:
    
    # Utilise la fonction generate_fdf de pdftk pour extraire les champs de formulaire, ensuite le formate ligne par ligne
    # dans un fichier txt facilement recherchable
    subprocess.run(["pdftk", fullFile, "generate_fdf", "output", tempFolder + "\\" + os.path.splitext(fileName)[0] + ".fdf"])
    f = open(tempFolder + "\\" + os.path.splitext(fileName)[0] + ".fdf", "r")
    filedata = f.read()
    f.close()
    clearData = [line for line in filedata.split('\n') if line.strip() != '']
    
    f = open(tempFolder + "\\" + os.path.splitext(fileName)[0] + ".txt", "w")
    for i in range(0, len(clearData)):
        f.write(clearData[i].replace(">>", ">>\n"))
    f.close()

    f = open(tempFolder + "\\" + os.path.splitext(fileName)[0] + ".txt", "r")
    filedata = f.readlines()
    f.close()

    # Affectation des lignes contenant les champs voulu
    noAppel = str([s for s in filedata if "Appel" in s])
    coName = str([s for s in filedata if "Nom Entreprise" in s])
    coAdresse = str([s for s in filedata if "Rue" in s])
    coVille = str([s for s in filedata if "Ville" in s])

    # Utilise la fonction getName pour extraire seulement que l'information souhaitée
    # dans nos variables, puis crée le fichier de texte associé.
    noAppel = getName(noAppel)
    coName = getName(coName)
    coAdresse = getName(coAdresse)
    coVille = getName(coVille)
    googleSearch = "https://www.google.ca/maps/search/" + (((coAdresse + " " + coVille).replace(" ", "+")).replace(",", "")).replace("++", "+")

    f = open(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt", "w")
    f.write(noAppel + "\n")
    f.write(coName + "\n")
    f.write(coAdresse + "\n")
    f.write(coVille + "\n")
    f.write(googleSearch + "\n")
    f.close()

alreadyExist = False
# Incrémentation avec la boucle for des appels de service, ensuite renommage du fichier pour "date - incrementation - nom du client.pdf"
for i in range (1, 9999):
    if len(glob.glob(dossierTermine + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier terminé
        if len(glob.glob(dossierArchive + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier archive
            if dateOfDay + " - " + str(i) + " - " + coName +  " - " + noAppel + ".pdf" == fileName:
                alreadyExist = True
                break
            elif len(glob.glob(fullDir + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier principal
                compteur = str(i)
                os.rename(fullFile, fullDir + "\\" + dateOfDay + " - " + str(i) + " - " + coName +  " - " + noAppel + ".pdf")
                # Création d'un fichier "ERROR" si le fichier "date - incrementation - nom du client.pdf" ne peut être créé
                if os.path.isfile(fullDir + "\\" + dateOfDay + " - " + str(i) + " - " + coName +  " - " + noAppel + ".pdf") == False:
                    coName = "ERROR"
                    compteur = str(i)
                    os.rename(fullFile, fullDir + "\\" + dateOfDay + " - " + str(i) + " - " + coName +  " - " + noAppel + ".pdf")
                    os.rename(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt", tempFolder + "\\" + dateOfDay + " - " + str(i) + " - " + coName + ".txt")
                    break
                os.rename(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt", appelEnCours + "\\" + dateOfDay + " - " + str(i) + " - " + coName +  " - " + noAppel + ".txt")
                break
            
if not alreadyExist:
    # Suppression des fichiers temporaire de travail
    if os.path.isdir(tempFolder): shutil.rmtree(tempFolder)


    # envoi d'un email (ou sms) a l'utilisateur qu'un appel de service lui a été envoyé
    blatPath = "Z:\\OneDrive\\_Logiciels\\Blat\\software\\full"
    sujet = "Appel de service"
    expediteur = "service.alm@megaburo.ca"
    smtp = "smtp.cgocable.ca"
    message = "Appel # " + compteur + " - " + coName + " - " + coAdresse + " - " + coVille + " \n- " + googleSearch
    subprocess.run(["blat", "-s", sujet, "-t", techEmail, "-f", expediteur, "-server", smtp, "-body", message])

    # Ajout du numéro d'appel dans le fichier List d'appels.
    with open(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt", "a") as myfile:
        myfile.write(noAppel + " - " + coName + " - " + coAdresse + " - " + coName + "\n")

