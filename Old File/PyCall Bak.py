# Aide-mÃ©moire des arguments Ã  appeller avec sys.argv[index]
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
fullDir = sys.argv[2]
fileName = sys.argv[3]
dateOfDay = sys.argv[4]
techEmail = sys.argv[5]

# Selon le mois, crée une variable qui écrit le mois au complet.

monthList = ["01 - Janvier", "02 - Février", "03 - Mars", "04 - Avril", "05 - Mai", "06 - Juin", "07 - Juillet", "08 - Août", "09 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Décembre"]
monthName = monthList[int(dateOfDay[5:7]) - 1]
year = dateOfDay[0:4]

# Affectation des dossiers et des stamps

formStamp = "Z:\\OneDrive\\_Logiciels\\stamp_Form.pdf"
textStamp = "Z:\\OneDrive\\_Logiciels\\stamp_text.pdf"
whiteMask = "Z:\\OneDrive\\_Logiciels\\whiteMask.png"
dossierAppel = "Appels de service"
cheminAppel = fullDir + "\\" + dossierAppel
tempFolder = fullDir + "\\" + os.path.splitext(fileName)[0]
dossierArchive = fullDir + "\\" + dossierAppel + "\\Archives" + "\\" + year + "\\" + monthName
dossierTermine = fullDir + "\\" + dossierAppel + "\\Terminé"
appelService = tempFolder + "\\appelservice.pdf"
pageDocument = tempFolder + "\\documents.pdf"
appelText = tempFolder + "\\Appel.pdf"

def stamp(inputFile, stampFile, outputFile):
    # Fonction pour utiliser la fonction Stamp de PDFtk
    # Prends comme arguements : Le fichier de base, le fichier "stamp" et l'emplacement du rÃ©sultat.
    subprocess.Popen(["pdftk", inputFile, "stamp", stampFile, "output", outputFile])
    
def fart(originalChar, newChar, fileName):
    f = open(tempFolder + "\\" + fileName,'r')
    filedata = f.read()
    f.close()

    newdata = filedata.replace(originalChar, newChar)

    f = open(tempFolder + "\\" + fileName, 'w')
    f.write(newdata)
    f.close()

#def appelLog(callNum, logPath):
 #   f = open(logPath,'a+')
  #  f.write(callNUm)
   # f.close()

# CrÃ©ation des fichiers de travails s'ils n'existent pas, Ã  inclure Ã©ventuellement
# Dans la section de traitement. (if do not exist, create and execute,
# if exist, just execute.
# Le traitement sera une fonction, pour faciliter l'utilisation.

if os.path.isdir(cheminAppel) == False : os.makedirs(cheminAppel)
if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
if os.path.isdir(dossierArchive) == False : os.makedirs(dossierArchive)
if os.path.isdir(dossierTermine) == False : os.makedirs(dossierTermine)

#
# Suggestion du consultant:
#
# Utilisation du GhostScript pour extraire le texte contenu dans l'appel de service PDF
# Si le fichier texte est vide (0 byte), donc appel de service NUMÉRISÉ, à traiter avec PDFTK / MAGICK / TESSERACT
# Si le fichier texte est non-vide, donc appel de service UNIX, traiter le contenu pour les informations du client
#
# Commande à utiliser:
# (la commande ci-dessous fonctionne avec GhostScript 9.19)
# "C:\Program Files (x86)\gs\gs9.19\bin\gswin32c.exe" -dSAFER -dBATCH -dNOPAUSE -dNOCACHE -dLastPage=1 -sDEVICE=txtwrite -sOutputFile=original-texte.txt original.pdf
# Le fichier original-texte.txt contiendra tout le texte extrait de l'appel de service venant de UNIX
#

#
# Début de traitement d'un appel de service PDF NUMÉRISÉ
#

# Division du fichier PDF en deux
subprocess.run(["pdftk", fullFile, "cat", "1", "output", appelService])
subprocess.run(["pdftk", fullFile, "cat", "2-end", "output", pageDocument])

# Conversion du pdf en png et application du masque, pour filtrer le OCR.
subprocess.run(["magick", "-density", "400", appelService, tempFolder + "\\img.png"])
subprocess.run(["magick", tempFolder + "\\img.png", whiteMask, "-composite", tempFolder + "\\mask.png"])

# Invocation de Tesseract OCR pour gÃ©nÃ©rer un fichier texte avec toutes les informations.
subprocess.run(["tesseract", tempFolder + "\\mask.png", tempFolder + "\\name", "-l", "meg"])

# Remplacement de tout les caractÃ¨res non autorisÃ©s.
fart("\\", "-", "name.txt")
fart("_", "-", "name.txt")
fart("/", "-", "name.txt")
fart(":", ";", "name.txt")
fart("*", " ", "name.txt")
fart("?", "!", "name.txt")
fart("\"", " ", "name.txt")
fart("<", "(", "name.txt")
fart(">", ")", "name.txt")
fart("|", " ", "name.txt")
fart("&", "ET", "name.txt")
fart("a", "-", "name.txt")
fart("Tel;", "Tel:", "name.txt")
fart("Te1;", "Tel:", "name.txt")
fart("â€”", "-", "name.txt")

# Modification du fichier afin de supprimer les lignes vides
f = open(tempFolder + "\\name.txt",'r')
filedata = f.read()
f.close()
filedata.upper()
newdata = ""
for i in range(0, len(filedata)):
    if filedata[i] == "0" and (filedata[i-1].isalpha() or filedata[i+1].isalpha()): # Si un 0 est entouré d'au moins une lettre, elle sera convertie en O
        newdata = newdata[0:i] + "O" 
    else:
        newdata = newdata[0:i] + filedata[i]
clearData = [line for line in newdata.split('\n') if line.strip() != '']
f = open(tempFolder + "\\name.txt",'w')
for i in range(0, len(clearData)):
    f.write(clearData[i] + "\n")
f.close()

# Aller chercher les lignes: num appel, nom entreprise, ville et adresse
f = open(tempFolder + "\\name.txt", "r")
filedata = f.read()
f.close()
myList = [line for line in filedata.split('\n') if line.strip() != '']
for i in range(0, len(myList)):
    if myList[i] == myList[1]:
        noAppel = myList[i]
    elif myList[i] == myList[5]:
        coName = myList[5]
    elif myList[i] == myList[6]:
        coAdresse = myList[6]
    elif myList[i] == myList[7]:
        coVille = myList[7]

#
# Fin de traitement d'un appel de service PDF NUMÉRISÉ
#

# Application du stamp de texte sur le document
subprocess.run(["pdftk", appelService, "stamp", textStamp, "output", appelText])

# Création du fichier pour google map
googleSearch = (((myList[6] + " " + myList[7]).replace(" ", "+")).replace(",", "")).replace("++", "+")

# Création du fichier de liste d'appel
if os.path.isfile(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt") == False:
    f = open(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt", "x")
    f.close()

# incrementation avec la boucle for des appels de service, ensuite renommage du fichier pour "date - incrementation - nom du client.pdf"
for i in range (1, 9999):
    if len(glob.glob(dossierTermine + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0:
        if len(glob.glob(dossierArchive + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0:
            if len(glob.glob(cheminAppel + "\\" + dateOfDay + " - " + str(i) + "*.pdf")) == 0:
                if os.path.isfile(pageDocument) == False:
                    subprocess.run(["pdftk", formStamp, "stamp", appelText, "output", cheminAppel + "\\" + dateOfDay + " - " + str(i) + " - " + coName + ".pdf"])
                    compteur = str(i)
                    # Création d'un fichier "ERROR" si le fichier "date - incrementation - nom du client.pdf" ne peut être créé
                    if os.path.isfile(cheminAppel + "\\" + dateOfDay + " - " + str(i) + " - " + coName + ".pdf") == False:
                        coName = ERROR
                        compteur = str(i)
                        subprocess.run(["pdftk", formStamp, "stamp", appelText, "output", cheminAppel + "\\" + dateOfDay + " - " + str(i) + " - " + coName + ".pdf"])
                    break
                subprocess.run(["pdftk", formStamp, "stamp", appelText, "output", tempFolder + "\\" + "Appel2.pdf"])
                subprocess.run(["pdftk", "A=" + tempFolder + "\\" + "Appel2.pdf", "B=" + pageDocument, "cat", "A", "B", "output", cheminAppel + "\\" + dateOfDay + " - " + str(i) + " - " + coName + ".pdf"])
                compteur = str(i)
                break

# Suppression des fichiers temporaire de travail
shutil.rmtree(tempFolder)
os.remove(fullFile)

# envoi d'un email (ou sms) a l'utilisateur qu'un appel de service lui a été envoyé
blatPath = "Z:\\OneDrive\\_Logiciels\\Blat\\software\\full"
sujet = "Appel de service"
expediteur = "service.alm@megaburo.ca"
smtp = "smtp.cgocable.ca"
message = "Appel # " + compteur + " - " + coName + " - " + coAdresse + " - " + coVille + " - https://www.google.ca/maps/search/" + googleSearch

subprocess.run(["blat", "-s", sujet, "-t", techEmail, "-f", expediteur, "-server", smtp, "-body", message])

with open(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt", "a") as myfile:
    myfile.write(noAppel + "\n")

