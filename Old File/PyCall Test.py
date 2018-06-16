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
#fullFile = "C:\\PyCall\\Monitoring\\20160623131510.pdf"
fullDir = sys.argv[2]
#fullDir = "C:\\PyCall\\Monitoring"
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

formStamp = "Z:\\OneDrive\\_Logiciels\\stamp_Form.pdf"
textStamp = "Z:\\OneDrive\\_Logiciels\\stamp_text.pdf"
whiteMask = "Z:\\OneDrive\\_Logiciels\\whiteMask.png"
dossierAppel = "Appels de service"
cheminAppel = fullDir + "\\" + dossierAppel
tempFolder = fullDir + "\\" + os.path.splitext(fileName)[0]
dossierArchive = fullDir + "\\" + dossierAppel + "\\Archives" + "\\" + year + "\\" + monthName
dossierTermine = fullDir + "\\" + dossierAppel + "\\Termine"
appelService = tempFolder + "\\appelservice.pdf"
pageDocument = tempFolder + "\\documents.pdf"
appelText = tempFolder + "\\Appel.pdf"
appelEnCours = fullDir + "\\.." + "\\Appels en cours"

def stamp(inputFile, stampFile, outputFile):
    # Fonction pour utiliser la fonction Stamp de PDFtk
    # Prends comme arguements : Le fichier de base, le fichier "stamp" et l'emplacement du résultat.
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
  #  f.write(callNum)
   # f.close()

# Création des fichiers de travails s'ils n'existent pas, à inclure éventuellement

if os.path.isdir(cheminAppel) == False : os.makedirs(cheminAppel)
if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
if os.path.isdir(dossierArchive) == False : os.makedirs(dossierArchive)
if os.path.isdir(dossierTermine) == False : os.makedirs(dossierTermine)
if os.path.isdir(appelEnCours) == False : os.makedirs(appelEnCours)

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

subprocess.run(["Z:\\OneDrive\\_Logiciels\\gs9.19\\bin\\gswin32c.exe", "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-dLastPage=1", "-sDEVICE=txtwrite", "-sOutputFile=" + tempFolder + "\\original-texte.txt", fullFile])

if os.stat(tempFolder + "\\original-texte.txt").st_size == 0:
    # Division du fichier PDF en deux
    subprocess.run(["pdftk", fullFile, "cat", "1", "output", appelService])
    subprocess.run(["pdftk", fullFile, "cat", "2-end", "output", pageDocument])

    # Conversion du pdf en png et application du masque, pour filtrer le OCR.
    subprocess.run(["magick", "-density", "400", appelService, tempFolder + "\\img.png"])
    subprocess.run(["magick", tempFolder + "\\img.png", whiteMask, "-composite", tempFolder + "\\mask.png"])

    # Invocation de Tesseract OCR pour générer un fichier texte avec toutes les informations.
    subprocess.run(["tesseract", tempFolder + "\\mask.png", tempFolder + "\\name", "-l", "meg"])

    # Remplacement de tout les caractères non autorisÃ©s.
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

#
# Traitement d'un appel UNIX
#

else:
    f = open(tempFolder + "\\original-texte.txt", "r")
    filedata = f.read()
    f.close()
    myList = [line for line in filedata.split('\n') if line.strip() != '']
    for i in range(0, len(myList)):
        if myList[i] == myList[1]:
            noAppel = myList[i][70:].strip()
        elif myList[i] == myList[10]:
            coName = myList[i][49:].strip()
        elif myList[i] == myList[11]:
            coAdresse = myList[i][49:].strip()
        elif myList[i] == myList[12]:
            coVille = myList[i][49:73].strip()


#
# Fin du traitement d'un appel UNIX
#

# Application du stamp de texte sur le document
subprocess.run(["pdftk", fullFile, "stamp", textStamp, "output", appelText])

# Création du fichier pour google map
googleSearch = "https://www.google.ca/maps/search/" + (((coAdresse + " " + coVille).replace(" ", "+")).replace(",", "")).replace("++", "+")

# Création du fichier de liste d'appel
if os.path.isfile(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt") == False:
    f = open(dossierArchive + "\\_Liste d'appels du " + dateOfDay + ".txt", "x")
    f.close()

# Création du fichier TXT contenant toutes les informations associées à l'appel
f = open(appelEnCours + "\\" + os.path.splitext(fileName)[0] + ".txt",'w')
f.write(noAppel + "\n")
f.write(coName + "\n")
f.write(coAdresse + "\n")
f.write(coVille + "\n")
f.write(googleSearch + "\n")
f.close()

# Application du stamp et déplacer dans le fichier du technicien, prêt à être renommé par PyRen
if os.path.isfile(pageDocument) == False:
    subprocess.run(["pdftk", formStamp, "stamp", appelText, "output", cheminAppel + "\\" + os.path.splitext(fileName)[0] + ".pdf"])
else:
    subprocess.run(["pdftk", formStamp, "stamp", appelText, "output", tempFolder + "\\" + "Appel2.pdf"])
    subprocess.run(["pdftk", "A=" + tempFolder + "\\" + "Appel2.pdf", "B=" + pageDocument, "cat", "A", "B", "output", cheminAppel + "\\" + os.path.splitext(fileName)[0] + ".pdf"])

# Suppression du fichier original
os.remove(fullFile)
