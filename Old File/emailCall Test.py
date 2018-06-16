# Aide-mémoire des arguments à  appeller avec sys.argv[index]
# [0] = Chemin du logiciel PyCall
# [1] = Chemin complet du fichier
# [2] = Chemin complet du dossier
# [3] = Nom du fichier
# [4] = Nom du tech
# [5] = True si retour, False si non

import os
import subprocess
import glob
import shutil
import sys

# Affectation des arguements

fullFile = sys.argv[1]
#fullFile = "C:\\PyCall\\Monitoring\\Appels de service\\Termine\\2016-09-27 - 1 - REMORQUAGE S.0.S SAGUENAY - 427223.pdf"
fullDir = sys.argv[2]
#fullDir = "C:\\PyCall\\Monitoring\\Appels de service\\Termine"
fileName = sys.argv[3]
#fileName = "2016-09-27 - 1 - REMORQUAGE S.0.S SAGUENAY - 427223.pdf"
techName = sys.argv[4]
#techName = "Tommy"
retour = sys.argv[5]
#retour = "False"

monthList = ["01 - Janvier", "02 - Fevrier", "03 - Mars", "04 - Avril", "05 - Mai", "06 - Juin", "07 - Juillet", "08 - Aout", "09 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Decembre"]
monthName = monthList[int(fileName[5:7]) - 1]
year = fileName[0:4]

ghostScriptPath = "Z:\\OneDrive\\_Logiciels\\gs9.19\\bin\\gswin32c.exe"
archivePath = fullDir + "\\" + ".." + "\\Archives" + "\\" + year + "\\" + monthName

if retour == "Termine":
    sujet = "Appel de service termine de " + techName
elif retour == "Retour":
    sujet = "***Retour requis*** Appel de service termine de " + techName
#destinataires = "jaclar@megaburo.ca,sabfil@megaburo.ca"
destinataires = "claduf@megaburo.ca" 
expediteur = "service.alm@megaburo.ca"
smtp = "smtp.cgocable.ca"
if retour == "Termine":
    message = "Voici l'appel de service " + fileName +  " termine par " + techName
elif retour == "Retour":
    message = "***Retour requis*** Voici l'appel de service " + fileName +  " termine par " + techName
    
#
# Suggestion du consultant:
# PDFTK avec l'option FLATTEN pour "aplatir" ton PDF avec les données du formulaire
# (la commande ci-dessous fonctionne avec GhostScript 9.19)
# "C:\Program Files (x86)\gs\gs9.19\bin\gswin32c.exe" -dSAFER -dBATCH -dNOPAUSE -dNOCACHE -sDEVICE=pdfwrite -sColorConversionStrategy=/LeaveColorUnchanged -dAutoFilterColorImages=true -dAutoFilterGrayImages=true -dDownsampleMonoImages=true -dDownsampleGrayImages=true -dDownsampleColorImages=true -sOutputFile=original_flattened.pdf original.pdf
#

subprocess.run(["blat", "-s", sujet, "-t", destinataires, "-f", expediteur, "-server", smtp, "-body", message, "-attach", fullFile])


if os.path.isdir(archivePath) == False : os.makedirs(archivePath)

# Vérifie si le fichier existe, si oui, le crée avec une incrémentation (i)
if os.path.isfile(archivePath + "\\" + fileName) == False:
    #subprocess.run([ghostScriptPath, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-sDEVICE=pdfwrite", "-sColorConversionStrategy=/LeaveColorUnchangedhanged", "-dAutoFilterColorImages=true", "-dAutoFilterGrayImages=true", "-dDownsampleMonoImages=true", "-dDownsampleGrayImages=true", "-dDownsampleColorImages=true", "-sOutputFile=" + fullFile, archivePath + "\\" + fileName])
    shutil.move(fullFile, archivePath)
else:
    fileNameRen = fileName
    i = 1
    while fileNameRen == fileName:
        if os.path.isfile(archivePath + "\\" + fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf") == False:
            fileNameRen = fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf"
            #subprocess.run([ghostScriptPath, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-sDEVICE=pdfwrite", "-sColorConversionStrategy=/LeaveColorUnchangedhanged", "-dAutoFilterColorImages=true", "-dAutoFilterGrayImages=true", "-dDownsampleMonoImages=true", "-dDownsampleGrayImages=true", "-dDownsampleColorImages=true", "-sOutputFile=" + fullFile, archivePath + "\\" + fileNameRen])
            os.rename(fullFile, fullDir + "\\" + fileNameRen)
            shutil.move(fullDir + "\\" + fileNameRen, archivePath)
        print(i)
        i = i + 1

# Suppression du fichier texte d'association.
if os.path.isfile(fullDir + "\\.." + "\\.." + "\\.." + "\\Appels en cours" + "\\" + os.path.splitext(fileName)[0] + ".txt"):
    os.remove(fullDir + "\\.." + "\\.." + "\\.." + "\\Appels en cours" + "\\" + os.path.splitext(fileName)[0] + ".txt")



#subprocess.run([ghostScriptPath, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-sDEVICE=pdfwrite", "-sColorConversionStrategy=/LeaveColorUnchangedhanged", "-dAutoFilterColorImages=true", "-dAutoFilterGrayImages=true", "-dDownsampleMonoImages=true", "-dDownsampleGrayImages=true", "-dDownsampleColorImages=true", "-sOutputFile=" + fullFile, archivePath + "\\" + fileNameRen])", "-dAutoFilterColorImages=true", "-dAutoFilterGrayImages=true", "-dDownsampleMonoImages=true", "-dDownsampleGrayImages=true", "-dDownsampleColorImages=true", "-sOutputFile=" + fullFile, archivePath + "\\" + fileNameRen])
#subprocess.run(["blat", "-s", sujet, "-t", destinataires, "-f", expediteur, "-server", smtp, "-body", message, "-attach", archivePath + "\\" + fileNameRen])
