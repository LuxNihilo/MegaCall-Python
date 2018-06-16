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
fullDir = sys.argv[2]
fileName = sys.argv[3]
techName = sys.argv[4]
retour = sys.argv[5]


monthList = ["01 - Janvier", "02 - Fevrier", "03 - Mars", "04 - Avril", "05 - Mai", "06 - Juin", "07 - Juillet", "08 - Aout", "09 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Decembre"]
monthName = monthList[int(fileName[5:7]) - 1]
year = fileName[0:4]

if retour == "False":
    sujet = "Appel de service termine de " + techName
elif retour == "True":
    sujet = "***Retour requis*** Appel de service termine de " + techName
destinataires = "jaclar@megaburo.ca,sabfil@megaburo.ca"
expediteur = "service.alm@megaburo.ca"
smtp = "smtp.cgocable.ca"
if retour == "False":
    message = "Voici l'appel de service " + fileName +  " termine par " + techName
elif retour == "True":
    message = "***Retour requis*** Voici l'appel de service " + fileName +  " termine par " + techName
    
#
# Suggestion du consultant:
# PDFTK avec l'option FLATTEN pour "aplatir" ton PDF avec les données du formulaire
# (la commande ci-dessous fonctionne avec GhostScript 9.19)
# "C:\Program Files (x86)\gs\gs9.19\bin\gswin32c.exe" -dSAFER -dBATCH -dNOPAUSE -dNOCACHE -sDEVICE=pdfwrite -sColorConversionStrategy=/LeaveColorUnchanged -dAutoFilterColorImages=true -dAutoFilterGrayImages=true -dDownsampleMonoImages=true -dDownsampleGrayImages=true -dDownsampleColorImages=true -sOutputFile=original_flattened.pdf original.pdf
#

subprocess.run(["blat", "-s", sujet, "-t", destinataires, "-f", expediteur, "-server", smtp, "-body", message, "-attach", fullFile])

archivePath = fullDir + "\\" + ".." + "\\Archives" + "\\" + year + "\\" + monthName
if os.path.isdir(archivePath) == False : os.makedirs(archivePath)

# Vérifie si le fichier existe, si oui, le crée avec une incrémentation (i)
if os.path.isfile(archivePath + "\\" + fileName) == False:
    shutil.move(fullFile, archivePath)
else:
    fileNameRen = fileName
    i = 1
    while fileNameRen == fileName:
        if os.path.isfile(archivePath + "\\" + fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf") == False:
            fileNameRen = fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf"
            os.rename(fullFile, fullDir + "\\" + fileNameRen)
            shutil.move(fullDir + "\\" + fileNameRen, archivePath)
        print(i)
        i = i + 1
        
#ghostScriptPath = "C:\Program Files (x86)\gs\gs9.19\bin\gswin32c.exe"
#subprocess.run([ghostScriptPath, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-sDEVICE=pdfwrite", "-sColorConversionStrategy=/LeaveColorUnchanged", "-dAutoFilterColorImages=true", "-dAutoFilterGrayImages=true", "-dDownsampleMonoImages=true", "-dDownsampleGrayImages=true", "-dDownsampleColorImages=true", "-sOutputFile=" + fullFile, archivePath + "\\" + fileNameRen])
#subprocess.run(["blat", "-s", sujet, "-t", destinataires, "-f", expediteur, "-server", smtp, "-body", message, "-attach", archivePath + "\\" + fileNameRen])
