# Importation des modules nécessaire
import os
import subprocess
import glob
import shutil
import sys
import datetime

# Importation des fichiers associé au logiciel
import megaVariable


# Cette fonction permet d'écrire dans un fichier XML de base de donnée associé à chaque appel.
# Si le fichier n'existe pas, elle se charge de le créer avec son modèle de base.
# Si toutefois le fichier existe, elle ne fait qu'ajouter l'information nécessaire.

# Prends en arguments : Le nom du champ sur lequel écrire, le contenu qu'on veut écrire,
# ainsi que le chemin complet du fichier XML
def ecrireXml(elementCible, contenu, fichierXml):
    from xml.dom import minidom, Node
    from xml.dom.minidom import parse
    from codecs import open
    if os.path.isfile(fichierXml) == False:

        doc = minidom.Document()
        root = doc.createElement("root")
        doc.appendChild(root)

        infoAppel = doc.createElement("infoAppel")
        root.appendChild(infoAppel)

        noAppel = doc.createElement("NumeroAppel")
        infoAppel.appendChild(noAppel)
        noAppel.appendChild(doc.createTextNode(""))

        noClient = doc.createElement("NumeroClient")
        infoAppel.appendChild(noClient)
        noClient.appendChild(doc.createTextNode(""))

        coName = doc.createElement("NomDeLaCompagnie")
        infoAppel.appendChild(coName)
        coName.appendChild(doc.createTextNode(""))

        coAdresse = doc.createElement("Adresse")
        infoAppel.appendChild(coAdresse)
        coAdresse.appendChild(doc.createTextNode(""))

        coVille = doc.createElement("Ville")
        infoAppel.appendChild(coVille)
        coVille.appendChild(doc.createTextNode(""))

        codePostal = doc.createElement("CodePostal")
        infoAppel.appendChild(codePostal)
        codePostal.appendChild(doc.createTextNode(""))

        noTel = doc.createElement("NumeroTelephone")
        infoAppel.appendChild(noTel)
        noTel.appendChild(doc.createTextNode(""))

        modele = doc.createElement("Modele")
        infoAppel.appendChild(modele)
        modele.appendChild(doc.createTextNode(""))
        
        noSerie = doc.createElement("NumeroSerie")
        infoAppel.appendChild(noSerie)
        noSerie.appendChild(doc.createTextNode(""))

        emplacement = doc.createElement("Emplacement")
        infoAppel.appendChild(emplacement)
        emplacement.appendChild(doc.createTextNode(""))

        nomContact = doc.createElement("NomContact")
        infoAppel.appendChild(nomContact)
        nomContact.appendChild(doc.createTextNode(""))

        prob1 = doc.createElement("ProblemeLigne1")
        infoAppel.appendChild(prob1)
        prob1.appendChild(doc.createTextNode(""))

        prob2 = doc.createElement("ProblemeLigne2")
        infoAppel.appendChild(prob2)
        prob2.appendChild(doc.createTextNode(""))

        prob3 = doc.createElement("ProblemeLigne3")
        infoAppel.appendChild(prob3)
        prob3.appendChild(doc.createTextNode(""))

        googleSearch = doc.createElement("GoogleSearch")
        infoAppel.appendChild(googleSearch)
        googleSearch.appendChild(doc.createTextNode(""))

        techName = doc.createElement("NomTechnicien")
        infoAppel.appendChild(techName)
        techName.appendChild(doc.createTextNode(""))

        fileName = doc.createElement("NomDuFichier")
        infoAppel.appendChild(fileName)
        fileName.appendChild(doc.createTextNode(""))

        etatAppel = doc.createElement("AppelEnAttente")
        infoAppel.appendChild(etatAppel)
        etatAppel.appendChild(doc.createTextNode("False"))
        donneeSupp = doc.createElement("DonneeDeSupression")
        root.appendChild(donneeSupp)

        fichier = doc.createElement("Fichier")
        donneeSupp.appendChild(fichier)

        dossier = doc.createElement("Dossier")
        donneeSupp.appendChild(dossier)

        with open(fichierXml, "w") as f:
            try:
                #doc.writexml(f, encoding="utf-8")
                f.write(doc.toprettyxml(indent="  "))
            finally:
                f.close()

    with open(fichierXml) as xml:
        doc = minidom.parse(xml)
    node = doc.getElementsByTagName(elementCible)[0]
    if node.tagName == "Fichier" or node.tagName == "Dossier":
        node.appendChild(doc.createTextNode(contenu + "\n"))
    elif node.childNodes.length == 0:
        node.appendChild(doc.createTextNode(contenu))
    else:
        node.firstChild.replaceWholeText(contenu)
    
    with open(fichierXml, "w") as f:
        try:
            doc.writexml(f, encoding="utf-8")
            #f.write(doc.toxml())
        finally:
            f.close()
    ecrireLog("activite", "FIN DE L'ECRITURE DE : %s DANS LA NODE : %s" % (contenu, elementCible))

# Permet de lire à l'intérieur du fichier XML de base de donnée.
# Prends en arguement, le champ désiré pour extraction et le chemin complet du fichier
# Retourne une liste, donc utiliser par défaut la notation myList[0] pour les champs à une variable.
def lireXml(elementCible, fichierXml):
    from xml.dom import minidom, Node

    dom1 = minidom.parse(fichierXml)
    node = dom1.getElementsByTagName(elementCible)[0]
    if node.childNodes.length == 0:
        myList = [""]
        return myList
    else:
        name = dom1.getElementsByTagName(elementCible)[0].firstChild.nodeValue

        # A activer seulement dans le cas d'une node a plusieurs lignes.
        myList = [line for line in name.split('\n') if line.strip() != '']

    return myList

# Supprime le contenu d'une node dans le fichier XML
def deleteXml(elementCible, fichierXml):
    from xml.dom import minidom, Node

    dom1 = minidom.parse(fichierXml)
    node = dom1.getElementsByTagName(elementCible)[0]
    y = node.childNodes[0]
    node.removeChild(y)

    f = open(fichierXml, "w")
    try:
        f.write(dom1.toxml())
    finally:
        f.close()



def replaceBadCharacter (inputString):

        inputString = inputString.replace("\\", "-")
        inputString = inputString.replace("_", "-")
        inputString = inputString.replace("/", "-")
        inputString = inputString.replace(":", ";")
        inputString = inputString.replace("*", "")
        inputString = inputString.replace("?", "!")
        inputString = inputString.replace("\"", "-")
        inputString = inputString.replace("<", "(")
        inputString = inputString.replace(">", ")")
        inputString = inputString.replace("|", "-")
        inputString = inputString.replace("&", "ET")
        inputString = inputString.replace("”", "")
        inputString = inputString.replace(",", "")

        return inputString

# Change le template de email pour envoyer l'information souhaitée
def mailTemplateChanger(cible, contenu, fichier):
    import os
    import sys
    import fileinput 

    
    #with fileinput.FileInput(fichier, inplace=True) as file:
    #    for line in file:
    #        print(line.replace(cible, contenu), end='')

    # Ouvrir le fichier dans la memoire
    with open(fichier, 'r') as file :
        filedata = file.read()

    # Replace the target string
    filedata = filedata.replace(cible, contenu)

    # Write the file out again
    with open(fichier, 'w') as file:
        file.write(filedata)
        
# Ecrit dans un fichier texte le rapport de route des appels a mesure qu'ils se ferment.
def ecrireRapportRoute(ville, nomClient, noAppel, compteur, techName):

    # Selon le mois, crée une variable qui écrit le mois au complet dans un tableau.
    # Exemple : [2016, 10 - Octobre]
    dateInfo = getDate()
    monthName = str(dateInfo[1])
    year = str(dateInfo[0])
    if dateInfo[2] < 10: mois = "0" + str(dateInfo[2])
    else: mois = str(dateInfo[2])
    if dateInfo[3] < 10: jour = "0" + str(dateInfo[3])
    else: jour = str(dateInfo[3])
    dateOfDay = year + "-" + mois + "-" + jour
    heure = str(datetime.datetime.now().hour)
    minute = str(datetime.datetime.now().minute)

    if len(heure) == 1:
        heure = "0" + heure
    if len(minute) == 1:
        minute = "0" + minute
    tempsAppel = heure + ":" + minute
    

    # Affectation des dossiers
    fileName = dateOfDay + ".txt"
    fullDir = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_RAPPORT_ROUTE
    fullFile = fullDir + "\\" + fileName
    
    if len(compteur) > 4: tabCompteur = "\t"
    else: tabCompteur = "\t\t"

    if len(ville) > 13: tabVille = "\t"
    else: tabVille = "\t\t"
    
    contenu = tempsAppel + " : " + noAppel + " | Odomètre : " + compteur + tabCompteur + "| " + ville + tabVille + "| " + nomClient
    
    from codecs import open

    with open(fullFile, "a") as f:
        try:
            f.write(contenu + "\n")
        finally:
            f.close()
    ecrireLog("activite", "FIN DE L'ECRITURE DE L'ENTREE DU RAPPORT DE ROUTE")


# Fonction utilisée pour terminer l'application de façon propre en supprimant tout les fichiers de traitement.
# Si notifEmail est à oui, 
def termProc(fullFile, fileName, notifEmail="oui"):

    # Aller chercher la liste des fichier/dossier à supprimer dans le XML
    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + os.path.splitext(fileName)[0] + ".xml"

    fichier = lireXml("Fichier", fichierXml)
    dossier = lireXml("Dossier", fichierXml)
    techName = lireXml("NomTechnicien", fichierXml)[0]
     
    callTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    if notifEmail == "oui":
        subprocess.run([megaVariable.BLAT_RUN, "*** ERREUR - APPEL ERONNÉ - ***", "-t", megaVariable.destService, "-f", megaVariable.expediteur, "-server", megaVariable.smtp, "-body", "L'appel envoyé à %s à: %s, à été rejeté par le logiciel\nVérifiez le fichier en attachement s'il est conforme au modèle accepté." % (techName, callTime), "-attach", fullFile])

    if len(fichier) is not 0:
        for i in range (0, len(fichier)):
            if os.path.isfile(fichier[i]):
                os.remove(fichier[i])
            
    if len(dossier) is not 0:
        for i in range (0, len(dossier)):
            if os.path.isdir(dossier[i]):
                shutil.rmtree(dossier[i])
        
    sys.exit()
    
# Permet d'ecrire dans un fichier log d'erreur ou d'activité. prends comme arguments :
# typeLog : erreur , activite
# messageEtat : Texte à écrire dans le log
def ecrireLog(typeLog, messageEtat):
    typeLog = typeLog.lower()

    logTime = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    if os.path.isdir(megaVariable.DOSSIER_LOG) == False : os.makedirs(megaVariable.DOSSIER_LOG)
    
    if typeLog == "erreur":
        
        if os.path.isfile(megaVariable.DOSSIER_LOG + "\\errorlog.txt") == False:
            f = open(megaVariable.DOSSIER_LOG + "\\errorlog.txt", "x")
            f.close()
        with open(megaVariable.DOSSIER_LOG + "\\errorlog.txt", "a") as myfile:
            myfile.write("[" + logTime + "]: " + messageEtat + "\n")
        
    elif typeLog == "activite":
        if megaVariable.logActivite == "oui":
            
            if os.path.isfile(megaVariable.DOSSIER_LOG + "\\activitylog.txt") == False:
                f = open(megaVariable.DOSSIER_LOG + "\\activitylog.txt", "x")
                f.close()
            with open(megaVariable.DOSSIER_LOG + "\\activitylog.txt", "a") as myfile:
                myfile.write("[" + logTime + "]: " + messageEtat + "\n")
    else:
        if os.path.isfile(megaVariable.DOSSIER_LOG + "\\errorlog.txt") == False:
            f = open(megaVariable.DOSSIER_LOG + "\\errorlog.txt", "x")
            f.close()
        with open(megaVariable.DOSSIER_LOG + "\\errorlog.txt", "a") as myfile:
            myfile.write("[" + logTime + "]: " + "ENTREE INCORRECTE DANS LE TYPE DU LOG." + "\n")
        

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

# Cette fonction retourne 
def getDate():
    now = datetime.datetime.now()
    currentMonthName = megaVariable.monthList[int(now.month) - 1]
    currentMonth = now.month
    currentYear = now.year
    currentDay = now.day
    return [currentYear, currentMonthName, currentMonth, currentDay]

def getEmail(techName):    
    for i in range(0, len(megaVariable.techTab)):
        for j in range(0, len(megaVariable.techTab[i])):
            if megaVariable.techTab[i][j] == techName:
                return megaVariable.techTab[i][1]

def stamp(inputFile, stampFile, outputFile):
        # Fonction pour utiliser la fonction Stamp de PDFtk
        # Prends comme arguements : Le fichier de base, le fichier "stamp" et l'emplacement du résultat.
        subprocess.run([megaVariable.PDFTK_RUN, inputFile, "stamp", stampFile, "output", outputFile])
        
def fart(originalChar, newChar, fileName, tempFolder):
    f = open(tempFolder + "\\" + fileName,'r')
    filedata = f.read()
    f.close()

    newData = filedata.replace(originalChar, newChar)

    f = open(tempFolder + "\\" + fileName, 'w')
    f.write(newData)
    f.close()

def createTxt(folder, increment, content):

    f = open(folder + "\\customList" + str(increment) + ".txt",'a')
    f.write(content)
    f.close()

def gererNouveau(fileName, techName):

    # Selon le mois, crée une variable qui écrit le mois au complet dans un tableau.
    # Exemple : [2016, 10 - Octobre]
    dateInfo = getDate()
    monthName = str(dateInfo[1])
    year = str(dateInfo[0])
    dateOfDay = str(datetime.datetime.now().year) + str(datetime.datetime.now().month) + str(datetime.datetime.now().day) + str(datetime.datetime.now().hour) + str(datetime.datetime.now().minute) + str(datetime.datetime.now().second)

    # Affectation des dossiers et des stamps
    fullDir = megaVariable.DOSSIER_TECH + "\\" + techName
    fullFile = fullDir + "\\" + fileName
    cheminAppel = fullDir + megaVariable.DOSSIER_APPEL
    tempFolder = megaVariable.DOSSIER_TEMPORAIRE + "\\" + os.path.splitext(fileName)[0]

    appelService = tempFolder + "\\appelservice.pdf"

    # Vérification et initialisation des dossier/fichier de travail

    if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)

    shutil.move(fullFile, tempFolder + "\\" + fileName)
    fullFile = tempFolder + "\\" + fileName

    ecrireLog("activite", "VARIABLES DE TRAVAILS CREE")
    

    #
    # Début de traitement d'un appel de service PDF NUMÉRISÉ
    #

    subprocess.run([megaVariable.GHOSTSCRIPT_RUN, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-dLastPage=1", "-sDEVICE=txtwrite", "-sOutputFile=" + tempFolder + "\\original-texte.txt", fullFile])

    # Utilisation de ghostscript pour extraire du texte sur l'image, si le fichier est vide, il est considéré comme un appel scanné, autrement c'est un appel UNIX
    if os.path.isfile(tempFolder + "\\original-texte.txt") == False:
        ecrireLog("erreur", "ERREUR PENDANT L'EXECUTION DE LA COMMANDE GHOSTSCRIPT")
        termProc(fullFile, fileName)
    else:
        ecrireLog("activite", "FIN DE VERIFICATION DES ARGUMENTS")

    # Detection si c'est un appel UNIX ou scanné, le traitement diffère dépendament.
    if os.stat(tempFolder + "\\original-texte.txt").st_size > 0:
        ecrireLog("activite", "APPEL UNIX DÉTECTÉ, DÉBUT DU TRAITEMENT")
        nomAppel = dateOfDay + "-" + "1" + ".pdf" 
        gererTraitement(fileName, os.path.splitext(fileName)[0], techName, "UNIX") 

    else:
        ecrireLog("activite", "L'APPEL À ÉTÉ DÉTECTÉ COMME UN APPEL SCANNÉ NON INITIALISÉ, DÉBUT DU TRAITEMENT")

        # Creation d'un dump pour aller chercher le nombre de pages du document
        subprocess.run([megaVariable.PDFTK_RUN, fullFile, "dump_data", "output", tempFolder + "\\" + "dump.txt"])

        f = open(tempFolder + "\\dump.txt", "r")
        filedata = f.read()
        f.close()
        myList = [line for line in filedata.split('\n') if line.strip() != '']
        for i in range(0, len(myList)):
            if myList[i][0:13] == "NumberOfPages":
                nombrePages = int(myList[i][15:].strip())

        # En fonction du nombre de page détecté, séparer le document en autant de fichiers

        for i in range(0, nombrePages):
            subprocess.run([megaVariable.PDFTK_RUN, fullFile, "cat", str(i + 1), "output", tempFolder + "\\appelservice" + str(i + 1) + ".pdf"])
                
        ecrireLog("activite", "SEPARATION DU DOCUMENT (APPELS ET HISTORIQUES) EFFECTUÉ AVEC SUCCÈS")

        # Pour chaque fichier, détermine s'il est un appel ou un historique, et les regroupe ensemble

        cntAppel = 1
        rotation = False
        listeFichier=[]
        listeAppel=[]
        listeNoAppelBase=[]
            
        for i in range(0, nombrePages):

            estAppel = False
            increment = str(i+1)
            appelService = tempFolder + "\\appelservice" + increment + ".pdf"
                
            # Conversion du pdf en png et application du masque, pour vérifier l'orientation
            subprocess.run([megaVariable.MAGICK_RUN, "-density", "400", appelService, tempFolder + "\\img" + increment + ".png"])
            subprocess.run([megaVariable.MAGICK_RUN, tempFolder + "\\img" + increment + ".png", megaVariable.maskOrientation, "-composite", tempFolder + "\\maskOrientation" + increment + ".png"])
            ecrireLog("activite", "IMAGE DE TRAITEMENT #" + increment + " CRÉÉ AVEC SUCCÈS")

            # Invocation de Tesseract OCR pour générer un fichier texte la phrase "BON REPARATION"
            
            subprocess.run([megaVariable.TESSERACT_RUN, tempFolder + "\\maskOrientation" + increment + ".png", tempFolder + "\\nameOrientation" + increment, "-l", "eng"])
            ecrireLog("activite", "DETECTION DU TEXTE #" + increment + " EFFECTUÉ AVEC SUCCÈS.")

            #
            # Utilisation d'un stamp permettant de vérifier si l'orientation du document est correct.
            #
            
            f = open(tempFolder + "\\nameOrientation" + increment + ".txt",'r')
            filedata = f.read()
            f.close()
            filedata = filedata.upper()

            compteur = 0

            ecrireLog("activite", "VERIFICATION DE L'ORIENTATION DU DOCUMENT #" + increment + ".")

            # Si au moins 6 caractères sur 13 sont détecté, le logiciel considèle l'appel comme bien orienté et continue le traitement.
            if len(filedata) > 10:
                    if len(filedata) > 0:
                        if filedata[0] == "B":
                            compteur += 1
                    if len(filedata) > 1:
                        if filedata[1] == "O":
                            compteur += 1
                    if len(filedata) > 2:
                        if filedata[2] == "N":
                            compteur += 1
                    if len(filedata) > 4:
                        if filedata[4] == "D":
                            compteur += 1
                    if len(filedata) > 5:
                        if filedata[5] == "E":
                            compteur += 1
                    if len(filedata) > 7:
                        if filedata[7] == "R":
                            compteur += 1
                    if len(filedata) > 8:
                        if filedata[8] == "E":
                            compteur += 1
                    if len(filedata) > 9:
                        if filedata[9] == "P":
                            compteur += 1
                    if len(filedata) > 10:
                        if filedata[10] == "A":
                            compteur += 1
                    if len(filedata) > 11:
                        if filedata[11] == "R":
                            compteur += 1
                    if len(filedata) > 12:
                        if filedata[12] == "A":
                            compteur += 1
                    if len(filedata) > 13:
                        if filedata[13] == "T":
                            compteur += 1
                    if len(filedata) > 14:
                        if filedata[14] == "I":
                            compteur += 1
                    if len(filedata) > 15:
                        if filedata[15] == "O":
                            compteur += 1
                    if len(filedata) > 16:
                        if filedata[16] == "N":
                            compteur += 1

            # Si moins de 6 caractères sont détecté, le logiciel considère que l'appel risque d'être mal orienté.
            # Il tentera une rotation à 180 degré et une nouvelle détection sera exécutée.
                        
            if compteur < 10:

                ecrireLog("activite", "AUCUNE DONNÉE TROUVÉE SUR LE FICHIER #" + increment + ", DÉBUT DU DEUXIÈME TRAITEMENT AVEC ROTATION.")
                
                subprocess.run([megaVariable.PDFTK_RUN, appelService, "rotate", "1down", "output", tempFolder + "\\rotated" + increment + ".pdf"])
                os.remove(appelService)
                os.rename(tempFolder + "\\rotated" + increment + ".pdf", appelService)
                subprocess.run([megaVariable.MAGICK_RUN, "-density", "400", appelService, tempFolder + "\\img" + increment + ".png"])
                subprocess.run([megaVariable.MAGICK_RUN, tempFolder + "\\img" + increment + ".png", megaVariable.maskOrientation, "-composite", tempFolder + "\\maskOrientation" + increment + ".png"])    
                subprocess.run([megaVariable.TESSERACT_RUN, tempFolder + "\\maskOrientation" + increment + ".png", tempFolder + "\\nameOrientation" + increment, "-l", "eng"])

                f = open(tempFolder + "\\nameOrientation" + increment + ".txt",'r')
                filedata = f.read()
                f.close()
                filedata = filedata.upper()

                compteur = 0

                if len(filedata) > 10:
                    if len(filedata) > 0:
                        if filedata[0] == "B":
                            compteur += 1
                    if len(filedata) > 1:
                        if filedata[1] == "O":
                            compteur += 1
                    if len(filedata) > 2:
                        if filedata[2] == "N":
                            compteur += 1
                    if len(filedata) > 4:
                        if filedata[4] == "D":
                            compteur += 1
                    if len(filedata) > 5:
                        if filedata[5] == "E":
                            compteur += 1
                    if len(filedata) > 7:
                        if filedata[7] == "R":
                            compteur += 1
                    if len(filedata) > 8:
                        if filedata[8] == "E":
                            compteur += 1
                    if len(filedata) > 9:
                        if filedata[9] == "P":
                            compteur += 1
                    if len(filedata) > 10:
                        if filedata[10] == "A":
                            compteur += 1
                    if len(filedata) > 11:
                        if filedata[11] == "R":
                            compteur += 1
                    if len(filedata) > 12:
                        if filedata[12] == "A":
                            compteur += 1
                    if len(filedata) > 13:
                        if filedata[13] == "T":
                            compteur += 1
                    if len(filedata) > 14:
                        if filedata[14] == "I":
                            compteur += 1
                    if len(filedata) > 15:
                        if filedata[15] == "O":
                            compteur += 1
                    if len(filedata) > 16:
                        if filedata[16] == "N":
                            compteur += 1

                if compteur < 10 and i == 0: # Si le premier fichier n'est pas un appel, le processus est arrêté immédiatement
                    ecrireLog("erreur", "LE DOCUMENT RECU N'EST PAS UN FORMAT ACCEPTÉ PAR LE LOGICIEL. VERIFIER QUE C'EST BIEN UN APPEL STANDARD ET QUE RIEN N'OBSTRUE LA FEUILLE.")
                    termProc(fullFile, fileName)
                        
                elif compteur < 10:  # Ce fichier est donc détecté comme un attachement et sera greffé au dernier appel détecté

                    estAppel = False
                            
                    if rotation == False:
                        subprocess.run([megaVariable.PDFTK_RUN, appelService, "rotate", "1down", "output", tempFolder + "\\rotated" + increment + ".pdf"])
                        os.remove(appelService)
                        os.rename(tempFolder + "\\rotated" + increment + ".pdf", appelService)

                    listeFichier.append(appelService)
                        
                        
                else:
                    cntAppel += 1
                    estAppel = True
                    rotation = True
                    listeAppel.append(listeFichier)
                    listeFichier=[]
                    listeFichier.append(appelService)
                    listeNoAppelBase.append(increment)

            elif i == 0:
                    
                listeFichier.append(appelService)
                rotation = False
                estAppel = True
                listeNoAppelBase.append(increment)

            else:
                cntAppel += 1
                listeAppel.append(listeFichier)
                listeFichier=[]
                listeFichier.append(appelService)
                rotation = False
                estAppel = True
                listeNoAppelBase.append(increment)

            
        #
        # Fin de la vérification de l'orientation du document.
        #

        listeAppel.append(listeFichier)
        
        ecrireLog("activite", "FIN DE SEPARATION DU DOCUMENT.")

        for i in range(0, len(listeAppel)):
            if len(listeAppel[i]) > 1:
                for j in range(0, len(listeAppel[i]) - 1):
                    appelBase = listeAppel[i][0]
                    attachement = listeAppel[i][j+1]
                    subprocess.run([megaVariable.PDFTK_RUN, appelBase, attachement, "cat", "output", tempFolder + "\\" + "newfile.pdf"])
                    os.remove(appelBase)
                    os.rename(tempFolder + "\\" + "newfile.pdf", appelBase)
                nomAppel = dateOfDay + "-" + str(i) + ".pdf"
                os.rename(appelBase, tempFolder + "\\" + nomAppel)
                gererTraitement(nomAppel, os.path.splitext(fileName)[0], techName, "SCAN")
                
            else:
                nomAppel = dateOfDay + "-" + str(i) + ".pdf"
                os.rename(listeAppel[i][0], tempFolder + "\\" + nomAppel) 
                gererTraitement(nomAppel, os.path.splitext(fileName)[0], techName, "SCAN")

            ecrireLog("activite", "APPEL #%s ENVOYÉ AU TRAITEMENT" % (i))
                
    if os.path.isdir(tempFolder): shutil.rmtree(tempFolder)

    ecrireLog("activite", "TRAITEMENT TERMINÉ, APPEL(S) CRÉÉ(S) AVEC SUCCÈS.")
        

    
def gererTraitement(fileName, batchFolder, techName, callType):

    # Selon le mois, crée une variable qui écrit le mois au complet dans un tableau.
    # Exemple : [2016, 10 - Octobre]
    dateInfo = getDate()
    monthName = str(dateInfo[1])
    year = str(dateInfo[0])

    # Affectation des dossiers et des stamps
    fullDir = megaVariable.DOSSIER_TEMPORAIRE + "\\" + batchFolder
    fullFile = fullDir + "\\" + fileName
    cheminAppel = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_APPEL
    tempFolder = fullDir + "\\" + os.path.splitext(fileName)[0]
    dossierArchive = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_ARCHIVE + "\\" + year + "\\" + monthName
    dossierTermine = megaVariable.DOSSIER_TERMINE
    dossierRetour = megaVariable.DOSSIER_RETOUR

    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + os.path.splitext(fileName)[0] + ".xml"
    appelService = tempFolder + "\\appelservice.pdf"
    pageDocument = tempFolder + "\\documents.pdf"
    appelText = tempFolder + "\\Appel.pdf"

    ecrireXml("AppelEnAttente", "False", fichierXml)
    ecrireXml("NomTechnicien", techName, fichierXml)
    ecrireXml("NomDuFichier", fileName, fichierXml)

    # Ajoute le fichier principal dans la liste des fichier à supprimer en cas d'erreur
    ecrireXml("Fichier", fullFile, fichierXml)
    ecrireXml("Fichier", fichierXml, fichierXml)
        
    # Affecte le email du tech correspondant dans la variable
    techEmail = getEmail(techName)
    if "@" not in techEmail:
        ecrireLog("erreur", "L'ADRESSE EMAIL : \"%s\", N'EST PAS UNE ADRESSE EMAIL VALIDE." % (techEmail))
        termProc(fullFile, fileName)


    # Vérification et initialisation des dossier/fichier de travail

    if os.path.isdir(cheminAppel) == False : os.makedirs(cheminAppel)
    if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
    if os.path.isdir(dossierArchive) == False : os.makedirs(dossierArchive)
    if os.path.isdir(dossierTermine) == False : os.makedirs(dossierTermine)
    if os.path.isdir(dossierRetour) == False : os.makedirs(dossierRetour)
    if os.path.isdir(megaVariable.DOSSIER_APPEL_EN_COURS) == False : os.makedirs(megaVariable.DOSSIER_APPEL_EN_COURS)

    shutil.move(fullFile, tempFolder + "\\" + fileName)
    fullFile = tempFolder + "\\" + fileName

    # Ajoute le dossier temporaire dans la liste des fichier a supprimer en cas d'erreur
    ecrireXml("Dossier", tempFolder, fichierXml)

    ecrireLog("activite", "VARIABLES DE TRAVAILS CREE")
    

    #
    # Traitement d'un appel UNIX
    #

    if callType == "UNIX":
        ecrireLog("activite", "L'APPEL À ÉTÉ DÉTECTÉ COMME UN APPEL UNIX, DÉBUT DU TRAITEMENT")

        subprocess.run([megaVariable.GHOSTSCRIPT_RUN, "-dSAFER", "-dBATCH", "-dNOPAUSE", "-dNOCACHE", "-dLastPage=1", "-sDEVICE=txtwrite", "-sOutputFile=" + tempFolder + "\\original-texte.txt", fullFile])

        # Utilisation de ghostscript pour extraire du texte sur l'image, si le fichier est vide, il est considéré comme un appel scanné, autrement c'est un appel UNIX
        if os.path.isfile(tempFolder + "\\original-texte.txt") == False:
            ecrireLog("erreur", "ERREUR PENDANT L'EXECUTION DE LA COMMANDE GHOSTSCRIPT")
            termProc(fullFile, fileName)

        #à enlever si tout fonctionne bien
        #fart("\\", "-", "original-texte.txt", tempFolder)
        #fart("_", "-", "original-texte.txt", tempFolder)
        #fart("/", "-", "original-texte.txt", tempFolder)
        #fart(":", ";", "original-texte.txt", tempFolder)
        #fart("*", "", "original-texte.txt", tempFolder)
        #fart("?", "!", "original-texte.txt", tempFolder)
        #fart("\"", "-", "original-texte.txt", tempFolder)
        #fart("<", "(", "original-texte.txt", tempFolder)
        #fart(">", ")", "original-texte.txt", tempFolder)
        #fart("|", "-", "original-texte.txt", tempFolder)
        #fart("&", "ET", "original-texte.txt", tempFolder)
        #fart("”", "", "original-texte.txt", tempFolder)
        #fart(",", "", "original-texte.txt", tempFolder)

        
        f = open(tempFolder + "\\original-texte.txt", "r")
        filedata = f.read()
        f.close()
        myList = [line for line in filedata.split('\n') if line.strip() != '']
        for i in range(0, len(myList)):
            if myList[i] == myList[1]:
                noAppel = myList[i][84:].strip()
                noAppel = replaceBadCharacter (noAppel)
                ecrireXml("NumeroAppel", noAppel, fichierXml)
            elif myList[i] == myList[7]:
                noClient = myList[i][65:].strip()
                noClient = replaceBadCharacter (noClient)
                ecrireXml("NumeroClient", noClient, fichierXml)
            elif myList[i] == myList[8]:
                coName = myList[i][54:].strip().lower().title()
                coName = replaceBadCharacter (coName)
                ecrireXml("NomDeLaCompagnie", coName, fichierXml)
            elif myList[i] == myList[9]:
                coAdresse = myList[i][54:].strip().lower().title()
                coAdresse = replaceBadCharacter (coAdresse)
                ecrireXml("Adresse", coAdresse, fichierXml)
            elif myList[i] == myList[10]:
                coVille = myList[i][54:].strip().lower().title()
                coVille = replaceBadCharacter (coVille)
                ecrireXml("Ville", coVille, fichierXml)
            elif myList[i] == myList[11]:
                codePostal = myList[i][54:].strip()
                codePostal = replaceBadCharacter (codePostal)
                ecrireXml("CodePostal", codePostal, fichierXml)
            elif myList[i] == myList[12]:
                noTel = myList[i][59:].strip()
                noTel = replaceBadCharacter (noTel)
                noTel = "(" + noTel[:3] + ") " + noTel[4:]
                ecrireXml("NumeroTelephone", noTel, fichierXml)
            elif myList[i] == myList[15]:
                infoMachine = myList[i][13:].strip()
                infoMachine = replaceBadCharacter (infoMachine)

                if infoMachine.find('#') >= 0:
                    diese = infoMachine.find('#')
                else:
                    ecrireLog("erreur", "INFORMATION SUR LA MACHINE IMPOSSIBLE A DÉTECTER")
                    continue
                
                modele = infoMachine[infoMachine.find(' '):diese - 1].strip()
                noSerie = infoMachine[diese + 2:].strip()
                ecrireXml("Modele", modele, fichierXml)
                ecrireXml("NumeroSerie", noSerie, fichierXml)
            elif myList[i] == myList[16]:
                problemeLigne1 = myList[i][69:99].strip().lower().title()
                problemeLigne1 = replaceBadCharacter (problemeLigne1)
                ecrireXml("ProblemeLigne1", problemeLigne1, fichierXml)
                emplacement = myList[i][17:50].strip().lower().title()
                emplacement = replaceBadCharacter (emplacement)
                ecrireXml("Emplacement", emplacement, fichierXml)
                
            elif myList[i] == myList[17]:
                problemeLigne2 = myList[i][64:99].strip().lower().title()
                problemeLigne2 = replaceBadCharacter (problemeLigne2)
                ecrireXml("ProblemeLigne2", problemeLigne2, fichierXml)
                nomContact = myList[i][13:50].strip().lower().title()
                nomContact = replaceBadCharacter (nomContact)
                ecrireXml("NomContact", nomContact, fichierXml)

            elif myList[i] == myList[18]:
                problemeLigne3 = myList[i][64:99].strip().lower().title()
                problemeLigne3 = replaceBadCharacter (problemeLigne3)
                ecrireXml("ProblemeLigne3", problemeLigne3, fichierXml)
                
                

        appelService = fullFile
                
        ecrireLog("activite", "APPEL UNIX TRAITÉ, LES VARIABLES SONT : %s, %s, %s, %s." % (noAppel, coName, coAdresse, coVille))

    #
    # Fin du traitement d'un appel UNIX
    #


    #
    # Lancement du traitement de l'appel scanné
    #
    
    elif callType == "SCAN":

        # Division du fichier PDF en deux
        subprocess.run([megaVariable.PDFTK_RUN, fullFile, "cat", "1", "output", appelService])
        subprocess.run([megaVariable.PDFTK_RUN, fullFile, "cat", "2-end", "output", pageDocument])
        ecrireLog("activite", "SEPARATION DU DOCUMENT (APPEL ET HISTORIQUE) EFFECTUÉ AVEC SUCCÈS")

        # Conversion du pdf en png et application du masque, pour vérifier l'orientation
        subprocess.run([megaVariable.MAGICK_RUN, "-density", "400", appelService, tempFolder + "\\img.png"])
        
        ecrireLog("activite", "L'APPEL À ÉTÉ DÉTECTÉ COMME UN APPEL SCANNÉ SEPARÉ, DÉBUT DU TRAITEMENT")
        
        subprocess.run([megaVariable.MAGICK_RUN, tempFolder + "\\img.png", megaVariable.whiteMask1, "-composite", tempFolder + "\\mask1.png"])
        subprocess.run([megaVariable.MAGICK_RUN, tempFolder + "\\img.png", megaVariable.whiteMask2, "-composite", tempFolder + "\\mask2.png"])
        subprocess.run([megaVariable.MAGICK_RUN, tempFolder + "\\img.png", megaVariable.whiteMask3, "-composite", tempFolder + "\\mask3.png"])
        subprocess.run([megaVariable.TESSERACT_RUN, tempFolder + "\\mask1.png", tempFolder + "\\name1", "-l", "eng"])
        subprocess.run([megaVariable.TESSERACT_RUN, tempFolder + "\\mask2.png", tempFolder + "\\name2", "-l", "eng"])
        subprocess.run([megaVariable.TESSERACT_RUN, tempFolder + "\\mask3.png", tempFolder + "\\name3", "-l", "eng"])

        # Remplacement de tout les caractères non autorisés dans le fichier 1.
        fart("\\", "-", "name1.txt", tempFolder)
        fart("_", "-", "name1.txt", tempFolder)
        fart("/", "-", "name1.txt", tempFolder)
        fart(":", ";", "name1.txt", tempFolder)
        fart("*", " ", "name1.txt", tempFolder)
        fart("?", "!", "name1.txt", tempFolder)
        fart("\"", " ", "name1.txt", tempFolder)
        fart("<", "(", "name1.txt", tempFolder)
        fart(">", ")", "name1.txt", tempFolder)
        fart("|", " ", "name1.txt", tempFolder)
        fart("&", "ET", "name1.txt", tempFolder)
        fart("a", "-", "name1.txt", tempFolder)
        fart("Tel;", "Tel:", "name1.txt", tempFolder)
        fart("Te1;", "Tel:", "name1.txt", tempFolder)
        fart("”", "", "name1.txt", tempFolder)
        fart("â€", "", "name1.txt", tempFolder)
        fart("€", "", "name1.txt", tempFolder)
        fart("Â", "", "name1.txt", tempFolder)
        fart("L", "1", "name1.txt", tempFolder)
        fart("I", "1", "name1.txt", tempFolder)
        fart("i", "1", "name1.txt", tempFolder)
        fart("l", "1", "name1.txt", tempFolder)
        fart("O", "0", "name1.txt", tempFolder)
        fart("o", "0", "name1.txt", tempFolder)
        fart("T", "7", "name1.txt", tempFolder)
        fart("Z", "2", "name1.txt", tempFolder)


        # Remplacement de tout les caractères non autorisés dans le fichier 2.
        fart("\\", "-", "name2.txt", tempFolder)
        fart("_", "-", "name2.txt", tempFolder)
        fart("/", "-", "name2.txt", tempFolder)
        fart(":", ";", "name2.txt", tempFolder)
        fart("*", " ", "name2.txt", tempFolder)
        fart("?", "!", "name2.txt", tempFolder)
        fart("\"", " ", "name2.txt", tempFolder)
        fart("<", "(", "name2.txt", tempFolder)
        fart(">", ")", "name2.txt", tempFolder)
        fart("|", " ", "name2.txt", tempFolder)
        fart("&", "ET", "name2.txt", tempFolder)
        fart("a", "-", "name2.txt", tempFolder)
        fart("Tel;", "Tel:", "name2.txt", tempFolder)
        fart("Te1;", "Tel:", "name2.txt", tempFolder)
        fart("”", "", "name2.txt", tempFolder)
        fart("â€", "", "name2.txt", tempFolder)
        fart("€", "", "name2.txt", tempFolder)
        fart("Â", "", "name2.txt", tempFolder)

        # Remplacement de tout les caractères non autorisés dans le fichier 3.
        fart("\\", "-", "name3.txt", tempFolder)
        fart("_", "-", "name3.txt", tempFolder)
        fart("/", "-", "name3.txt", tempFolder)
        fart(":", ";", "name3.txt", tempFolder)
        fart("*", " ", "name3.txt", tempFolder)
        fart("?", "!", "name3.txt", tempFolder)
        fart("\"", " ", "name3.txt", tempFolder)
        fart("<", "(", "name3.txt", tempFolder)
        fart(">", ")", "name3.txt", tempFolder)
        fart("|", " ", "name3.txt", tempFolder)
        fart("&", "ET", "name3.txt", tempFolder)
        fart("a", "-", "name3.txt", tempFolder)
        fart("Tel;", "Tel:", "name3.txt", tempFolder)
        fart("Te1;", "Tel:", "name3.txt", tempFolder)
        fart("”", "", "name3.txt", tempFolder)
        fart("â€", "", "name3.txt", tempFolder)
        fart("€", "", "name3.txt", tempFolder)
        fart("Â", "", "name3.txt", tempFolder)
        
        
        # Modification du fichier afin de supprimer les lignes vides dans le fichier 1
        f = open(tempFolder + "\\name1.txt",'r')
        filedata = f.read()
        f.close()
        filedata = filedata.upper()

        clearData = [line for line in filedata.split('\n') if line.strip() != '']
        f = open(tempFolder + "\\name1.txt",'w')
        for i in range(0, len(clearData)):
            f.write(clearData[i] + "\n")
        f.close()

        # Modification du fichier afin de supprimer les lignes vides dans le fichier 2
        f = open(tempFolder + "\\name2.txt",'r')
        filedata = f.read()
        f.close()
        filedata = filedata.upper()
        newdata = ""
        for i in range(0, len(filedata)):
            if filedata[i] == "0" and (filedata[i-1].isalpha() or filedata[i+1].isalpha()): # Si un 0 est entouré d'au moins une lettre, elle sera convertie en O
                newdata = newdata[0:i] + "O" 
            else:
                newdata = newdata[0:i] + filedata[i]
        clearData = [line for line in newdata.split('\n') if line.strip() != '']
        f = open(tempFolder + "\\name2.txt",'w')
        for i in range(0, len(clearData)):
            f.write(clearData[i] + "\n")
        f.close()

        # Modification du fichier afin de supprimer les lignes vides dans le fichier 3
        f = open(tempFolder + "\\name3.txt",'r')
        filedata = f.read()
        f.close()
        filedata = filedata.upper()
        newdata = ""
        for i in range(0, len(filedata)):
            if filedata[i] == "0" and (filedata[i-1].isalpha() or filedata[i+1].isalpha()): # Si un 0 est entouré d'au moins une lettre, elle sera convertie en O
                newdata = newdata[0:i] + "O" 
            else:
                newdata = newdata[0:i] + filedata[i]
        clearData = [line for line in newdata.split('\n') if line.strip() != '']
        f = open(tempFolder + "\\name3.txt",'w')
        for i in range(0, len(clearData)):
            f.write(clearData[i] + "\n")
        f.close()

        # Aller chercher les lignes: num appel
        f = open(tempFolder + "\\name1.txt", "r")
        filedata = f.read()
        f.close()

        myList = [line for line in filedata.split('\n') if line.strip() != '']
        if len(myList) < 1:
            ecrireLog("erreur", "LE FICHIER \"name1.txt\" N'EST PAS CONFORME, CERTAINES INFORMATIONS SONT MANQUANTES. L'APPEL DE SERVICE EST-IL SCANNÉ DROIT OU COMPORTE-T-IL DES IRREGULARITE?")
            termProc(fullFile, fileName)
        else:
            for i in range(0, len(myList)):
                if myList[i] == myList[0]:
                    noAppel = myList[i]
                    ecrireXml("NumeroAppel", myList[i], fichierXml)


        # Aller chercher les lignes: nom entreprise, ville et adresse
        f = open(tempFolder + "\\name2.txt", "r")
        filedata = f.read()
        f.close()

        myList = [line for line in filedata.split('\n') if line.strip() != '']
        if len(myList) < 4:
            ecrireLog("erreur", "LE FICHIER \"name2.txt\" N'EST PAS CONFORME, CERTAINES INFORMATIONS SONT MANQUANTES. L'APPEL DE SERVICE EST-IL SCANNÉ DROIT OU COMPORTE-T-IL DES IRREGULARITE?")
            termProc(fullFile, fileName)
        else:
            for i in range(0, len(myList)):
                if myList[i] == myList[0]:
                    coName = myList[i]
                    ecrireXml("NomDeLaCompagnie", myList[i], fichierXml)
                elif myList[i] == myList[1]:
                    coAdresse = myList[i]
                    ecrireXml("Adresse", myList[i], fichierXml)
                elif myList[i] == myList[2]:
                    coVille = myList[i]
                    ecrireXml("Ville", myList[i], fichierXml)

        # Aller chercher la description du probleme
        f = open(tempFolder + "\\name3.txt", "r")
        filedata = f.read()
        f.close()

        myList = [line for line in filedata.split('\n') if line.strip() != '']
        if len(myList) < 1:
            ecrireLog("erreur", "LE FICHIER \"name3.txt\" N'EST PAS CONFORME, CERTAINES INFORMATIONS SONT MANQUANTES. L'APPEL DE SERVICE EST-IL SCANNÉ DROIT OU COMPORTE-T-IL DES IRREGULARITE?")
            termProc(fullFile, fileName)
        else:
            problemeLigne1 = ""
            problemeLigne2 = ""
            problemeLigne3 = ""
            for i in range(0, len(myList)):
                if myList[i] == myList[0]:
                    problemeLigne1 = myList[i]
                    ecrireXml("ProblemeLigne1", myList[i], fichierXml)
                
                if len(myList) >= 2:
                    if myList[i] == myList[1]:
                        problemeLigne2 = myList[i]
                        ecrireXml("ProblemeLigne2", myList[i], fichierXml)
                if len(myList) >= 3:
                    if myList[i] == myList[2]:
                        problemeLigne3 = myList[i]
                        ecrireXml("ProblemeLigne3", myList[i], fichierXml)
                        
        ecrireLog("activite", "FIN DE TRAITEMENT D'UN APPEL DE SERVICE PDF NUMERISE")

    else:
        termProc(fullFile, fileName)

    #
    # Fin de traitement d'un appel de service PDF NUMÉRISÉ
    #

    
    # Application du stamp de texte sur le document
    subprocess.run([megaVariable.PDFTK_RUN, appelService, "stamp", megaVariable.textStamp, "output", appelText])
    ecrireLog("activite", "STAMP APPLIQUE AVEC SUCCES")
    

    # Création du fichier pour google map
    googleSearch = "https://www.google.ca/maps/search/" + (((coAdresse + " " + coVille).replace(" ", "+")).replace(",", "")).replace("++", "+")
    ecrireXml("GoogleSearch", googleSearch, fichierXml)
    ecrireLog("activite", "VARIABLE GOOGLE MAP AVEC SUCCES")

    # Application du stamp et déplacer dans le fichier du technicien, prêt à être renommé par PyRen
    if os.path.isfile(pageDocument) == False:
        subprocess.run([megaVariable.PDFTK_RUN, megaVariable.formStamp, "stamp", appelText, "output", cheminAppel + "\\" + fileName])
    else:
        subprocess.run([megaVariable.PDFTK_RUN, megaVariable.formStamp, "stamp", appelText, "output", tempFolder + "\\" + "Appel2.pdf"])
        subprocess.run([megaVariable.PDFTK_RUN, "A=" + tempFolder + "\\" + "Appel2.pdf", "B=" + pageDocument, "cat", "A", "B", "output", cheminAppel + "\\" + fileName])
    ecrireLog("activite", "FICHIER DÉPLACÉ AVEC SUCCÈS, PRÊT POUR LE RENOMMAGE")

    # Ajoute le fichier traité dans la liste des fichier a supprimer en cas d'erreur
    ecrireXml("Fichier", cheminAppel + "\\" + fileName, fichierXml)

    # Vérification si le fichier traité à bien été créé.
    if os.path.isfile(cheminAppel + "\\" + fileName) == False:
        ecrireLog("erreur", "UNE ERREUR EST SURVENUE LORS DE LA CRÉATION DU FICHIER FINAL TRAITÉ.")
        termProc(fullFile, fileName)

    # Suppression de la liste de fichier a supprimer en cas d'erreurs
    ecrireLog("activite", "SUPRESSION DE LA LISTE DES FICHIERS A SUPPRIMER")
    #os.system("pause")
    #deleteXml("Dossier", fichierXml)
    #os.system("pause")
    #deleteXml("Fichier", fichierXml)
    ecrireLog("activite", "LISTE DES FICHIERS A SUPPRIMER EFFACEE")

def gererRenommer(fileName, techName):
    
    # Selon le mois, crée une variable qui écrit le mois au complet dans un tableau.
    # Exemple : [2016, 10 - Octobre]
    dateInfo = getDate()
    monthName = str(dateInfo[1])
    year = str(dateInfo[0])
    if dateInfo[2] < 10: mois = "0" + str(dateInfo[2])
    else: mois = str(dateInfo[2])
    if dateInfo[3] < 10: jour = "0" + str(dateInfo[3])
    else: jour = str(dateInfo[3])
    dateOfDay = year + "-" + mois + "-" + jour

    # Affectation des dossiers et des stamps
    fullDir = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_APPEL
    fullFile = fullDir + "\\" + fileName
    tempFolder = megaVariable.DOSSIER_TEMPORAIRE + "\\" + os.path.splitext(fileName)[0]
    dossierArchive = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_ARCHIVE + "\\" + year + "\\" + monthName
    dossierTermine = megaVariable.DOSSIER_TERMINE
    dossierRetour = megaVariable.DOSSIER_RETOUR
    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + os.path.splitext(fileName)[0] + ".xml"

    # Vérification et initialisation des dossier/fichier de travail
    if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
    if os.path.isdir(megaVariable.DOSSIER_APPEL_EN_COURS) == False : os.makedirs(megaVariable.DOSSIER_APPEL_EN_COURS)

    if os.path.isfile(fichierXml):
        ecrireXml("AppelEnAttente", "False", fichierXml)
        ecrireXml("NomTechnicien", techName, fichierXml)
        
        # Ajoute le fichier principal dans la liste des fichier à supprimer en cas d'erreur
        ecrireXml("Fichier", fullFile, fichierXml)
        ecrireXml("Dossier", tempFolder, fichierXml)
    
    # Affecte le email du tech correspondant dans la variable
    techEmail = getEmail(techName)
    if "@" not in techEmail:
        ecrireLog("erreur", "L'ADRESSE EMAIL : \"%s\", N'EST PAS UNE ADRESSE EMAIL VALIDE." % (techEmail))
        termProc(fullFile, fileName)

    # Lecture du fichier texte associé à l'appel, si existant, puis crée les variables
    # de travail à partir des informations obtenues
    if os.path.isfile(fichierXml):
        noAppel = lireXml("NumeroAppel", fichierXml)[0]
        coName = lireXml("NomDeLaCompagnie", fichierXml)[0]
        coAdresse = lireXml("Adresse", fichierXml)[0]
        coVille = lireXml("Ville", fichierXml)[0]
        codePostal = lireXml("CodePostal", fichierXml)[0]
        noTel = lireXml("NumeroTelephone", fichierXml)[0]
        modele = lireXml("Modele", fichierXml)[0]
        noSerie = lireXml("NumeroSerie", fichierXml)[0]
        emplacement = lireXml("Emplacement", fichierXml)[0]
        nomContact = lireXml("NomContact", fichierXml)[0]
        googleSearch = lireXml("GoogleSearch", fichierXml)[0]
        problemeLigne1 = lireXml("ProblemeLigne1", fichierXml)[0]
        problemeLigne2 = lireXml("ProblemeLigne2", fichierXml)[0]
        problemeLigne3 = lireXml("ProblemeLigne3", fichierXml)[0]
        googleSearchHTML = "<a href=\"" + googleSearch + "\"> Afficher la carte </a>"
        
    # Si le fichier texte associé est inexistant, le logiciel traite l'appel comme si c'était un
    # appel fait à la main. Il execute la fonction fdf de pdftk afin d'extraire les champs voulu,
    # pour ensuite créer les variables associées.
    else:
        ecrireXml("AppelEnAttente", "False", fichierXml)
        ecrireXml("NomTechnicien", techName, fichierXml)
        
        # Ajoute le fichier principal dans la liste des fichier à supprimer en cas d'erreur
        ecrireXml("Fichier", fullFile, fichierXml)
        ecrireXml("Dossier", tempFolder, fichierXml)
        
        # Utilise la fonction generate_fdf de pdftk pour extraire les champs de formulaire, ensuite le formate ligne par ligne
        # dans un fichier txt facilement recherchable
        subprocess.run([megaVariable.PDFTK_RUN, fullFile, "generate_fdf", "output", tempFolder + "\\" + os.path.splitext(fileName)[0] + ".fdf"])
        if os.path.isfile(tempFolder + "\\" + os.path.splitext(fileName)[0] + ".fdf") == False:
            ecrireLog("erreur", "ERREUR PENDANT LA CRÉATION DU FICHIER FDF")
            termProc(fullFile, fileName)
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
        problemeLigne1 = str([s for s in filedata if "Probleme" in s])

        # Utilise la fonction getName pour extraire seulement que l'information souhaitée
        # dans nos variables, puis crée le fichier de texte associé.
        noAppel = getName(noAppel)
        coName = getName(coName)
        coAdresse = getName(coAdresse)
        coVille = getName(coVille)
        problemeLigne1 = getName(problemeLigne1)
        problemeLigne2 = ""
        problemeLigne3 = ""
        
        googleSearch = "https://www.google.ca/maps/search/" + (((coAdresse + " " + coVille).replace(" ", "+")).replace(",", "")).replace("++", "+")
        googleSearchHTML = "<a href=\"" + googleSearch + "\"> Afficher la carte </a>"

        ecrireXml("NumeroAppel", noAppel, fichierXml)
        ecrireXml("NomDeLaCompagnie", coName, fichierXml)
        ecrireXml("Adresse", coAdresse, fichierXml)
        ecrireXml("Ville", coVille, fichierXml)
        ecrireXml("ProblemeLigne1", problemeLigne1, fichierXml)
        ecrireXml("ProblemeLigne2", problemeLigne2, fichierXml)
        ecrireXml("ProblemeLigne3", problemeLigne3, fichierXml)
        ecrireXml("GoogleSearch", googleSearch, fichierXml)
        

    alreadyExist = False
    # Incrémentation avec la boucle for des appels de service, ensuite renommage du fichier pour "date - incrementation - nom du client.pdf"
    for i in range (1, 9999):
        if i < 10: cntOrdre = "0" + str(i)
        else: cntOrdre = str(i)
        if len(glob.glob(dossierTermine + "\\" + dateOfDay + " - " + cntOrdre + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier terminé
            if len(glob.glob(dossierArchive + "\\" + dateOfDay + " - " + cntOrdre + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier archive
                if dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".pdf" == fileName:
                    alreadyExist = True
                    break
                elif len(glob.glob(fullDir + "\\" + dateOfDay + " - " + cntOrdre + "*.pdf")) == 0: # Si le numéro d'ordre n'existe pas dans le dossier principal
                    compteur = str(i)
                    os.rename(fullFile, fullDir + "\\" + dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".pdf")
                    ecrireXml("NomDuFichier", dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".pdf", fichierXml)
                    if os.path.isfile(megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".xml"):
                        os.remove(megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".xml")
                    os.rename(fichierXml, megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".xml")
                    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + dateOfDay + " - " + cntOrdre + " - " + coName +  " - " + noAppel + ".xml"
                    break
                
    if not alreadyExist:

        # envoi d'un email (ou sms) a l'utilisateur qu'un appel de service lui a été envoyé
        
        from distutils.dir_util import copy_tree
        templateFolder = tempFolder + "\\MailTemplate"
        copy_tree(megaVariable.DOSSIER_MAIL_TEMPLATE, templateFolder)
        templateHtml = templateFolder + "\\index.htm"
        templateLogo = templateFolder + "\\images" + "\\Logo-Megaburo1.PNG"

        if nomContact != "" and noTel != "":
            infoContact = nomContact + " - " + noTel
        elif nomContact != "" and noTel == "":
            infoContact = nomContact
        elif nomContact == "" and noTel != "":
            infoContact = noTel
        else:
            infoContact = ""
            
        if modele != "" and noSerie != "":
            infoMachine = modele + " - " + noSerie
        elif modele != "" and noSerie == "":
            infoMachine = modele
        elif modele == "" and noSerie != "":
            infoMachine = noSerie
        else:
            infoMachine = ""
      
        mailTemplateChanger("compteurAppel", "Appel " + str(compteur), templateHtml)
        mailTemplateChanger("nomClient", coName, templateHtml)
        mailTemplateChanger("coAdresse" , coAdresse, templateHtml)
        mailTemplateChanger("coVille", coVille, templateHtml)
        mailTemplateChanger("noAppel", noAppel, templateHtml)
        mailTemplateChanger("emplacementAppel", emplacement, templateHtml)
        mailTemplateChanger("nomContact", infoContact, templateHtml)
        mailTemplateChanger("modeleMachine", infoMachine, templateHtml)
        mailTemplateChanger("ligneProbleme", problemeLigne1 + "<br />" + problemeLigne2 + "<br />" + problemeLigne3, templateHtml)
        mailTemplateChanger("lienGoogleMap", googleSearch, templateHtml)

        sujet = "Appel de service"
        subprocess.run([megaVariable.BLAT_RUN, templateHtml, "-s", sujet, "-t", techEmail, "-f", megaVariable.expediteur, "-server", megaVariable.smtp, "-html", "-embed", templateLogo])

    
    # Suppression de la liste de fichier a supprimer en cas d'erreurs
    # Suppression des fichiers temporaire de travail
    if os.path.isdir(tempFolder): shutil.rmtree(tempFolder)
    ecrireLog("activite", "SUPRESSION DE LA LISTE DES FICHIERS A SUPPRIMER")
    deleteXml("Dossier", fichierXml)
    deleteXml("Fichier", fichierXml)
    ecrireLog("activite", "LISTE DES FICHIERS A SUPPRIMER EFFACEE")

def gererEmail(fileName, etatFinCall):
    from codecs import open
    # Selon le mois, crée une variable qui écrit le mois au complet dans un tableau.
    # Exemple : [2016, 10 - Octobre]
    dateInfo = getDate()
    monthName = str(dateInfo[1])
    year = str(dateInfo[0])
    if dateInfo[2] < 10: mois = "0" + str(dateInfo[2])
    else: mois = str(dateInfo[2])
    if dateInfo[3] < 10: jour = "0" + str(dateInfo[3])
    else: jour = str(dateInfo[3])
    dateOfDay = year + "-" + mois + "-" + jour

    #Aller chercher les informations dans la database
    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + os.path.splitext(fileName)[0] + ".xml"
    noAppel = lireXml("NumeroAppel", fichierXml)[0]
    noClient = lireXml("NumeroClient", fichierXml)[0]
    nomClient = lireXml("NomDeLaCompagnie", fichierXml)[0]
    adresse = lireXml("Adresse", fichierXml)[0]
    ville = lireXml("Ville", fichierXml)[0]
    codePostal = lireXml("CodePostal", fichierXml)[0]
    noTel = lireXml("NumeroTelephone", fichierXml)[0]
    modele = lireXml("Modele", fichierXml)[0]
    noSerie = lireXml("NumeroSerie", fichierXml)[0]
    emplacement = lireXml("Emplacement", fichierXml)[0]
    nomContact = lireXml("NomContact", fichierXml)[0]
    techName = lireXml("NomTechnicien", fichierXml)[0]

    # Affectation des dossiers et des stamps
    fullDir = megaVariable.DOSSIER_ACTION + "\\" + etatFinCall
    fullFile = fullDir + "\\" + fileName
    tempFolder = megaVariable.DOSSIER_TEMPORAIRE + "\\" + os.path.splitext(fileName)[0]
    tempFile = tempFolder + "\\" + "rapport.txt"
    dossierArchive = megaVariable.DOSSIER_TECH + "\\" + techName + megaVariable.DOSSIER_ARCHIVE + "\\" + year + "\\" + monthName
    dossierTermine = megaVariable.DOSSIER_TERMINE
    dossierRetour = megaVariable.DOSSIER_RETOUR
    
    # Vérification et initialisation des dossier/fichier de travail

    if os.path.isdir(tempFolder) == False : os.makedirs(tempFolder)
    
    # Écriture du compteur dans le log journalier
    subprocess.run([megaVariable.PDFTK_RUN, fullFile, "dump_data_fields", "output", tempFile])

    f = open(tempFile, "r")
    filedata = f.read()
    f.close()
    myListOdo = [line for line in filedata.split('\n') if line.strip() != '']
    for i in range(0, len(myListOdo)):
        if myListOdo[i] == myListOdo[4]:
            if myListOdo[i][0:11] == "FieldValue:":
                compteur = myListOdo[i][12:].strip()
            else:
                compteur = "-"
    retourErds = False
    myListErds = [line for line in filedata.split('\n') if line.strip() != '']
    for i in range(0, len(myListErds)):
        if myListErds[i] == "FieldName: ERDS":
            lineCheck = i + 2
            
            if myListErds[lineCheck][12:].strip() == "Yes":
                retourErds = True
            else:
                retourErds = False

    
    # Suppression des fichiers temporaire de travail
    if os.path.isdir(tempFolder): shutil.rmtree(tempFolder)

    # Ecriture du rapport de route
    ecrireRapportRoute(ville, nomClient, str(noAppel), compteur, techName)
    ecrireLog("activite", "FIN D'ECRITURE DU RAPPORT")

    # Ecriture du fichier ERDS si applicable
    if retourErds == True:
        if noClient == "":
            erdsFileName = dateOfDay + " - " + "N/A" + " - " + str(time.time()) + ".txt"
        else:
            erdsFileName = dateOfDay + " - " + noSerie + ".txt"
        erdsFullDir = megaVariable.DOSSIER_ERDS
        erdsFullFile = erdsFullDir + "\\" + erdsFileName
        if noClient == "":
            contenu = nomClient + "\n" + adresse + "\n" + ville + "\n"
        else:
            contenu = noClient + "\n" + nomClient + "\n" + adresse + "\n" + ville + "\n" + codePostal + "\n" + noTel + "\n" + modele + "\n" + noSerie

        if os.path.isdir(erdsFullDir) == False : os.makedirs(erdsFullDir)

        with open(erdsFullFile, "a") as f:
            try:
                f.write(contenu)
            finally:
                f.close()
        ecrireLog("activite", "FIN DE L'ECRITURE DE L'ENTREE DU FICHIER ERDS")
    
    # Envoi du email au service, le message diffère selon la commande donnée "Termine" ou "Retour"

    if etatFinCall == "Termine":
        sujet = "Appel de service termine de " + techName
        message = "Voici l'appel de service " + fileName +  " termine par " + techName
    elif etatFinCall == "Retour":
        sujet = "***Retour requis*** Appel de service termine de " + techName
        message = "***Retour requis*** Voici l'appel de service " + fileName +  " termine par " + techName
    
    subprocess.run([megaVariable.BLAT_RUN, "-s", sujet, "-t", megaVariable.destService, "-f", megaVariable.expediteur, "-server", megaVariable.smtp, "-body", message, "-attach", fullFile])

    if os.path.isdir(dossierArchive) == False : os.makedirs(dossierArchive)

    # Vérifie si le fichier existe, si oui, le crée avec une incrémentation (i)
    if os.path.isfile(dossierArchive + "\\" + fileName) == False:
        shutil.move(fullFile, dossierArchive)
    else:
        fileNameRen = fileName
        i = 1
        while fileNameRen == fileName:
            if os.path.isfile(dossierArchive + "\\" + fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf") == False:
                fileNameRen = fileName[:len(fileName) - 4] + " (" + str(i) + ").pdf"
                os.rename(fullFile, fullDir + "\\" + fileNameRen)
                shutil.move(fullDir + "\\" + fileNameRen, dossierArchive)
            i += 1

    # Suppression du fichier texte d'association.
    if os.path.isfile(fichierXml):
        os.remove(fichierXml)
        

def gererSuppression(fileName):
    ecrireLog("activite", "SUPPRESSION D'UN APPEL ET SON ATTACHEMENT")
    
    #Aller chercher les informations dans la database
    fichierXml = megaVariable.DOSSIER_APPEL_EN_COURS + "\\" + os.path.splitext(fileName)[0] + ".xml"
    
    # Affectation des dossiers et des stamps
    fullDir = megaVariable.DOSSIER_SUPPRESSION
    fullFile = fullDir + "\\" + fileName
    
    if os.path.isfile(fichierXml):
        os.remove(fichierXml)
    if os.path.isfile(fullFile):
        os.remove(fullFile)
    if os.path.isfile(fichierXml) == False and os.path.isfile(fullFile) == False:
        ecrireLog("activite", "APPEL SUPPRIMÉ AVEC SUCCÈS")
        
    else:
        if os.path.isfile(fichierXml):
            ecrireLog("erreur", "LE FICHIER : " + fichierXml + "\nN'A PU ÊTRE SUPPRIMÉ, VÉRIFIER S'IL N'EST PAS OUVERT OU PROTÉGÉ EN ÉCRITURE")
        elif os.path.isfile(fullFile):
            ecrireLog("erreur", "LE FICHIER : " + fullFile + "\nN'A PU ÊTRE SUPPRIMÉ, VÉRIFIER S'IL N'EST PAS OUVERT OU PROTÉGÉ EN ÉCRITURE")
                       
