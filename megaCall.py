import sys

# Importation des fichiers associé au logiciel
import megaFunction 
import megaVariable

# Affectation des paramètres aux variables correspondantes.
if len(sys.argv) == 2:
    commande = (sys.argv[1]).upper()
    fileName = "email13578.pdf"
    techName = "Tommy"
elif len(sys.argv) < 4:
    #megaFunction.ecrireLog("erreur", "Nombre de paramètres insuffisant. Paramètres obligatoires : commande, nom du fichier, nom du tech.")
    #sys.exit()
    commande = "TERMINE"
    fileName = "2018-06-05 - 01 - Vente Comptant Service - 472395.pdf"
    techName = "Tommy"

    #commande = "MENAGE"
    #fileName = "blank.pdf"
    #techName = "Test"
else:
    commande = (sys.argv[1]).upper()
    fileName = sys.argv[2]
    techName = ((sys.argv[3]).lower()).title()  
#
# Vérification si tout les paramètres sont correctement écrit et dans le bon ordre.
#

# Regarde si la commande envoyé est existante
if commande != "NOUVEAU" and \
   commande != "RENOMMER" and \
   commande != "RETOUR" and \
   commande != "TERMINE" and \
   commande != "MENAGE" and \
   commande != "SUPPRESSION":
    megaFunction.ecrireLog("erreur", "NOM DE COMMANDE INVALIDE : \"%s\", N'EST PAS UNE COMMANDE VALIDE." % (commande))
    sys.exit()
# Vérifie si l'arguement donné est bel et bien un fichier PDF
if fileName[len(fileName)-4:len(fileName)] != ".pdf":
    megaFunction.ecrireLog("erreur", "\"%s\", N'EST PAS UN FICHIER VALIDE." % (fileName))
    sys.exit()
# Vérifie si le nom envoyé est existant et n'existe qu'une seule fois.
nomCorrespond = 0
for i in range(0, len(megaVariable.techTab)):
    if techName == megaVariable.techTab[i][0]:
        nomCorrespond = nomCorrespond + 1
if nomCorrespond == 0:
    megaFunction.ecrireLog("erreur", "AUCUN NOM CORRESPONDANT : \"%s\", N'EST PAS UN NOM DE TECH VALIDE." % (techName))
    sys.exit()
elif nomCorrespond > 1:
    megaFunction.ecrireLog("erreur", "\"%s\" EXISTE A PLUSIEURS REPRISES. AVEZ-VOUS VERIFIE SI DEUX TECH AVAIT LE MEME NOM? SI OUI, UTILISEZ LA FORME SUIVANTE : \"tommya\" \"tommyb\"" % (techName))
    sys.exit()
megaFunction.ecrireLog("activite", "FIN DE VERIFICATION DES ARGUMENTS")
#
# Fin de la vérification des paramètres
#

    
if commande == "NOUVEAU":
    megaFunction.ecrireLog("activite", "DEBUT DE TRAITEMENT D'UN NOUVEL APPEL")
    megaFunction.gererNouveau(fileName, techName)
elif commande == "RENOMMER":
    megaFunction.ecrireLog("activite", "APPEL TRAITE - DEBUT DU PROCESSUS DE RENOMMAGE")
    megaFunction.gererRenommer(fileName, techName)
elif commande == "RETOUR":
    megaFunction.ecrireLog("activite", "APPEL TERMINE, RETOUR REQUIS")
    megaFunction.gererEmail(fileName, "Retour")
elif commande == "TERMINE":
    megaFunction.ecrireLog("activite", "APPEL TERMINE, PRET A FERMER")
    megaFunction.gererEmail(fileName, "Termine")
elif commande == "MENAGE":
    megaFunction.ecrireLog("activite", "DEBUT DE PROCEDURE DE NETTOYAGE DES DOSSIERS")
    megaFunction.gererMenage()
elif commande == "SUPPRESSION":
    megaFunction.ecrireLog("activite", "DEBUT DE PROCEDURE DE SUPPRESSION")
    megaFunction.gererSuppression(fileName)
else:
    megaFunction.ecrireLog("erreur", "ERREUR -  COMMANDE NON VALIDE")
    
    
