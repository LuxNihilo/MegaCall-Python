# Si cette variable est a oui un log sera rempli détaillant l'activité du logiciel.
# Pour désactiver, le mettre à non.
logActivite = "oui"

#
# Affectation des dossiers de travail
#

### Ceci est la racine du logiciel. Tout les autres dossiers seront créé selon cette base
DOSSIER_RACINE = "Z:\\OneDrive\\_MegaCall"
###

DOSSIER_LOG = DOSSIER_RACINE + "\\Log"

DOSSIER_TECH = DOSSIER_RACINE + "\\Technicien"
DOSSIER_APPEL = "\\Appel\\Appels de service"
DOSSIER_RAPPORT_ROUTE = "\\Appel\\Rapport de route"
DOSSIER_ARCHIVE = "\\Appel\\Archive"

DOSSIER_ACTION = DOSSIER_RACINE + "\\Action"
DOSSIER_TERMINE = DOSSIER_ACTION + "\\Termine"
DOSSIER_RETOUR = DOSSIER_ACTION + "\\Retour"
DOSSIER_SUPPRESSION = DOSSIER_ACTION + "\\Suppression"

DOSSIER_CONSULTATION = DOSSIER_RACINE + "\\Consultation"
DOSSIER_ERDS = DOSSIER_CONSULTATION + "\\ERDS"

DOSSIER_LOGICIEL = DOSSIER_RACINE + "\\_Logiciels"
DOSSIER_LOGICIEL_TIERS = DOSSIER_LOGICIEL + "\\Tiers"
DOSSIER_TEMPLATE = DOSSIER_LOGICIEL + "\\Template"
DOSSIER_STAMP = DOSSIER_TEMPLATE + "\\Stamp"
DOSSIER_MAIL_TEMPLATE = DOSSIER_TEMPLATE + "\\MailTemplate"

DOSSIER_TEMPORAIRE = DOSSIER_RACINE + "\\Temp"
DOSSIER_APPEL_EN_COURS = DOSSIER_TEMPORAIRE + "\\_Appels en cours"



# Affectation des logiciels

GHOSTSCRIPT_RUN = DOSSIER_LOGICIEL_TIERS + "\\gs9.19\\bin\\gswin32c.exe"
PDFTK_RUN = DOSSIER_LOGICIEL_TIERS + "\\PDFtk\\bin\\pdftk.exe"
MAGICK_RUN = "magick"
TESSERACT_RUN = "tesseract"
BLAT_RUN = DOSSIER_LOGICIEL_TIERS + "\\Blat\\software\\full\\blat.exe"

# Affectation des stamps

formStamp = DOSSIER_STAMP + "\\stamp_Form.pdf"
textStamp = DOSSIER_STAMP + "\\stamp_text.pdf"
whiteMask1 = DOSSIER_STAMP + "\\whiteMask1.png"
whiteMask2 = DOSSIER_STAMP + "\\whiteMask2.png"
whiteMask3 = DOSSIER_STAMP + "\\whiteMask3.png"
maskOrientation = DOSSIER_STAMP + "\\maskOrientation.png"


#
# Tableau des techs
#

techTab = [["Tommy", "tomgir@outlook.com"],
           ["Claude", "claduf@megaburo.ca"],
           ["Denis", "contact_02@hotmail.com"],
           ["Eric", "eribou@megaburo.ca"],
           ["Laurier", "lausim@megaburo.ca"],
           ["Stephane", "stecot@megaburo.ca"],
           ["Carl", "carlal@megaburo.ca"],
           ["Test", "tomgir@outlook.com"],
           ["Sylvain", "sylbau@megaburo.ca"],
           ["David", "davfou@megaburo.ca"],
           ["Dummy", "dummy@megaburo.ca"]] #Utilisé pour les fonction générale

#
# Information courriel
#

destService = "jaclar@megaburo.ca,sabfil@megaburo.ca"

expediteur = "service.alm@megaburo.ca"
smtp = "smtp.cgocable.ca"

#
# Liste des mois pour utilisation dans les noms de dossier d'archive.
#

monthList = ["01 - Janvier", "02 - Fevrier", "03 - Mars", "04 - Avril", "05 - Mai", "06 - Juin", "07 - Juillet", "08 - Aout", "09 - Septembre", "10 - Octobre", "11 - Novembre", "12 - Decembre"]
