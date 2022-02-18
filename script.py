from tkinter.font import names
import pandas as pd
import smtplib
from os.path import basename
from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.utils import COMMASPACE, formatdate
from email import encoders
from email.mime.base import MIMEBase




SenderAddress = "amine@sv1project.tech"
password = "d73faed5208d28af08f3926dc650aef6-c3d1d1eb-91717d3f"
print('working... DO NOT CLOSE THIS WINDOW')
e = pd.read_excel("C:\\Users\\HPr\\Desktop\\Bulk-Email-Sender-master\\email.xlsx")
emails = e['Emails'].values
cnames = e['cnames'].values
server = smtplib.SMTP("smtp.mailgun.org", 587)
server.starttls()
server.login(SenderAddress, password)




for email, cname in zip(emails, cnames):
    body = MIMEMultipart()
    msg = "A Monsieur le directeur de {cname},\n\nObjet : demande de sponsoring \n\nLe projet “SuperVehiculeOne” (https://fr.sv1project.tech) a été créé par une équipe de lycéens marocains de différentes villes, dont l'objectif est d'augmenter le taux d'innovation dans la région MENA et d'atteindre les 9e et 13e objectifs du développement durable. Le projet a été et sera entièrement conçu et construit par des élèves du secondaire, qui ont fait preuve d'un très grand intérêt pour le projet et d'une intelligence extraordinaire qui les qualifie pour travailler sur des choses aussi remarquables.\n Compte tenu de la communauté et de la mission de{cname}, qui consiste à donner aux jeunes les moyens d'adopter des modes de vie plus développés et plus intelligents, nous sommes convaincus que notre projet aura un impact considérable s'il est mis en œuvre de manière réaliste, avec le soutien de {cname} en tant que sponsor.\n À ce titre, votre entreprise apportera une contribution financière ou des communiqués de presse pour faire triompher ce projet jeune et original. En tant que sponsor officiel, nous promouvrons votre entreprise sur tous les canaux de médias sociaux et ajouterons votre logo sur notre site web. Nous prendrons également des dispositions spéciales pour souligner l'implication de {cname} dans la création de SV1. Votre équipe peut également choisir parmi les niveaux de parrainage allant de 10 000 MAD jusqu’à 15 000 MAD détaillés dans le document PDF ci-joint. Tous les parrainages sont négociables et peuvent être discutés plus en détail.\n Si cela vous intéresse, nous aimerions poursuivre notre conversation. Si vous êtes disponible pour planifier un appel à votre meilleure convenance pour discuter de plus de détails , veuillez nous contacter à l'adresse contact@sv1project.tech \nNous vous remercions, \n\nÉquipe de contact au projet SV1."
    message = msg.format(cname=cname)
    subject = "Demande de sponsoring"
    body["From"] = SenderAddress
    body["To"] = email
    body["Subject"] = subject
    body.attach(MIMEText(message, 'plain'))
    # Open PDF file in binary mode
    filename = 'C:\\Users\\HPr\\Desktop\\Bulk-Email-Sender-master\\sponsorship_fr.pdf'
    with open(filename, "rb") as attachment:
    # Add file as application/octet-stream
    # Email client can usually download this automatically as attachment
      part = MIMEBase("application", "octet-stream")
      part.set_payload(attachment.read())

    # Encode file in ASCII characters to send by email    
    encoders.encode_base64(part)

    # Add header as key/value pair to attachment part
    part.add_header(
    "Content-Disposition",
    f"attachment; filename= {'sponsorship_fr.pdf'}",
    )

    # Add attachment to message and convert message to string
    body.attach(part)
    fullemail = body.as_string()
    
    print('sending to :'+ email)
    server.sendmail(SenderAddress, email, fullemail)
server.quit()

print('done... you can close this window')
input()