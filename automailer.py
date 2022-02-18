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


#write the path to the folder where you put this code.
folderpath = 'C:\\Users\\HPr\\Desktop\\automail\\'
SenderAddress = "your@email"
password = "YOUR_PASSWORD"


#DO NOT TOUCH THE CODE BELOW !!!!!!

print('working... DO NOT CLOSE THIS WINDOW YOU MORON!!')
e = pd.read_excel(folderpath+"email.xlsx")
emails = e['Emails'].values
cnames = e['cnames'].values
langs = e['language'].values
server = smtplib.SMTP("smtp.mailgun.org", 587)
server.starttls()
server.login(SenderAddress, password)


  
for email, cname, lang in zip(emails, cnames, langs):
    if lang in ['FR', 'fr', 'Fr', 'fR', 'french']:
      body = MIMEMultipart()
      msg = "A Monsieur le directeur de {cname},\n\nObjet : demande de sponsoring \n\nLe projet “SuperVehiculeOne” (https://fr.sv1project.tech) a été créé par une équipe de lycéens marocains de différentes villes, dont l'objectif est d'augmenter le taux d'innovation dans la région MENA et d'atteindre les 9e et 13e objectifs du développement durable. Le projet a été et sera entièrement conçu et construit par des élèves du secondaire, qui ont fait preuve d'un très grand intérêt pour le projet et d'une intelligence extraordinaire qui les qualifie pour travailler sur des choses aussi remarquables.\n Compte tenu de la communauté et de la mission de{cname}, qui consiste à donner aux jeunes les moyens d'adopter des modes de vie plus développés et plus intelligents, nous sommes convaincus que notre projet aura un impact considérable s'il est mis en œuvre de manière réaliste, avec le soutien de {cname} en tant que sponsor.\n À ce titre, votre entreprise apportera une contribution financière ou des communiqués de presse pour faire triompher ce projet jeune et original. En tant que sponsor officiel, nous promouvrons votre entreprise sur tous les canaux de médias sociaux et ajouterons votre logo sur notre site web. Nous prendrons également des dispositions spéciales pour souligner l'implication de {cname} dans la création de SV1. Votre équipe peut également choisir parmi les niveaux de parrainage allant de 10 000 MAD jusqu’à 15 000 MAD détaillés dans le document PDF ci-joint. Tous les parrainages sont négociables et peuvent être discutés plus en détail.\n Si cela vous intéresse, nous aimerions poursuivre notre conversation. Si vous êtes disponible pour planifier un appel à votre meilleure convenance pour discuter de plus de détails , veuillez nous contacter à l'adresse contact@sv1project.tech \nNous vous remercions, \n\nÉquipe de contact au projet SV1."
      message = msg.format(cname=cname)
      subject = "Demande de sponsoring"
      body["From"] = SenderAddress
      body["To"] = email
      body["Subject"] = subject
      body.attach(MIMEText(message, 'plain'))
      # Open PDF file in binary mode
      filename = folderpath+'sponsorship_fr.pdf'
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

  
    elif lang in ['EN', 'en', 'En', 'eN', 'ENG', 'eng','english']:
      bodyf = MIMEMultipart()
      msgf = "To {cname},\n\nSubject : sponsoring request \n\nSuperVehiculeOne Project (https://sv1project.tech) is created by a team of highschool students from all over Morocco, dedicated to raising the rate of innovation in the MENA Region as well as achieving the 9th and 13th SDGs.The project has been and will be fully designed and built by high school students that have displayed exceedingly big interest in the project and extraordinary intellect that qualifies them to work on such remarkable things.\nWith {cname}’s community and mission to empower young people to have more developed and smarter lifestyles, we strongly believe our project will create a great impact if it turns out to be realistically created, with {cname}’s support as our corporate sponsor.\nIn this role, your company would provide monetary contributions or media and press releases to make this youthfully fresh project triumphant. As our official sponsor, we will promote your company on all social media channels and add your logo to the car and to our website. We will also make special arrangements to highlight {cname}’s involvement in creating SV1, as well as feature interested employees as mentors or keynote speakers. Alternatively, your team could choose from sponsorship tiers ranging from 10 000 MAD to 15 000 MAD, detailed in the attached PDF document. All sponsorships are negotiable and can be further discussed.\nIf this excites your team, we’d like to continue our conversation. Would you be available to schedule a call at your earliest convenience to discuss further details? Please reach out at contact@sv1project.tech. Looking forward to hearing from you soon!\nThank you, \n\nOutreach Team @ SV1 Project"
      messagef = msg.format(cname=cname)
      subjectf = "Sponsorship request"
      bodyf["From"] = SenderAddress
      bodyf["To"] = email
      bodyf["Subject"] = subjectf
      bodyf.attach(MIMEText(messagef, 'plain'))
      # Open PDF file in binary mode
      filename = folderpath+'sponsorship_en.pdf'
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
      f"attachment; filename= {'sponsorship_en.pdf'}",
      )

      # Add attachment to message and convert message to string
      bodyf.attach(part)
      fullemailf = bodyf.as_string()
    
      print('sending to :'+ email)
      server.sendmail(SenderAddress, email, fullemailf)
      

    else:
      print('language error!... check the language you chose you idiot!')

server.quit()



print('done... you can close this window <3')
print('THANK YOU FOR YOUR SERVICE!')
input()