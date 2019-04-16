
# -*- coding: utf-8 -*-
# Python 3
 
# python -m zipfile -c NEGO.zip Fichiers_Chargement_cycle_2_NEGO

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import zipfile as myzip
import os



for element in os.listdir('data'):
	myzip.write(element)

fromaddr = "controlecycle.neg.etd.tvx@gmail.com"
toaddr = "ahmed.boussoufa@ixin.net"

msg = MIMEMultipart()

msg['From'] = fromaddr
msg['To'] = toaddr
msg['Subject'] = "SUBJECT OF THE EMAIL"

body = "TEXT YOU WANT TO SEND"

msg.attach(MIMEText(body, 'plain'))

filename = "Fichier_DPL_OPT.zip"
attachment = open("Fichier_DPL_OPT.zip", "rb")

part = MIMEBase('application', 'octet-stream')
part.set_payload((attachment).read())
encoders.encode_base64(part)
part.add_header('Content-Disposition', "attachment; filename= %s" % filename)

msg.attach(part)

server = smtplib.SMTP('smtp.gmail.com', 587)
server.starttls()
server.login(fromaddr, "ixinadm01")
text = msg.as_string()
server.sendmail(fromaddr, toaddr, text)
server.quit()