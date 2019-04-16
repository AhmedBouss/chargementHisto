import smtplib
from email.mime.text import MIMEText

message = MIMEText('Ceci est un test !')
message['Subject'] = 'Objet du message'

message['From'] = 'controlecycle.neg.etd.tvx@gmail.com'
message['To'] = 'ahmed.boussoufa@ixin.net'
message['CC'] = 'ahmed.boussoufa@gmail.com'
server = smtplib.SMTP('smtp.gmail.com:587')
server.starttls()
server.login('controlecycle.neg.etd.tvx@gmail.com','ixinadm01')
server.send_message(message)
server.quit()