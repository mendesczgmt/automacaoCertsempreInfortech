import smtplib

from email.mime.multipart import  MIMEMultipart
from email.mime.text import MIMEText
import smtplib
import email.message

def enviar_email():  
    corpo_email = """
    <p>Oi</p>
    <p>Este é um e-mail teste da automação</p>
    """

    msg = email.message.Message()
    msg['Subject'] = "Assunto"
    msg['From'] = 'notafiscalcertsempre@gmail.com'
    msg['To'] = 'Gabryelmendescz@gmail.com'
    password = 'ksbzriefhjehzjkl' 
    msg.add_header('Content-Type', 'text/html')
    msg.set_payload(corpo_email )

    s = smtplib.SMTP('smtp.gmail.com: 587')
    s.starttls()
    # Login Credentials for sending the mail
    s.login(msg['From'], password)
    s.sendmail(msg['From'], [msg['To']], msg.as_string().encode('utf-8'))
    print('Email enviado')




enviar_email()
# host = 'smtp.gmail.com'
# port  = '587'
# login = 'notafiscalcertsempre@gmail.com'
# senha = 'cert@1230'

# server = smtplib.SMTP(host, port)

# server.ehlo()
# server.starttls()

# server.login(login, senha)




 