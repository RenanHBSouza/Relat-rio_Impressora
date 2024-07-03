import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os

# Configurações do servidor SMTP
SMTP_SERVER = 'smtp.gmail.com'
SMTP_PORT = 587
EMAIL = os.environ['EMAIL']       #Variaveis de ambiente por questão de segurança
PASSWORD = os.environ['PASS']

# Detalhes do e-mail a ser enviado
TO_EMAIL = 'destino.email@gmail.com'
SUBJECT = 'Assunto'
BODY = 'Corpo do email'
ATTACHMENT_PATH = 'caminho/do/arquivo/a/ser/anexado'  # Caminho do arquivo a ser anexado

def enviar_email():
    # Configuração do MIMEMultipart
    msg = MIMEMultipart()
    msg['From'] = EMAIL
    msg['To'] = TO_EMAIL
    msg['Subject'] = SUBJECT

    # Adicionar corpo do e-mail
    msg.attach(MIMEText(BODY, 'plain'))

    # Anexar arquivo
    if ATTACHMENT_PATH:
        anexo = open(ATTACHMENT_PATH, "rb")     #Abre o arquivo em modo leitura binária      
        part = MIMEBase('application', 'octet-stream')
        part.set_payload(anexo.read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition', f"attachment; filename= {os.path.basename(ATTACHMENT_PATH)}")
        msg.attach(part)    #Anexa a parte codificada na mensagem
        anexo.close()

    # Conectar ao servidor SMTP e enviar o e-mail
    try:
        server = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        server.starttls()
        server.login(EMAIL, PASSWORD)
        texto = msg.as_string()
        server.sendmail(EMAIL, TO_EMAIL, texto)
        print('Email enviado com sucesso!')
    except Exception as e:
        print(f'Falha ao enviar email: {e}')
    finally:
        server.quit()

if __name__ == "__main__":      #Verifica se o escript está sendo executado independentemente
    enviar_email()
