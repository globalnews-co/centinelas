import imaplib
import os
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import time
from datetime import datetime, date
import logging

os.system ("title Centinela BotEmail")

load_dotenv()

MAIL_SERVER = os.getenv("MAIL_SERVER")
PASSWORD_MAIL = os.getenv("PASSWORD_MAIL")
ALERT_EMAIL = os.getenv("ALERT_EMAIL")
ALERT_FROM_EMAIL = os.getenv("ALERT_FROM_EMAIL")
PASS_ALERT = os.getenv("PASS_ALERT")

def read_json_file(json_path):
    with open(json_path, 'r', encoding='utf-8') as file:
        data = json.load(file)
    return data

def contar_correos(mail_address, folder_name="INBOX"):
    try:
        mail = imaplib.IMAP4("mail.globalnews.com.co", 143)
        mail.login(mail_address, PASSWORD_MAIL)

        status, _ = mail.select(folder_name)
        if status != 'OK':
            print(f"[ERROR] Error al seleccionar la carpeta {folder_name} para {mail_address}")
            return None

        result, data = mail.search(None, "ALL")
        mail_ids = data[0]

        id_list = mail_ids.split()
        numero_de_correos = len(id_list)

        mail.logout()

        return numero_de_correos

    except imaplib.IMAP4.error as e:
        print(f"[ERROR] Error al conectar con el servidor IMAP para {mail_address}: {e}")
    except Exception as e:
        print(f"[ERROR] Error obteniendo la cantidad de correos en la bandeja de: {mail_address} - {e}")

    return None

def enviar_alerta_alertas_no_enviadas(alertas_no_enviadas):
    try:
        
        if not alertas_no_enviadas:
            return

        msg = MIMEMultipart()
        msg['From'] = ALERT_FROM_EMAIL
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = 'Alerta Correo Bot WhatsApp - Alertas sin enviar'

        body = '''
        <html>
            <body>
                <table border="1" style="border-collapse: collapse; padding:3px;">
                    <tr>
                        <th>Correo</th>
                        <th>Cantidad de Alertas No Enviadas</th>
                    </tr>
        '''
        for correo, cantidad in alertas_no_enviadas.items():
            body += f'''
                    <tr>
                        <td>{correo}</td>
                        <td>{cantidad}</td>
                    </tr>
            '''
        body += '''
                </table>
            </body>
        </html>
        '''
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(MAIL_SERVER)
        server.login(ALERT_FROM_EMAIL, PASS_ALERT)
        text = msg.as_string()
        server.sendmail(ALERT_FROM_EMAIL, ALERT_EMAIL, text)
        server.quit()

        print("[INFO] Alerta de alertas no enviadas enviada por correo")

    except Exception as e:
        print(f"[ERROR] Error al enviar la alerta por correo: {e}")

def enviar_alerta_inbox(resultados, limite):
    try:
        if not resultados:
            return

        msg = MIMEMultipart()
        msg['From'] = ALERT_FROM_EMAIL
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = 'Alerta Correo Bot WhatsApp - Bandeja de entrada llena'

        body = '''
        <html>
            <body>
                <table border="1" style="border-collapse: collapse; padding:3px;">
                    <tr>
                        <th>Correo</th>
                        <th>Cantidad de Correos</th>
                        <th>Límite</th>
                    </tr>
        '''
        for correo, cantidad in resultados.items():
            if cantidad > limite:
                body += f'''
                    <tr>
                        <td>{correo}</td>
                        <td>{cantidad}</td>
                        <td>{limite}</td>
                    </tr>
                '''
        body += '''
                </table>
            </body>
        </html>
        '''
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(MAIL_SERVER)
        server.login(ALERT_FROM_EMAIL, PASS_ALERT)
        text = msg.as_string()
        server.sendmail(ALERT_FROM_EMAIL, ALERT_EMAIL, text)
        server.quit()

        print("[INFO] Alerta de correos excediendo límite enviada por correo")

    except Exception as e:
        print(f"[ERROR] Error al enviar la alerta por correo: {e}")
        
def generar_log():
    log_dir = r'C:\Users\Administrador\Documents\botEmail\logs'
    
    if not os.path.exists(log_dir):
        os.makedirs(log_dir)
        
    fecha = datetime.today().date()
    
    logger = logging.getLogger('centinelaBW_log')
    logger.setLevel(logging.DEBUG)
    
    ch = logging.StreamHandler()
    ch.setLevel(logging.DEBUG)
    
    log_file_path = os.path.join(log_dir, f'log_{fecha}.log')
    
    fh = logging.FileHandler(log_file_path)
    fh.setLevel(logging.DEBUG)
    
    formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
    
    ch.setFormatter(formatter)
    fh.setFormatter(formatter)
    
    logger.addHandler(ch)
    logger.addHandler(fh)
    
    return logger
       
while True: 
    logger = generar_log()
    resultados = {}
    alertas_no_enviadas = {}

    config = read_json_file(r'C:\Users\Administrador\Documents\botEmail\correos.json')
    correos_enviados = config[0]['correos']
    limite = config[0]['limite']
    
    for correo in correos_enviados:
        # revisa carpeta alertasNoEnviadas
        print(f'[INFO] Correo: {correo} - Revisando carpetas')
        
        numero_alertasNoEnviadas = contar_correos(correo, "alertasNoEnviadas")
        
        if numero_alertasNoEnviadas and numero_alertasNoEnviadas > 0:
            alertas_no_enviadas[correo] = numero_alertasNoEnviadas
            print(f'[WARNING] {correo} tiene alertas sin enviar')
            
        # revisa bandeja de entrada
        numero_de_correos = contar_correos(correo, "INBOX")
        if numero_de_correos is not None:
            if numero_de_correos > limite:
                resultados[correo] = numero_de_correos
                print(f'[WARNING] {correo} bandeja llena')
 
    enviar_alerta_alertas_no_enviadas(alertas_no_enviadas)
    enviar_alerta_inbox(resultados, limite)

    with open('resultados_correos.json', 'w') as json_file:
        json.dump(resultados, json_file, indent=4)
        
    print(f'[{datetime.now()}] Esperando 15 minutos para la proxima ejecucion...')
    time.sleep(900)
    
    
    
    