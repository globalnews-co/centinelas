import imaplib
import os
import json
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from dotenv import load_dotenv
import time
from datetime import datetime, timedelta
import logging
import pyodbc
import subprocess
import sys

load_dotenv()

MAIL_SERVER = os.getenv("MAIL_SERVER")
PASSWORD_MAIL = os.getenv("PASSWORD_MAIL")
ALERT_EMAIL = os.getenv("ALERT_EMAIL")
ALERT_FROM_EMAIL = os.getenv("ALERT_FROM_EMAIL")
PASS_ALERT = os.getenv("PASS_ALERT")

SERVER = '192.168.1.210'
DATABASE = 'Multiencoder'
USERNAME = 'SOPORTE'
PASSWORD = 'T3cn0l0g14'

def check_pm2_service():
    try:
        pm2_commands = ['pm2', 'npx pm2', r'C:\Users\TECNOLOGIA\AppData\Roaming\npm\pm2.cmd']
        
        pm2_found = False
        pm2_cmd_working = None
        for pm2_cmd in pm2_commands:
            try:
                if isinstance(pm2_cmd, str) and pm2_cmd.startswith('C:'):
                    result = subprocess.run([pm2_cmd, 'jlist'], capture_output=True, text=True, timeout=10, encoding='utf-8', errors='ignore')
                else:
                    result = subprocess.run([pm2_cmd, 'jlist'], capture_output=True, text=True, timeout=10, encoding='utf-8', errors='ignore')
                
                if result.returncode == 0:
                    pm2_found = True
                    pm2_cmd_working = pm2_cmd
                    break
            except:
                continue
        
        if not pm2_found:
            return True
        
        if result.returncode == 0:
            pm2_data = json.loads(result.stdout)
            
            for process in pm2_data:
                process_name = process.get('name', '')
                if 'vt-palabras' in process_name or process_name == 'vt-palabras':
                    status = process.get('pm2_env', {}).get('status', 'unknown')
                    print(f"Encontrado proceso: {process_name} - Estado: {status}")
                    
                    if status in ['stopped', 'errored', 'error']:
                        print(f"Reiniciando servicio {process_name}...")
                        restart_result = restart_pm2_service(pm2_cmd_working, process_name)
                        return restart_result
                    elif status == 'online':
                        print(f"Servicio {process_name} funcionando correctamente")
                        return True
                    else:
                        return False
            
            print("Servicio vt-palabras no encontrado")
            return True
        else:
            return False
            
    except Exception as e:
        return True

def restart_pm2_service(pm2_cmd='pm2', process_name='vt-palabras'):
    try:
        print(f"pm2 restart {process_name}")
        
        if isinstance(pm2_cmd, str) and pm2_cmd.startswith('C:'):
            result = subprocess.run([pm2_cmd, 'restart', process_name], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='ignore')
        else:
            result = subprocess.run([pm2_cmd, 'restart', process_name], capture_output=True, text=True, timeout=30, encoding='utf-8', errors='ignore')
        
        return True
            
    except Exception as e:
        return False

def get_medios_activos():
    conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        
        dia_actual_numero = str(datetime.now().isoweekday())
        hora_actual = datetime.now().time()
        
        query = """
            SELECT [ID_Registro], [ID_Medio], [Start_Time], [End_Time], [NombreMedio], [Dias], [CercaDe]
            FROM [Videoteca_dev].[dbo].[VT_Horarios]
            WHERE ([Dias] LIKE ? OR [Dias] = '1,2,3,4,5,6,7' OR [Dias] = 'todos')
            AND ([CercaDe] IS NULL OR [CercaDe] != 'Deshabilitado')
        """
        
        patron_busqueda = f'%{dia_actual_numero}%'
        cursor.execute(query, (patron_busqueda,))
        horarios_del_dia = cursor.fetchall()
        
        medios_activos = []
        
        for horario in horarios_del_dia:
            try:
                dias_config = str(horario.Dias).split(',') if horario.Dias else []
                dia_incluido = dia_actual_numero in dias_config or horario.Dias in ['todos', '1,2,3,4,5,6,7']
                
                if not dia_incluido:
                    continue
                
                if isinstance(horario.Start_Time, str):
                    start_time = datetime.strptime(horario.Start_Time, '%H:%M:%S').time()
                else:
                    start_time = horario.Start_Time
                    
                if isinstance(horario.End_Time, str):
                    end_time = datetime.strptime(horario.End_Time, '%H:%M:%S').time()
                else:
                    end_time = horario.End_Time
                
                if start_time <= end_time:
                    if start_time <= hora_actual <= end_time:
                        medios_activos.append(horario)
                else:
                    if hora_actual >= start_time or hora_actual <= end_time:
                        medios_activos.append(horario)
                        
            except Exception as e:
                continue
            
        print('Medios: ', medios_activos)
        
        return medios_activos
        
    except Exception as e:
        return []
    finally:
        if 'conn' in locals():
            conn.close()

def check_alertas_por_medio():
    medios_activos = get_medios_activos()
    
    if not medios_activos:
        return True
    
    print(f"Verificando {len(medios_activos)} medios activos")
    print(medios_activos)
    
    conn_str = f'DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={SERVER};DATABASE={DATABASE};UID={USERNAME};PWD={PASSWORD}'
    medios_sin_alertas = []
    
    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()
        hace_una_hora = datetime.now() - timedelta(minutes=60)
        ahora = datetime.now()
        
        for medio in medios_activos:
            medio_id = medio.ID_Medio
            nombre_medio = medio.NombreMedio
            start_time = medio.Start_Time
            end_time = medio.End_Time
            
            query = """
                SELECT TOP (1) [IdMedio],[NombreMedio],[FechaInsersion] as FechaRgistro
                FROM [Videoteca_dev].[dbo].[dev_voicetotext]
                WHERE [IdMedio] = ?
                ORDER BY FechaInsersion DESC
            """
            
            cursor.execute(query, (medio_id,))
            row = cursor.fetchone()
            
            # Validar si la última transcripción está dentro del horario de hoy
            alerta = False
            fecha_ultima_transcripcion = row.FechaRgistro if row else None

            if fecha_ultima_transcripcion:
                # Si la última transcripción es de hoy
                if fecha_ultima_transcripcion.date() == ahora.date():
                    hora_registro = fecha_ultima_transcripcion.time()
                    # Está dentro del horario actual
                    if start_time <= end_time:
                        # Horario normal
                        if not (start_time <= hora_registro <= end_time and fecha_ultima_transcripcion >= hace_una_hora):
                            alerta = True
                    else:
                        # Horario que cruza medianoche
                        if not ((hora_registro >= start_time or hora_registro <= end_time) and fecha_ultima_transcripcion >= hace_una_hora):
                            alerta = True
                else:
                    # Última transcripción es de un día anterior
                    alerta = True
            else:
                # No hay transcripciones
                alerta = True

            if alerta:
                medios_sin_alertas.append({
                    'id': medio_id,
                    'nombre': nombre_medio,
                    'ultima_transcripcion': fecha_ultima_transcripcion if fecha_ultima_transcripcion else "Sin registros",
                    'horario': f"{start_time} - {end_time}",
                    'dias': medio.Dias
                })
        
        if medios_sin_alertas:
            print(f"Medios sin alertas: {len(medios_sin_alertas)}")
            enviar_alerta_medios_sin_registros(medios_sin_alertas)
            return False
        else:
            return True
            
    except Exception as e:
        return False
    finally:
        if 'conn' in locals():
            conn.close()

def enviar_alerta_medios_sin_registros(medios_sin_alertas):
    try:
        msg = MIMEMultipart()
        msg['From'] = ALERT_FROM_EMAIL
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = 'Alerta: Medios sin Registros de Transcripciones'

        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        tabla_medios = ""
        for medio in medios_sin_alertas:
            tabla_medios += f"""
            <tr>
                <td style="border: 1px solid #ddd; padding: 8px;">{medio['nombre']}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">{medio['id']}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">{medio['horario']}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">{medio['dias']}</td>
                <td style="border: 1px solid #ddd; padding: 8px;">{medio['ultima_transcripcion']}</td>
            </tr>
            """
        
        body = f'''
        <html>
            <body>
                <h2 style="color:#d9534f;">Alertas Voz a Texto - Medios sin Transcripción</h2>
                <p>Los siguientes medios activos <b>no tienen registros de alertas</b> en la última hora:</p>
                
                <table style="border-collapse: collapse; width: 100%; margin: 20px 0;">
                    <thead>
                        <tr style="background-color: #f2f2f2;">
                            <th style="border: 1px solid #ddd; padding: 8px;">Medio</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">ID</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Horario</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Días</th>
                            <th style="border: 1px solid #ddd; padding: 8px;">Última Transcripción</th>
                        </tr>
                    </thead>
                    <tbody>
                        {tabla_medios}
                    </tbody>
                </table>
                
                <p><b>Fecha y hora de la verificación:</b> {fecha_actual}</p>
                
                <small><b>Comandos para verificar el servicio:</b>
                <br/><br/>
                <b style="color:#d9534f;">
                - pm2 list<br/>
                - pm2 restart vt-palabras<br/>
                - pm2 logs vt-palabras
                </b></small>
                <hr>
            </body>
        </html>
        '''
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(MAIL_SERVER)
        server.login(ALERT_FROM_EMAIL, PASS_ALERT)
        text = msg.as_string()
        server.sendmail(ALERT_FROM_EMAIL, ALERT_EMAIL, text)
        server.quit()

        print("Alerta de medios sin registros enviada")

    except Exception as e:
        pass

def enviar_alerta_servicio_reiniciado():
    try:
        msg = MIMEMultipart()
        msg['From'] = ALERT_FROM_EMAIL
        msg['To'] = ALERT_EMAIL
        msg['Subject'] = 'Alerta: Servicio vt-palabras Reiniciado Automáticamente'

        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        body = f'''
        <html>
            <body>
                <h2 style="color:#f0ad4e;">Servicio PM2 Reiniciado</h2>
                <p>El servicio <b>vt-palabras</b> fue encontrado en estado <b>detenido/error</b> y ha sido <b>reiniciado automáticamente</b>.</p>
                <p><b>Fecha y hora del reinicio:</b> {fecha_actual}</p>
                
                <p><b>Acción realizada:</b> pm2 restart vt-palabras</p>
                
                <small><b>Comandos para monitoreo:</b>
                <br/><br/>
                <b style="color:#f0ad4e;">
                - pm2 list<br/>
                - pm2 logs vt-palabras --lines 50<br/>
                - pm2 monit
                </b></small>
                <hr>
            </body>
        </html>
        '''
        msg.attach(MIMEText(body, 'html'))

        server = smtplib.SMTP(MAIL_SERVER)
        server.login(ALERT_FROM_EMAIL, PASS_ALERT)
        text = msg.as_string()
        server.sendmail(ALERT_FROM_EMAIL, ALERT_EMAIL, text)
        server.quit()

        print("Alerta de reinicio enviada")

    except Exception as e:
        pass

def main():
    
    while True:
        try:
            print(f'Iniciando verificacion - {datetime.now().strftime("%H:%M:%S")}')
            
            servicio_ok = check_pm2_service()
            
            if not servicio_ok:
                enviar_alerta_servicio_reiniciado()
            
            alertas_ok = check_alertas_por_medio()
            
            if servicio_ok and alertas_ok:
                print("Sistema funcionando correctamente")
            
            print(f'Esperando 1 hora - {datetime.now().strftime("%H:%M:%S")}')
            
        except KeyboardInterrupt:
            print("Deteniendo monitor...")
            break
        except Exception as e:
            print(f"Error: {e}")
            logger.error(f"Error: {e}")
        
        time.sleep(3600)

if __name__ == "__main__":
    main()