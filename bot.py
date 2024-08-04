import os
import requests
import json
import openai
import pandas as pd
import nltk
from nltk.tokenize import word_tokenize
import glob
import pandas as pd
import numpy as np
import time
from datetime import datetime, timedelta
from zoneinfo import ZoneInfo
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import NoSuchElementException
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google.auth.exceptions import RefreshError
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
from googleapiclient.errors import HttpError
import io
import os
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.chrome.options import Options
import shutil
from selenium.webdriver.support.ui import Select
import sys
import platform
import glob
from unidecode import unidecode
from google.cloud import bigquery
from google.oauth2 import service_account
import csv
import threading
import json
import re
import ssl
from threading import Lock

import subprocess

def get_config_value(key, filename='config.json'):
    """
    Lee el archivo JSON y devuelve el valor correspondiente a la clave especificada.

    :param key: La clave cuyo valor se desea obtener.
    :param filename: El nombre del archivo JSON (por defecto 'config.json').
    :return: El valor asociado a la clave, o None si la clave no existe.
    """
    try:
        with open(filename, 'r') as file:
            config_data = json.load(file)
            return config_data.get(key, None)
    except FileNotFoundError:
        print(f"El archivo {filename} no se encontró.")
    except json.JSONDecodeError:
        print("Error al decodificar el archivo JSON.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")


def download_google_sheet_as_excel(service, id_file, download_path, filename):
    """
    Descarga un archivo Google Sheets como Excel desde Google Drive.

    Parámetros:
        service: El cliente de servicio de la API de Google Drive.
        id_file: El ID del archivo de Google Sheets.
        download_path: La ruta del directorio donde se guardará el archivo.
        filename: El nombre bajo el cual se guardará el archivo descargado.
    """
    try:
        request = service.files().export(fileId=id_file, mimeType='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while not done:
            status, done = downloader.next_chunk()
            print("Descarga %d%%." % int(status.progress() * 100))

        # Guardar el contenido del buffer en un archivo en el sistema de archivos
        with open(os.path.join(download_path, filename + '.xlsx'), 'wb') as f:
            f.write(fh.getvalue())
        print(f"Archivo descargado y guardado como: {os.path.join(download_path, filename + '.xlsx')}")

    except Exception as error:
        print(f"Error al descargar el archivo: {error}")

def load_asesores_from_excel(file_path):
    # Leer el archivo Excel
    df = pd.read_excel(file_path, engine='openpyxl')
    
    # Verificar que existan las columnas necesarias
    if 'Nombre' in df.columns and 'Prioridad' in df.columns:
        # Ordenar el DataFrame por la columna 'Prioridad' de menor a mayor
        df_sorted = df.sort_values(by='Prioridad', ascending=True)
        # Extraer la lista de nombres ordenados
        options_to_click = df_sorted['Nombre'].tolist()
    else:
        # Imprimir un mensaje de error si alguna columna no está presente
        missing_cols = []
        if 'Nombre' not in df.columns:
            missing_cols.append("Nombre")
        if 'Prioridad' not in df.columns:
            missing_cols.append("Prioridad")
        print(f"Las columnas {', '.join(missing_cols)} no se encuentran en el archivo Excel.")
        options_to_click = []

    return options_to_click

def download_audio():

    usuario = get_config_value("ccvox_user") 
    contrasena =  get_config_value("ccvox_password")

    # Configura la ruta de descarga
    current_dir = os.getcwd()
    nombre_carpeta = "audios"
    ruta_carpeta = os.path.join(current_dir, nombre_carpeta)
    archivo_json = os.path.join(ruta_carpeta, 'Audios descargados.json')

    if not os.path.exists(ruta_carpeta):
        os.makedirs(ruta_carpeta)
        print(f"Carpeta '{nombre_carpeta}' en disco {ruta_carpeta} creada con éxito.")

    # Lee el archivo JSON para saber qué archivos ya han sido descargados
    if os.path.exists(archivo_json):
        with open(archivo_json, 'r') as f:
            archivos_descargados = set(json.load(f))
    else:
        archivos_descargados = set()

    options = Options()
    options.add_argument("--disable-extensions")
    options.add_argument("--disable-application-cache")
    options.add_argument("--disable-session-crashed-bubble")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--start-maximized")

    # Establecer las preferencias para la ruta de descarga
    prefs = {
        "download.default_directory": ruta_carpeta,
        "download.prompt_for_download": False,  # Para evitar que se muestre la ventana de confirmación de descarga
        "download.directory_upgrade": True,  # Para usar la ruta de descarga global en lugar de la del perfil
        "safebrowsing.enabled": True  # Activar la protección de navegación segura
    }
    options.add_experimental_option("prefs", prefs)

    # Crea el controlador de Chrome con las opciones configuradas
    driver = webdriver.Chrome(options=options)

    # Abrir la página web
    url = get_config_value("ccvox_url")
    driver.get(url)

    ##### Inicia Sesion
    usuario_input = driver.find_element(By.ID, "login")
    usuario_input.send_keys(usuario)
    contrasena_input = driver.find_element(By.ID, "pass")
    contrasena_input.send_keys(contrasena)
    elemento_boton = driver.find_element(By.XPATH, "//input[@type='image' and @src='img/imagen_boton.png']")
    elemento_boton.click()

    grabaciones_link = driver.find_element(By.LINK_TEXT, "Grabaciones")
    grabaciones_link.click()

    # Cambia al iframe que contiene el select element
    iframe = WebDriverWait(driver, 10).until(
        EC.frame_to_be_available_and_switch_to_it((By.ID, 'idIframe'))
    )

    # Espera hasta que el elemento select esté presente y luego crea una instancia de Select
    select_element = WebDriverWait(driver, 10).until(
        EC.presence_of_element_located((By.ID, "lstAgentes"))
    )
    select = Select(select_element)

    # Construye la lista de opciones a clickear
    options_to_click = []
    all_options = select_element.find_elements(By.TAG_NAME, "option")
    for option in all_options:
        options_to_click.append(option.text)

    time.sleep(3)

    options_to_click = load_asesores_from_excel('Asesores.xlsx')
    
    max_llamadas = (20-1)/len(options_to_click) +1

    for option_text in options_to_click:
        selected_option = driver.find_element(By.XPATH, f"//option[text()='{option_text}']")

        time.sleep(5)  # Add a brief pause to allow the selection and scrolling to settle

        # Double-click the selected option
        ActionChains(driver).double_click(selected_option).perform()
        time.sleep(3)    
        input_element = driver.find_element(By.ID, "btnnuevoregistro")

        # Click the input element
        input_element.click()
        driver.execute_script("listarg();")
        
        try:
            dwld = WebDriverWait(driver, 3).until(
                EC.presence_of_element_located((By.ID, "download"))
            )
            
            #tabla = driver.find_element(By.ID, "listado")
            # Inicio la variable tabla consiguiendo el punto de jerarquia del listado generado, incluyendo la tabla_resultado.
            #tabla_cabeza = tabla.find_element(By.ID, "green")
            # Obtengo la cabecera conteredora de los botones para avanzar las paginas.
            
            #try:
            #    indicar_pagina = tabla_cabeza.find_element(By.XPATH, "/div[contains(@class,'short')]/input[contains(@type,'text')]")
            #    avanzar_pagina_boton = tabla_cabeza.find_element(By.XPATH, "/div[contains(@class,'short')]/input[contains(@type,'button')]")
            #except:
            #    print("El Xpath no encuentra")
            
            #tabla = tabla.find_elements(By.CLASS_NAME, "tabla_resultado_grabacion")[1]
            tabla = driver.find_element(By.ID, "mt")
            # Sobreescribo la variable tabla a lo que es el contenido de tabla_resultado.
            print("Tabla obtenida")

            tabla_cuerpo = tabla.find_element(By.TAG_NAME, "tbody")
            #tabla_cuerpo = tabla.find_element(By.XPATH,"/tbody")
            print("Cuerpo de Tabla obtenido")

            filas = tabla_cuerpo.find_elements(By.TAG_NAME, "tr")
            long_filas = len(filas)
            
            #Hacer que comience desde el primero hasta la llamada más reciente hecha.
            #filas.reverse()
            
            print(f"Revisando llamadas de asesor {option_text}")
            llamadas = 0

            for i in range(1,long_filas):
                
                if llamadas > max_llamadas:
                    break
                
                if (not filas[i].is_displayed()):
                    driver.execute_script("arguments[0].style.display = 'none';", filas[i-15])
                    driver.execute_script("arguments[0].style.display = 'table-row';", filas[i])
                    time.sleep(0.05)
                
                time.sleep(5)
                duracion=filas[i].find_elements(By.TAG_NAME,"td")[3].text
                mm=int(duracion[3:5])
                ss=int(duracion[6:8])
                if mm<= 20 and mm >= 3:
                    elemento_descargar=filas[i].find_elements(By.TAG_NAME,"td")[-1]
                    href_link=elemento_descargar.find_element(By.TAG_NAME,"a")
                    archivo_nombre = href_link.get_attribute('href').split('/')[-1]
                    archivo_nombre = archivo_nombre[39:-17]
                    archivo_nombre = archivo_nombre.split('&')[0]
                    print(f'Nombre de archivo detectado: {archivo_nombre}')
                    if archivo_nombre not in archivos_descargados:
                        href_link.click()
                        archivos_descargados.add(archivo_nombre)
                        #print(f"Descargando llamada {i+1} de asesor {option_text} con duración {mm}:{ss}")
                        print(f"Descargando llamada {i+1} de asesor {option_text} con duración {mm}m{ss}s")

                        # Actualizar el JSON después de cada descarga
                        with open(archivo_json, 'w') as f:
                            json.dump(list(archivos_descargados), f)
                        
                        llamadas= llamadas+1
        except:
            print(f"Asesor {option_text} sin llamadas")
    time.sleep(5)
    driver.quit()


def check_ping(host='api.openai.com'):
    """Realiza un ping al host especificado para comprobar la conectividad."""
    try:
        # En algunos sistemas, puede ser necesario especificar el número total de ECHO_REQUEST a enviar.
        flag = '-n' # '-c' en Linux, en Windows usar '-n'
        response = subprocess.run(['ping', flag, '4', host], stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)
        
        if response.returncode == 0:
            print("Ping exitoso.")
            return True
        else:
            print(f"Fallo al alcanzar {host}. Código de salida: {response.returncode}")
            print("Salida del comando:", response.stdout)
            print("Error del comando:", response.stderr)
            return False
    except Exception as e:
        print("Excepción al ejecutar el ping:", str(e))
        return False

# Crea una instancia de Lock que será compartida por todos los hilos
plataforma_lock = Lock()

def get_google_drive_api(TOKEN_PATH = 'token_drive.json', CREDENTIALS_PATH = 'drive_credentials.json'):
    SCOPES = [
    'https://www.googleapis.com/auth/drive',          # Para operaciones generales en Drive (lectura y escritura)
    'https://www.googleapis.com/auth/drive.file',     # Para acceder a archivos creados o abiertos por la app
    'https://www.googleapis.com/auth/drive.readonly', # Para acceso de sólo lectura a Drive
    'https://www.googleapis.com/auth/drive.install',
    'https://www.googleapis.com/auth/bigquery',
    'https://www.googleapis.com/auth/cloud-platform',
    'https://www.googleapis.com/auth/bigquery.readonly',
    'https://www.googleapis.com/auth/bigquery.insertdata',
    'https://www.googleapis.com/auth/datastore',
    'https://www.googleapis.com/auth/logging.admin',
    'https://www.googleapis.com/auth/logging.read',
    'https://www.googleapis.com/auth/docs',
    'https://www.googleapis.com/auth/drive.apps.readonly',
    'https://www.googleapis.com/auth/activity',
    'https://www.googleapis.com/auth/drive.scripts'
    ]
    
    creds = None
    # El archivo token.json almacena los tokens de acceso y actualización del usuario, y se
    # crea automáticamente cuando el flujo de autorización se completa por primera vez.
    
    if os.path.exists(TOKEN_PATH):
        creds = Credentials.from_authorized_user_file(TOKEN_PATH, SCOPES)
    # Si no hay credenciales válidas disponibles, deja que el usuario se loguee.
    if not creds or not creds.valid:
        try:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    CREDENTIALS_PATH, SCOPES)
                creds = flow.run_local_server(port=0)
            # Guarda las credenciales para la próxima ejecución
            with open(TOKEN_PATH, 'w') as token:
                token.write(creds.to_json())
        except RefreshError:
            # Si hay un error de actualización, es probable que el token de acceso haya expirado
            # Solicita autenticacion
            print('Token expirado, requiere volver a autenticarse')
            flow = InstalledAppFlow.from_client_secrets_file(
                CREDENTIALS_PATH, SCOPES)
            creds = flow.run_local_server(port=0)
            # Guarda las credenciales para la próxima ejecución
            with open(TOKEN_PATH, 'w') as token:
                token.write(creds.to_json())

    # Llama a la API de Drive v3
    service = build('drive', 'v3', credentials=creds)
    return service
    
def encontrar_id_carpeta(sala_numero, tipo_carpeta):
    """Encuentra el ID de carpeta usando la posición directa basada en el número de sala."""
    # Cargar los datos de correspondencia
    try:
        with open('correspondencia.json') as f:
            correspondencias = json.load(f)
        # Acceso directo usando sala_numero - 1 como índice
        entrada = correspondencias[sala_numero - 1]
        return entrada[f'{tipo_carpeta}_ID']
    except (IndexError, KeyError):
        print("Error: No se encontró la entrada correspondiente o el tipo de carpeta es incorrecto.")
        return None

def listar_archivos(service, folder_id, detail=False):
    """Devuelve detalles de archivos en una carpeta específica si 'detail' es True."""
    query = f"'{folder_id}' in parents"
    response = service.files().list(q=query, fields="files(id, name)").execute()
    items = response.get('files', [])
    if detail:
        return {item['name']: item['id'] for item in items}
    else:
        return [item['name'] for item in items]

def archivos_nuevos_y_descarga(service, folder_id, sala_numero, tipo_carpeta, download_path):
    archivo_json = f'S{sala_numero}_{tipo_carpeta}.json'
    archivos_info = listar_archivos(service, folder_id, detail=True)  # Ajustar para obtener detalles
    archivos_actuales = set(archivos_info.keys())

    if os.path.exists(archivo_json):
        with open(archivo_json, 'r') as f:
            archivos_anteriores = set(json.load(f))
    else:
        archivos_anteriores = set()

    nuevos_archivos = archivos_actuales - archivos_anteriores
    numero = 0
    
    # Guardar el estado actual de archivos para futuras comparaciones
    with open(archivo_json, 'w') as f:
        json.dump(list(archivos_actuales), f)

    archivos = os.listdir(os.getcwd())

    # Filtra los archivos que coinciden con el patrón deseado
    patron_archivo = fr'S{sala_numero}_{tipo_carpeta}_\d+\.xlsx$'
    archivos_filtrados = [archivo for archivo in archivos if re.match(patron_archivo, archivo)]

    # Si no se encuentran archivos, proceder a una acción alternativa o lanzar un error
    if archivos_filtrados:
        archivo_mayor_secuencia = max(archivos_filtrados, key=lambda x: int(re.search(r'(\d+)\.xlsx$', x).group(1)))
        numero_mayor_secuencia = int(re.search(r'(\d+)\.xlsx$', archivo_mayor_secuencia).group(1))
        numero = numero_mayor_secuencia + 1  # Comienza con el siguiente número en la secuencia
        print(f'El próximo número de secuencia es: {numero}')
    else:
        numero = 0  # Si no hay archivos, comienza desde 0
    
    for archivo_nombre in nuevos_archivos:
        file_id = archivos_info[archivo_nombre]
        file_name_descarga = f'S{sala_numero}_{tipo_carpeta}_{numero}.xlsx'
        descargar_archivo(service, file_id, file_name_descarga, download_path)
        numero += 1

    return list(nuevos_archivos), list(archivos_actuales)

def descargar_archivo(service, file_id, file_name, download_path):
    """Descarga un archivo específico de Google Drive."""
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()  # Usamos un buffer en memoria para almacenar temporalmente el archivo descargado
    
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    try:
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)  # Vuelve al inicio del buffer
        
        # Escribe el contenido del buffer en un archivo en el sistema de archivos
        # Solo si la descarga se completó exitosamente
        with open(os.path.join(download_path, file_name), 'wb') as f:
            f.write(fh.read())
        print(f"Archivo {file_name} descargado en {download_path}.")
        
    except ssl.SSLError as e:
        print(f"Error SSL al intentar descargar el archivo {file_name}: {e}")
    except Exception as e:
        print(f"Error inesperado al intentar descargar el archivo {file_name}: {e}")


def guardar_ids_en_json(ids_dict, file_path="ids_carpeta.json"):
    with open(file_path, "w") as json_file:
        json.dump(ids_dict, json_file)

def leer_ids_desde_json(file_path="ids_carpeta.json"):
    
    if os.path.exists(file_path):
        with open(file_path, "r") as json_file:
            return json.load(json_file)
    return None

def buscar_id_carpeta(service, nombre_carpeta, carpeta_padre_id=None):
    """
    Busca el ID de una carpeta por su nombre en Google Drive.

    Parámetros:
    - service: El servicio de Google Drive autenticado.
    - nombre_carpeta: El nombre de la carpeta que estás buscando.
    - carpeta_padre_id: Opcional. El ID de la carpeta padre si deseas buscar en una carpeta específica.

    Retorna:
    - El ID de la carpeta encontrada o None si no se encuentra.
    """
    query = f"name = '{nombre_carpeta}' and mimeType = 'application/vnd.google-apps.folder'"
    if carpeta_padre_id:
        query += f" and '{carpeta_padre_id}' in parents"
    
    try:
        response = service.files().list(q=query,
                                        spaces='drive',
                                        fields='files(id, name)').execute()
        files = response.get('files', [])
        
        if not files:
            print(f"No se encontró la carpeta: {nombre_carpeta}")
            return None
        else:
            # Asume que el primer archivo encontrado es la carpeta deseada
            carpeta = files[0]
            print(f"Encontrado: {carpeta['name']} (ID: {carpeta['id']})")
            return carpeta['id']
    except Exception as error:
        print(f"Error al buscar la carpeta: {error}")
        return None


def check_transcribed_files(folder_path_d, filename):
    """ Verifica si un archivo ya ha sido transcrito anteriormente."""
    transcribed_json = os.path.join(folder_path_d, 'Archivos transcritos.json')
    if os.path.exists(transcribed_json):
        with open(transcribed_json, 'r') as f:
            transcribed_files = json.load(f)
    else:
        transcribed_files = []

    if filename in transcribed_files:
        return True
    return False

def update_transcribed_files(folder_path_d, filename):
    """ Actualiza el archivo JSON con los nombres de los archivos ya transcritos."""
    transcribed_json = os.path.join(folder_path_d, 'Archivos transcritos.json')
    if os.path.exists(transcribed_json):
        with open(transcribed_json, 'r') as f:
            transcribed_files = json.load(f)
    else:
        transcribed_files = []

    transcribed_files.append(filename)
    with open(transcribed_json, 'w') as f:
        json.dump(transcribed_files, f)



def transcribe_audio(api_key, folder_path_s, folder_path_d, ext, filename):
    if check_transcribed_files(folder_path_d, filename):
        print(f"Archivo {filename} ya ha sido transcrito.")
        return

    model_endpoint = 'https://api.openai.com/v1/audio/transcriptions'
    drive_folder_id = get_config_value("transcripted_calls_drive_folder_id")  # ID de la carpeta de destino en Drive
    headers = {
        'Authorization': f'Bearer {api_key}'
    }

    data = {
        'model': 'whisper-1',
        'language': 'es'  # Set language to Spanish
    }

    with open(os.path.join(folder_path_s, f'{filename}{ext}'), 'rb') as audio_file:
        files = {
            'file': audio_file
        }
        response = requests.post(model_endpoint, headers=headers, files=files, data=data)
        response_json = response.json() #None
        
        # Guardar la transcripción localmente y subirla a Google Drive
        output_filename = f'{filename}.txt'
        output_path = os.path.join(folder_path_d, output_filename)
        process_response(response_json, output_path)
        
        # Subir el archivo Excel a Google Drive
        service = get_google_drive_api()
        upload_file_to_drive(service, drive_folder_id, output_path, output_filename, 'text/plain')
        
        # Actualizar el registro de archivos transcritos
        update_transcribed_files(folder_path_d, filename)
        #Esperar 1/2 min
        time.sleep(30)
        

def process_response(response, filename):
    if response is None:
        transcription=get_text_ex()
        write_to_file(transcription, filename)
    elif 'error' in response:
        print("Error:", response['error']['message'])
    elif 'text' in response:
        transcription = response['text']
        write_to_file(transcription, filename)
    else:
        print("No transcription available.")

def write_to_file(text, filename):
    with open(filename, 'w', encoding='utf-8') as file:
        file.write(text)
    print(f"Output written to {filename}")

def transcribe_folder(folder_path_s, folder_path_d,valid_extensions):
    if not check_ping():
        print("Conectividad fallida, terminando ejecución.")
        return

    for filename in os.listdir(folder_path_s):
        print(f'Revisando {filename}')
        api_key=get_api_key()
        # Check if the file's extension is in the list of valid extensions
        if any(filename.endswith(ext) for ext in valid_extensions):
            output_filename = os.path.splitext(filename)[0]
            ext = os.path.splitext(filename)[1]
            transcribe_audio(api_key, folder_path_s,folder_path_d, ext, output_filename)
            
        
def interact_with_openai(text, model, mandate, context, api_key):
    """
    Función para interactuar con la API de OpenAI utilizando el modelo gpt-4-1106-preview.

    :param text: El texto a ser analizado.
    :param model: El modelo a ser usado.
    :param mandate: El mandato basado en la información.
    :param api_key: La clave de API para autenticar la solicitud.
    :return: La respuesta de la API de OpenAI.
    """
    
    # Construir la carga útil (payload) para la solicitud JSON
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": context},
            {"role": "user", "content": mandate},
            {"role": "user", "content": text}
        ]
    }

    # Definir los encabezados para la solicitud
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}"
    }

    # Realizar la solicitud POST a la API de OpenAI
    response = requests.post("https://api.openai.com/v1/chat/completions", headers=headers, data=json.dumps(payload))
    
    # Manejar la respuesta de la API
    if response.status_code == 200:
        response_data = response.json()
        return response_data["choices"][0]["message"]["content"]
    else:
        raise Exception(f"Error: {response.status_code}, {response.text}")            

def get_api_key():

    return get_config_value('OPENAI_API_KEY')

def get_text_ex():
    return '''
    Hola, Victor, ¿cómo está? Buenos días. Estoy llamando de Empresa inc, La cuna del conocimiento. Ah, ya, ok. Lo que pasa es que estaba interesada en nuestra carrera de Comunicación Audiovisual Multimedia. Quería saber si aún se encuentra interesada. Ah, sí, sí, sí, comento. Genial, perfecto. ¿Esta carrera es para ti, es para algún familiar? Para mí. ¿Para ti? ¿Y ya terminaste quinto o secundaria? ¿Recién estás por terminar? No, no, no, ya he terminado. Ya, genial. ¿En cuál de nuestras sedes te gustaría estudiar? En la sede, si no me equivoco, en San Borja. Ah, en San Borja. No, pero no tenemos uno en San Borja, tenemos uno en Magdalena, que está por la avenida Javier Prado Oeste y tenemos otro en la avenida Primavera, que está en Surquillo, en Surco, perdón, está en Chacaría del Estante, en Surco. Ah, perfecto. Sí, genial. Sí tenemos todavía vacantes para el inicio de clases en esta sede. Bien, ahora, el día de hoy estaría iniciando un nuevo ciclo, pero también tenemos el nuevo ciclo iniciando el 19 de agosto. El 19 de agosto iniciamos un segundo ciclo, perdón, otro ciclo académico y estamos iniciando y estamos adelantando el curso de creatividad el día 3 de junio. El 3 de junio adelantamos clases con el curso de creatividad. Va a durar desde el día 3 hasta el día 22, donde normalmente tú recibes en el primer ciclo seis cursos. De estos seis cursos estarías adelantando uno de junio a julio, para que cuando ya termines, el 19 de agosto, inicies tus clases con tus cinco cursos restantes. O sea, estarías adelantando cursos y, lo mejor de todo, es que cuando tú adelantas un curso, porque se ven confundidos, lo que hacen es cobrarte por el adelanto. Sin embargo, para esta ocasión, lo único que tendrías que hacer para el adelanto es que se cobra la primera cuota de tu carrera y se cobra el desafío, que es el examen que realizas, que reemplazas el examen. Se estaría cobrando para que puedas matricularte y pagarías una adicional por el adelanto de clases. ¿Te interesa adelantar cursos? Bueno, no es obligatorio, ¿cierto? No, puedes tú matricularte para el día 19, o sea, puedes de frente iniciar el día 19 o puedes iniciar el día 3 de junio adelantando tus clases con el Inicia Tu Luz. Lo que sí puedo indicarte también es que estás sujeto a vacantes. Como tú estás interesada en estudiar en la sede de Javier Prado, ¿me podrías confirmar en qué colegio has terminado quinto secundario para ver temas de vacantes? En Juventud Científica. Juventud Científica, ¿pero de qué distrito es? ¿Dónde queda la Juventud Científica? En el Agustino. En el Agustino, aquí está. ¿Es un colegio particular o es un colegio estatal? Es particular. Juventud Científica. O si no, Científica. Científica. Un momentito, por favor. Estoy buscando a tu colegio ya. Aquí está. Tengo Juventud Científica, uno está en Zárate, el otro está en... Bueno, dos están en San Juan del Urigancho y uno está en Villa María del Triunfo. Uno está en Zárate. Sí, sí, lo que pasa es que no... ...dice que no es para mi colegio. O sea, la sede de mi colegio... Ah, ya. ¿Pero cuál de las dos la pongo? ¿De Zárate o San Juan del Urigancho? Bueno, luego vamos a San Juan del Urigancho. Sí, es que uno de esos dos, o sea, usted está en el Agustino. La dirección se lo podría mandar por WhatsApp o desde su dirección en WhatsApp. Ya, está bien. ¿En qué año terminaste, me dice, en el 2018? Sí. ¿Tu secundaria terminaste en el 2018? Deciré. Sí, sí. Genial, genial. A ver, un momento. Un momento, por favor. Un momento, no me cuelguen. Ya, genial. Entonces, si sería para el día... Ya, genial. Entonces, si sería para el día 19 de agosto, sería en Javier Prado, con la carrera de Comunicación Audiovisual Multimedia, que tiene el grado de bachiller, que es equivalente al nivel universitario. Mira, me figura exactamente tres vacantes disponibles para el día 19. Para el día 19. O sea, de igual forma, tendríamos que realizar la separación de la vacante, en el caso te interese iniciar el 19. Sí, así de repente te gustaría llevar el adelanto de curso, o en el caso de que no quieras adelantar curso, sí, también tendríamos que separar la vacante para poder iniciar el 19 de agosto. Me interesa más el inicio de agosto. Claro, pero repito, solo tengo tres vacantes para el inicio de agosto. Como nosotros somos una institución privada, en realidad nuestras vacantes o nuestras matrículas se realizan de muchos meses anteriores. Hemos iniciado desde octubre del año pasado. Entonces, por esa razón, me quedan tres vacantes como total, en cantidad de vacantes. Entonces, tú puedes, una de estas vacantes, si tú deseas activar el curso de creatividad para el 3 de junio, sería genial, esperando una de estas vacantes para iniciar el 19 de agosto. Entonces, tengo que hacer los pagos. Sí, tendrías que matricularte. Lo que sí puedo también indicar, por lo que estoy revisando, es que normalmente a tu colegio le toca 1.190 soles. Y esto normalmente se divide en seis cuotas de 865 soles. Ahora, si te matriculas, aprovechando las tres vacantes, te atendría a una recategorización más baja. Y 6.530 soles dividido en seis cuotas de 850. Es una cuota de 430. Y se empezaría exonerando el costo de desafío. Entonces, 850 soles. ¿Te interesa? Ahorita mismo estoy en el micro. No lo escucho muy bien, entonces intercorto la llamada. Pero, digo, esa es la información que me acaba de brindar. ¿Me podría mandar tu whatsapp? ¿Dentro de cuánto tiempo me confirmaría si es interesada? Ok. Déjame mandarte entonces toda la info que necesitas y estamos en contacto a las seis, ¿sí? No sé si tienes alguna otra consulta adicional. Podría ser ese... ¿No te matriculas? Sí. En el caso de que te matricules, te estaríamos exonerando el costo de matrícula de 430 soles y el desafío de 150 soles. Esos dos montos se te estaría exonerando. Solo te matricularías pagando 855, que es la primera cuota de tu carrera. ¿Cómo? Ah, sí, sí. La carrera tiene una duración de cuatro años con el grado de bachiller, que es equivalente al nivel universitario. ¿Y este desafío de qué se trata? El desafío es una evaluación que reemplace el examen de admisión. Es para conocerte. No es eliminatorio, no es desaprobatorio. Vas a realizar una serie de preguntas, vas a realizar un dibujo y esto pasa por una revisión psicológica. Entonces, dependiendo de esa información que arroje, los psicólogos pueden hacerte una entrevista por Zoom o personal. Va a depender de la respuesta final que tenga tu examen en sí. Pero normalmente eso es, en algunos casos, porque normalmente a algunos chicos les sale su desafío para conocer sus actitudes, para conocer sus actitudes psicológicas, sus habilidades, y normalmente es para conocerlos. En el caso de que, por ejemplo, ¿qué pasa si de repente tú estás pasando por una etapa de duelo o de repente tienes alguna deficiencia? '''

def get_text(file_path):
    # Abrir el archivo en modo de lectura y obtener el texto
    with open(file_path, 'r', encoding='utf-8') as file:
        text = file.read()
    return text

def get_context(knowledge_base_path_1,knowledge_base_path_2=None):
    context = '''
    GPT is a good spanish assistant that interprets Spanish transcribed calls and responds concisely "Si"/"No"/"N/A" when asked about meeting criterion. For each criterion GPT will respond "Si" when it is sure that the criterion is being met in the transcript provided, if GPT is not sure about find that the criterion is beign met, it will respond "N/A", If GPT does not find criterion is beign met or is not present in transcription provided, it will respond "No" . GPT answers paragraphs other than "Si"/"No"/"N/A" only if it is told to provide feedback, otherwise it always answers "Si"/"No"/"N/A", for each question, it is the most appropriate way it works. Transcripted calls always have two interlocutors, so GPT knows that must identify text presents on examples on criterions to meet the criteria. GPT has the following knowledge base for understating what to do:
    '''
    
    df = pd.read_excel(knowledge_base_path_1, usecols=[0, 1, 2, 3], names=['N_etapa', 'Criterio', 'Recomendaciones', 'Ejemplos'])
    # Itera sobre las filas del DataFrame
    for index, row in df.iterrows():
    context += f"--\n\nEtapa {row['N_etapa']} - {row['Criterio']}\nRecomendaciones: {row['Recomendaciones']}\nEjemplos: {row['Ejemplos']}"

    # Genera una lista de respuestas aleatorias "Sí", "No" o "N/A"
    responses = ["Sí", "No", "N/A"]
    random_responses = [random.choice(responses) for _ in range(len(df))]

    # Convierte la lista de respuestas en una cadena separada por '|'
    example_response = '|'.join(random_responses)

    # Añade la parte final al contexto
    context += f'''\n
    --
    When GPT is asked to analyze the call based on the criteria, GPT responds "Si"/"No"/"N/A" for each of all {len(df)} criteria, all criteria must be answered, each response must be separated by a pipe (|). Example response:
    "{example_response}"
    GPT always responds using the output format, GPT knows his response for each criterion will be used by a bot and this response is useful for people.
    '''
    
    return context
    
def get_questions(file_path):
    # Lee el archivo .xlsx y convierte la primera columna en un DataFrame
    df = pd.read_excel(file_path, usecols=[0], names=['Criterio'])
    return df
    
def generate_embeddings(text, model="text-embedding-ada-002"):
    response = openai.Embedding.create(
        input=text,
        model=model
    )
    return response['data'][0]['embedding']

def calculate_tokens(text):

    nltk.download('punkt')  # Descargar datos necesarios para el tokenizador
    tokens = word_tokenize(text)
    t_num=len(tokens)
    return t_num

def gather_responses(file_path):
    preguntas_df = get_questions('Preguntas.xlsx')

    api_key = get_api_key() 
    context = get_context('Fases de venta.xlsx')
    
    print(context)
    models = ['gpt-3.5-turbo-0125', 'gpt-4o-2024-05-13', 'gpt-3.5-turbo', 'gpt-4-1106-preview']
    mandate = 'Analiza la llamada basado en los criterios'
    model = models[3]
    

    text = get_text(file_path)
    print(text)

    print(f"Tokens entrada: {calculate_tokens(context) + calculate_tokens(text)}")

    print(f"Pregunta: {mandate}")  # Esto imprime la pregunta (para verificación)
    #'No|Sí|Sí|Sí|No|Sí|Sí|Sí|Sí|Sí|Sí|No|N/A|Sí|Sí'
    response = interact_with_openai(text, model, mandate, context, api_key)
    
    print(f"Respuesta: {response}")
    
    scores = '2%,3%,5%,10%,5%,10%,5%,5%,10%,5%,10%,6%,7%,5%,2%,4%,3%,3%'
    scores = [float(score.strip('%')) / 100 for score in scores.split(',')]
    scores_df = pd.DataFrame(scores, columns=['Puntaje esperado'])
    preguntas_df = pd.concat([preguntas_df, scores_df], axis=1)
    
    print(f'Tokens salida: {calculate_tokens(response)}')
    respuestas = response.split('|')

    # Crear un DataFrame con las preguntas y las respuestas
    respuestas_df = pd.DataFrame(respuestas, columns=['Respuesta'])
    result_df = pd.concat([preguntas_df, respuestas_df], axis=1)

    # Asignar puntaje
    result_df['Puntaje asignado'] = result_df.apply(lambda row: row['Puntaje esperado'] if row['Respuesta'] != 'No' else 0, axis=1)

    return result_df


def save_responses_to_csv(df, file_path):
    """
    Guarda el DataFrame de respuestas en un archivo CSV.
    
    Args:
    df (pd.DataFrame): El DataFrame que contiene las preguntas y respuestas.
    file_path (str): La ruta donde se guardará el archivo CSV.
    """
    df.to_csv(file_path, index=False, encoding='utf-8')

def check_analyzed_files(formats_folder, filename):
    """ Verifica si un archivo ya ha sido analizado anteriormente."""
    analyzed_json = os.path.join(formats_folder, 'Archivos analizados.json')
    if os.path.exists(analyzed_json):
        with open(analyzed_json, 'r') as f:
            analyzed_files = json.load(f)
    else:
        analyzed_files = []

    if filename in analyzed_files:
        return True
    return False

def update_analyzed_files(formats_folder, filename):
    """ Actualiza el archivo JSON con los nombres de los archivos ya analizados."""
    analyzed_json = os.path.join(formats_folder, 'Archivos analizados.json')
    if os.path.exists(analyzed_json):
        with open(analyzed_json, 'r') as f:
            analyzed_files = json.load(f)
    else:
        analyzed_files = []

    analyzed_files.append(filename)
    with open(analyzed_json, 'w') as f:
        json.dump(analyzed_files, f)



def analyze_transcriptions(transcriptions_folder, formats_folder):
    if not check_ping():
        print("Conectividad fallida, terminando ejecución.")
        return

    drive_folder_id = get_config_value("transcripted_calls_drive_folder_id")

    # Listar todos los archivos en la carpeta de transcripciones
    for filename in os.listdir(transcriptions_folder):
        # Comprobar si el archivo es un archivo de texto
        if filename.endswith('.txt'):
            print(f'Analizando transcripción: {filename}')
            if check_analyzed_files(formats_folder, filename):
                print(f"Archivo {filename} ya ha sido analizado.")
                continue
            # Ruta completa al archivo de transcripción
            file_path = os.path.join(transcriptions_folder, filename)

            # Obtener las respuestas basadas en la transcripción
            df = gather_responses(file_path)

            # Nombre del archivo de salida basado en el nombre del archivo de transcripción
            output_filename = filename[:-4] + '.xlsx'  # Remueve '.txt' y agrega '.xlsx'
            output_path = os.path.join(formats_folder, output_filename)

            # Guardar el DataFrame en un archivo Excel
            save_responses_to_excel(df, output_path)

            # Subir el archivo Excel a Google Drive
            service = get_google_drive_api()
            upload_file_to_drive(service, drive_folder_id, output_path, output_filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            
            # Actualizar el registro de archivos analizados
            update_analyzed_files(formats_folder, filename)
            
            #Esperar 2 min
            time.sleep(120)

def upload_file_to_drive(service, folder_id, file_path, file_name, mimetype):
    """
    Sube un archivo a Google Drive en la carpeta especificada.
    """
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(file_path,
                            mimetype=mimetype,
                            resumable=True)
    try:
        file = service.files().create(body=file_metadata, media_body=media, fields='id').execute()
        print(f'Archivo {file_name} subido exitosamente con ID {file["id"]}')
    except Exception as error:
        print(f'Error al subir el archivo a Google Drive: {error}')
    return file["id"]

def save_responses_to_excel(df, filename):
    # Usar ExcelWriter para guardar el DataFrame en formato Excel
    with pd.ExcelWriter(filename, engine='xlsxwriter') as writer:
        df.to_excel(writer, index=False)
        print(f'Respuestas guardadas en {filename}')

def listar_archivos(service, folder_id, detail=False):
    """Devuelve detalles de archivos en una carpeta específica si 'detail' es True."""
    query = f"'{folder_id}' in parents"
    response = service.files().list(q=query, fields="files(id, name)").execute()
    items = response.get('files', [])
    if detail:
        return {item['name']: item['id'] for item in items}
    else:
        return [item['name'] for item in items]


def listar_archivos_locales(download_path):
    """Lista los archivos en una carpeta local."""
    return set(os.listdir(download_path))

def descargar_archivos_nuevos(service, folder_id, download_path):
    archivo_json = os.path.join(download_path, 'Audios descargados.json')
    archivos_info = listar_archivos(service, folder_id, detail=True)
    archivos_drive = set(archivos_info.keys())
    archivos_locales = listar_archivos_locales(download_path)

    nuevos_archivos = archivos_drive - archivos_locales
    if not nuevos_archivos:
        print("No se encontraron archivos nuevos.")
        return list(archivos_drive), list(nuevos_archivos)

    if os.path.exists(archivo_json):
        with open(archivo_json, 'r') as f:
            archivos_descargados = set(json.load(f))
    else:
        archivos_descargados = set()

    for archivo in nuevos_archivos:
        file_id = archivos_info[archivo]
        descargar_archivo(service, file_id, archivo, download_path)
        archivos_descargados.add(archivo)

        # Actualizar el archivo JSON después de cada descarga
        with open(archivo_json, 'w') as f:
            json.dump(list(archivos_descargados), f)

    return list(archivos_drive), list(nuevos_archivos)

def descargar_archivo(service, file_id, file_name, download_path):
    """Descarga un archivo específico de Google Drive."""
    request = service.files().get_media(fileId=file_id)
    fh = io.BytesIO()

    downloader = MediaIoBaseDownload(fh, request)
    done = False
    try:
        while not done:
            status, done = downloader.next_chunk()
        fh.seek(0)
        with open(os.path.join(download_path, file_name), 'wb') as f:
            f.write(fh.read())
        print(f"Archivo {file_name} descargado en {download_path}.")
    except ssl.SSLError as e:
        print(f"Error SSL al intentar descargar el archivo {file_name}: {e}")
    except Exception as e:
        print(f"Error inesperado al intentar descargar el archivo {file_name}: {e}")

def main():
    folder_path = '.'  # Current directory
    valid_extensions = ['.mp3', '.wav', '.m4a', '.flac','.MP3', '.WAV', '.M4A', '.FLAC'] 

    service=get_google_drive_api()
    download_google_sheet_as_excel(service, get_config_value("list_of_salesmen"),'.', 'Asesores')
    download_google_sheet_as_excel(service, get_config_value("list_of_criteria"),'.', 'Fases de venta')
    download_audio()
    descargar_archivos_nuevos(service, get_config_value("recorded_calls_drive_folder_id"),os.path.join(folder_path, 'audios') )
    transcribe_folder(os.path.join(folder_path, 'audios'),os.path.join(folder_path, 'transcriptions'), valid_extensions)
    analyze_transcriptions(os.path.join(folder_path, 'transcriptions'),os.path.join(folder_path, 'formatos'))

if __name__ == "__main__":
    main()



