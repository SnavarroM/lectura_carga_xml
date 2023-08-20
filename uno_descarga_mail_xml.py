import time
import os
import email
import datetime
from imbox import Imbox

# Configuración de conexión IMAP
host = "imap.gmail.com"
username = "copiafactura@cenabast.cl"
password = 'Cena@2023'
download_folder = r'C:\Users\chilo\Downloads\ejecucion_xml-20230819T172426Z-001\ejecucion_xml\script xml\data'

# Verificar si la carpeta de destino existe, si no, crearla
if not os.path.exists(download_folder):
    try:
        os.makedirs(download_folder)
        print("Carpeta de destino creada correctamente.")
    except Exception as e:
        print(f"Error al crear la carpeta de destino: {str(e)}")

# Variable para mantener el contador autoincremental
counter = 1

while True:
    # Obtener la fecha y hora actual
    now = datetime.datetime.now()

    # Verificar si es la hora de eliminar los archivos
    if now.hour == 0 and now.minute == 1:
        print("Eliminando archivos de la carpeta...")
        try:
            for file in os.listdir(download_folder):
                file_path = os.path.join(download_folder, file)
                if os.path.isfile(file_path):
                    os.remove(file_path)
            print("Archivos eliminados correctamente.")
        except Exception as e:
            print(f"Error al eliminar los archivos: {str(e)}")

    # Conexión al servidor IMAP
    with Imbox(host, username=username, password=password, ssl=True, ssl_context=None, starttls=False) as mail:
        # Obtener todos los correos electrónicos del día actual
        messages = mail.messages(date__on=datetime.date.today())
        print("Obteniendo correos electrónicos del día...")

        # Recorrer todos los correos electrónicos
        for uid, message in messages:
            print(f"Procesando correo electrónico con UID: {uid}")
            # Verificar si el correo electrónico tiene archivos adjuntos
            if message.attachments:
                print("Correo electrónico tiene archivos adjuntos...")
                for attachment in message.attachments:
                    print("Encontrado archivo adjunto...")
                    # Obtener el nombre del archivo adjunto
                    filename = attachment.get('filename')
                    # Decodificar el nombre del archivo adjunto
                    decoded_filename = email.header.decode_header(filename)[0][0]
                    if isinstance(decoded_filename, bytes):
                        decoded_filename = decoded_filename.decode()
                    # Verificar si el archivo adjunto tiene extensión .xml o .XML
                    if decoded_filename.lower().endswith(('.xml', '.XML')):
                        # Generar el nombre único para el archivo XML
                        file_date = datetime.date.today().strftime("%d_%m_%Y")
                        unique_name = f"{file_date}_{counter}.xml"
                        counter += 1

                        # Guardar el archivo adjunto en la carpeta de destino con el nombre único
                        filepath = os.path.join(download_folder, unique_name)
                        print(f"Guardando archivo adjunto en: {filepath}")

                        try:
                            with open(filepath, 'wb') as f:
                                f.write(attachment.get('content').read())
                            print("Archivo adjunto guardado exitosamente.")
                        except FileNotFoundError:
                            print("Error: No se pudo encontrar la carpeta de destino.")
                        except Exception as e:
                            print(f"Error al guardar el archivo adjunto: {str(e)}")
                    else:
                        print("El archivo adjunto no es un archivo XML. No se descargará.")
            else:
                print("El correo electrónico no tiene archivos adjuntos.")

    # Esperar 60 segundos antes de repetir el proceso
    time.sleep(60)