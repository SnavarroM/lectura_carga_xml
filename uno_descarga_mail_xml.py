import time
import os
import email
import datetime
import imaplib
import base64

# Configuración de conexión IMAP
host = "imap.gmail.com"
username = "pruebafactura@cenabast.cl"
password = 'mbbv kxua vrsk wcdo'
download_folder = r'C:\Users\snavarro\Desktop\proyecto\factura_xml-main_2\carga_produccion\data'

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
    mail = imaplib.IMAP4_SSL(host)
    mail.login(username, password)

    # Seleccionar la bandeja de entrada
    mail.select("inbox")

    # Buscar todos los correos electrónicos no leídos
    status, messages = mail.search(None, "UNSEEN")

    # Recorrer todos los correos electrónicos
    for num in messages[0].split():
        print(f"Procesando correo electrónico con número: {num}")

        # Obtener el correo electrónico
        status, message = mail.fetch(num, "(RFC822)")

        # Decodificar el correo electrónico
        message = email.message_from_bytes(message[0][1])

        # Verificar si el correo electrónico tiene archivos adjuntos
        if message.get_content_maintype() == 'multipart':
            for part in message.walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue

                filename = part.get_filename()

                # Decodificar el nombre del archivo adjunto
                if filename:
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
                                f.write(base64.b64decode(part.get_payload()))
                            print("Archivo adjunto guardado exitosamente.")

                            # Marcar el correo electrónico como leído
                            mail.store(num, '+FLAGS', '\\Seen')
                        except FileNotFoundError:
                            print("Error: No se pudo encontrar la carpeta de destino.")
                        except Exception as e:
                            print(f"Error al guardar el archivo adjunto: {str(e)}")
                    else:
                        print("El archivo adjunto no es un archivo XML. No se descargará.")
        else:
            print("El correo electrónico no tiene archivos adjuntos.")

    # Cerrar la conexión al servidor IMAP
    mail.close()
    mail.logout()

    # Esperar 60 segundos antes de repetir el proceso
    time.sleep(60)
