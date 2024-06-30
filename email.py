import smtplib # Para enviar correos electrónicos
import email.mime.multipart # Para crear mensajes de correo electrónico
import email.mime.base # Para adjuntar archivos al correo electrónico
import csv # Para leer archivos CSV
import asyncio # Para leer archivos CSV
import locale # Para formatear fechas en español
from email.mime.text import MIMEText # Para crear mensajes de texto plano
from docx import Document # Para leer y escribir archivos Word
from docx2pdf import convert # Para convertir archivos Word a PDF
from datetime import datetime # Para trabajar con fechas
from email.mime.image import MIMEImage # Para adjuntar imágenes al correo electrónico
from decouple import config # Para leer variables de entorno desde un archivo .env (opcional)
from email.mime.multipart import MIMEMultipart # Para crear mensajes de correo electrónico con varias partes
from email.mime.application import MIMEApplication # Para adjuntar archivos al correo electrónico

async def leer_csv(ruta_csv):
    destinatarios = []
    datos_csv = []
    with open(ruta_csv, 'r', encoding='utf-8-sig') as archivo_csv:
        lector_csv = csv.DictReader(archivo_csv, delimiter=';') 
        print(lector_csv.fieldnames)  # Imprimir los encabezados
        for fila in lector_csv:
            empresa = fila['Empresa']
            correo = fila['Email titular']
            nombre = fila['Nombre titular']
            identificacion = fila['Identificacion']
            nit = fila['Nit']
            apartamento = fila['Apartamento']
            torre = fila['Torre']
            macroproyecto = fila['Macroproyecto']
            valor = fila['Valor']
            clientes = fila['Clientes']
            
            destinatarios.append((nombre, correo))
            datos_csv.append((empresa, nombre, correo, identificacion, nit, apartamento, torre, macroproyecto, valor, clientes))
    return destinatarios, datos_csv


def formatear_fecha(fecha):
    # Convertir la cadena de fecha en un objeto datetime
    fecha_obj = datetime.strptime(fecha, '%d/%m/%Y')  # Ahora asume que la fecha está en el formato 'DD/MM/AAAA'

    # Formatear la fecha
    mes = fecha_obj.strftime('%B')  # Mes como nombre completo ('Enero')
    dia = fecha_obj.strftime('%d')  # Día del mes como número decimal
    año = fecha_obj.strftime('%Y')  # Año con siglo como número decimal

    return mes, dia, año

def obtener_fecha_actual():
    # Establecer la localidad a español
    locale.setlocale(locale.LC_TIME, 'es_ES.UTF-8')

    # Obtener la fecha actual
    fecha_actual = datetime.now()

    # Formatear la fecha
    mes_exp = fecha_actual.strftime('%B')  # Mes como nombre completo ('Enero')
    dia_exp = fecha_actual.strftime('%d')  # Día del mes como número decimal
    año_exp = fecha_actual.strftime('%Y')  # Año con siglo como número decimal

    return mes_exp, dia_exp, año_exp

def crear_cuerpo_mensaje(nombre):
    cuerpo = f"""
    <!DOCTYPE html>
    <html>
    <head>
    <style>
        a {{
            display: block;
            margin-left: auto;
            margin-right: auto;
        }}
    </style>
    </head>
    <body>
        <div style="display: flex; justify-content: center;">
            <a href="https://static-pages.cusezar.com/notificaciones-sac/">
                <img src="cid:imagen" alt="Cusezar" width="597" height="1134" style="border: none;">
            </a>
        </div>
    </body>
    </html>
    """  # noqa: F541
    return cuerpo

def crear_mensaje_html(cuerpo):
    mensaje = MIMEMultipart("alternative")
    parte_html = MIMEText(cuerpo, "html")
    mensaje.attach(parte_html)
    return mensaje


def agregar_imagen(mensaje, ruta_imagen):
    with open(ruta_imagen, 'rb') as file:
        img_data = file.read()
    if ruta_imagen.lower().endswith('.jpg') or ruta_imagen.lower().endswith('.jpeg'):
        img = MIMEImage(img_data, 'jpeg')
    elif ruta_imagen.lower().endswith('.png'):
        img = MIMEImage(img_data, 'png')
    else:
        raise ValueError('Unsupported image type')
    img.add_header('Content-ID', '<imagen>')  # se puede usar cualquier nombre en el ID en lugar de 'imagen'
    mensaje.attach(img)
    
def crear_mensaje(remitente, destinatario, mensaje_html, ruta_archivo, nombre_archivo, ruta_imagen):
    mensaje = email.mime.multipart.MIMEMultipart()
    mensaje['From'] = remitente
    mensaje['To'] = destinatario
    mensaje['Subject'] = "Fe de Erratas: Certificado Anual de Declaración de Renta Cusezar"
    mensaje.attach(mensaje_html)  # Adjuntar el mensaje HTML
    

    # Adjuntar el archivo PDF
    try:
        with open(ruta_archivo, 'rb') as archivo:
            adjunto = MIMEApplication(archivo.read(),_subtype="pdf")
            adjunto.add_header('Content-Disposition', 'attachment', filename=str(nombre_archivo))
            mensaje.attach(adjunto)
            agregar_imagen(mensaje, ruta_imagen)
    except Exception as e:
        print(f"Error al adjuntar el archivo {nombre_archivo}: {e}")

    return mensaje.as_string()



async def enviar_correo(server, remitente, destinatario, mensaje):
    try:
        server.sendmail(remitente, destinatario, mensaje)
    except smtplib.SMTPServerDisconnected:
        # La conexión con el servidor SMTP se ha perdido, reconectar y reintentar
        server = smtplib.SMTP('smtp-mail.outlook.com', 587)
        server.starttls()
        server.login(config('email_cuz'), config('pass_cuz'))
        server.sendmail(remitente, destinatario, mensaje)
    except smtplib.SMTPRecipientsRefused as e:
        print(f"No se pudo enviar el correo a {e.recipients}")



async def reemplazar_datos_doc(doc, *args):
    # args es una lista de argumentos que contiene pares de texto antiguo y nuevo
    for i in range(0, len(args), 2):
        old_text = args[i]
        new_text = args[i+1]
        for paragraph in doc.paragraphs:
            if old_text in paragraph.text:
                paragraph.text = paragraph.text.replace(old_text, new_text)

    return doc

def guardar_resultados(resultados):
    # Guardar resultados de envío de correos en un archivo CSV
    with open('cvs\\resultados_envio_correos.csv', 'w', newline='') as file:
        writer = csv.writer(file, delimiter=';')
        writer.writerow(["Nombre titular", "Correo", "Estado"])  # Escribir los encabezados de las columnas

        for nombre, correo, estado in resultados:
            writer.writerow([nombre, correo, estado])  # Escribir los datos

async def main():
    # Crea la conexión SMTP
    server = smtplib.SMTP('smtp-mail.outlook.com', 587)
    server.connect('smtp-mail.outlook.com', 587)

    correo = config('email_cuz')
    pas = config('pass_cuz')

    # Inicia sesión en el correo
    server.starttls()
    server.login(correo, pas)

    # Definir el remitente del correo electrónico
    remitente = config('email_cuz')

    # Leer el archivo CSV con los destinatarios
    ruta_csv = 'cvs\\BASE DECLARACIÓN DE RENTA4.csv'
    destinatarios, datos_csv = await leer_csv(ruta_csv)


    # Lista para almacenar las tareas de envío de correos electrónicos
    # tareas_envio_correos = []

    # Iterar sobre cada fila del archivo CSV y enviar un correo a cada destinatario
    for i, (empresa, nombre, correo, identificacion, nit, apartamento, torre, macroproyecto, valor, clientes) in enumerate(datos_csv): 
        
        # Leer el archivo Word original en cada iteración
        doc = Document('docs\\plantilla.docx')
        mes_exp, dia_exp, año_exp = obtener_fecha_actual() # Obtener la fecha actual
        
        
        # marcadores de posición con valores correspondientes
        doc = await reemplazar_datos_doc(doc, '{{empresa}}', empresa, '{{nombre}}', nombre, '{{clientes}}', clientes, '{{identificacion}}', identificacion, '{{nit}}', nit, '{{apartamento}}', apartamento,'{{torre}}', torre,'{{proyecto}}', macroproyecto, '{{dia_exp}}', dia_exp, '{{mes_exp}}', mes_exp, '{{año_exp}}', año_exp, '{{valor}}', valor)
        
        # Guardar el archivo Word modificado con un nombre único
        doc.save(f'docs\\modified_plantilla.docx')  # noqa: F541
        
        # Convertir el archivo Word modificado a PDF
        nombre_pdf = nombre.replace(" ", "_") # Reemplazar espacios en blanco con guiones bajos
        empresa_pdf = empresa.replace(" ", "_")[:2] # Reemplazar espacios en blanco con guiones bajos
        proyecto_pdf = macroproyecto.replace(" ", "_") # Reemplazar espacios en blanco con guiones bajos y tomar los dos primeros caracteres
        agruapcion_pdf = apartamento.replace(" ", "_") # Reemplazar espacios en blanco con guiones bajos
        nombre_archivo_pdf = f'{empresa_pdf}_{nombre_pdf}_{proyecto_pdf}_{agruapcion_pdf}.pdf' # Nombre del archivo PDF
        ruta_archivo_pdf = f'docs\\pdf\\{nombre_archivo_pdf}' # Ruta del archivo PDF
        convert(f'docs\\modified_plantilla.docx', ruta_archivo_pdf) # Convertir y guardar el archivo PDF  # noqa: F541

        # blob=open(f'docs\\pdf\\plantilla_{nombre_pdf}.pdf','r',encoding='latin-1')
        # blob.read()

        # requests.post() # generar token, hacer post a la api sharepoint

        # Crear un nuevo mensaje para cada destinatario
        ruta_archivo = ruta_archivo_pdf # Utilizar la ruta del archivo PDF recién creado para leer el archivo
        ruta_imagen = 'img\\Declaración de rentas.jpg'
        cuerpo = crear_cuerpo_mensaje(nombre) 
        mensaje_html = crear_mensaje_html(cuerpo) 
        mensaje = crear_mensaje(remitente, correo, mensaje_html, ruta_archivo, nombre_archivo_pdf, ruta_imagen) # Pasar el nombre del archivo sin la ruta al crear el mensaje


        # Agregar la tarea de envío de correo a la lista
        await enviar_correo(server, remitente, correo, mensaje)
        

    # Esperar a que todas las tareas de envío de correos terminen y recoger los resultados
    # resultados_envio_correos = await asyncio.gather(*tareas_envio_correos)
    
    # Cerrar la conexión SMTP
    # Cerrar la conexión SMTP
    try:
        server.quit()
    except smtplib.SMTPServerDisconnected:
        pass

    # llamo la funcion para guardar el envío de correos en un archivo CSV 
    # guardar_resultados([(nombre, correo, estado) for (nombre, correo), estado in zip(destinatarios, resultados_envio_correos)])
    # se guarda resultados en un archivo CSV con los datos de los destinatarios y el estado del envío del correo

# Ejecutar la función principal
asyncio.run(main())