# ==============================================================================
# Sistema de Gestión y Envío de Correos Institucionales - IEEE
# Autor: Julio Ricardo Barrios Amador
# Descripción: Aplicación Flask para gestionar contactos institucionales
#              y enviar correos electrónicos de manera individual y controlada.
#              Usa archivo Excel como base de datos directa.
# ==============================================================================

import os
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash

# ==============================================================================
# CONFIGURACIÓN DEL SISTEMA
# ==============================================================================

# Credenciales SMTP - Modificar antes de usar en producción
EMAIL_REMITENTE = ""
EMAIL_PASSWORD = ""

# Configuración del servidor SMTP (por defecto Gmail con TLS)
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587

# Rutas de archivos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "contactos.xlsx")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
IEEE_LOGO_PATH = os.path.join(IMAGES_DIR, "ieee.png")
IEEE_CS_LOGO_PATH = os.path.join(IMAGES_DIR, "ieee cs imagen.png")

# Clave secreta para sesiones Flask (cambiar en producción)
SECRET_KEY = "ieee-correos-institucionales-clave-secreta-2026"

# ==============================================================================
# INICIALIZACIÓN DE LA APLICACIÓN FLASK
# ==============================================================================

app = Flask(__name__)
app.secret_key = SECRET_KEY


# ==============================================================================
# FUNCIONES DE MANEJO DE EXCEL
# ==============================================================================

def inicializar_excel():
    """
    Verifica que el archivo Excel exista y tenga las columnas necesarias.
    Si no existe, crea un archivo vacío con la estructura correcta.
    """
    if not os.path.exists(EXCEL_PATH):
        # Crear archivo vacío con estructura
        df = pd.DataFrame(columns=["empresa", "contacto", "correo", "enviado", "fecha_envio"])
        df.to_excel(EXCEL_PATH, index=False, engine="openpyxl")
        return True
    
    # Verificar y agregar columnas faltantes
    try:
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
        
        # Agregar columna "enviado" si no existe
        if "enviado" not in df.columns:
            df["enviado"] = "no"
        
        # Agregar columna "fecha_envio" si no existe
        if "fecha_envio" not in df.columns:
            df["fecha_envio"] = ""
        
        # Asegurar que las columnas básicas existan
        columnas_requeridas = ["empresa", "contacto", "correo"]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = ""
        
        # Reordenar columnas
        df = df[["empresa", "contacto", "correo", "enviado", "fecha_envio"]]
        
        # Limpiar valores NaN
        df = df.fillna("")
        
        # Asegurar que "enviado" solo tenga "si" o "no"
        df["enviado"] = df["enviado"].apply(lambda x: "si" if str(x).lower() in ["si", "sí", "yes", "1"] else "no")
        
        # Guardar cambios
        df.to_excel(EXCEL_PATH, index=False, engine="openpyxl")
        return True
        
    except Exception as e:
        print(f"Error al inicializar Excel: {e}")
        return False


def leer_contactos():
    """
    Lee todos los contactos del archivo Excel.
    Retorna un DataFrame de pandas.
    """
    if not os.path.exists(EXCEL_PATH):
        return pd.DataFrame(columns=["empresa", "contacto", "correo", "enviado", "fecha_envio"])
    
    try:
        df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
        df = df.fillna("")
        
        # Asegurar que las columnas existan
        if "enviado" not in df.columns:
            df["enviado"] = "no"
        if "fecha_envio" not in df.columns:
            df["fecha_envio"] = ""
        
        return df
    except Exception as e:
        print(f"Error al leer Excel: {e}")
        return pd.DataFrame(columns=["empresa", "contacto", "correo", "enviado", "fecha_envio"])


def guardar_contactos(df):
    """
    Guarda el DataFrame completo en el archivo Excel.
    Sobrescribe el archivo existente.
    """
    try:
        df = df.fillna("")
        df.to_excel(EXCEL_PATH, index=False, engine="openpyxl")
        return True
    except Exception as e:
        print(f"Error al guardar Excel: {e}")
        return False


def obtener_contacto_por_indice(indice):
    """
    Obtiene un contacto específico por su índice (número de fila).
    Retorna un diccionario con los datos del contacto.
    """
    df = leer_contactos()
    if indice < 0 or indice >= len(df):
        return None
    
    contacto = df.iloc[indice].to_dict()
    contacto["indice"] = indice
    return contacto


def agregar_contacto(empresa, contacto, correo):
    """
    Agrega un nuevo contacto al archivo Excel.
    Verifica duplicados por correo.
    """
    df = leer_contactos()
    
    # Verificar duplicados
    if correo.lower() in df["correo"].str.lower().values:
        return False
    
    # Crear nuevo registro
    nuevo = pd.DataFrame([{
        "empresa": empresa.strip(),
        "contacto": contacto.strip(),
        "correo": correo.strip().lower(),
        "enviado": "no",
        "fecha_envio": ""
    }])
    
    # Concatenar y guardar
    df = pd.concat([df, nuevo], ignore_index=True)
    return guardar_contactos(df)


def actualizar_estado_envio(indice, enviado, fecha=None):
    """
    Actualiza el estado de envío de un contacto específico.
    enviado: "si" o "no"
    fecha: fecha y hora del envío (opcional)
    """
    df = leer_contactos()
    
    if indice < 0 or indice >= len(df):
        return False
    
    df.at[indice, "enviado"] = enviado
    df.at[indice, "fecha_envio"] = fecha if fecha else ""
    
    return guardar_contactos(df)


def eliminar_contacto(indice):
    """
    Elimina un contacto del archivo Excel por su índice.
    """
    df = leer_contactos()
    
    if indice < 0 or indice >= len(df):
        return False
    
    df = df.drop(indice).reset_index(drop=True)
    return guardar_contactos(df)


def obtener_estadisticas():
    """
    Calcula estadísticas de los contactos.
    """
    df = leer_contactos()
    
    total = len(df)
    enviados = len(df[df["enviado"].str.lower() == "si"])
    pendientes = len(df[df["enviado"].str.lower() == "no"])
    
    return {
        "total": total,
        "enviados": enviados,
        "pendientes": pendientes,
        "errores": 0  # No usamos errores, solo si/no
    }


# ==============================================================================
# FUNCIONES DE VALIDACIÓN
# ==============================================================================

def validar_correo(correo):
    """
    Valida el formato de un correo electrónico usando expresión regular.
    Retorna True si el formato es válido, False en caso contrario.
    """
    patron = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return re.match(patron, correo) is not None


# ==============================================================================
# FUNCIONES DE GENERACIÓN DE CORREO
# ==============================================================================

def generar_saludo(empresa, contacto):
    """
    Genera el saludo dinámico según la lógica institucional (lenguaje neutro):
    - Si contacto tiene valor: "Saludos, {contacto},"
    - Si contacto está vacío: "Saludos,"
    """
    if contacto and contacto.strip():
        return f"Saludos, {contacto.strip()},"
    else:
        return f"Saludos,"


def generar_cuerpo_html(saludo, empresa, incluir_logo_ieee=False, incluir_logo_ieee_cs=False):
    """
    Genera el cuerpo del correo en formato HTML con diseño institucional.
    Incluye saludo dinámico, párrafo institucional y cierre formal.
    """
    logo_ieee_html = ""
    if incluir_logo_ieee:
        logo_ieee_html = """
            <img src=\"cid:logo_ieee\" alt=\"Logo IEEE\" style=\"display: block; width: 100%; max-width: 160px; height: auto;\">
        """

    logo_ieee_cs_html = ""
    if incluir_logo_ieee_cs:
        logo_ieee_cs_html = """
            <img src=\"cid:logo_ieee_cs\" alt=\"Logo IEEE Computer Society\" style=\"display: block; width: 100%; max-width: 160px; height: auto;\">
        """

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            @media only screen and (max-width: 620px) {{
                .email-card {{
                    width: 100% !important;
                }}

                .header-cell {{
                    padding: 20px 16px !important;
                }}

                .logo-cell-left {{
                    padding-left: 8px !important;
                    padding-right: 12px !important;
                }}

                .logo-cell-right {{
                    padding-left: 12px !important;
                    padding-right: 8px !important;
                }}

                .logo-cell-left img,
                .logo-cell-right img {{
                    max-width: 110px !important;
                }}
            }}
        </style>
    </head>
    <body style="margin: 0; padding: 0; background-color: #f4f6f9; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
        <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="background-color: #f4f6f9; padding: 30px 0;">
            <tr>
                <td align="center">
                    <table role="presentation" class="email-card" width="600" cellspacing="0" cellpadding="0" border="0" style="background-color: #ffffff; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 12px rgba(0,0,0,0.08);">
                        <!-- Header institucional -->
                        <tr>
                            <td class="header-cell" style="background: linear-gradient(135deg, #00629B 0%, #004A7C 100%); padding: 28px 40px; text-align: center;">
                                <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 16px;">
                                    <tr>
                                        <td class="logo-cell-left" align="left" style="width: 50%; vertical-align: middle; padding-left: 12px; padding-right: 18px;">
                                            {logo_ieee_html}
                                        </td>
                                        <td class="logo-cell-right" align="right" style="width: 50%; vertical-align: middle; padding-left: 18px; padding-right: 12px;">
                                            {logo_ieee_cs_html}
                                        </td>
                                    </tr>
                                </table>
                                <h1 style="margin: 0; color: #ffffff; font-size: 22px; font-weight: 600; letter-spacing: 0.5px;">
                                    IEEE Computer Society
                                </h1>
                                <p style="margin: 6px 0 0 0; color: rgba(255,255,255,0.85); font-size: 13px; letter-spacing: 0.3px;">
                                    Instituto Tecnológico de Costa Rica &mdash; Campus Central Cartago
                                </p>
                            </td>
                        </tr>
                        <!-- Cuerpo del correo -->
                        <tr>
                            <td style="padding: 36px 40px 20px 40px; background-color: #0b3f66;">
                                <p style="margin: 0 0 20px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    {saludo}
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Reciba un cordial saludo.
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Mi nombre es Julio Barrios y formo parte de la Junta Directiva del IEEE Computer Society 
                                    en el Instituto Tecnológico de Costa Rica, Campus Central Cartago. Actualmente estamos impulsando distintas iniciativas y 
                                    actividades dirigidas a estudiantes de todas las carreras del TEC, con el objetivo de 
                                    fortalecer su <strong>desarrollo profesional</strong>, <strong>liderazgo</strong> y <strong>crecimiento integral</strong>.
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Nos gustaría invitar a <strong>{empresa.strip()}</strong> a participar como <strong>patrocinador</strong> de estas iniciativas estudiantiles. 
                                    Buscamos establecer una relación colaborativa que permita a su organización tener mayor presencia dentro 
                                    del Tecnológico y crear una conexión más cercana entre la empresa y la comunidad estudiantil. 
                                    <strong style="color: #ffd166;">Específicamente, estamos interesados en:</strong>
                                </p>
                                <table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin: 0 0 18px 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;">
                                    <tr>
                                        <td style="width: 18px; vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffd166; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0; font-weight: 700;">&#8226;</td>
                                        <td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0;"><strong>Charlas técnicas o talleres</strong> impartidos por profesionales de su empresa</td>
                                    </tr>
                                    <tr>
                                        <td style="width: 18px; vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffd166; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0; font-weight: 700;">&#8226;</td>
                                        <td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0;"><strong>Participación en ferias de empleo</strong> con stands institucionales</td>
                                    </tr>
                                    <tr>
                                        <td style="width: 18px; vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffd166; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0; font-weight: 700;">&#8226;</td>
                                        <td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; padding: 0 0 8px 0;"><strong>Actividades de networking</strong> y acercamiento con talento joven comprometido</td>
                                    </tr>
                                    <tr>
                                        <td style="width: 18px; vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffd166; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; font-weight: 700;">&#8226;</td>
                                        <td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;"><strong>Apoyo en eventos</strong> y competencias académicas</td>
                                    </tr>
                                </table>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Nuestro objetivo es formar una relación duradera que beneficie tanto a los estudiantes como a {empresa.strip()}.
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    En caso de que esta solicitud deba ser gestionada por otra persona o departamento, le 
                                    agradeceríamos si pudiera <strong>indicarnos el contacto correspondiente</strong> o <strong>reenviar este mensaje</strong>.
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Con gusto podemos <strong>coordinar una reunión</strong> para compartir más detalles sobre el alcance de las 
                                    actividades y las oportunidades de participación.
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Para mayor información pueden responder a este correo o contactarnos directamente a:
                                </p>
                                <p style="margin: 0 0 20px 0; font-size: 15px; line-height: 1.7; color: #00d4ff; font-weight: 700; text-align: center;">
                                    sbc-tec-cs@ieee.org
                                </p>
                                <p style="margin: 0 0 18px 0; font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    Agradecemos su tiempo y quedamos atentos.
                                </p>
                            </td>
                        </tr>
                        <!-- Cierre formal -->
                        <tr>
                            <td style="padding: 0 40px 36px 40px; background-color: #0b3f66;">
                                <p style="margin: 0 0 4px 0; font-size: 15px; color: #ffffff;">
                                    Atentamente,
                                </p>
                                <br>
                                <p style="margin: 0 0 2px 0; font-size: 15px; font-weight: 700; color: #ffffff;">
                                    Julio Ricardo Barrios Amador
                                </p>
                                <p style="margin: 0 0 1px 0; font-size: 13px; color: #dce8f4; letter-spacing: 0.3px;">
                                    Section Student Representative (SSR) | IEEE Costa Rica Section
                                </p>
                                <p style="margin: 0 0 1px 0; font-size: 13px; color: #dce8f4; letter-spacing: 0.3px;">
                                    Vocal 2 | IEEE Computer Society - Instituto Tecnológico de Costa Rica
                                </p>
                                <p style="margin: 0; font-size: 12px; color: #dce8f4; letter-spacing: 0.3px;">
                                    IEEE Member: 101781510
                                </p>
                                <p style="margin: 6px 0 0 0; font-size: 12px; color: #00d4ff; letter-spacing: 0.3px; font-weight: 700;">
                                    julio.barrios@ieee.org
                                </p>
                            </td>
                        </tr>
                        <!-- Footer -->
                        <tr>
                            <td style="background-color: #f8f9fa; padding: 18px 40px; text-align: center; border-top: 1px solid #e9ecef;">
                                <p style="margin: 0; font-size: 11px; color: #adb5bd; letter-spacing: 0.2px;">
                                    &copy; {datetime.now().year} IEEE Computer Society – Instituto Tecnológico de Costa Rica &mdash; Todos los derechos reservados.
                                </p>
                            </td>
                        </tr>
                    </table>
                </td>
            </tr>
        </table>
    </body>
    </html>
    """
    return html


# ==============================================================================
# FUNCIONES DE ENVÍO DE CORREO
# ==============================================================================

def enviar_correo(destinatario, empresa, contacto):
    """
    Envía un correo electrónico individual usando SMTP con TLS.
    Genera el saludo dinámico y el cuerpo HTML institucional.
    Retorna un diccionario con el resultado del envío.
    """
    # Verificar disponibilidad de logos para incrustarlos en el HTML.
    incluir_logo_ieee = os.path.exists(IEEE_LOGO_PATH)
    incluir_logo_ieee_cs = os.path.exists(IEEE_CS_LOGO_PATH)

    # Generar contenido del correo
    saludo = generar_saludo(empresa, contacto)
    cuerpo_html = generar_cuerpo_html(
        saludo,
        empresa,
        incluir_logo_ieee=incluir_logo_ieee,
        incluir_logo_ieee_cs=incluir_logo_ieee_cs
    )

    # Usamos multipart/related para soportar imágenes inline referenciadas por CID.
    mensaje = MIMEMultipart("related")
    mensaje["From"] = EMAIL_REMITENTE
    mensaje["To"] = destinatario
    mensaje["Subject"] = "Invitación a colaborar - Iniciativas estudiantiles TEC"

    # El HTML vive dentro de multipart/alternative para compatibilidad con clientes.
    parte_alternativa = MIMEMultipart("alternative")
    parte_html = MIMEText(cuerpo_html, "html", "utf-8")
    parte_alternativa.attach(parte_html)
    mensaje.attach(parte_alternativa)

    # Adjuntar logos inline si están disponibles.
    if incluir_logo_ieee:
        with open(IEEE_LOGO_PATH, "rb") as logo_ieee:
            img_ieee = MIMEImage(logo_ieee.read())
            img_ieee.add_header("Content-ID", "<logo_ieee>")
            img_ieee.add_header("Content-Disposition", "inline", filename="ieee.png")
            mensaje.attach(img_ieee)

    if incluir_logo_ieee_cs:
        with open(IEEE_CS_LOGO_PATH, "rb") as logo_ieee_cs:
            img_ieee_cs = MIMEImage(logo_ieee_cs.read())
            img_ieee_cs.add_header("Content-ID", "<logo_ieee_cs>")
            img_ieee_cs.add_header("Content-Disposition", "inline", filename="ieee_cs.png")
            mensaje.attach(img_ieee_cs)

    try:
        # Conexión SMTP con TLS
        servidor = smtplib.SMTP(SMTP_SERVER, SMTP_PORT)
        servidor.ehlo()
        servidor.starttls()
        servidor.ehlo()
        servidor.login(EMAIL_REMITENTE, EMAIL_PASSWORD)
        servidor.sendmail(EMAIL_REMITENTE, destinatario, mensaje.as_string())
        servidor.quit()

        return {"exito": True, "mensaje": f"Correo enviado exitosamente a {destinatario}"}

    except smtplib.SMTPAuthenticationError:
        return {"exito": False, "mensaje": "Error de autenticación SMTP. Verifique las credenciales."}
    except smtplib.SMTPRecipientsRefused:
        return {"exito": False, "mensaje": f"El destinatario {destinatario} fue rechazado por el servidor."}
    except smtplib.SMTPServerDisconnected:
        return {"exito": False, "mensaje": "El servidor SMTP se desconectó inesperadamente."}
    except smtplib.SMTPException as e:
        return {"exito": False, "mensaje": f"Error SMTP: {str(e)}"}
    except ConnectionRefusedError:
        return {"exito": False, "mensaje": "No se pudo conectar al servidor SMTP. Verifique la configuración."}
    except Exception as e:
        return {"exito": False, "mensaje": f"Error inesperado al enviar correo: {str(e)}"}


# ==============================================================================
# RUTAS DE LA APLICACIÓN
# ==============================================================================

@app.route("/")
def dashboard():
    """
    Ruta principal - Dashboard de gestión de contactos.
    Muestra la tabla de contactos, estadísticas y formulario de agregar.
    """
    df = leer_contactos().fillna("")

    # Parámetros de filtro y paginación
    busqueda = request.args.get("q", "").strip()
    estado = request.args.get("estado", "todos").strip().lower()
    ordenar_por = request.args.get("sort", "empresa").strip().lower()
    direccion = request.args.get("dir", "asc").strip().lower()

    try:
        pagina = int(request.args.get("page", "1"))
    except ValueError:
        pagina = 1
    pagina = max(1, pagina)

    # Construir dataset de trabajo con ID y estado legible
    df = df.reset_index().rename(columns={"index": "id"})
    df["estado_envio"] = df["enviado"].astype(str).str.lower().apply(
        lambda x: "Enviado" if x == "si" else "Pendiente"
    )

    # Filtrado por nombre/empresa/correo
    if busqueda:
        patron = re.escape(busqueda)
        mascara_busqueda = (
            df["empresa"].astype(str).str.contains(patron, case=False, na=False)
            | df["contacto"].astype(str).str.contains(patron, case=False, na=False)
            | df["correo"].astype(str).str.contains(patron, case=False, na=False)
        )
        df = df[mascara_busqueda]

    # Filtrado por estado
    if estado in ("pendiente", "enviado"):
        df = df[df["estado_envio"].str.lower() == estado]
    else:
        estado = "todos"

    # Ordenamiento
    if direccion not in ("asc", "desc"):
        direccion = "asc"

    ascendente = direccion == "asc"
    if ordenar_por == "estado":
        df = df.sort_values(by=["estado_envio", "empresa"], ascending=[ascendente, True], kind="mergesort")
    elif ordenar_por == "fecha_envio":
        # Para fecha vacía, usamos un extremo para que queden al final según el orden.
        fecha_tmp = pd.to_datetime(df["fecha_envio"], errors="coerce")
        if ascendente:
            fecha_tmp = fecha_tmp.fillna(pd.Timestamp.max)
        else:
            fecha_tmp = fecha_tmp.fillna(pd.Timestamp.min)
        df = df.assign(_fecha_sort=fecha_tmp).sort_values(by=["_fecha_sort", "empresa"], ascending=[ascendente, True], kind="mergesort").drop(columns=["_fecha_sort"])
    else:
        ordenar_por = "empresa"
        df = df.sort_values(by=["empresa", "id"], ascending=[ascendente, True], kind="mergesort")

    total_filtrados = len(df)
    tam_pagina = 50
    total_paginas = max(1, (total_filtrados + tam_pagina - 1) // tam_pagina)
    pagina = min(pagina, total_paginas)

    inicio = (pagina - 1) * tam_pagina
    fin = inicio + tam_pagina
    df_pagina = df.iloc[inicio:fin]

    # Convertir DataFrame filtrado/paginado a lista de diccionarios
    contactos = []
    for _, row in df_pagina.iterrows():
        contacto = row.to_dict()
        contactos.append(contacto)

    inicio_paginas = max(1, pagina - 2)
    fin_paginas = min(total_paginas, pagina + 2)

    estadisticas = obtener_estadisticas()
    return render_template(
        "index.html",
        contactos=contactos,
        estadisticas=estadisticas,
        filtros={
            "q": busqueda,
            "estado": estado,
            "sort": ordenar_por,
            "dir": direccion,
        },
        paginacion={
            "pagina_actual": pagina,
            "tam_pagina": tam_pagina,
            "total_filtrados": total_filtrados,
            "total_paginas": total_paginas,
            "tiene_anterior": pagina > 1,
            "tiene_siguiente": pagina < total_paginas,
            "pagina_anterior": pagina - 1,
            "pagina_siguiente": pagina + 1,
            "paginas_visibles": list(range(inicio_paginas, fin_paginas + 1)),
        },
    )


@app.route("/agregar", methods=["POST"])
def agregar():
    """
    Ruta POST para agregar una nueva empresa/contacto manualmente.
    Valida los datos antes de insertar en el Excel.
    """
    empresa = request.form.get("empresa", "").strip()
    contacto = request.form.get("contacto", "").strip()
    correo = request.form.get("correo", "").strip().lower()

    # Validar campos obligatorios
    if not empresa:
        flash("El campo 'Empresa' es obligatorio.", "error")
        return redirect(url_for("dashboard"))

    if not correo:
        flash("El campo 'Correo' es obligatorio.", "error")
        return redirect(url_for("dashboard"))

    # Validar formato de correo
    if not validar_correo(correo):
        flash("El formato del correo electrónico no es válido.", "error")
        return redirect(url_for("dashboard"))

    # Intentar agregar
    if agregar_contacto(empresa, contacto, correo):
        flash(f"Contacto '{empresa}' agregado exitosamente.", "success")
    else:
        flash(f"El correo '{correo}' ya existe en el archivo.", "warning")

    return redirect(url_for("dashboard"))


@app.route("/enviar/<int:id>")
def enviar(id):
    """
    Ruta para enviar correo individual a un contacto específico.
    Actualiza el estado de envío según el resultado.
    """
    contacto = obtener_contacto_por_indice(id)

    if not contacto:
        flash("Contacto no encontrado.", "error")
        return redirect(url_for("dashboard"))

    # Ejecutar envío de correo
    resultado = enviar_correo(
        destinatario=contacto["correo"],
        empresa=contacto["empresa"],
        contacto=contacto["contacto"]
    )

    if resultado["exito"]:
        # Actualizar estado a "si" con fecha actual
        fecha_actual = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        actualizar_estado_envio(id, "si", fecha_actual)
        flash(resultado["mensaje"], "success")
    else:
        # No cambiamos el estado si falla
        flash(resultado["mensaje"], "error")

    return redirect(url_for("dashboard"))


@app.route("/reimportar")
def reimportar():
    """
    Ruta para reinicializar el archivo Excel.
    Verifica las columnas necesarias.
    """
    if inicializar_excel():
        flash("Archivo Excel actualizado correctamente.", "success")
    else:
        flash("Error al actualizar el archivo Excel.", "error")

    return redirect(url_for("dashboard"))


@app.route("/eliminar/<int:id>")
def eliminar(id):
    """
    Ruta para eliminar un contacto específico del Excel.
    """
    contacto = obtener_contacto_por_indice(id)

    if not contacto:
        flash("Contacto no encontrado.", "error")
        return redirect(url_for("dashboard"))

    if eliminar_contacto(id):
        flash(f"Contacto '{contacto['empresa']}' eliminado correctamente.", "success")
    else:
        flash("Error al eliminar el contacto.", "error")

    return redirect(url_for("dashboard"))


# ==============================================================================
# INICIALIZACIÓN DEL SISTEMA
# ==============================================================================

def inicializar_sistema():
    """
    Inicializa el archivo Excel con las columnas necesarias.
    Se ejecuta una sola vez al arrancar la aplicación.
    """
    print("=" * 60)
    print("  Sistema de Correos Institucionales - IEEE")
    print("  Inicializando...")
    print("=" * 60)

    # Verificar e inicializar Excel
    if inicializar_excel():
        print(f"[OK] Archivo Excel inicializado: {EXCEL_PATH}")
        df = leer_contactos()
        print(f"[OK] Total de contactos: {len(df)}")
    else:
        print("[ERROR] No se pudo inicializar el archivo Excel")

    print("=" * 60)
    print("  Sistema listo. Acceda a http://127.0.0.1:5000")
    print("=" * 60)


# ==============================================================================
# PUNTO DE ENTRADA
# ==============================================================================

if __name__ == "__main__":
    inicializar_sistema()
    # NOTA: Cambiar debug=False para entorno de producción
    app.run(host="127.0.0.1", port=5000, debug=False)
