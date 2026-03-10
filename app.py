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
import shutil
import tempfile
import threading
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.image import MIMEImage
from datetime import datetime
import json

import pandas as pd
from flask import Flask, render_template, request, redirect, url_for, flash, jsonify
from dotenv import load_dotenv

# Cargar variables de entorno desde archivo .env si existe
load_dotenv()

# ==============================================================================
# CONFIGURACIÓN DEL SISTEMA
# ==============================================================================

# Rutas de archivos
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
EXCEL_PATH = os.path.join(BASE_DIR, "contactos.xlsx")
CONFIG_PATH = os.path.join(BASE_DIR, "configuracion.json")
IMAGES_DIR = os.path.join(BASE_DIR, "images")
IEEE_LOGO_PATH = os.path.join(IMAGES_DIR, "ieee.png")
IEEE_CS_LOGO_PATH = os.path.join(IMAGES_DIR, "ieee cs imagen.png")
IEEE_CR_LOGO_PATH = os.path.join(IMAGES_DIR, "ieee costa rica.png")

# Clave secreta para sesiones Flask
SECRET_KEY = os.getenv(
    "FLASK_SECRET_KEY", "dev-secret-key-change-in-production")

# ==============================================================================
# FUNCIONES DE CONFIGURACIÓN
# ==============================================================================


def cargar_configuracion():
    """
    Carga la configuración desde el archivo JSON.
    Si no existe, crea uno con valores por defecto.
    """
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"Error al cargar configuración: {e}")

    # Configuración por defecto
    return {
        "email_remitente": os.getenv("EMAIL_REMITENTE", ""),
        "email_password": os.getenv("EMAIL_PASSWORD", "").replace(" ", ""),
        "smtp_server": os.getenv("SMTP_SERVER", "smtp.gmail.com"),
        "smtp_port": int(os.getenv("SMTP_PORT", "587")),
        "alias_emails": [
            {
                "nombre": "Principal",
                "email": os.getenv("EMAIL_REMITENTE", ""),
                "predeterminado": True
            }
        ],
        "mensajes_predeterminados": {
            "asunto_defecto": "Invitación a colaborar - Iniciativas estudiantiles TEC",
            "encabezado_titulo": "IEEE Computer Society",
            "encabezado_subtitulo": "Instituto Tecnológico de Costa Rica — Campus Central Cartago",
            "copyright": "© 2026 IEEE Computer Society – Instituto Tecnológico de Costa Rica — Todos los derechos reservados."
        },
        "colores_defecto": {
            "color_encabezado": "#00629B",
            "color_cuerpo": "#0b3f66",
            "firma_color": "#dce8f4",
            "viñetas_color": "#ffd166"
        }
    }


def guardar_configuracion(config):
    """
    Guarda la configuración en el archivo JSON.
    """
    try:
        with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=2, ensure_ascii=False)
        return True
    except Exception as e:
        print(f"Error al guardar configuración: {e}")
        return False


def obtener_configuracion():
    """
    Obtiene la configuración actual del sistema.
    """
    return cargar_configuracion()


# Cargar configuración desde JSON o .env
_CONFIG = obtener_configuracion()

# Credenciales SMTP: usa configuración JSON o variables de entorno
EMAIL_REMITENTE = os.getenv(
    "EMAIL_REMITENTE", _CONFIG.get("email_remitente", "")).strip()
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD", _CONFIG.get(
    "email_password", "")).replace(" ", "").strip()

# Configuración del servidor SMTP (por defecto Gmail con TLS)
SMTP_SERVER = os.getenv("SMTP_SERVER", _CONFIG.get(
    "smtp_server", "smtp.gmail.com")).strip()
SMTP_PORT = int(os.getenv("SMTP_PORT", str(_CONFIG.get("smtp_port", 587))))

# ==============================================================================
# INICIALIZACIÓN DE LA APLICACIÓN FLASK
# ==============================================================================

app = Flask(__name__)
app.secret_key = SECRET_KEY

# Estructura base del archivo Excel y control de logs de error repetidos.
EXCEL_COLUMNS = ["empresa", "contacto", "correo", "enviado", "fecha_envio"]
_excel_error_reportado = False
EXCEL_LOCK = threading.RLock()


def _parametros_contactos_desde_request():
    """
    Conserva filtros y paginación para redirecciones de vuelta a /contactos.
    """
    return {
        "q": request.args.get("q", "").strip(),
        "estado": request.args.get("estado", "todos").strip().lower(),
        "sort": request.args.get("sort", "excel").strip().lower(),
        "dir": request.args.get("dir", "asc").strip().lower(),
        "page": request.args.get("page", "1").strip() or "1",
    }


def _escribir_excel_atomico(df):
    """
    Escribe el Excel de forma atómica para evitar archivos parciales/corruptos.
    """
    archivo_tmp = None
    try:
        excel_dir = os.path.dirname(EXCEL_PATH) or BASE_DIR
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx", dir=excel_dir) as tmp:
            archivo_tmp = tmp.name

        df.to_excel(archivo_tmp, index=False, engine="openpyxl")
        os.replace(archivo_tmp, EXCEL_PATH)
        return True
    finally:
        if archivo_tmp and os.path.exists(archivo_tmp):
            try:
                os.remove(archivo_tmp)
            except Exception:
                pass


# ==============================================================================
# FUNCIONES DE MANEJO DE EXCEL
# ==============================================================================

def inicializar_excel():
    """
    Verifica que el archivo Excel exista y tenga las columnas necesarias.
    Si no existe, crea un archivo vacío con la estructura correcta.
    """
    with EXCEL_LOCK:
        if not os.path.exists(EXCEL_PATH):
            # Crear archivo vacío con estructura
            df = pd.DataFrame(columns=EXCEL_COLUMNS)
            return _escribir_excel_atomico(df)

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
            df = df[EXCEL_COLUMNS]

            # Limpiar valores NaN
            df = df.fillna("")

            # Asegurar que "enviado" solo tenga "si" o "no"
            df["enviado"] = df["enviado"].apply(lambda x: "si" if str(x).lower() in [
                                                "si", "sí", "yes", "1"] else "no")

            # Guardar cambios
            return _escribir_excel_atomico(df)

        except Exception as e:
            print(f"Error al inicializar Excel: {e}")

            # Si el Excel está corrupto, lo respaldamos y regeneramos para que
            # la aplicación pueda seguir operando con un archivo válido.
            if es_error_excel_corrupto(e):
                return reparar_excel_corrupto()

            return False


def es_error_excel_corrupto(error):
    """
    Detecta mensajes comunes de corrupción/integridad en archivos XLSX.
    """
    texto = f"{type(error).__name__}: {error}".lower()
    patrones = [
        "invalid block type",
        "invalid distance too far back",
        "while decompressing data",
        "badzipfile",
        "file is not a zip file",
        "truncated file header",
        "bad crc-32",
    ]
    return any(p in texto for p in patrones)


def reparar_excel_corrupto():
    """
    Respaldar el archivo corrupto y crear uno nuevo con estructura válida.
    """
    with EXCEL_LOCK:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = f"{EXCEL_PATH}.corrupto_{timestamp}.bak"

        try:
            if os.path.exists(EXCEL_PATH):
                shutil.copy2(EXCEL_PATH, backup_path)
                print(f"[WARN] Excel corrupto respaldado en: {backup_path}")

            df = pd.DataFrame(columns=EXCEL_COLUMNS)
            if _escribir_excel_atomico(df):
                print(f"[OK] Se regeneró el archivo Excel: {EXCEL_PATH}")
                return True

            return False
        except Exception as e:
            print(f"[ERROR] No se pudo reparar el archivo Excel: {e}")
            return False


def leer_contactos():
    """
    Lee todos los contactos del archivo Excel.
    Retorna un DataFrame de pandas.
    """
    global _excel_error_reportado

    with EXCEL_LOCK:
        if not os.path.exists(EXCEL_PATH):
            return pd.DataFrame(columns=EXCEL_COLUMNS)

        try:
            df = pd.read_excel(EXCEL_PATH, engine="openpyxl")
            df = df.fillna("")

            # Asegurar que las columnas existan
            if "enviado" not in df.columns:
                df["enviado"] = "no"
            if "fecha_envio" not in df.columns:
                df["fecha_envio"] = ""

            _excel_error_reportado = False
            return df
        except Exception as e:
            # Evita inundar la consola con el mismo error en cada request.
            if not _excel_error_reportado:
                print(f"Error al leer Excel: {e}")
                _excel_error_reportado = True

            if es_error_excel_corrupto(e):
                if reparar_excel_corrupto():
                    try:
                        df = pd.read_excel(
                            EXCEL_PATH, engine="openpyxl").fillna("")
                        _excel_error_reportado = False
                        return df
                    except Exception as e2:
                        print(
                            f"[ERROR] Falló la lectura después de reparar Excel: {e2}")

            return pd.DataFrame(columns=EXCEL_COLUMNS)


def guardar_contactos(df):
    """
    Guarda el DataFrame completo en el archivo Excel.
    Sobrescribe el archivo existente.
    """
    with EXCEL_LOCK:
        try:
            df = df.fillna("")
            return _escribir_excel_atomico(df)
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
    with EXCEL_LOCK:
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
    with EXCEL_LOCK:
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
    with EXCEL_LOCK:
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
    if not EMAIL_REMITENTE or not EMAIL_PASSWORD:
        return {
            "exito": False,
            "mensaje": "Faltan credenciales SMTP. Configure EMAIL_REMITENTE y EMAIL_PASSWORD en variables de entorno."
        }

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
            img_ieee.add_header("Content-Disposition",
                                "inline", filename="ieee.png")
            mensaje.attach(img_ieee)

    if incluir_logo_ieee_cs:
        with open(IEEE_CS_LOGO_PATH, "rb") as logo_ieee_cs:
            img_ieee_cs = MIMEImage(logo_ieee_cs.read())
            img_ieee_cs.add_header("Content-ID", "<logo_ieee_cs>")
            img_ieee_cs.add_header("Content-Disposition",
                                   "inline", filename="ieee_cs.png")
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
        return {
            "exito": False,
            "mensaje": "Error de autenticacion SMTP (535). Verifique EMAIL_REMITENTE, EMAIL_PASSWORD (contrasena de aplicacion vigente) y SMTP_SERVER/SMTP_PORT."
        }
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


def procesar_listas_para_email(html_content, viñetas_color="#ffd166"):
    """
    Procesa el HTML generado por Quill para convertir listas <ul> y <ol> 
    en un formato compatible con clientes de correo electrónico.
    Reemplaza listas con tablas HTML que tienen estilos inline.
    """
    import re

    def convertir_ul(match):
        """Convierte <ul>...</ul> en tablas HTML con viñetas coloreadas."""
        contenido = match.group(1)
        # Extraer cada <li>...</li>
        items = re.findall(r'<li[^>]*>(.*?)</li>', contenido, re.DOTALL)

        tablas = []
        for item in items:
            tabla = (
                f'<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin: 0 0 8px 0;">'
                f'<tr>'
                f'<td style="width: 18px; vertical-align: top; font-size: 15px; line-height: 1.7; color: {viñetas_color}; font-weight: 700; padding: 0;">&#8226;</td>'
                f'<td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; padding: 0;">{item.strip()}</td>'
                f'</tr>'
                f'</table>'
            )
            tablas.append(tabla)

        return '\n'.join(tablas)

    def convertir_ol(match):
        """Convierte <ol>...</ol> en tablas HTML con números coloreados."""
        contenido = match.group(1)
        # Extraer cada <li>...</li>
        items = re.findall(r'<li[^>]*>(.*?)</li>', contenido, re.DOTALL)

        tablas = []
        for i, item in enumerate(items, start=1):
            tabla = (
                f'<table role="presentation" width="100%" cellspacing="0" cellpadding="0" border="0" style="margin: 0 0 8px 0;">'
                f'<tr>'
                f'<td style="width: 28px; vertical-align: top; font-size: 15px; line-height: 1.7; color: {viñetas_color}; font-weight: 700; padding: 0;">{i}.</td>'
                f'<td style="vertical-align: top; font-size: 15px; line-height: 1.7; color: #ffffff; padding: 0;">{item.strip()}</td>'
                f'</tr>'
                f'</table>'
            )
            tablas.append(tabla)

        return '\n'.join(tablas)

    # Procesar listas no ordenadas
    html_content = re.sub(r'<ul[^>]*>(.*?)</ul>',
                          convertir_ul, html_content, flags=re.DOTALL)

    # Procesar listas ordenadas
    html_content = re.sub(r'<ol[^>]*>(.*?)</ol>',
                          convertir_ol, html_content, flags=re.DOTALL)

    return html_content


def generar_cuerpo_html_personalizado(detalle, firma, imagenes_seleccionadas, encabezado_titulo="",
                                      encabezado_subtitulo="", firma_color="#dce8f4",
                                      firma_estilos=None, viñetas_color="#ffd166",
                                      color_encabezado="#00629B", color_cuerpo="#0b3f66",
                                      copyright=""):
    """
    Genera el cuerpo del correo personalizado en formato HTML con diseño institucional.

    Args:
        detalle: El contenido HTML principal del correo (puede contener etiquetas HTML con estilos)
        firma: La firma personalizada (puede ser vacía para usar la predeterminada)
        imagenes_seleccionadas: Lista de imágenes a incluir ['ieee', 'ieee_cr', 'ieee_cs']
        encabezado_titulo: Título personalizado del encabezado
        encabezado_subtitulo: Subtítulo personalizado del encabezado
        firma_color: Color del texto de la firma
        firma_estilos: Lista con estilos de firma ['negrita', 'cursiva']
        viñetas_color: Color de los bullet points y números de lista
        color_encabezado: Color de fondo del encabezado
        color_cuerpo: Color de fondo del cuerpo del correo
        copyright: Texto de copyright personalizado para el footer
    """
    if firma_estilos is None:
        firma_estilos = []

    # Usar valores predeterminados si no se proporcionan
    if not encabezado_titulo:
        encabezado_titulo = "IEEE Computer Society"
    # El subtítulo puede ir vacío, no forzamos un valor por defecto

    # Procesar el HTML del detalle para convertir listas en formato compatible con email
    detalle_procesado = procesar_listas_para_email(detalle, viñetas_color)

    # Procesar el HTML del detalle para convertir listas en formato compatible con email
    detalle_procesado = procesar_listas_para_email(detalle, viñetas_color)

    # Generar estilos CSS para la firma
    estilo_firma = ""
    if 'negrita' in firma_estilos:
        estilo_firma += "font-weight: 700; "
    if 'cursiva' in firma_estilos:
        estilo_firma += "font-style: italic; "

    # Generar HTML de las imágenes según la cantidad seleccionada
    imagenes_html = ""

    if len(imagenes_seleccionadas) == 1:
        # Una sola imagen, a la izquierda
        img_src = f"cid:logo_{imagenes_seleccionadas[0]}"
        imagenes_html = f"""
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 16px;">
                <tr>
                    <td align="left" style="width: 100%; vertical-align: middle; padding-left: 0; padding-right: 0;">
                        <img src="{img_src}" alt="{imagenes_seleccionadas[0]}" style="display: block; width: 100%; max-width: 180px; height: auto;">
                    </td>
                    <td style="width: 100%;"></td>
                </tr>
            </table>
        """
    elif len(imagenes_seleccionadas) == 2:
        # Dos imágenes, una a la izquierda y otra a la derecha
        img_src_1 = f"cid:logo_{imagenes_seleccionadas[0]}"
        img_src_2 = f"cid:logo_{imagenes_seleccionadas[1]}"
        imagenes_html = f"""
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 16px;">
                <tr>
                    <td align="left" style="width: 50%; vertical-align: middle; padding-left: 12px; padding-right: 18px;">
                        <img src="{img_src_1}" alt="{imagenes_seleccionadas[0]}" style="display: block; width: 100%; max-width: 160px; height: auto;">
                    </td>
                    <td align="right" style="width: 50%; vertical-align: middle; padding-left: 18px; padding-right: 12px;">
                        <img src="{img_src_2}" alt="{imagenes_seleccionadas[1]}" style="display: block; width: 100%; max-width: 160px; height: auto;">
                    </td>
                </tr>
            </table>
        """
    elif len(imagenes_seleccionadas) == 3:
        # Tres imágenes distribuidas: izquierda, centro, derecha
        img_src_1 = f"cid:logo_{imagenes_seleccionadas[0]}"
        img_src_2 = f"cid:logo_{imagenes_seleccionadas[1]}"
        img_src_3 = f"cid:logo_{imagenes_seleccionadas[2]}"
        imagenes_html = f"""
            <table role="presentation" cellspacing="0" cellpadding="0" border="0" width="100%" style="margin-bottom: 16px;">
                <tr>
                    <td align="left" style="width: 33.33%; vertical-align: middle; padding-left: 12px; padding-right: 8px;">
                        <img src="{img_src_1}" alt="{imagenes_seleccionadas[0]}" style="display: block; width: 100%; max-width: 130px; height: auto;">
                    </td>
                    <td align="center" style="width: 33.33%; vertical-align: middle; padding-left: 8px; padding-right: 8px;">
                        <img src="{img_src_2}" alt="{imagenes_seleccionadas[1]}" style="display: block; width: 100%; max-width: 130px; height: auto; margin: 0 auto;">
                    </td>
                    <td align="right" style="width: 33.33%; vertical-align: middle; padding-left: 8px; padding-right: 12px;">
                        <img src="{img_src_3}" alt="{imagenes_seleccionadas[2]}" style="display: block; width: 100%; max-width: 130px; height: auto;">
                    </td>
                </tr>
            </table>
        """

    # Firma predeterminada si no se proporciona una personalizada
    if not firma or not firma.strip():
        firma = """Julio Ricardo Barrios Amador
Section Student Representative (SSR) | IEEE Costa Rica Section
Vocal 2 | IEEE Computer Society - Instituto Tecnológico de Costa Rica
IEEE Member: 101781510
julio.barrios@ieee.org"""

    # Convertir la firma en párrafos HTML con estilos personalizados
    lineas_firma = firma.strip().split('\n')
    firma_html = ""
    for i, linea in enumerate(lineas_firma):
        if i == 0:
            # Primera línea (nombre) siempre en negrita
            firma_html += f"""
                <p style="margin: 0 0 2px 0; font-size: 15px; font-weight: 700; color: {firma_color}; {estilo_firma}">
                    {linea}
                </p>
            """
        elif '@' in linea and 'ieee.org' in linea:
            # Línea con correo en color especial
            firma_html += f"""
                <p style="margin: 6px 0 0 0; font-size: 12px; color: #00d4ff; letter-spacing: 0.3px; font-weight: 700;">
                    {linea}
                </p>
            """
        else:
            # Resto de líneas
            firma_html += f"""
                <p style="margin: 0 0 1px 0; font-size: 13px; color: {firma_color}; letter-spacing: 0.3px; {estilo_firma}">
                    {linea}
                </p>
            """

    # El detalle ahora ya viene como HTML del editor contenteditable
    # No es necesario convertir saltos de línea

    html = f"""
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <style>
            /* Estilos para enlaces visibles sobre fondo oscuro */
            a {{
                color: #00d4ff;
                text-decoration: underline;
            }}
            
            a:hover {{
                color: #66e3ff;
            }}
            
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
                            <td class="header-cell" style="background-color: {color_encabezado}; padding: 28px 40px; text-align: center;">
                                {imagenes_html}
                                <h1 style="margin: 0; color: #ffffff; font-size: 22px; font-weight: 600; letter-spacing: 0.5px;">
                                    {encabezado_titulo}
                                </h1>
                                {f'<p style="margin: 6px 0 0 0; color: rgba(255,255,255,0.85); font-size: 13px; letter-spacing: 0.3px;">{encabezado_subtitulo}</p>' if encabezado_subtitulo else ''}
                            </td>
                        </tr>
                        <!-- Cuerpo del correo -->
                        <tr>
                            <td style="padding: 36px 40px 20px 40px; background-color: {color_cuerpo};">
                                <div style="font-size: 15px; line-height: 1.7; color: #ffffff;">
                                    {detalle_procesado}
                                </div>
                            </td>
                        </tr>
                        <!-- Cierre formal -->
                        <tr>
                            <td style="padding: 0 40px 36px 40px; background-color: {color_cuerpo};">
                                <p style="margin: 0 0 4px 0; font-size: 15px; color: #ffffff;">
                                    Atentamente,
                                </p>
                                <br>
                                {firma_html}
                            </td>
                        </tr>
                        <!-- Footer -->
                        <tr>
                            <td style="background-color: #f8f9fa; padding: 18px 40px; text-align: center; border-top: 1px solid #e9ecef;">
                                <p style="margin: 0; font-size: 11px; color: #adb5bd; letter-spacing: 0.2px;">
                                    {copyright if copyright else f'&copy; {datetime.now().year} IEEE Computer Society – Instituto Tecnológico de Costa Rica &mdash; Todos los derechos reservados.'}
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


def enviar_correo_personalizado_smtp(destinatario, asunto, detalle, firma, imagenes_seleccionadas,
                                     encabezado_titulo="", encabezado_subtitulo="",
                                     firma_color="#dce8f4", firma_estilos=None, viñetas_color="#ffd166",
                                     color_encabezado="#00629B", color_cuerpo="#0b3f66", cc_list=None, copyright="", alias_email=""):
    """
    Envía un correo electrónico personalizado usando SMTP con TLS.
    Permite seleccionar las imágenes del encabezado y personalizar estilos de firma.
    Soporta envío desde alias de correo.
    El contenido HTML es generado directamente del editor (contenteditable).
    Retorna un diccionario con el resultado del envío.
    """
    # La autenticación SMTP debe hacerse con la cuenta principal.
    # El alias se usa solo como remitente visible (header From).
    smtp_login_user = EMAIL_REMITENTE
    email_remitente = alias_email if alias_email else EMAIL_REMITENTE
    email_password = EMAIL_PASSWORD
    smtp_server = SMTP_SERVER
    smtp_port = SMTP_PORT

    if not email_remitente or not email_password:
        return {
            "exito": False,
            "mensaje": "Faltan credenciales SMTP. Configure EMAIL_REMITENTE y EMAIL_PASSWORD en variables de entorno."
        }

    if firma_estilos is None:
        firma_estilos = []

    # Generar contenido del correo
    cuerpo_html = generar_cuerpo_html_personalizado(
        detalle=detalle,
        firma=firma,
        imagenes_seleccionadas=imagenes_seleccionadas,
        encabezado_titulo=encabezado_titulo,
        encabezado_subtitulo=encabezado_subtitulo,
        firma_color=firma_color,
        firma_estilos=firma_estilos,
        viñetas_color=viñetas_color,
        color_encabezado=color_encabezado,
        color_cuerpo=color_cuerpo,
        copyright=copyright
    )

    # Usamos multipart/related para soportar imágenes inline referenciadas por CID.
    mensaje = MIMEMultipart("related")
    mensaje["From"] = email_remitente
    mensaje["To"] = destinatario
    # Añadir CC en cabecera si se proporcionó
    if cc_list:
        mensaje["Cc"] = ", ".join(cc_list)
    mensaje["Subject"] = asunto

    # El HTML vive dentro de multipart/alternative para compatibilidad con clientes.
    parte_alternativa = MIMEMultipart("alternative")
    parte_html = MIMEText(cuerpo_html, "html", "utf-8")
    parte_alternativa.attach(parte_html)
    mensaje.attach(parte_alternativa)

    # Mapeo de nombres de imágenes a rutas de archivo
    mapeo_imagenes = {
        'ieee': IEEE_LOGO_PATH,
        'ieee_cr': IEEE_CR_LOGO_PATH,
        'ieee_cs': IEEE_CS_LOGO_PATH
    }

    # Adjuntar las imágenes seleccionadas
    for img_key in imagenes_seleccionadas:
        img_path = mapeo_imagenes.get(img_key)
        if img_path and os.path.exists(img_path):
            with open(img_path, "rb") as img_file:
                img_data = MIMEImage(img_file.read())
                img_data.add_header("Content-ID", f"<logo_{img_key}>")
                img_data.add_header("Content-Disposition",
                                    "inline", filename=f"{img_key}.png")
                mensaje.attach(img_data)

    try:
        # Conexión SMTP con TLS
        servidor = smtplib.SMTP(smtp_server, smtp_port)
        servidor.ehlo()
        servidor.starttls()
        servidor.ehlo()
        servidor.login(smtp_login_user, email_password)
        # Enviar a destinatario principal y a los CC (si los hay)
        destinatarios_envio = [destinatario]
        if cc_list:
            destinatarios_envio += cc_list
        servidor.sendmail(
            email_remitente, destinatarios_envio, mensaje.as_string())
        servidor.quit()

        return {"exito": True, "mensaje": f"Correo enviado exitosamente a {destinatario}"}

    except smtplib.SMTPAuthenticationError:
        return {
            "exito": False,
            "mensaje": "Error de autenticación SMTP (535). Verifique EMAIL_REMITENTE, EMAIL_PASSWORD (contraseña de aplicación vigente) y SMTP_SERVER/SMTP_PORT."
        }
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
def dashboard_principal():
    """
    Ruta principal - Dashboard con opciones principales.
    Muestra botones para gestión de contactos y correos personalizados.
    """
    return render_template("dashboard.html", year=datetime.now().year)


@app.route("/contactos")
def contactos():
    """
    Ruta de gestión de contactos - Dashboard de gestión de contactos.
    Muestra la tabla de contactos, estadísticas y formulario de agregar.
    """
    df = leer_contactos().fillna("")

    # Parámetros de filtro y paginación
    busqueda = request.args.get("q", "").strip()
    estado = request.args.get("estado", "todos").strip().lower()
    ordenar_por = request.args.get("sort", "excel").strip().lower()
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
            df["empresa"].astype(str).str.contains(
                patron, case=False, na=False)
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
    if ordenar_por == "excel":
        # Respeta el orden original del Excel usando el índice (id).
        df = df.sort_values(by=["id"], ascending=[
                            ascendente], kind="mergesort")
    elif ordenar_por == "estado":
        df = df.sort_values(by=["estado_envio", "empresa"], ascending=[
                            ascendente, True], kind="mergesort")
    elif ordenar_por == "fecha_envio":
        # Para fecha vacía, usamos un extremo para que queden al final según el orden.
        fecha_tmp = pd.to_datetime(df["fecha_envio"], errors="coerce")
        if ascendente:
            fecha_tmp = fecha_tmp.fillna(pd.Timestamp.max)
        else:
            fecha_tmp = fecha_tmp.fillna(pd.Timestamp.min)
        df = df.assign(_fecha_sort=fecha_tmp).sort_values(by=["_fecha_sort", "empresa"], ascending=[
            ascendente, True], kind="mergesort").drop(columns=["_fecha_sort"])
    elif ordenar_por == "empresa":
        df = df.sort_values(by=["empresa", "id"], ascending=[
                            ascendente, True], kind="mergesort")
    else:
        ordenar_por = "excel"
        df = df.sort_values(by=["id"], ascending=[
                            ascendente], kind="mergesort")

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
        return redirect(url_for("contactos"))

    if not correo:
        flash("El campo 'Correo' es obligatorio.", "error")
        return redirect(url_for("contactos"))

    # Validar formato de correo
    if not validar_correo(correo):
        flash("El formato del correo electrónico no es válido.", "error")
        return redirect(url_for("contactos"))

    # Intentar agregar
    if agregar_contacto(empresa, contacto, correo):
        flash(f"Contacto '{empresa}' agregado exitosamente.", "success")
    else:
        flash(f"El correo '{correo}' ya existe en el archivo.", "warning")

    return redirect(url_for("contactos"))


@app.route("/enviar/<int:id>")
def enviar(id):
    """
    Ruta para enviar correo individual a un contacto específico.
    Actualiza el estado de envío según el resultado.
    """
    contacto = obtener_contacto_por_indice(id)
    return_params = _parametros_contactos_desde_request()

    if not contacto:
        flash("Contacto no encontrado.", "error")
        return redirect(url_for("contactos", **return_params))

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

    return redirect(url_for("contactos", **return_params))


@app.route("/vista_previa/<int:id>")
def vista_previa(id):
    """
    Ruta para obtener la vista previa del correo de un contacto específico.
    Retorna JSON con asunto y cuerpo HTML.
    """
    contacto = obtener_contacto_por_indice(id)

    if not contacto:
        return jsonify({"error": "Contacto no encontrado"}), 404

    # Generar saludo y cuerpo del correo
    saludo = generar_saludo(contacto["empresa"], contacto["contacto"])
    cuerpo_html = generar_cuerpo_html(
        saludo,
        contacto["empresa"],
        incluir_logo_ieee=os.path.exists(IEEE_LOGO_PATH),
        incluir_logo_ieee_cs=os.path.exists(IEEE_CS_LOGO_PATH)
    )

    return jsonify({
        "asunto": "Invitación a colaborar - Iniciativas estudiantiles TEC",
        "cuerpo": cuerpo_html,
        "destinatario": contacto["correo"],
        "empresa": contacto["empresa"],
        "contacto": contacto.get("contacto", "")
    })


@app.route("/reimportar")
def reimportar():
    """
    Ruta para reinicializar el archivo Excel.
    Verifica las columnas necesarias.
    """
    return_params = _parametros_contactos_desde_request()

    if inicializar_excel():
        flash("Archivo Excel actualizado correctamente.", "success")
    else:
        flash("Error al actualizar el archivo Excel.", "error")

    return redirect(url_for("contactos", **return_params))


@app.route("/eliminar/<int:id>")
def eliminar(id):
    """
    Ruta para eliminar un contacto específico del Excel.
    """
    contacto = obtener_contacto_por_indice(id)
    return_params = _parametros_contactos_desde_request()

    if not contacto:
        flash("Contacto no encontrado.", "error")
        return redirect(url_for("contactos", **return_params))

    if eliminar_contacto(id):
        flash(
            f"Contacto '{contacto['empresa']}' eliminado correctamente.", "success")
    else:
        flash("Error al eliminar el contacto.", "error")

    return redirect(url_for("contactos", **return_params))


@app.route("/configuracion", methods=["GET"])
def configuracion():
    """
    Página de configuración del sistema.
    Permite personalizar alias de correos, mensajes y colores por defecto.
    """
    config = obtener_configuracion()
    return render_template("configuracion.html", config=config)


@app.route("/guardar_configuracion", methods=["POST"])
def guardar_config():
    """
    Guarda la configuración del sistema.
    """
    try:
        config = obtener_configuracion()

        # Actualizar configuración general
        config["email_remitente"] = request.form.get(
            "email_remitente", "").strip()
        config["email_password"] = request.form.get(
            "email_password", "").strip().replace(" ", "")
        config["smtp_server"] = request.form.get(
            "smtp_server", "smtp.gmail.com").strip()
        config["smtp_port"] = int(request.form.get("smtp_port", "587"))

        # Actualizar mensajes predeterminados
        config["mensajes_predeterminados"]["asunto_defecto"] = request.form.get(
            "asunto_defecto", "").strip()
        config["mensajes_predeterminados"]["encabezado_titulo"] = request.form.get(
            "encabezado_titulo", "").strip()
        config["mensajes_predeterminados"]["encabezado_subtitulo"] = request.form.get(
            "encabezado_subtitulo", "").strip()
        config["mensajes_predeterminados"]["copyright"] = request.form.get(
            "copyright", "").strip()

        # Actualizar colores por defecto
        config["colores_defecto"]["color_encabezado"] = request.form.get(
            "color_encabezado", "#00629B").strip()
        config["colores_defecto"]["color_cuerpo"] = request.form.get(
            "color_cuerpo", "#0b3f66").strip()
        config["colores_defecto"]["firma_color"] = request.form.get(
            "firma_color", "#dce8f4").strip()
        config["colores_defecto"]["viñetas_color"] = request.form.get(
            "viñetas_color", "#ffd166").strip()

        # Procesar alias de emails
        alias_emails = []
        alias_count = int(request.form.get("alias_count", "0"))

        for i in range(alias_count):
            alias_nombre = request.form.get(f"alias_nombre_{i}", "").strip()
            alias_email = request.form.get(f"alias_email_{i}", "").strip()
            alias_predeterminado = request.form.get(
                f"alias_predeterminado_{i}") == "on"

            if alias_nombre and alias_email:
                alias_emails.append({
                    "nombre": alias_nombre,
                    "email": alias_email,
                    "predeterminado": alias_predeterminado
                })

        if alias_emails:
            config["alias_emails"] = alias_emails

        # Guardar configuración
        if guardar_configuracion(config):
            flash("Configuración guardada correctamente.", "success")
            # Actualizar variables globales
            global EMAIL_REMITENTE, EMAIL_PASSWORD, SMTP_SERVER, SMTP_PORT, _CONFIG
            _CONFIG = config
            EMAIL_REMITENTE = config.get("email_remitente", "")
            EMAIL_PASSWORD = config.get("email_password", "")
            SMTP_SERVER = config.get("smtp_server", "smtp.gmail.com")
            SMTP_PORT = config.get("smtp_port", 587)
        else:
            flash("Error al guardar la configuración.", "error")
    except Exception as e:
        flash(f"Error: {str(e)}", "error")

    return redirect(url_for("configuracion"))


@app.route("/api/configuracion", methods=["GET"])
def api_configuracion():
    """
    API para obtener la configuración actual (usada por JavaScript).
    """
    config = obtener_configuracion()
    return jsonify({
        "mensajes_predeterminados": config.get("mensajes_predeterminados", {}),
        "colores_defecto": config.get("colores_defecto", {}),
        "alias_emails": config.get("alias_emails", [])
    })


@app.route("/api/preview_correo_personalizado", methods=["POST"])
def api_preview_correo_personalizado():
    """
    API para generar vista previa del correo personalizado.
    Genera el HTML exactamente como se generaría para el envío real.
    Retorna JSON con el HTML del correo completo y metadatos.
    """
    try:
        # Obtener datos del formulario
        destinatario = request.json.get("destinatario", "").strip()
        cc = request.json.get("cc", "").strip()
        asunto = request.json.get("asunto", "").strip()
        detalle = request.json.get("detalle", "").strip()
        firma = request.json.get("firma", "").strip()

        # Campos personalizables
        encabezado_titulo = request.json.get(
            "encabezado_titulo", "IEEE Computer Society").strip()
        encabezado_subtitulo = request.json.get(
            "encabezado_subtitulo", "").strip()

        # Colores
        firma_color = request.json.get("firma_color", "#dce8f4").strip()
        viñetas_color = request.json.get("viñetas_color", "#ffd166").strip()
        color_encabezado = request.json.get(
            "color_encabezado", "#00629B").strip()
        color_cuerpo = request.json.get("color_cuerpo", "#0b3f66").strip()

        # Copyright
        copyright = request.json.get("copyright", "").strip()

        # Estilos de firma y imágenes
        firma_estilos = request.json.get("firma_estilos", [])
        imagenes_seleccionadas = request.json.get("imagenes", [])

        # Generar HTML usando la misma función que para envío real
        html_correo = generar_cuerpo_html_personalizado(
            detalle=detalle,
            firma=firma,
            imagenes_seleccionadas=imagenes_seleccionadas,
            encabezado_titulo=encabezado_titulo,
            encabezado_subtitulo=encabezado_subtitulo,
            firma_color=firma_color,
            firma_estilos=firma_estilos,
            viñetas_color=viñetas_color,
            color_encabezado=color_encabezado,
            color_cuerpo=color_cuerpo,
            copyright=copyright
        )

        # Retornar respuesta JSON con el HTML y metadatos
        return jsonify({
            "success": True,
            "html": html_correo,
            "metadata": {
                "destinatario": destinatario,
                "cc": cc,
                "asunto": asunto
            }
        })

    except Exception as e:
        return jsonify({
            "success": False,
            "error": str(e)
        }), 400


@app.route("/correo_personalizado", methods=["GET"])
def correo_personalizado():
    """
    Ruta para mostrar el formulario de correo personalizado.
    Permite crear y enviar correos con formato institucional personalizable.
    """
    return render_template("correo_personalizado.html")


@app.route("/enviar_correo_personalizado", methods=["POST"])
def enviar_correo_personalizado():
    """
    Ruta POST para procesar y enviar un correo personalizado.
    Valida los datos y envía el correo con el formato institucional personalizable.
    """
    destinatario = request.form.get("destinatario", "").strip().lower()
    asunto = request.form.get("asunto", "").strip()
    detalle = request.form.get("detalle", "").strip()
    firma = request.form.get("firma", "").strip()

    # Alias de correo desde el que se enviaría
    alias_email = request.form.get("alias_email", "").strip()

    # Nuevos campos personalizables
    encabezado_titulo = request.form.get(
        "encabezado_titulo", "IEEE Computer Society").strip()
    encabezado_subtitulo = request.form.get("encabezado_subtitulo", "").strip()

    # Colores de firma
    firma_color = request.form.get("firma_color", "#dce8f4").strip()
    viñetas_color = request.form.get("viñetas_color", "#ffd166").strip()

    # Colores de encabezado y cuerpo del correo
    color_encabezado = request.form.get("color_encabezado", "#00629B").strip()
    color_cuerpo = request.form.get("color_cuerpo", "#0b3f66").strip()

    # Copyright personalizado del footer
    copyright = request.form.get("copyright", "").strip()

    # Estilos de firma
    firma_estilos = request.form.getlist("firma_estilos")

    # Obtener imágenes seleccionadas (pueden ser múltiples)
    imagenes_seleccionadas = request.form.getlist("imagenes")

    # Campo CC (opcional) - puede contener múltiples direcciones separadas por comas o punto y coma
    cc_raw = request.form.get("cc", "").strip()
    # Parsear lista de CC separada por comas o punto y coma
    cc_list = []
    if cc_raw:
        cc_list = [c.strip().lower()
                   for c in re.split(r'[;,]', cc_raw) if c.strip()]

    # Validar campos obligatorios
    if not destinatario:
        flash("El campo 'Destinatario' es obligatorio.", "error")
        return redirect(url_for("correo_personalizado"))

    if not asunto:
        flash("El campo 'Asunto' es obligatorio.", "error")
        return redirect(url_for("correo_personalizado"))

    if not detalle:
        flash("El campo 'Mensaje' es obligatorio.", "error")
        return redirect(url_for("correo_personalizado"))

    # Validar formato de correo
    if not validar_correo(destinatario):
        flash("El formato del correo electrónico del destinatario no es válido.", "error")
        return redirect(url_for("correo_personalizado"))

    # Validar formato de CC (si se proporcionaron)
    for cc in cc_list:
        if not validar_correo(cc):
            flash(
                f"El correo en CC '{cc}' no tiene un formato válido.", "error")
            return redirect(url_for("correo_personalizado"))

    # Validar que haya al menos una imagen seleccionada
    if not imagenes_seleccionadas:
        flash("Debe seleccionar al menos una imagen para el encabezado.", "warning")
        return redirect(url_for("correo_personalizado"))

    # Ejecutar envío de correo con los parámetros personalizables
    resultado = enviar_correo_personalizado_smtp(
        destinatario=destinatario,
        asunto=asunto,
        detalle=detalle,
        firma=firma,
        imagenes_seleccionadas=imagenes_seleccionadas,
        encabezado_titulo=encabezado_titulo,
        encabezado_subtitulo=encabezado_subtitulo,
        firma_color=firma_color,
        firma_estilos=firma_estilos,
        viñetas_color=viñetas_color,
        color_encabezado=color_encabezado,
        color_cuerpo=color_cuerpo,
        copyright=copyright,
        cc_list=cc_list,
        alias_email=alias_email
    )

    if resultado["exito"]:
        flash(f"Correo enviado exitosamente a {destinatario}", "success")
        return redirect(url_for("dashboard_principal"))
    else:
        flash(resultado["mensaje"], "error")
        return redirect(url_for("correo_personalizado"))


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
    puerto = int(os.getenv("FLASK_PORT", "5001"))
    print(f"  Sistema listo. Acceda a http://127.0.0.1:{puerto}")
    print("=" * 60)


# ==============================================================================
# PUNTO DE ENTRADA
# ==============================================================================

if __name__ == "__main__":
    inicializar_sistema()
    # NOTA: Cambiar debug=False para entorno de producción
    port = int(os.getenv("FLASK_PORT", "5001"))
    app.run(host="0.0.0.0", port=port, debug=False)
