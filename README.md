# 📧 Sistema de Gestión y Envío de Correos Institucionales

## IEEE Computer Society - Instituto Tecnológico de Costa Rica

<div align="center">

**Aplicación Flask profesional para gestionar contactos empresariales y enviar correos electrónicos institucionales personalizados con diseño corporativo IEEE.**

[![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Flask](https://img.shields.io/badge/Flask-3.0.0-green?logo=flask&logoColor=white)](https://flask.palletsprojects.com/)
[![License](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

</div>

---

## 📋 Tabla de Contenidos

- [Descripción](#-descripción)
- [Características Principales](#-características-principales)
- [Tecnologías Utilizadas](#-tecnologías-utilizadas)
- [Requisitos del Sistema](#-requisitos-del-sistema)
- [Instalación](#-instalación)
  - [Opción A: Entorno Local](#opción-a-entorno-local-windowsmacoslinux)
  - [Opción B: GitHub Codespaces](#opción-b-github-codespaces)
- [Configuración](#-configuración)
  - [Configuración SMTP](#configuración-smtp)
  - [Configuración de Correos Alias](#configuración-de-correos-alias)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Uso de la Aplicación](#-uso-de-la-aplicación)
- [Funcionalidades Detalladas](#-funcionalidades-detalladas)
- [Seguridad y Buenas Prácticas](#-seguridad-y-buenas-prácticas)
- [Solución de Problemas](#-solución-de-problemas)
- [Contribución](#-contribución)
- [Licencia](#-licencia)

---

## 🎯 Descripción

Sistema web desarrollado en Flask para el **IEEE Computer Society del Instituto Tecnológico de Costa Rica, Campus Central Cartago**, diseñado para facilitar la gestión de contactos empresariales y el envío masivo de correos electrónicos profesionales con formato institucional personalizable.

### Casos de Uso Principales

- **Gestión de patrocinios**: Contactar empresas para solicitar patrocinios de eventos estudiantiles
- **Invitaciones a charlas**: Enviar invitaciones profesionales a expertos del sector
- **Networking corporativo**: Mantener comunicación con empresas colaboradoras
- **Ferias de empleo**: Coordinar participación de empresas en eventos de reclutamiento
- **Correos personalizados**: Crear y enviar correos con editor de texto enriquecido (WYSIWYG)
- **Correos desde alias**: Enviar correos desde múltiples direcciones institucionales (requiere configuración)

---

## ✨ Características Principales

### 🗂️ Gestión Avanzada de Contactos

- ✅ **Agregar contactos** con validación automática de correos electrónicos
- ✅ **Almacenamiento en Excel** (formato `.xlsx`) para fácil exportación y respaldo
- ✅ **Prevención de duplicados** basada en dirección de correo electrónico
- ✅ **Eliminación segura** de contactos con confirmación
- ✅ **Importación/Reimportación** de datos desde archivo Excel
- ✅ **Registro automático** de estado de envío y fecha/hora

### 🔍 Sistema de Búsqueda y Filtros

- 🔎 **Búsqueda en tiempo real** por empresa, contacto o correo electrónico
- 🎯 **Filtros por estado**: Todos, Pendientes, Enviados
- 📊 **Ordenamiento flexible**:
  - Por orden original de Excel
  - Por nombre de empresa (A-Z / Z-A)
  - Por estado de envío
  - Por fecha de envío (más reciente / más antiguo)
- 📄 **Paginación inteligente**: 50 contactos por página

### 📨 Envío de Correos Electrónicos

#### Correo Predeterminado Institucional
- 📧 Plantilla HTML profesional con diseño IEEE
- 🖼️ **Logos institucionales inline**: IEEE, IEEE Computer Society, IEEE Costa Rica
- 🎨 **Diseño responsive** compatible con todos los clientes de correo
- 📝 **Contenido predefinido** para solicitud de patrocinio
- ✅ **Saludo dinámico** personalizado por contacto
- 📅 **Registro automático** de fecha y hora de envío

#### Correo Personalizado con Editor Rico
- ✏️ **Editor WYSIWYG** (Quill.js) con barra de herramientas completa
- 📬 **Correos desde alias**: Selecciona desde qué dirección de correo enviar
- 🎨 **Personalización completa**:
  - Título y subtítulo del encabezado editables
  - Colores personalizables (encabezado y cuerpo del correo)
  - Selección de logos a incluir (1, 2 o 3 imágenes)
  - Color de viñetas y números de lista
  - Firma personalizable con estilos (negrita, cursiva, color)
- 📝 **Formato de texto enriquecido**:
  - Negritas, cursivas, subrayado, tachado
  - Encabezados (H1-H6)
  - Listas con viñetas y numeradas (con colores personalizados)
  - Alineación de texto (izquierda, centro, derecha, justificado)
  - Enlaces web
  - Colores de texto y fondo
  - Código inline y bloques de código
  - Citas (blockquote)
  - Sangría y formato avanzado
- 👁️ **Vista previa en tiempo real** antes de enviar
- ✅ **Compatible con clientes de correo**: Gmail, Outlook, Apple Mail, etc.
- 📬 **Campo CC (Copia con Copia)**: Enviar copia del correo a múltiples destinatarios
  - Soporta múltiples direcciones separadas por comas o punto y coma
  - Validación automática de formato de cada correo
  - Los destinatarios en CC aparecen en la cabecera del correo
  - Opcional: no es obligatorio para enviar

### ⚙️ Panel de Configuración

- 🔧 **Gestión de correos alias**: Agrega, edita o elimina correos desde los cuales puedes enviar
- 🎨 **Personalización visual**: Configura colores predeterminados para encabezados y cuerpos
- 📝 **Mensajes predeterminados**: Define mensajes estándar para reutilizar
- 📋 **Copyright personalizable**: Configura el texto del pie de página de los correos
- 💾 **Persistencia en JSON**: Todas las configuraciones se guardan en `configuracion.json`
- 🔐 **Integración con .env**: Las credenciales SMTP se cargan desde variables de entorno

### 📊 Panel de Control (Dashboard)

- 📈 **Estadísticas en tiempo real**:
  - Total de contactos
  - Correos pendientes
  - Correos enviados
  - Errores de envío
- 🎯 **Acceso rápido** a las funcionalidades principales:
  - Gestión de Contactos
  - Correo Personalizado
  - Configuración del Sistema
- 🖥️ **Interfaz limpia** con diseño institucional IEEE

### 🔒 Seguridad

- 🔐 **Credenciales protegidas**: Uso de variables de entorno (archivo `.env`)
- 🔑 **Contraseñas de aplicación**: Soporte para Gmail App Passwords
- 📁 **Archivos sensibles ignorados**: `.gitignore` configurado correctamente
- 🛡️ **Validación de datos**: Sanitización de inputs del usuario
- 🔒 **Conexión SMTP segura**: TLS/STARTTLS habilitado
- 🔐 **Autenticación de alias**: Los alias se envían usando la cuenta principal para autenticación

---

## 🛠️ Tecnologías Utilizadas

### Backend
- **[Python 3.8+](https://www.python.org/)**: Lenguaje de programación principal
- **[Flask 3.0.0](https://flask.palletsprojects.com/)**: Framework web ligero y potente
- **[python-dotenv](https://github.com/theskumar/python-dotenv)**: Carga de variables de entorno desde archivo `.env`
- **[smtplib](https://docs.python.org/3/library/smtplib.html)**: Envío de correos electrónicos vía SMTP

### Gestión de Datos
- **[Pandas](https://pandas.pydata.org/)**: Manipulación de datos y DataFrames
- **[OpenPyXL 3.1.2](https://openpyxl.readthedocs.io/)**: Lectura/escritura de archivos Excel (`.xlsx`)

### Frontend
- **HTML5**: Estructura semántica moderna
- **CSS3**: Estilos personalizados con diseño institucional IEEE
- **JavaScript (Vanilla)**: Interactividad sin dependencias pesadas
- **[Quill.js 1.3.6](https://quilljs.com/)**: Editor de texto enriquecido (WYSIWYG)

### Email
- **MIME Multipart**: Correos con imágenes inline (CID)
- **HTML Email Templates**: Diseño responsive compatible con clientes de correo

---

## 📦 Requisitos del Sistema

### Software Requerido

- **Python**: Versión 3.8 o superior
- **pip**: Gestor de paquetes de Python (incluido con Python)
- **Navegador web moderno**: Chrome, Firefox, Edge o Safari
- **Cuenta de correo SMTP**: Gmail, Outlook, o servidor SMTP personalizado

### Requisitos de Red

- Conexión a Internet para envío de correos
- Puerto SMTP 587 (TLS) o 465 (SSL) abierto en el firewall

### Sistema Operativo

- ✅ **Windows** 10/11 (PowerShell 5.1+)
- ✅ **macOS** 10.15+
- ✅ **Linux** (Ubuntu 20.04+, Debian 10+, etc.)
- ✅ **GitHub Codespaces** (Debian GNU/Linux)

---

## 🚀 Instalación

### Opción A: Entorno Local (Windows/macOS/Linux)

#### 1. Clonar el repositorio

```bash
git clone https://github.com/CS-TEC-IEEE-Costa-Rica/Correos.git
cd Correos
```

#### 2. Crear y activar entorno virtual

**Windows (PowerShell):**
```powershell
python -m venv .venv
& ".venv\Scripts\Activate.ps1"
```

**macOS / Linux:**
```bash
python3 -m venv .venv
source .venv/bin/activate
```

#### 3. Instalar dependencias

```bash
pip install -r requirements.txt
```

#### 4. Configurar archivo de contactos

```bash
# Copia el archivo de ejemplo y renómbralo
cp contactos.xlsx.example contactos.xlsx
```

O crea un archivo `contactos.xlsx` vacío (la aplicación lo inicializará automáticamente).

#### 5. Configurar variables de entorno

Copia el archivo de ejemplo como base:

```bash
cp .env.example .env
```

Luego edita `.env` con tus credenciales reales (ver sección [Configuración SMTP](#configuración-smtp)).

#### 6. Iniciar la aplicación

```bash
python app.py
```

La aplicación se iniciará en:
- 🌐 **http://127.0.0.1:5001**
- 🌐 **http://localhost:5001**

---

### Opción B: GitHub Codespaces

GitHub Codespaces proporciona un entorno de desarrollo en la nube completo y preconfigurado.

#### 1. Crear el Codespace

1. Ve al repositorio en GitHub: [CS-TEC-IEEE-Costa-Rica/Correos](https://github.com/CS-TEC-IEEE-Costa-Rica/Correos)
2. Haz clic en el botón verde **`Code`**
3. Selecciona la pestaña **`Codespaces`**
4. Haz clic en **`Create codespace on main`**

GitHub creará y configurará automáticamente el entorno. Esto puede tomar 1-3 minutos.

#### 2. Instalar dependencias

Una vez que el Codespace esté listo, abre la terminal integrada (Terminal → New Terminal) y ejecuta:

```bash
pip3 install -r requirements.txt
```

#### 3. Configurar archivo de contactos

```bash
# Copia el archivo de ejemplo
cp contactos.xlsx.example contactos.xlsx
```

#### 4. Configurar variables de entorno

Copia el archivo de ejemplo como base:

```bash
cp .env.example .env
```

Abre el archivo `.env` y completa las credenciales (ver sección [Configuración SMTP](#configuración-smtp)).

#### 5. Iniciar la aplicación

```bash
python3 app.py
```

#### 6. Abrir la aplicación en el navegador

Cuando Flask se inicie en el puerto **5001**, GitHub Codespaces mostrará una notificación automática:

- **Opción 1**: Haz clic en la notificación **"Open in Browser"**
- **Opción 2**: Manual:
  1. Ve a la pestaña **`PORTS`** (al lado de TERMINAL)
  2. Busca el puerto **5001**
  3. Haz clic en el ícono 🌐 o en **"Open in Browser"**

La aplicación se abrirá en una nueva pestaña del navegador.

> 💡 **Tip**: El puerto se redirige automáticamente a una URL pública como `https://xxx-5001.preview.app.github.dev`

---

### Verificación de Instalación

Verifica que todas las dependencias estén correctamente instaladas:

```bash
python -c "import flask, pandas, openpyxl, dotenv; print('✅ Todas las dependencias instaladas correctamente')"
```

Si ves el mensaje de éxito, la instalación está completa.

---

## 🔧 Configuración

### Configuración SMTP

La aplicación utiliza **python-dotenv** para cargar las credenciales SMTP de forma segura desde un archivo `.env`.

#### Paso 1: Crear el archivo `.env`

Crea un archivo llamado `.env` en la raíz del proyecto:

```ini
# Credenciales SMTP (cuenta principal)
EMAIL_REMITENTE=tu-correo@gmail.com
EMAIL_PASSWORD=tu-contraseña-de-aplicacion

# Configuración del servidor SMTP
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587

# Clave secreta de Flask (genera una aleatoria)
FLASK_SECRET_KEY=tu-clave-secreta-super-segura-y-aleatoria-aqui
```

> ⚠️ **Importante**: El archivo `.env` está en `.gitignore` y **nunca se subirá a GitHub**.

#### Paso 2: Obtener Contraseña de Aplicación (Gmail)

Para usar **Gmail**, no uses tu contraseña normal. Debes crear una **Contraseña de Aplicación**:

1. Ve a [Seguridad de tu Cuenta de Google](https://myaccount.google.com/security)
2. Activa la **Verificación en 2 pasos** (si no la tienes activada)
3. Busca **"Contraseñas de aplicaciones"** o ve directamente a:
   - [https://myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
4. Selecciona:
   - **Aplicación**: Correo
   - **Dispositivo**: Computadora Windows (o el que corresponda)
5. Haz clic en **"Generar"**
6. **Copia la contraseña de 16 caracteres** (puede tener espacios)
7. Pégala en tu archivo `.env` en `EMAIL_PASSWORD`

**Ejemplo de `.env` completo:**
```ini
EMAIL_REMITENTE=julio.barrios@gmail.com
EMAIL_PASSWORD=abcd efgh ijkl mnop
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=587
FLASK_SECRET_KEY=mi-clave-super-secreta-12345-xyz
```

#### Configuración para otros proveedores

**Outlook / Hotmail:**
```ini
SMTP_SERVER=smtp-mail.outlook.com
SMTP_PORT=587
```

**Yahoo Mail:**
```ini
SMTP_SERVER=smtp.mail.yahoo.com
SMTP_PORT=587
```

**Servidor SMTP personalizado:**
```ini
SMTP_SERVER=mail.tu-dominio.com
SMTP_PORT=587  # o 465 para SSL
```

---

### Configuración de Correos Alias

La aplicación permite enviar correos desde múltiples direcciones (alias). Esto es útil para cuentas institucionales compartidas.

#### ¿Cómo funcionan los alias?

- **Autenticación**: Siempreuse la cuenta principal (`EMAIL_REMITENTE`) para autenticarse con el servidor SMTP
- **Remitente visible**: El correo aparece como enviado desde el alias seleccionado
- **Requisito**: El alias debe estar configurado en Gmail (o tu proveedor) como "Enviar como"

#### Configurar alias en Gmail

1. Ve a **Gmail → Configuración → Cuentas e importación**
2. En **"Enviar correo como"**, haz clic en **"Añadir otra dirección de correo electrónico"**
3. Ingresa la dirección del alias (ej: `sbc-tec-cs@ieee.org`)
4. Sigue las instrucciones de verificación
5. Una vez verificado, puedes enviar correos desde ese alias

#### Configurar alias en la aplicación

1. Inicia la aplicación y abre el navegador en `http://127.0.0.1:5001`
2. En el dashboard, haz clic en el ícono de **⚙️ Configuración** (esquina superior derecha del encabezado)
3. En la sección **"Correos Electrónicos (Alias)"**:
   - Haz clic en **"+ Agregar Alias"**
   - Ingresa la dirección de correo del alias
   - Marca como **"Predeterminado"** si deseas que sea el remitente por defecto
   - Haz clic en **"Guardar Configuración"**

Todos los alias configurados aparecerán en el selector de correos al crear mensajes personalizados.

#### Estructura del archivo `configuracion.json`

La configuración se guarda automáticamente en `configuracion.json`:

```json
{
  "email_remitente": "julio.barrios@gmail.com",
  "alias_emails": [
    {
      "email": "julio.barrios@ieee.org",
      "predeterminado": true
    },
    {
      "email": "sbc-tec-cs@ieee.org",
      "predeterminado": false
    }
  ],
  "mensajes_predeterminados": {
    "bienvenida": "Hola, somos IEEE Computer Society..."
  },
  "colores_defecto": {
    "encabezado": "#0066CC",
    "cuerpo": "#003366"
  },
  "copyright": "© 2024 IEEE Computer Society TEC"
}
```

> 💡 **Nota**: Las credenciales SMTP (EMAIL_PASSWORD) **nunca** se guardan en `configuracion.json`, solo en `.env`

---

## 📁 Estructura del Proyecto

```text
Correos/
│
├── 📄 app.py                          # Aplicación Flask principal (~1600 líneas)
│   ├── Sistema de configuración (JSON + .env)
│   ├── Funciones de manejo de Excel
│   ├── Funciones de validación de emails
│   ├── Generación de correos HTML
│   ├── Procesamiento de listas para email
│   ├── Envío de correos vía SMTP (con soporte de alias)
│   └── 12 rutas Flask (endpoints)
│
├── 📄 requirements.txt                # Dependencias de Python
│   ├── Flask==3.0.0
│   ├── pandas==2.1.4; python_version < "3.14"
│   ├── pandas>=2.2; python_version >= "3.14"
│   ├── openpyxl==3.1.2
│   └── python-dotenv==1.0.0
│
├── 📄 .env                            # Variables de entorno (NO se sube a Git)
│   ├── EMAIL_REMITENTE               # Cuenta principal SMTP
│   ├── EMAIL_PASSWORD                # Contraseña de aplicación
│   ├── SMTP_SERVER                   # Servidor SMTP (smtp.gmail.com)
│   ├── SMTP_PORT                     # Puerto SMTP (587)
│   └── FLASK_SECRET_KEY              # Clave secreta de Flask
│
├── 📄 .env.example                    # Template de variables de entorno (sí se sube a Git)
│   └── (Mismas variables que .env con valores de ejemplo)
│
├── 📄 configuracion.json              # Configuración persistente (NO se sube a Git)
│   ├── alias_emails[]                # Lista de correos alias
│   ├── mensajes_predeterminados{}    # Mensajes predefinidos
│   ├── colores_defecto{}             # Colores para correos
│   └── copyright                     # Texto de copyright
│
├── 📄 contactos.xlsx                  # Base de datos Excel (NO se sube a Git)
├── 📄 contactos.xlsx.example          # Archivo de ejemplo para contactos
│
├── 📄 README.md                       # Este archivo - Documentación completa
├── 📄 PREVIEW_SYSTEM.md               # Documentación del sistema de vista previa
├── 📄 PRINT_SYSTEM.md                 # Documentación del sistema de impresión/PDF
├── 📄 .gitignore                      # Archivos excluidos de Git
│
├── 📂 templates/                      # Plantillas HTML (Jinja2)
│   ├── index.html                     # Gestión de contactos con tabla paginada
│   ├── dashboard.html                 # Dashboard principal con estadísticas
│   ├── correo_personalizado.html     # Editor WYSIWYG de correos (Quill.js)
│   └── configuracion.html             # Panel de configuración del sistema
│
├── 📂 static/                         # Archivos estáticos (CSS)
│   └── style.css                      # Estilos CSS institucionales IEEE (~1000 líneas)
│
├── 📂 images/                         # Logos IEEE para correos
│   ├── ieee.png                       # Logo principal IEEE
│   ├── ieee cs imagen.png            # Logo IEEE Computer Society
│   └── ieee costa rica.png           # Logo IEEE Costa Rica
│
└── 📂 .venv/                          # Entorno virtual Python (NO se sube a Git)
```

### Descripción de Archivos Clave

#### `app.py` - Aplicación Principal

**Sistema de configuración:**
- `cargar_configuracion()`: Carga configuración desde `configuracion.json` y `.env`
- `guardar_configuracion()`: Guarda configuración en `configuracion.json`
- `obtener_configuracion()`: Retorna configuración actual (para API)

**Manejo de contactos:**
- `inicializar_excel()`: Crea/verifica estructura del archivo Excel
- `leer_contactos()`: Lee contactos desde Excel como DataFrame
- `guardar_contactos()`: Guarda DataFrame en Excel
- `agregar_contacto()`: Agrega nuevo contacto con validación de duplicados
- `actualizar_estado_envio()`: Marca contacto como enviado con fecha
- `eliminar_contacto()`: Elimina contacto por índice
- `validar_correo()`: Valida formato de email con regex

**Generación de correos:**
- `generar_saludo()`: Crea saludo dinámico personalizado
- `generar_cuerpo_html()`: Genera HTML del correo predeterminado
- `generar_cuerpo_html_personalizado()`: Genera HTML del correo personalizado
- `procesar_listas_para_email()`: Convierte listas HTML a formato compatible con clientes de correo

**Envío SMTP:**
- `enviar_correo()`: Envía correo predeterminado vía SMTP
- `enviar_correo_personalizado_smtp()`: Envía correo personalizado con soporte de alias

**Rutas (Endpoints):**

| Ruta | Método | Descripción |
|------|--------|-------------|
| `/` | GET | Dashboard principal con estadísticas |
| `/contactos` | GET | Vista de gestión de contactos con tabla paginada |
| `/agregar` | POST | Agrega nuevo contacto al Excel |
| `/enviar/<id>` | GET | Envía correo predeterminado a un contacto |
| `/vista_previa/<id>` | GET | Muestra vista previa del correo (JSON) |
| `/reimportar` | GET | Reinicia datos desde Excel |
| `/eliminar/<id>` | GET | Elimina contacto del Excel |
| `/correo_personalizado` | GET | Formulario de correo personalizado con editor |
| `/api/preview_correo_personalizado` | POST | API para generar vista previa del correo |
| `/api/configuracion` | GET | API para obtener configuración actual |
| `/enviar_correo_personalizado` | POST | Procesa y envía correo personalizado |
| `/configuracion` | GET | Panel de configuración del sistema |
| `/guardar_configuracion` | POST | Guarda configuración (alias, colores, mensajes) |
| `/api/configuracion` | GET | API JSON con configuración actual |

#### `templates/` - Plantillas HTML

- **`dashboard.html`**: Pantalla de inicio con acceso a funcionalidades principales
- **`index.html`**: Gestión completa de contactos (CRUD, búsqueda, filtros, paginación)
- **`correo_personalizado.html`**: Editor rico con Quill.js y selector de alias
- **`configuracion.html`**: Configuración de alias, colores, mensajes y copyright

#### `static/style.css` - Estilos Institucionales

- Variables CSS para colores IEEE (`--ieee-blue`, `--ieee-red`, etc.)
- Componentes reutilizables (botones, formularios, tarjetas)
- Estilos para header, footer, tablas responsivas
- Animaciones y transiciones suaves
- Diseño mobile-first

#### `.env` - Variables de Entorno

**Archivo crítico que NO se sube a Git.** Contiene:
- Credenciales SMTP (EMAIL_REMITENTE, EMAIL_PASSWORD)
- Configuración del servidor SMTP
- Clave secreta de Flask

#### `configuracion.json` - Configuración Persistente

**Archivo que NO se sube a Git.** Almacena:
- Lista de correos alias configurados
- Mensajes predeterminados reutilizables
- Colores por defecto para correos
- Texto del copyright institucional

---

## 🎮 Uso de la Aplicación

### 1. Iniciar el Servidor

Con dependencias instaladas y `.env` configurado:

```bash
# En Local
python app.py

# En Codespaces
python3 app.py
```

### 2. Acceder a la Aplicación

- **Local**: http://127.0.0.1:5001 o http://localhost:5001
- **Codespaces**: Abre el puerto 5001 desde la pestaña PORTS

### 3. Navegar por el Dashboard

En la pantalla principal verás las opciones principales:

- **📇 Gestión de Contactos**: Administrar lista de contactos
- **✉️ Correo Personalizado**: Crear y enviar correos personalizados
- **⚙️ Configuración**: Configurar alias, colores y mensajes (ícono en header)

---

## 🔥 Funcionalidades Detalladas

### 📇 Gestión de Contactos

#### Agregar Nuevo Contacto

1. En la página `/contactos`, completa el formulario:
   - **Empresa** (obligatorio): Nombre de la organización
   - **Contacto** (opcional): Nombre de la persona
   - **Correo** (obligatorio): Email válido
2. Click en "Agregar Contacto"
3. El sistema valida:
   - Formato de email correcto
   - No duplicados (por correo)
4. Contacto agregado con estado "Pendiente"

#### Buscar y Filtrar Contactos

- **Búsqueda en tiempo real**: Escribe en la barra de búsqueda para filtrar por empresa, contacto o correo
- **Filtros por estado**:
  - Todos: Muestra todos los contactos
  - Pendientes: Solo contactos sin envío
  - Enviados: Solo contactos con correo enviado

#### Ordenar Contactos

Haz clic en los encabezados de columna para ordenar:
- **Por Excel**: Orden original del archivo
- **Por Empresa**: Orden alfabético (A-Z / Z-A)
- **Por Estado**: Agrupa por pendiente/enviado
- **Por Fecha de Envío**: Del más reciente al más antiguo

#### Acciones sobre Contactos

- **✉️ Enviar**: Envía correo predeterminado institucional
- **👁️ Vista Previa**: Muestra cómo se verá el correo antes de enviar
- **🗑️ Eliminar**: Elimina el contacto permanentemente (con confirmación)

### ✉️ Correo Personalizado con Editor Rico

#### Paso 1: Acceder al Editor

Desde el dashboard, haz clic en **"✉️ Correo Personalizado"**

#### Paso 2: Configurar el Correo

**Información del Destinatario:**
- **Correo electrónico** (obligatorio)
- **CC - Copia con copia** (opcional): múltiples correos separados por comas
- **Alias remitente**: Selecciona desde qué correo enviar (carga desde configuración)
- **Asunto del correo** (obligatorio)

**Personalización Visual:**
- **Título del encabezado**: Por defecto "IEEE Computer Society"
- **Subtítulo**: Información adicional (opcional)
- **Color del encabezado**: Selector de color
- **Color del cuerpo**: Selector de color
- **Logos a incluir**: Marca 1, 2 o 3 logos (IEEE, IEEE CR, IEEE CS)

#### Paso 3: Redactar el Mensaje

Usa el editor Quill.js con las siguientes herramientas:

- **Formato de texto**: Negrita, cursiva, subrayado, tachado
- **Encabezados**: H1, H2, H3
- **Listas**: Con viñetas o numeradas (con color personalizable)
- **Alineación**: Izquierda, centro, derecha, justificado
- **Enlaces**: Inserta URLs
- **Colores**: Texto y fondo
- **Código**: Bloques de código o inline
- **Citas**: Blockquote
- **Limpiar formato**: Elimina todos los estilos

#### Paso 4: Configurar Firma

- Escribe tu firma en el campo correspondiente
- Usa los controles para aplicar formato (negrita, cursiva, color)

#### Paso 5: Vista Previa y Envío

- Haz clic en **"Vista Previa"** para ver cómo se verá el correo
- Verifica que todo esté correcto
- Haz clic en **"Enviar Correo"**
- Confirma el envío en el diálogo

### ⚙️ Panel de Configuración

#### Acceder a Configuración

Haz clic en el ícono **⚙️** en la esquina superior derecha del encabezado (header).

#### Gestión de Alias de Correos

- **Agregar alias**: Haz clic en "+ Agregar Alias"
- **Configurar como predeterminado**: Marca el checkbox para definir el remitente por defecto
- **Eliminar alias**: Haz clic en "Eliminar" junto al alias no deseado

> ⚠️ **Importante**: Los alias deben estar configurados en Gmail como "Enviar como" para funcionar correctamente

#### Personalización de Colores

- **Color del encabezado**: Define el color predeterminado para encabezados de correos
- **Color del cuerpo**: Define el color predeterminado para el cuerpo de correos

#### Mensajes Predeterminados

- Configura mensajes estándar que uses frecuentemente
- Útil para plantillas de invitaciones, solicitudes, etc.

#### Copyright

- Define el texto que aparecerá en el pie de página de todos los correos
- Ejemplo: "© 2024 IEEE Computer Society TEC - Todos los derechos reservados"

#### Guardar Cambios

- Haz clic en **"Guardar Configuración"**
- Los cambios se guardan en `configuracion.json`
- Verás una confirmación de éxito

---

## 🔒 Seguridad y Buenas Prácticas

### Protección de Credenciales

✅ **Hacer**:
- Usa siempre contraseñas de aplicación (no tu contraseña real)
- Guarda credenciales en archivo `.env` (nunca en el código)
- Verifica que `.env` esté en `.gitignore`
- Usa variables de entorno en producción

❌ **No hacer**:
- Nunca subas el archivo `.env` a GitHub
- Nunca compartas tu contraseña de aplicación
- No uses tu contraseña de Gmail real en la aplicación
- No hardcodees credenciales en `app.py`

### Validación de Datos

La aplicación valida automáticamente:
- ✅ Formato de correos electrónicos (regex)
- ✅ Duplicados en la base de datos
- ✅ Campos obligatorios en formularios
- ✅ Formato de correos en campo CC

### Conexión SMTP Segura

- Todas las conexiones use TLS/STARTTLS (puerto 587)
- Las credenciales se cifran en tránsito
- Los alias se autentican con la cuenta principal

### Archivos Sensibles

El `.gitignore` está configurado para excluir:
- `.env` (credenciales)
- `configuracion.json` (configuración con alias)
- `contactos.xlsx` (datos personales)
- `.venv/` (entorno virtual)
- `__pycache__/` (archivos compilados de Python)

---

## 🛠️ Solución de Problemas

### Error: "Address already in use" (Puerto 5001 ocupado)

**Causa**: Otra instancia de la aplicación o proceso está usando el puerto 5001.

**Solución**:

```bash
# Linux/macOS
lsof -ti:5001 | xargs kill -9

# Windows
netstat -ano | findstr :5001
taskkill /PID <PID> /F
```

O cambia el puerto en `app.py`:
```python
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5002)  # Usa puerto 5002
```

### Error: "SMTPAuthenticationError (535)"

**Causa**: Gmail está rechazando las credenciales.

**Soluciones**:

1. **Regenera la contraseña de aplicación**:
   - Ve a [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Elimina la contraseña antigua
   - Genera una nueva y cópiala a `.env`

2. **Verifica la verificación en 2 pasos**:
   - Debe estar activa en tu cuenta de Google

3. **Revisa el archivo `.env`**:
   - Asegúrate que `EMAIL_PASSWORD` tiene la contraseña correcta
   - Quita espacios innecesarios o usa comillas si es necesario

### Error: "No module named 'flask'" o similares

**Causa**: Dependencias no instaladas.

**Solución**:

```bash
# Asegúrate de tener el entorno virtual activado
source .venv/bin/activate  # Linux/macOS
.venv\Scripts\Activate.ps1  # Windows

# Reinstala dependencias
pip install -r requirements.txt
```

### Error: "Permission denied" al crear archivos

**Causa**: Permisos insuficientes en el directorio.

**Solución (Linux/macOS)**:

```bash
chmod -R 755 /workspaces/Correos
chmod 644 contactos.xlsx
```

### Los alias no aparecen en el selector

**Causa**: No has configurado alias en la página de configuración.

**Solución**:

1. Ve a `http://127.0.0.1:5001` → ícono ⚙️
2. Agrega alias en "Correos Electrónicos (Alias)"
3. Guarda configuración
4. Recarga la página de correo personalizado

### Los correos no se envían desde el alias

**Causa**: El alias no está configurado en Gmail como "Enviar como".

**Solución**:

1. Ve a Gmail → Configuración → Cuentas e importación
2. En "Enviar correo como", agrega el alias
3. Verifica el alias siguiendo las instrucciones de Gmail
4. Intenta enviar nuevamente

### Error: "Credential loading" o configuracion.json corrupto

**Causa**: El archivo `configuracion.json` tiene formato JSON inválido.

**Solución**:

```bash
# Elimina el archivo corrupto
rm configuracion.json

# La aplicación creará uno nuevo al iniciar
python app.py
```

### En Codespaces: El puerto no se abre automáticamente

**Solución manual**:

1. Ve a la pestaña **PORTS** en VS Code Web
2. Si no ves el puerto 5001, haz clic en **"Forward a Port"**
3. Ingresa **5001**
4. Haz clic en el ícono 🌐 para abrir en el navegador

### Las imágenes (logos) no se muestran en los correos

**Causa**: Los archivos de logos no están en la carpeta `images/`.

**Solución**:

```bash
# Verifica que existan los logos
ls images/
# Deberías ver: ieee.png, ieee cs imagen.png, ieee costa rica.png

# Si faltan, asegúrate de tener los archivos correctos
```

---

## 🤝 Contribución

¡Las contribuciones son bienvenidas! Si deseas mejorar este proyecto:

1. **Fork** el repositorio
2. Crea una **rama** para tu funcionalidad (`git checkout -b feature/nueva-funcionalidad`)
3. **Commit** tus cambios (`git commit -m 'Agrega nueva funcionalidad'`)
4. **Push** a la rama (`git push origin feature/nueva-funcionalidad`)
5. Abre un **Pull Request**

### Áreas de Mejora

- 🔄 Integración con bases de datos SQL (SQLite, PostgreSQL)
- 📊 Exportación de reportes en PDF
- 📧 Plantillas adicionales de correos
- 🌐 Internacionalización (i18n)
- 🧪 Tests unitarios y de integración
- 📱 Diseño responsive mejorado para móviles
- 🔐 Autenticación de usuarios
- 📨 Envío de correos en lotes (batch)

---

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.

---

## 👤 Autores

**IEEE Computer Society - Instituto Tecnológico de Costa Rica**

- 🌐 Organización: [CS-TEC-IEEE-Costa-Rica](https://github.com/CS-TEC-IEEE-Costa-Rica)
- 📧 Contacto: ieee-cs@tec.ac.cr
- 🌍 Sitio web: [IEEE Costa Rica](https://ieee.org/costarica)

---

## 🙏 Agradecimientos

- **[Flask](https://flask.palletsprojects.com/)**: Por el excelente framework web
- **[Quill.js](https://quilljs.com/)**: Por el editor de texto enriquecido
- **[Pandas](https://pandas.pydata.org/)**: Por la manipulación de datos
- **IEEE**: Por el soporte institucional

---

<div align="center">

**Desarrollado con ❤️ por IEEE Computer Society TEC**

⭐ Si te gusta este proyecto, por favor dale una estrella en GitHub ⭐

</div>
