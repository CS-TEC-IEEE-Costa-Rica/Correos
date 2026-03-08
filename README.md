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
- [Configuración SMTP](#-configuración-smtp)
- [Estructura del Proyecto](#-estructura-del-proyecto)
- [Uso de la Aplicación](#-uso-de-la-aplicación)
- [Funcionalidades Detalladas](#-funcionalidades-detalladas)
- [Seguridad y Buenas Prácticas](#-seguridad-y-buenas-prácticas)
- [Solución de Problemas](#-solución-de-problemas)
- [Publicar en GitHub](#-publicar-en-github)
- [Contribución](#-contribución)
- [Autor](#-autor)

---

## 🎯 Descripción

Sistema web desarrollado en Flask para el **IEEE Computer Society del Instituto Tecnológico de Costa Rica, Campus Central Cartago**, diseñado para facilitar la gestión de contactos empresariales y el envío masivo de correos electrónicos profesionales con formato institucional personalizable.

### Casos de Uso Principales

- **Gestión de patrocinios**: Contactar empresas para solicitar patrocinios de eventos estudiantiles
- **Invitaciones a charlas**: Enviar invitaciones profesionales a expertos del sector
- **Networking corporativo**: Mantener comunicación con empresas colaboradoras
- **Ferias de empleo**: Coordinar participación de empresas en eventos de reclutamiento
- **Correos personalizados**: Crear y enviar correos con editor de texto enriquecido (WYSIWYG)

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

### 📊 Panel de Control (Dashboard)

- 📈 **Estadísticas en tiempo real**:
  - Total de contactos
  - Correos pendientes
  - Correos enviados
  - Errores de envío
- 🎯 **Acceso rápido** a las funcionalidades principales
- 🖥️ **Interfaz limpia** con diseño institucional IEEE

### 🔒 Seguridad

- 🔐 **Credenciales protegidas**: Uso de variables de entorno
- 🔑 **Contraseñas de aplicación**: Soporte para Gmail App Passwords
- 📁 **Archivos sensibles ignorados**: `.gitignore` configurado correctamente
- 🛡️ **Validación de datos**: Sanitización de inputs del usuario
- 🔒 **Conexión SMTP segura**: TLS/STARTTLS habilitado

---

## 🛠️ Tecnologías Utilizadas

### Backend
- **[Python 3.8+](https://www.python.org/)**: Lenguaje de programación principal
- **[Flask 3.0.0](https://flask.palletsprojects.com/)**: Framework web ligero y potente
- **[smtplib](https://docs.python.org/3/library/smtplib.html)**: Envío de correos electrónicos vía SMTP

### Gestión de Datos
- **[Pandas](https://pandas.pydata.org/)**: Manipulación de datos y DataFrames (`2.1.4` para Python < 3.14, `2.2+` para Python >= 3.14)
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

---

## 🚀 Instalación

### 1. Clonar o Descargar el Repositorio

#### Opción A: Clonar con Git

```powershell
git clone <URL_DEL_REPOSITORIO>
cd "carpeta"
```

#### Opción B: Descargar ZIP

1. Descargar el archivo ZIP desde GitHub
2. Extraer en la ubicación deseada
3. Abrir terminal en la carpeta del proyecto

### 2. Crear Entorno Virtual

Un entorno virtual aísla las dependencias del proyecto de tu instalación global de Python.

```powershell
# Windows (PowerShell)
python -m venv .venv
```

```bash
# macOS / Linux
python3 -m venv .venv
```

### 3. Activar Entorno Virtual

#### Windows (PowerShell)

```powershell
& ".venv\Scripts\Activate.ps1"
```

Si obtienes un error de permisos, ejecuta:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

#### macOS / Linux

```bash
source .venv/bin/activate
```

### 4. Instalar Dependencias

Con el entorno virtual activado:

```powershell
pip install -r requirements.txt
```

**Dependencias instaladas:**
- `Flask==3.0.0`: Framework web
- `pandas==2.1.4; python_version < "3.14"`: Manipulación de datos (Python 3.8 a 3.13)
- `pandas>=2.2; python_version >= "3.14"`: Compatibilidad con Python 3.14+
- `openpyxl==3.1.2`: Lectura/escritura de Excel

### 5. Verificar Instalación

```powershell
python -c "import flask, pandas, openpyxl; print('✅ Todas las dependencias instaladas correctamente')"
```

---

## 🔧 Configuración SMTP

### Variables de Entorno (Recomendado)

#### Windows (PowerShell)

```powershell
# Configurar variables de entorno para la sesión actual
$env:EMAIL_REMITENTE = "tu-correo@ieee.org"
$env:EMAIL_PASSWORD = "tu-contrasena-de-aplicacion"
$env:FLASK_SECRET_KEY = "clave-secreta-super-larga-y-aleatoria-12345"

# Para persistencia entre sesiones (opcional)
[System.Environment]::SetEnvironmentVariable('EMAIL_REMITENTE', 'tu-correo@ieee.org', 'User')
[System.Environment]::SetEnvironmentVariable('EMAIL_PASSWORD', 'tu-contrasena-de-aplicacion', 'User')
[System.Environment]::SetEnvironmentVariable('FLASK_SECRET_KEY', 'clave-secreta-super-larga-y-aleatoria-12345', 'User')
```

#### macOS / Linux (Bash/Zsh)

```bash
# Agregar al archivo ~/.bashrc o ~/.zshrc para persistencia
export EMAIL_REMITENTE="tu-correo@ieee.org"
export EMAIL_PASSWORD="tu-contrasena-de-aplicacion"
export FLASK_SECRET_KEY="clave-secreta-super-larga-y-aleatoria-12345"

# Recargar configuración
source ~/.bashrc  # o source ~/.zshrc
```

### Configuración en el Código (No Recomendado para Producción)

Si no usas variables de entorno, edita `app.py` (líneas 26-32):

```python
EMAIL_REMITENTE = "tu-correo@ieee.org"
EMAIL_PASSWORD = "tu-contrasena-de-aplicacion"
SMTP_SERVER = "smtp.gmail.com"  # Cambiar según tu proveedor
SMTP_PORT = 587  # 587 para TLS, 465 para SSL
```

⚠️ **ADVERTENCIA**: Nunca subas credenciales reales a repositorios públicos.

### Configuración para Gmail

1. **Habilitar verificación en 2 pasos**:
   - Ve a [myaccount.google.com/security](https://myaccount.google.com/security)
   - Activa "Verificación en 2 pasos"

2. **Crear contraseña de aplicación**:
   - Ve a [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)
   - Selecciona "Correo" y "Windows Computer" (o dispositivo correspondiente)
   - Copia la contraseña de 16 caracteres generada
   - Usa esta contraseña en `EMAIL_PASSWORD` (sin espacios)

3. **Configuración SMTP de Gmail**:
   ```python
   SMTP_SERVER = "smtp.gmail.com"
   SMTP_PORT = 587
   ```

### Configuración para Outlook/Hotmail

```python
SMTP_SERVER = "smtp-mail.outlook.com"
SMTP_PORT = 587
```

### Configuración para Servidor SMTP Personalizado

```python
SMTP_SERVER = "mail.tuservidor.com"
SMTP_PORT = 587  # o el puerto que use tu servidor
```

---

## 📁 Estructura del Proyecto

```text
Programa contactos/
│
├── 📄 app.py                          # Aplicación Flask principal (1200+ líneas)
│   ├── Configuración SMTP
│   ├── Funciones de manejo de Excel
│   ├── Funciones de validación
│   ├── Generación de correos HTML
│   ├── Procesamiento de listas para email
│   ├── Envío de correos vía SMTP
│   └── 9 rutas Flask (endpoints)
│
├── 📄 requirements.txt                # Dependencias de Python
│   ├── Flask==3.0.0
│   ├── pandas==2.1.4; python_version < "3.14"
│   ├── pandas>=2.2; python_version >= "3.14"
│   └── openpyxl==3.1.2
│
├── 📄 README.md                       # Este archivo
├── 📄 .gitignore                      # Archivos excluidos de Git
│
├── 📂 templates/                      # Plantillas HTML (Jinja2)
│   ├── index.html                     # Vista principal de gestión de contactos
│   ├── dashboard.html                 # Dashboard principal con estadísticas
│   └── correo_personalizado.html     # Editor de correos personalizados
│
├── 📂 static/                         # Archivos estáticos
│   └── style.css                      # Estilos CSS institucionales IEEE (~1000 líneas)
│
├── 📂 images/                         # Logos IEEE para correos
│   ├── ieee.png                       # Logo principal IEEE
│   ├── ieee cs imagen.png            # Logo IEEE Computer Society
│   └── ieee costa rica.png           # Logo IEEE Costa Rica
│
├── 📂 .venv/                          # Entorno virtual (no se sube a Git)
├── 📄 contactos.xlsx                  # Base de datos Excel (no se sube a Git)
└── 📄 database.db                     # Base de datos SQLite (no se sube a Git)
```

### Descripción de Archivos Clave

#### `app.py` - Lógica del Servidor

**Funciones principales:**

- `inicializar_excel()`: Crea/verifica estructura del archivo Excel
- `leer_contactos()`: Lee contactos desde Excel como DataFrame
- `guardar_contactos()`: Guarda DataFrame en Excel
- `agregar_contacto()`: Agrega nuevo contacto con validación de duplicados
- `actualizar_estado_envio()`: Marca contacto como enviado con fecha
- `eliminar_contacto()`: Elimina contacto por índice
- `validar_correo()`: Valida formato de email con regex
- `generar_saludo()`: Crea saludo dinámico personalizado
- `generar_cuerpo_html()`: Genera HTML del correo predeterminado
- `generar_cuerpo_html_personalizado()`: Genera HTML del correo personalizado
- `procesar_listas_para_email()`: Convierte listas HTML a formato compatible con clientes de correo
- `enviar_correo()`: Envía correo predeterminado vía SMTP
- `enviar_correo_personalizado_smtp()`: Envía correo personalizado vía SMTP

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
| `/enviar_correo_personalizado` | POST | Procesa y envía correo personalizado |

#### `templates/index.html` - Vista Principal

- Tabla de contactos con búsqueda, filtros y paginación
- Formulario de agregar contacto
- Tarjetas de estadísticas
- Botones de acción (enviar, eliminar, vista previa)
- Diseño responsive con CSS Grid/Flexbox

#### `templates/dashboard.html` - Dashboard

- Pantalla de inicio con dos opciones principales
- Tarjetas con enlaces a gestión de contactos y correos personalizados
- Diseño limpio y profesional

#### `templates/correo_personalizado.html` - Editor de Correos

- Editor Quill.js con barra de herramientas completa
- Selectores de colores para personalización visual
- Checkboxes para selección de logos
- Vista previa modal en tiempo real
- Validación de campos obligatorios

#### `static/style.css` - Estilos Institucionales

- Variables CSS para colores IEEE (`--ieee-blue`, etc.)
- Estilos para header, footer, botones, formularios
- Cards de estadísticas con colores según estado
- Tabla responsive con diseño profesional
- Animaciones y transiciones suaves
- Diseño mobile-first

---

## 🎮 Uso de la Aplicación

### 1. Iniciar el Servidor

Con el entorno virtual activado y las variables de entorno configuradas:

```powershell
python app.py
```

Deberías ver:

```
============================================================
  Sistema de Correos Institucionales - IEEE
============================================================
✅ Archivo Excel inicializado correctamente
Servidor en ejecución: http://127.0.0.1:5000
Presiona Ctrl+C para detener el servidor
============================================================
```

### 2. Acceder a la Aplicación

Abre tu navegador en:

- **Principal**: [http://127.0.0.1:5000](http://127.0.0.1:5000)
- **Alternativo**: [http://localhost:5000](http://localhost:5000)

### 3. Navegar por el Dashboard

En la pantalla principal verás dos opciones:

- **📇 Gestión de Contactos**: Administrar lista de contactos
- **✉️ Correo Personalizado**: Crear y enviar correos personalizados

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

#### Buscar Contactos

- Usa la barra de búsqueda en la parte superior
- Busca por: nombre de empresa, contacto o correo
- La búsqueda es **en tiempo real** (sin necesidad de presionar Enter)

#### Filtrar por Estado

- **Todos**: Muestra todos los contactos
- **Pendientes**: Solo contactos sin envío
- **Enviados**: Solo contactos con correo enviado

#### Ordenar Contactos

Haz clic en los encabezados de columna:

- **Por Excel**: Orden original del archivo
- **Por Empresa**: Orden alfabético (A-Z / Z-A)
- **Por Estado**: Agrupa por pendiente/enviado
- **Por Fecha de Envío**: Del más reciente al más antiguo

#### Enviar Correo Predeterminado

1. Click en el botón "✉️ Enviar" junto al contacto
2. El sistema:
   - Genera el saludo dinámico
   - Crea el HTML con logos inline
   - Envía vía SMTP
   - Actualiza estado a "Enviado"
   - Registra fecha y hora

#### Vista Previa del Correo

1. Click en "👁️ Vista Previa"
2. Se abre una ventana modal con el correo renderizado
3. Puedes verificar el contenido antes de enviar

#### Eliminar Contacto

1. Click en "🗑️ Eliminar"
2. Confirmación del navegador
3. Contacto eliminado permanentemente del Excel

### ✉️ Correo Personalizado con Editor Rico

#### Acceder al Editor

Desde el dashboard, click en "✉️ Correo Personalizado"

#### Configurar el Correo

**1. Información del Destinatario**

- **Correo electrónico** (obligatorio)
- **Asunto del correo** (obligatorio)

**2. Encabezado Personalizable**

- **Título**: Por defecto "IEEE Computer Society" (editable)
- **Subtítulo**: Información adicional (opcional)
- **Color del encabezado**: Selector de color (por defecto azul IEEE)
- **Color del cuerpo**: Selector de color (por defecto azul oscuro)

**3. Configuración de Imágenes**

Selecciona 1, 2 o 3 logos para el encabezado:

- ☑️ **IEEE Logo**: Logo principal
- ☑️ **IEEE Costa Rica**: Logo sección nacional
- ☑️ **IEEE Computer Society**: Logo sociedad técnica

Las imágenes se distribuyen automáticamente:
- 1 imagen: Alineada a la izquierda
- 2 imágenes: Izquierda y derecha
- 3 imágenes: Izquierda, centro y derecha

**4. Contenido del Correo con Editor Rico**

Usa la barra de herramientas de Quill.js:

| Función | Descripción |
|---------|-------------|
| **B** | Texto en negrita |
| **I** | Texto en cursiva |
| **U** | Subrayado |
| **S** | Tachado |
| **H1-H6** | Encabezados (tamaños) |
| **🎨** | Color de texto |
| **🖌️** | Color de fondo de texto |
| **• Lista** | Lista con viñetas |
| **1. Lista** | Lista numerada |
| **⬅️ ➡️** | Alineación (izq, centro, der, justificado) |
| **🔗** | Insertar enlace |
| **" "** | Cita (blockquote) |
| **{ }** | Bloque de código |
| **🧹** | Limpiar formato |

**5. Personalización de Estilo de Firma**

- **Color de firma**: Selector de color
- **Negrita**: Checkbox para aplicar negrita
- **Cursiva**: Checkbox para aplicar cursiva
- **Texto de firma** (opcional): Usa firma predeterminada si se deja vacío

**6. Color de Viñetas**

- Selector de color para bullets (•) y números (1., 2., etc.)
- Compatible con todos los clientes de correo electrónico

#### Vista Previa del Correo

1. Click en "👁️ Vista Previa"
2. Se abre un modal con:
   - Encabezado "Para:" y "Asunto:"
   - Renderizado exacto del correo
   - Logos simulados (cajas de colores en preview)
3. Click en "×" para cerrar

#### Enviar Correo Personalizado

1. Completa todos los campos obligatorios
2. Click en "✉️ Enviar Correo"
3. El sistema:
   - Valida campos obligatorios
   - Procesa el HTML del editor
   - Convierte listas a formato compatible con email
   - Adjunta logos como imágenes inline (CID)
   - Envía vía SMTP con TLS
4. Redirección al dashboard con mensaje de éxito/error

### 📊 Estadísticas en Tiempo Real

En `/contactos`, la barra superior muestra:

- **Total Contactos**: Contador total
- **Pendientes**: Correos no enviados (color naranja)
- **Enviados**: Correos enviados exitosamente (color verde)
- **Con Error**: Errores de envío (color rojo) - actualmente sin uso

---

## 🔒 Seguridad y Buenas Prácticas

### Protección de Credenciales

✅ **Hazlo:**
- Usa variables de entorno para credenciales
- Genera contraseñas de aplicación (no uses tu contraseña principal)
- Mantén actualizado el `.gitignore`
- Usa HTTPS en producción (detrás de un proxy inverso)

❌ **No lo hagas:**
- Subir credenciales reales a Git/GitHub
- Compartir contraseñas en texto plano
- Usar tu contraseña personal de Gmail
- Exponer el servidor Flask directamente a Internet en producción

### Archivos Sensibles (`.gitignore`)

El `.gitignore` ya excluye:

```gitignore
# Credenciales
.env
.env.*

# Datos locales
contactos.xlsx
database.db

# Entornos virtuales
.venv/
venv/
```

### Validación de Datos

El sistema implementa:

- ✅ Validación de formato de email con regex
- ✅ Sanitización de inputs HTML
- ✅ Prevención de duplicados
- ✅ Límites de longitud en campos
- ✅ Validación de tipos de datos

### Conexión SMTP Segura

- 🔒 **TLS**: Conexión encriptada usando `starttls()`
- 🔐 **Autenticación**: Login con credenciales antes de enviar
- ⚠️ **Manejo de errores**: Captura de excepciones SMTP específicas

---

## 🛠️ Solución de Problemas

### Error: "No se pudo conectar al servidor SMTP"

**Causas posibles:**
- Firewall bloqueando puerto 587/465
- Servidor SMTP incorrecto
- Sin conexión a Internet

**Soluciones:**
1. Verifica conexión a Internet: `ping smtp.gmail.com`
2. Revisa configuración de `SMTP_SERVER` y `SMTP_PORT`
3. Prueba con otro puerto (465 para SSL)

### Error: "Error de autenticación SMTP (535)"

**Causas posibles:**
- Contraseña incorrecta
- No has creado contraseña de aplicación (Gmail)
- Verificación en 2 pasos no habilitada

**Soluciones:**
1. Verifica que `EMAIL_PASSWORD` sea la contraseña de aplicación (16 caracteres sin espacios)
2. Habilita verificación en 2 pasos en Gmail
3. Genera una nueva contraseña de aplicación: [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords)

### Error: "No module named 'flask'"

**Causa:** Dependencias no instaladas

**Solución:**
```powershell
# Asegúrate de tener el entorno virtual activado
& ".venv\Scripts\Activate.ps1"
pip install -r requirements.txt
```

### Error: "No module named 'pandas'" en Python 3.14+

**Causa:** `pandas==2.1.4` no es compatible con Python 3.14 y la instalación puede fallar.

**Solución:** usar `requirements.txt` actualizado con marcadores por versión de Python:

```powershell
pip install -r requirements.txt
```

Si instalaste manualmente una versión fija de pandas antes, reinstala con:

```powershell
python -m pip install --upgrade "pandas>=2.2"
```

### Error: "Permission denied" al activar entorno virtual

**Causa:** Política de ejecución de PowerShell restringida

**Solución:**
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Los colores de las viñetas no se ven en Gmail/Outlook

✅ **Este problema ya está solucionado** en la versión actual.

La función `procesar_listas_para_email()` convierte las listas HTML a tablas compatibles con clientes de correo, asegurando que los colores personalizados se vean correctamente.

### El archivo Excel está corrupto

**Solución:**
1. Renombra `contactos.xlsx` a `contactos_backup.xlsx`
2. Reinicia el servidor (`python app.py`)
3. El sistema creará un nuevo archivo Excel vacío
4. Importa manualmente los contactos desde el backup si es necesario

### Puerto 5000 ya está en uso

**Causa:** Otro proceso usando el puerto 5000

**Solución:**

```powershell
# Windows: Encontrar proceso usando puerto 5000
netstat -ano | findstr :5000
taskkill /PID <PID> /F

# O cambiar el puerto en app.py (última línea):
app.run(host="0.0.0.0", port=8080, debug=True)
```

---

## 📤 Publicar en GitHub

### 1. Inicializar Repositorio Git

```powershell
# Inicializar repositorio local (si no existe)
git init
```

### 2. Crear Repositorio en GitHub

1. Ve a [github.com/new](https://github.com/new)
2. Nombre: `ieee-email-management-system`
3. Descripción: "Sistema de gestión y envío de correos institucionales IEEE"
4. Visibilidad: Público o Privado
5. **No inicialices con README** (ya tienes uno)
6. Click en "Create repository"

### 3. Configurar Git Local

```powershell
# Configurar nombre y correo (primera vez)
git config --global user.name "Tu Nombre"
git config --global user.email "tu-correo@ieee.org"
```

### 4. Agregar Archivos al Staging

```powershell
# Ver archivos a subir
git status

# Agregar todos los archivos (respeta .gitignore)
git add .

# Ver qué se subirá
git status
```

### 5. Crear Commit Inicial

```powershell
git commit -m "Initial commit: Sistema de correos institucionales IEEE

- Gestión de contactos con Excel
- Envío de correos SMTP con plantillas HTML
- Editor de correos personalizados con Quill.js
- Dashboard con estadísticas en tiempo real
- Búsqueda, filtros y paginación
- Diseño responsive institucional IEEE"
```

### 6. Vincular Repositorio Remoto

```powershell
# Reemplaza <TU_USUARIO> con tu nombre de usuario de GitHub
git remote add origin https://github.com/<TU_USUARIO>/ieee-email-management-system.git

# Verificar URL remoto
git remote -v
```

### 7. Subir a GitHub

```powershell
# Renombrar rama a 'main' (estándar actual de GitHub)
git branch -M main

# Push inicial
git push -u origin main
```

### 8. Verificar en GitHub

1. Refresca tu repositorio en GitHub
2. Deberías ver todos los archivos excepto los del `.gitignore`
3. El README.md se renderizará automáticamente

### Flujo de Trabajo Futuro

```powershell
# Después de hacer cambios en el código:

# 1. Ver archivos modificados
git status

# 2. Agregar cambios
git add .

# 3. Crear commit con mensaje descriptivo
git commit -m "Descripción del cambio"

# 4. Subir a GitHub
git push
```

---

## 🤝 Contribución

### Cómo Contribuir

1. **Fork** el repositorio
2. **Crea una rama** para tu funcionalidad:
   ```bash
   git checkout -b feature/nueva-funcionalidad
   ```
3. **Commit** tus cambios:
   ```bash
   git commit -m "Add: Nueva funcionalidad X"
   ```
4. **Push** a tu fork:
   ```bash
   git push origin feature/nueva-funcionalidad
   ```
5. **Abre un Pull Request** en GitHub

### Convenciones de Código

- **PEP 8**: Sigue las convenciones de estilo de Python
- **Comentarios**: Documenta funciones complejas
- **Nombres descriptivos**: Variables y funciones con nombres claros
- **Español**: Código y comentarios en español

### Ideas para Contribuir

- 📱 Modo oscuro (dark mode)
- 📊 Exportar estadísticas a CSV/PDF
- 📅 Programar envíos de correos (scheduling)
- 📧 Plantillas de correo guardadas
- 🔍 Búsqueda avanzada con múltiples criterios
- 📱 Aplicación móvil (Progressive Web App)
- 🔔 Notificaciones push al enviar correos
- 📈 Gráficos de estadísticas con Chart.js
- 🌐 Internacionalización (i18n) inglés/español
- 🔒 Autenticación de usuarios (login)

---

## 👨‍💻 Autor

**Julio Ricardo Barrios Amador**

- 🎓 **Rol**: Section Student Representative (SSR) | IEEE Costa Rica Section
- 🏫 **Institución**: IEEE Computer Society - Instituto Tecnológico de Costa Rica
- 📧 **Email**: julio.barrios@ieee.org
- 🆔 **IEEE Member**: 101781510

---

## 📄 Licencia

Este proyecto fue desarrollado para uso interno del IEEE Computer Society del Instituto Tecnológico de Costa Rica.

**Nota sobre el uso**: Si deseas usar este código para tu organización, siéntete libre de hacerlo. Se agradece mantener la atribución al autor original.

---

## 🙏 Agradecimientos

- **IEEE Computer Society**: Por la oportunidad de desarrollar este sistema
- **Instituto Tecnológico de Costa Rica**: Por el apoyo institucional
- **Comunidad de Flask**: Por la excelente documentación
- **Quill.js**: Por el potente editor WYSIWYG

---

<div align="center">

**⭐ Si este proyecto te fue útil, considera darle una estrella en GitHub ⭐**

Desarrollado con ❤️ para IEEE Computer Society TEC

© 2026 Julio Ricardo Barrios Amador | IEEE Costa Rica Section

</div>
