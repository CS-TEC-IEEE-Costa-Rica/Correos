# Sistema de Gestion y Envio de Correos Institucionales - IEEE Computer Society

Aplicacion Flask para gestionar contactos empresariales y enviar correos institucionales del IEEE Computer Society del Instituto Tecnologico de Costa Rica, Campus Central Cartago.

## Requisitos

- Python 3.8 o superior
- Cuenta de correo con acceso SMTP (recomendado: contrasena de aplicacion)

## Instalacion

1. Crear entorno virtual:

```powershell
python -m venv .venv
```

2. Activar entorno virtual:

```powershell
& ".venv\Scripts\Activate.ps1"
```

3. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

## Configuracion SMTP

En `app.py`, actualiza estas variables:

```python
EMAIL_REMITENTE = "tu-correo@ieee.org"
EMAIL_PASSWORD = "tu-contrasena-de-aplicacion"
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587
```

Nota:
- No subas credenciales reales al repositorio.
- Si usas Gmail, genera una contrasena de aplicacion para SMTP.

## Ejecucion

Con el entorno virtual activado:

```powershell
python app.py
```

Abrir en navegador:

- `http://127.0.0.1:5000`
- `http://localhost:5000`

## Estructura del proyecto

```text
Programa contactos/
|- app.py
|- requirements.txt
|- README.md
|- .gitignore
|- images/
|- static/
|  |- style.css
|- templates/
   |- index.html
```

## Funcionalidades principales

- Gestion de contactos (agregar, eliminar, listar)
- Filtros por texto y estado
- Ordenamiento por empresa, estado y fecha de envio
- Paginacion de 50 contactos por pagina
- Envio individual de correo HTML con logos IEEE inline
- Registro de estado de envio y fecha

## Datos locales

El sistema usa archivos locales para datos de trabajo:

- `contactos.xlsx`
- `database.db`

Estos archivos estan ignorados por `.gitignore` para evitar subir informacion sensible o temporal.

## Publicar en GitHub

1. Inicializar repositorio (si aun no existe):

```powershell
git init
```

2. Agregar archivos:

```powershell
git add .
```

3. Crear commit inicial:

```powershell
git commit -m "Initial commit"
```

4. Vincular repositorio remoto y subir:

```powershell
git remote add origin <URL_DEL_REPO>
git branch -M main
git push -u origin main
```

## Autor

Julio Ricardo Barrios Amador
