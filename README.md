# PROGRAMA-RCV

Aplicación local de solo lectura para explorar la hoja RCV de un libro Excel. Los registros se leen en este equipo y se muestran en una interfaz de navegador servida únicamente en 127.0.0.1. El tablero incluye gráficos agregados de sexo, régimen y zonificación; los porcentajes incluyen los registros sin dato.

La aplicación no edita el libro, no guarda copias de sus registros y no realiza solicitudes a servicios externos. Los datos se cargan en memoria mientras el servidor local está encendido.

## Iniciar en Windows

Desde PowerShell, en esta carpeta:

    powershell -ExecutionPolicy Bypass -File .\run_local.ps1

La ruta inicial apunta a `%USERPROFILE%\Downloads\RCVSEP2026 (1).xlsx`. Para usar otro libro:

    powershell -ExecutionPolicy Bypass -File .\run_local.ps1 -WorkbookPath "C:\ruta\al\archivo.xlsx"

El servidor se detiene con Ctrl+C. Abre el navegador en http://127.0.0.1:8765/.

## Requisitos

- Python 3.10 o posterior.
- openpyxl, instalable con python -m pip install -r requirements.txt.

La aplicación usa la hoja RCV, conserva las fórmulas guardadas como valores por Excel y muestra las referencias #REF! como incidencias. No recalcula medidas ni aplica umbrales clínicos. El libro conserva otras hojas, pero la primera versión solo presenta la tabla de registros de RCV.

## Privacidad

El servidor escucha solo en 127.0.0.1, no en la red local. No se registran términos de búsqueda en la consola. Mantén el equipo y la sesión protegidos; cualquier persona con acceso a este equipo y a la dirección local podría consultar los registros mientras el servidor esté activo.
