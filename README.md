<img width="1117" height="626" alt="Captura de pantalla 2025-09-04 222627" src="https://github.com/user-attachments/assets/ffed3000-c4cb-4438-a003-8bbd13e0b81f" />

🧩 Office AIO FULL 1.0
Office AIO FULL es un asistente automatizado en PowerShell que permite descargar e instalar versiones originales de Microsoft Office ProPlus x64 en español, directamente desde los servidores CDN de Microsoft. El sistema incluye soporte para múltiples versiones, integración de Project y Visio, y activación opcional vía MAS.

🚀 Características principales
- Instalación automatizada de:
- Office 365, 2024, 2021, 2019, 2016, 2013
- Project y Visio para cada versión
- Descarga directa desde Microsoft (sin intermediarios ni modificaciones)
- Interfaz visual en consola con menú interactivo
- Validación de integridad del archivo descargado
- Activación opcional vía MAS
- Limpieza automática de archivos temporales al finalizar

📦 Estructura del script
El proyecto está contenido en un único archivo: MenuGIT.ps1, que incluye:
- Función Script-Logo: encabezado visual
- Función MostrarSubmenu: selector por versión
- Función DescargarYEjecutar: descarga y ejecución del instalador
- Función EsperarAccionPostDescarga: control por usuario antes de instalar
- Función Script-End: cierre visual del asistente
- Diccionario $linksDeOffice: enlaces CDN por producto y versión

🧪 Ejecución remota desde PowerShell
Para ejecutar el script directamente desde GitHub vía PowerShell:
```powershell
irm "https://raw.githubusercontent.com/Hellowen6060/Office-AIO-FULL-1.0/refs/heads/main/MenuGIT.ps1" | iex
```

🛡️ Requisitos
- PowerShell 5.1 o superior
- Conexión a internet estable
- Permisos de ejecución (Set-ExecutionPolicy RemoteSigned si es necesario)

📁 Organización del repositorio
Office-AIO-FULL-1.0/
├── MenuGIT.ps1          # Script principal unificado
├── README.md            # Documentación del proyecto

👤 Autor
Diego Garcia
Arquitectura modular, trazabilidad visual y ejecución remota sin errores.
Compilado con precisión y estilo técnico.
