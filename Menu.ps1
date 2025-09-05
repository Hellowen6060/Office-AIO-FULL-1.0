# 🧩 Configuración visual inicial
$host.UI.RawUI.WindowTitle = "Office AIO FULL 1.0"
Script-Logo
# Comentario delimitador: Carga de módulo de enlaces
$linksPath = Join-Path $PSScriptRoot "Links.ps1"
if (-not (Test-Path $linksPath)) {
    Write-Host "❌ No se encontró el archivo Links.ps1 en $PSScriptRoot" -ForegroundColor Red
    Pause
    exit
}
. $linksPath

if (-not $linksDeOffice -or $linksDeOffice.Count -eq 0) {
    Write-Host "❌ La variable 'linksDeOffice' no está definida o está vacía." -ForegroundColor Red
    Pause
    exit
}

# VARIABLES
$inf1         = "Diego Garcia"
$fechaActual  = Get-Date -Format "dd/MM/yyyy"
$usuarioActual = $env:USERNAME

# FUNCIÓN: LOGO
function Script-Logo {
    Write-Host ""
    Write-Host "===================================================================================" -ForegroundColor White
    Write-Host " ███╗  ███╗    ██████╗             ██████╗   █████╗  ███╗   ███╗ ███████╗ ██████╗" -ForegroundColor DarkRed
    Write-Host " ══██╗ ══██╗   ██╔══██╗           ██╔════╝  ██╔══██╗ ████╗ ████║ ██╔════╝ ██╔══██╗" -ForegroundColor DarkRed
    Write-Host "    ██╗   ██╗  ██║  ██║           ██║  ███╗ ███████║ ██╔████╔██║ █████╗   ██████╔╝" -ForegroundColor DarkRed
    Write-Host "   ██╔╝  ██╔╝  ██╚══██║           ██║   ██║ ██╔══██║ ██║╚██╔╝██║ ██╔══╝   ██╔══██╗" -ForegroundColor DarkRed
    Write-Host " ███╔╝ ███╔╝   ██████╔╝  ███╗     ╚██████╔╝ ██║  ██║ ██║ ╚═╝ ██║ ███████╗ ██║  ██║" -ForegroundColor DarkRed
    Write-Host " ═══╝  ═══╝    ╚═════╝   ╚══╝      ╚═════╝  ╚═╝  ╚═╝ ╚═╝     ╚═╝ ╚══════╝ ╚═╝  ╚═╝" -ForegroundColor DarkRed
    Write-Host "===================================================================================" -ForegroundColor White
    Write-Host ""
}

# FUNCIÓN: FIN VISUAL
function Script-End {
    Write-Host "                             " -NoNewline; Write-Host "      ██████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "    ██████████████" -ForegroundColor DarkYellow
    Write-Host "    Compilado realizado por: " -NoNewline; Write-Host "   █████████████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "  ███████████" -ForegroundColor DarkYellow
    Write-Host "    " -NoNewline
    Write-Host "$inf1      " -ForegroundColor DarkRed -NoNewline
    Write-Host "       " -NoNewline; Write-Host "  █████████" -ForegroundColor DarkYellow -NoNewline
    Write-Host "         " -NoNewline; Write-Host "██" -ForegroundColor DarkBlue -NoNewline
    Write-Host "   " -NoNewline; Write-Host "██" -ForegroundColor Green -NoNewline
    Write-Host "   " -NoNewline; Write-Host "██" -ForegroundColor DarkRed
    Write-Host "                             " -NoNewline; Write-Host "  ███████████" -ForegroundColor DarkYellow
    Write-Host "" -NoNewline; Write-Host "    D.GAMER" -ForegroundColor DarkRed -NoNewline
    Write-Host "                  " -NoNewline; Write-Host "   █████████████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "    ██████████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "      ██████████" -ForegroundColor DarkYellow
    Write-Host ""
}

# 🧩 Función principal: MostrarSubmenu
function MostrarSubmenu {
    param (
        [string]$versionTexto,
        [string]$versionCorta,
        [string]$color
    )

    $tempFolder = "$env:TEMP\officeAIO"
    if (-not (Test-Path $tempFolder)) {
        New-Item -Path $tempFolder -ItemType Directory | Out-Null
    }

    # 🧩 Función: EsperarAccionPostDescarga (corregida)
    function EsperarAccionPostDescarga {
    param (
        [string]$nombre,
        [string]$versionTexto,
        [string]$destino
    )

    Clear-Host
    Script-Logo

    Write-Host "`n✅ Instalador Microsoft $nombre $versionTexto descargado con éxito.`n" -ForegroundColor Green
    Write-Host "Presiona ENTER para comenzar con la instalación o ESC para regresar al menú principal." -ForegroundColor Yellow

    $instalacionCompletada = $false
    $regresarAlMenu = $false

    while (-not $instalacionCompletada) {
        try {
            $keyInfo = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
            $keyCode = $keyInfo.VirtualKeyCode
            $char    = $keyInfo.Character
            # "`n🔎 Tecla capturada: Código = $keyCode | Carácter = '$char'" -ForegroundColor Cyan
        } catch {
            Write-Host "❌ Error al capturar tecla: $_" -ForegroundColor Red
            break
        }

        switch ($keyCode) {
            13 {  # ENTER
                if (-not (Test-Path $destino)) {
                    Write-Host "❌ El archivo no existe: $destino" -ForegroundColor Red
                    break
                }

                if ((Get-Item $destino).Length -eq 0) {
                    Write-Host "❌ El archivo está vacío. No se puede ejecutar." -ForegroundColor Red
                    break
                }

                try {
                    Write-Host "`n🚀 Ejecutando instalador: `"$destino`"" -ForegroundColor Green
                    Start-Process -FilePath "`"$destino`"" -Wait -ErrorAction Stop
                    Write-Host "`n✅ Instalación finalizada." -ForegroundColor Green
                } catch {
                    Write-Host "❌ Error al ejecutar el instalador: $_" -ForegroundColor Red
                    Pause
                }

                $instalacionCompletada = $true
            }
            27 {  # ESC
                Write-Host "`n↩ Regresando al menú principal..." -ForegroundColor DarkGray
                $regresarAlMenu = $true
                $instalacionCompletada = $true
            }
            default {
                Write-Host "`nTecla no válida. Presiona ENTER o ESC." -ForegroundColor DarkRed
            }
        }
    }

    Start-Sleep -Seconds 1
    Clear-Host
    Script-Logo

    return $regresarAlMenu
}
    # 🧩 Fin función EsperarAccionPostDescarga
    

    
# 🧩 Función: DescargarYEjecutar (versión final sin pausas)
function DescargarYEjecutar {
    param (
        [string]$nombre,
        [string]$url,
        [string]$versionTexto,
        [string]$color
    )

    Clear-Host
    Script-Logo
    
    $tempFolder = "$env:TEMP\officeAIO"
    if (-not (Test-Path $tempFolder)) {
        New-Item -Path $tempFolder -ItemType Directory | Out-Null
        Write-Host "`n🧩 Paso 1: Carpeta temporal creada con éxito" -ForegroundColor Cyan
    } else {
        Write-Host "`n🧩 Paso 1: Carpeta temporal existente" -ForegroundColor Cyan
    }

    $archivo = "$nombre$versionTexto.exe"
    $destino = Join-Path $tempFolder $archivo
    
    Write-Host "`nDescargando $nombre $versionTexto..." -ForegroundColor $color

    try {
        Invoke-WebRequest -Uri $url -OutFile $destino -ErrorAction Stop
        Write-Host "`n🧩 Paso 2: Descarga completada desde $url" -ForegroundColor Cyan
    } catch {
        Write-Host "❌ Error al descargar ${nombre} ${versionTexto}: $_" -ForegroundColor Red
        Start-Sleep -Seconds 3
        return
    }

    # 🔍 Verificación extendida del archivo descargado
    if (Test-Path $destino) {
        $tamano = (Get-Item $destino).Length
        Write-Host "`n📦 Archivo detectado: $destino" -ForegroundColor DarkGreen
        Write-Host "📏 Tamaño del archivo: $tamano bytes" -ForegroundColor DarkGreen

        if ($tamano -gt 0) {
            $regresar = EsperarAccionPostDescarga -nombre $nombre -versionTexto $versionTexto -destino $destino
            return $regresar
        } else {
            Write-Host "❌ El archivo está vacío. No se puede ejecutar." -ForegroundColor Red
            return $true
        }
    } else {
        Write-Host "❌ El archivo no fue encontrado en la ruta esperada." -ForegroundColor Red
        return $true
    }
}
# 🧩 Fin función DescargarYEjecutar
    
    $salirSubmenu = $false
    while (-not $salirSubmenu) {
        Clear-Host
        Script-Logo
        Write-Host "Selecciona el paquete Microsoft Office $versionTexto ProPlus a instalar:" -ForegroundColor White
        Write-Host ""

        Write-Host "1." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office $versionTexto ProPlus" -ForegroundColor $color
        Write-Host "2." -ForegroundColor White -NoNewline; Write-Host " Microsoft Project $versionTexto Pro" -ForegroundColor $color
        Write-Host "3." -ForegroundColor White -NoNewline; Write-Host " Microsoft Visio $versionTexto Pro" -ForegroundColor $color
        Write-Host "4." -ForegroundColor White -NoNewline; Write-Host " Regresar" -ForegroundColor Red

        $subopcion = Read-Host "`nSelecciona una opción del submenú (1-4):"

        switch ($subopcion) {
            "1" {
                $key = "Office$versionCorta"
                if (-not $linksDeOffice.ContainsKey($key)) {
                    Write-Host "❌ Clave '$key' no encontrada en linksDeOffice" -ForegroundColor Red
                    Pause
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Office" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
if ($regresar) {
    $salirSubmenu = $true  # Solo salir si el usuario lo pidió
}
            }
            "2" {
                $key = "Project$versionCorta"
                if (-not $linksDeOffice.ContainsKey($key)) {
                    Write-Host "❌ Clave '$key' no encontrada en linksDeOffice" -ForegroundColor Red
                    Pause
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Office" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
if ($regresar) {
    $salirSubmenu = $true  # Solo salir si el usuario lo pidió
}
            }
            "3" {
                $key = "Visio$versionCorta"
                if (-not $linksDeOffice.ContainsKey($key)) {
                    Write-Host "❌ Clave '$key' no encontrada en linksDeOffice" -ForegroundColor Red
                    Pause
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Office" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
if ($regresar) {
    $salirSubmenu = $true  # Solo salir si el usuario lo pidió
}
            }
            "4" {
                $salirSubmenu = $true
            }
            Default {
                Write-Host "`nOpción inválida. Intenta nuevamente." -ForegroundColor DarkRed
                Start-Sleep -Seconds 2
            }
        }
    }
}
# 🧩 Fin función MostrarSubmenu

# 🧩 Bucle del menú principal
do {
    Clear-Host
    Script-Logo
    Write-Host "Bienvenido $usuarioActual a Office AIO FULL v.1, actualización $fechaActual"
	Write-Host ""
	Write-Host "Este asistente te ayudará a descargar e instalar el paquete en español de"
	Write-Host "Microsoft Office ProPlus x64 de tu preferencia automáticamente."
    Write-Host ""
	Write-Host "Todo el proceso de descarga e instalación se realiza directamente desde"
	Write-host "los servidore CDN de Microsoft, por lo cual, todo es original y sin modificaciones."
	Write-Host ""
    Write-Host "Para iniciar, selecciona una opción y espera que el proceso de instalación termine:"
    Write-Host ""
    Write-Host "1." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 365 ProPlus" -ForegroundColor Cyan
    Write-Host "2." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 2024 ProPlus" -ForegroundColor Magenta
    Write-Host "3." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 2021 ProPlus" -ForegroundColor Blue
    Write-Host "4." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 2019 ProPlus" -ForegroundColor Green
    Write-Host "5." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 2016 ProPlus" -ForegroundColor Yellow
    Write-Host "6." -ForegroundColor White -NoNewline; Write-Host " Microsoft Office 2013 ProPlus" -ForegroundColor DarkYellow
    Write-Host "7." -ForegroundColor White -BackgroundColor DarkGreen -NoNewline; Write-Host " Microsoft Activation Scripts (MAS)" -ForegroundColor White -BackgroundColor DarkGreen
    Write-Host "8." -ForegroundColor White -NoNewline; Write-Host " Finalizar" -ForegroundColor DarkRed

    $opcion = Read-Host "`nSelecciona una opción disponible y presiona ENTER (1-8)"

        switch ($opcion) {
        "1" { MostrarSubmenu -versionTexto "365" -versionCorta "365" -color "Cyan" }
        "2" { MostrarSubmenu -versionTexto "2024" -versionCorta "24" -color "Magenta" }
        "3" { MostrarSubmenu -versionTexto "2021" -versionCorta "21" -color "Blue" }
        "4" { MostrarSubmenu -versionTexto "2019" -versionCorta "19" -color "Green" }
        "5" { MostrarSubmenu -versionTexto "2016" -versionCorta "16" -color "Yellow" }
        "6" { MostrarSubmenu -versionTexto "2013" -versionCorta "13" -color "DarkYellow" }
        "7" {
            $valor = $linksDeOffice["MAS"]
            if ($valor -like "COMMAND:*") {
                $comando = $valor -replace "^COMMAND:", ""
                Invoke-Expression $comando
            } else {
                Write-Host "`nDescargando MAS..." -ForegroundColor White -BackgroundColor DarkGreen
                Invoke-WebRequest -Uri $valor -OutFile "$env:TEMP\MAS.zip"
            }
        }
        "8" {
    # 🧼 Limpieza visual y cierre del sistema
    Clear-Host
    Script-Logo
    Script-End

    Write-Host "`n🧹 Estado de limpieza: Eliminando carpeta temporal..." -ForegroundColor Cyan
    $carpetaTemporal = "$env:TEMP\officeAIO"

    if (Test-Path $carpetaTemporal) {
        try {
            Remove-Item -Path $carpetaTemporal -Recurse -Force
            Write-Host "✅ Carpeta temporal eliminada: $carpetaTemporal" -ForegroundColor Green
        } catch {
            Write-Host "❌ Error al eliminar carpeta temporal: $_" -ForegroundColor Red
        }
    } else {
        Write-Host "ℹ️ No se encontró carpeta temporal para eliminar." -ForegroundColor DarkYellow
    }

    Start-Sleep -Seconds 5
    exit
}
        Default {
            Write-Host "`nOpción inválida. Intenta nuevamente." -ForegroundColor DarkRed
            Start-Sleep -Seconds 2
        }
    }
} while ($true)
# 🧩 Fin del bucle del menú principal
