$host.UI.RawUI.WindowTitle = "Office AIO FULL 1.0"

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

function Script-End {
    Write-Host "                             " -NoNewline; Write-Host "      ██████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "    ██████████████" -ForegroundColor DarkYellow
    Write-Host "    Compilado realizado por: " -NoNewline; Write-Host "   █████████████████" -ForegroundColor DarkYellow
    Write-Host "                             " -NoNewline; Write-Host "  ███████████" -ForegroundColor DarkYellow
    Write-Host "    " -NoNewline
    Write-Host "Diego Garcia      " -ForegroundColor DarkRed -NoNewline
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

$linksDeOffice = @{
    "Office365" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=O365ProPlusRetail&platform=x64&language=es-es&version=O16GA"
    "Project365" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectProRetail&platform=x64&language=es-es&version=O16GA"
    "Visio365"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioProRetail&platform=x64&language=es-es&version=O16GA"
    "Office24" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2024Retail&platform=x64&language=es-es&version=O16GA"
    "Project24" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectPro2024Retail&platform=x64&language=es-es&version=O16GA"
    "Visio24"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioPro2024Retail&platform=x64&language=es-es&version=O16GA"
    "Office21" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=es-es&version=O16GA"
    "Project21" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectPro2021Retail&platform=x64&language=es-es&version=O16GA"
    "Visio21"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioPro2021Retail&platform=x64&language=es-es&version=O16GA"
    "Office19" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2019Retail&platform=x64&language=es-es&version=O16GA"
    "Project19" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectPro2019Retail&platform=x64&language=es-es&version=O16GA"
    "Visio19"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioPro2019Retail&platform=x64&language=es-es&version=O16GA"
    "Office16" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=es-es&version=O16GA"
    "Project16" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectProRetail&platform=x64&language=es-es&version=O16GA"
    "Visio16"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioProRetail&platform=x64&language=es-es&version=O16GA"
    "Office13" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlusRetail&platform=x64&language=es-es&version=O15GA"
    "Project13" = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProjectProRetail&platform=x64&language=es-es&version=O15GA"
    "Visio13"  = "https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=VisioProRetail&platform=x64&language=es-es&version=O15GA"
    "MAS" = "COMMAND:Invoke-Expression (Invoke-RestMethod -Uri 'https://get.activated.win')"
}

$inf1 = "Diego Garcia"
$fechaActual = Get-Date -Format "dd/MM/yyyy"
$usuarioActual = $env:USERNAME

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
            } catch {
                Write-Host "❌ Error al capturar tecla: $_" -ForegroundColor Red
                break
            }

            switch ($keyCode) {
                13 {
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
                    }

                    $instalacionCompletada = $true
                }
                27 {
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
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Office" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
                if ($regresar) { $salirSubmenu = $true }
            }
            "2" {
                $key = "Project$versionCorta"
                if (-not $linksDeOffice.ContainsKey($key)) {
                    Write-Host "❌ Clave '$key' no encontrada en linksDeOffice" -ForegroundColor Red
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Project" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
                if ($regresar) { $salirSubmenu = $true }
            }
            "3" {
                $key = "Visio$versionCorta"
                if (-not $linksDeOffice.ContainsKey($key)) {
                    Write-Host "❌ Clave '$key' no encontrada en linksDeOffice" -ForegroundColor Red
                    return
                }
                $regresar = DescargarYEjecutar -nombre "Visio" -url $linksDeOffice[$key] -versionTexto $versionTexto -color $color
                if ($regresar) { $salirSubmenu = $true }
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

do {
    Clear-Host
    Script-Logo
    Write-Host "Bienvenido $usuarioActual a Office AIO FULL v.1, actualización $fechaActual"
    Write-Host ""
    Write-Host "Este asistente te ayudará a descargar e instalar el paquete en español de"
    Write-Host "Microsoft Office ProPlus x64 de tu preferencia automáticamente."
    Write-Host ""
    Write-Host "Todo el proceso de descarga e instalación se realiza directamente desde"
    Write-Host "los servidores CDN de Microsoft, por lo cual, todo es original y sin modificaciones."
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