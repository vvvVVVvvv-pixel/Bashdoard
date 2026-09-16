<#
.SYNOPSIS
    Automates N1 customer data extraction, token management, and ticket formatting for CRM workflows.
.DESCRIPTION
    Queries backend APIs, parses technical fiber/ONT parameters, and generates a standardized markdown ticket.
#>

Add-Type -AssemblyName System.Windows.Forms
$wshell = New-Object -ComObject WScript.Shell
# Session configuration and authentication cookies (Tokens expire ~every 6 hours)
class SessionConfig {
    [string]$PathWork = 'path'
    [hashtable]$Cookies = @{
        'incapsession' = ''
        'incapkey'     = ''
        'dashsyskey'   = ''
        'jsessionkey'  = ''
    }
}
# new instance of active cookies
$config = [SessionConfig]::new()
# client data template
class DatosCliente {
    [string]$acc
    [string]$suc
    [string]$nomclient
    [string]$falla
    [string]$plan
    [string]$dir
    [string]$desc
    [string]$con
    [string]$connum
    [string]$conmail
    [string]$convis
    [string]$conseg
    [string]$consegnum
    [string]$consegmail
    [string]$conseghor
    [string]$perm
    [string]$entrega
    [hashtable]$dattec = @{}
}

$key = [DatosCliente]::new()

# DEFINE FUNCTIONS
# function to extend the session, it just opens a new instance of all the platforms used in the workflow, prints the date to aid keeping track of expiring sessions
function Reload {
  Start-Process -FilePath "chrome.exe" -ArgumentList "--new-tab", "(removed urls to preserve privacy)"
  Start-Process -FilePath "chrome.exe" -ArgumentList "--new-tab", "(removed urls to preserve privacy)"
  Start-Process -FilePath "chrome.exe" -ArgumentList "--new-tab", "(removed urls to preserve privacy)"
  date
}
# change the terminal text, for aesthetic purposes, and to know bashdoard is ready to use
function prompt {
    "Bashdoard>>  "
}
# reset
function Finish {
    $script:key = [DatosCliente]::new()
    Write-Host "`n `n Gracias por usar Bashdoard! `n `n"
    $host.UI.RawUI.ForegroundColor = 'White'
}

function Ascii-Art {
$ascii = @"
 __    _    __  _   _ __   _    _    __  __
| _ )  / \  / __|| | | |  _ \ / _ \  / \  |  _ \|  _ \
|  _ \ / _ \ \_ \| || | | | | | | |/ _ \ | |) | | | |
| |) / __ \ _) |  _  | || | || / _ \|  _ <| |_| |
|_//   \\_/|| ||_/ \_//   \\| \\_/ v6.0
"@

$art = @"
`t▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒░░▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒░▒▓▓▓
`t▓▓▓▓▓▓▓▓▓▓▓▓▒▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒▓░▒▓
`t▓▓▓▓▓▓▓▓▓▓░▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒░
`t▒▒▒▒▒▒▒▒░░▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▒
`t░░░░░░░░▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓▓▒▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓▓▒▓▓▒▓▒▓▓▓▓▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓▒▒▓▓░▓▓░▒▓▓▓▒▒▓▓▓▓▓▓▓▓▒▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓░▒▒▓░░▓░▓░▒▓▓▓░▒▓▓▓▓▓▓▒▒▒▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓░▒░▒▓▒░▒▒▒▓▓▒░░▒▒░▒▒▓▓▓▓▒░░▒▒▓▓▓▓▓▓▓
`t▓▓▓▒▒▒░▒▒▓▒░▒░░▒▓▓▓▓▓▓▓▒▓▓▓▓▒▓▒▒▓▒▒▓▓▓▓▓
`t▓▓▓▓▓▓▓░▒░▓▒▒▓▓▓▓▓▓▓▓▓▓░▓▓▓▒░░░░▒▓░▓▓▓▓▓
`t▓▓▓▓▓▓▓▓░▒░░░░░░▒▓▓▓▓▓▓░▓▓▓░░▒▒▒▒▓░▓▓▓▓▓
`t▓▓▓▓▓▓▓▓▓▓▒░▓░░▓▓▓▓▓▓▓▓░▒▓▓░░░░▒▓▒░▓▓▓▓▓
`t▓▓▓▓▓▓▓▓▓▓▓▓▒▓░▒░▓▓▓▓▓▓░▓░▓░▒▒▓▓░▒▓▓▓▓▒▒
`t▓▓▓▓▓▓▓▓▓▓▓░▓▓▓▓▓▓▓▓▓▓░▒░▒▓░▓▓▓▒░▓▓▓▓▒░▓
`t▓▓▓▓▓▓▓▓▓▒░▓▓▓▓▓▓▓▓▓▓▓▓▒▒▓░▓▓▒░▒▓▓▓▓▓░░▓
`t▓▓▓▓▓▓▓▓▓▒░░▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓░▒▓▒▓▓▒░░▒
`t▓▓▓▓▓▓▓▓▓▓▓▓░▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓░▒▓▒▓▓░░░▒
`t▓▓▓▓▓▓▓▓▓▓▓▓▓░░▓▓▓▓▓▓▓▓▓▒░▓▓▒░▒░▓▒░▓░░▓▒
`t▓▓▓▓▓▓▓▓▓▓▓░▒▓░▒▒▓▓▓▓▓▒░░▓▓▒▒▒░░▓▒▒▒░░▒▒
`t▓▓▓▓▓▓▓▓▓▒░░░▓▓░▓░░▒░░▒▒▒░▒░░▓▓▓░▒░░░░░▒
`t▒░░░░░▒▒░▒▓▒░▒▓░▒▒▒▒▒▒▒░░▓▓▓▓▓░░░░▒░▓▓▒░
`t▒░░░░░░░▒░▓▓░▓▓▒░▒▒▒▒░░░▒░▓▓░░▓▓▓▓▓▓▒▒▓░
`t▓▓▓▓▓░░▓▓▓▒▓▓▓▓▒░▓▓▓▓▒░▒░░▓░▒▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▒▒▓▒▓▓▓▓▓▓▓▒░▒▓▓▓▓▒░▒▒░▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓░▒▒▓▓▓▓▓▓▓▓▒▒▒▓▓▓░░▒░▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓░▒▓▓▓▓▓▓▓▓▒░▓▓░░▒▓░░▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t▓▓▓▓▓░▒▓▓▓▓▓▓▓▓░▒▒░▒▓▓▒░▒▓▓▓▓▓▓▓▓▓▓▓▓▓▓▓
`t `t
"@

    $intro = "Herramienta de extracción de datos para automatización de plantillas"
    $host.UI.RawUI.ForegroundColor = 'Green'
    Write-Host $ascii `n
    $host.UI.RawUI.ForegroundColor = 'white'
    $host.UI.RawUI.backgroundColor = 'black'
    Write-Host $art `n
    $host.UI.RawUI.ForegroundColor = 'green'
    $host.UI.RawUI.backgroundColor = 'black'
    Write-Host $intro
}
# the cookies expire every 6~ hours, and must be reinjected from the application tab in the dev console, due to security concerns
function Refresh-Cookies {
    param([SessionConfig]$Session)

    $config.Cookies['incapsession'] = Read-Host -Prompt "Enter incap_ses_"
    $config.Cookies['incapkey']     = Read-Host -Prompt "Enter incapkey"
    $config.Cookies['dashsyskey']   = Read-Host -Prompt "Enter dashSysInfo"
    $config.Cookies['jsessionkey']  = Read-Host -Prompt "Enter JSESSIONID"
}

function Input-Account {
    Read-Host -Prompt "cuenta: "
}

function Load-Url {
    foreach ($acc in $key.acc) {
        $url = "(removed urls to preserve privacy)"
        Start-Process chrome.exe $url
    }
}
# 
function Get-DashData {
    try {
        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        $session.Cookies.Add((New-Object System.Net.Cookie("$($config.cookies.incapsession)", "$($config.cookies.incapkey)", "/", ".(removed urls to preserve privacy)")))
        $session.Cookies.Add((New-Object System.Net.Cookie("dashSysInfo", "$($config.cookies.dashsyskey)", "/", "(removed urls to preserve privacy)")))
        $session.Cookies.Add((New-Object System.Net.Cookie("JSESSIONID", "$($config.cookies.jsessionkey)", "/", "(removed urls to preserve privacy)")))
        $dashrawdata = Invoke-WebRequest -UseBasicParsing -Uri "(removed urls to preserve privacy)" `
        -Method "POST" `
        -WebSession $session `
        -Headers @{
        } `
        -ContentType "application/json; charset=UTF-8" `
        -Body "{"cuenta":"$($key.acc)"}"
        return $dashrawdata
    }
    catch {
        Write-Warning "Fracaso al intentar conectar al API DASHBOARD en la cuenta $($key.acc). Error: $_"
        return $null
    }
}

function Clean-DashData {
    $clientdata = Get-DashData
    if ($null -eq $clientdata) { return }
    $jsononly = $clientdata.rawcontent.substring($clientdata.rawcontent.indexof('{'))
    $outer = $jsonOnly | ConvertFrom-Json
    $inner = $outer.result   | ConvertFrom-Json
    $key.plan = $inner.lcr.CuentaPadre.plan
    $key.suc = $inner.lcr.CuentaPadre.nombreSitio
    $key.dir = $inner.lcr.CuentaPadre.direccion_instalacion
    $key.entrega = $inner.lcr.CuentaPadre.medioAcceso
    $key.nomclient = $inner.lcr.CuentaPadre.nombre
    return $key
}

function Scan-Medium {
    $stringToFind = "Sólo fibra"
    if ($key.entrega -match $stringToFind) {
        Detected-Fiber
        } else {
        Detected-Etc
    }
}

function Get-TPUXData {
    try {
        $url = "(removed urls to preserve privacy)"
        $technicaldata = curl.exe $url `
        -b "idSession=vavav34qdfafegth34"
        $outer = $technicaldata | convertfrom-json
        $key.dattec["SN"] = "$($outer.s.SN)"
        $key.dattec["IPOLT"] = "$($outer.s.IPOLT)"
        $key.dattec["OLT"] = "$($outer.s.OLT)"
        $key.dattec["IDOLT"] = "$($outer.s.IDOLT)"
        $key.dattec["MODELO ONT"] = "$($outer.s.strMODEL)"
        $key.dattec["F/S/P/ONT ID"] = "$($outer.s.FRAME)/$($outer.s.SLOT)/$($outer.s.PORT)/$($outer.s.ONT_ID)"
        $key.dattec["IP ONT"] = "$($outer.s.strIpONT)"
        $key.dattec["ETIQUETA"] = "$($outer.s.ETIQUETA)"
        $host.UI.RawUI.ForegroundColor = 'white'
        $key.dattec | Format-Table -AutoSize
        $host.UI.RawUI.ForegroundColor = 'green'
        Start-Sleep -Milliseconds 50
        $proc = Get-Process -Name MobaXterm -ErrorAction Stop
        $wshell.AppActivate($proc.Id)
       }
    catch {
        Write-Warning "Fracaso al intentar extraer datos técnicos TPUX en la cuenta $($key.acc). Error: $_"
    }
}

function Detected-Fiber {
    $host.UI.RawUI.ForegroundColor = 'cyan'
    Write-Host "Medio 'Sólo Fibra' detectado; se intenta anexar datos técnicos de UX, si este proceso falla revisar aprovisionamiento"
    $host.UI.RawUI.ForegroundColor = 'green'
    $key.entrega = "MEDIO DE ENTREGA: SOLO FIBRA"
    Get-TPUXData
    # fun formatting for autosniffer in paramiko
    $olt_ip = $key.dattec.IPOLT
    $ont_sn = $key.dattec.SN
    $fsp    = $key.dattec.'F/S/P/ONT ID'

    $parts  = $fsp -split '/'
    $fs     = "$($parts[0])/$($parts[1])"
    $pi     = "$($parts[2]) $($parts[3])"
    set-clipboard "$olt_ip,$ont_sn,$fs,$pi"
}

function Detected-Etc {
    $key.entrega = "SM, MICROONDAS"
    $wshell.AppActivate("POWERSHELL")
    $host.UI.RawUI.ForegroundColor = 'Magenta'
    Write-Host $key.entrega
    $key.dattec = @{ Status = "Validar datos tecnicos" }
    $host.UI.RawUI.ForegroundColor = 'Green'
}
# client data required for coordinating operations, for easy collection during contact
function Probe {
    $key.conseg = read-host -prompt "¿Quien es el contacto de seguimiento?"
    $key.consegnum = read-host -prompt "¿ Numero telefónico ?"
    $key.consegmail = read-host -prompt "¿ email ?"
    $key.conseghor = read-host -prompt "¿ horario ?"
    $key.con = read-host -prompt "¿ contacto en sitio ?"
    $key.connum = read-host -prompt "¿ numero ?"
    $key.convis = read-host -prompt "¿ horario de asistencia ?"
    $key.perm = read-host -prompt "¿ requie  re de permisos especiales para visita en sitio (DC3, SUA, LISTA DE IDC POR MAIL?"
    $key.falla = read-host -prompt "¿Cuál problema presenta el cliente? (este es un buen momento para verificar UX y descartar FALLA MASIVA)"
    $key.desc = read-host -prompt "Breve descripcion del problema"
}

function Create-Template {
    # Convert the hashtable to a formatted string table
    $dattecFormatted = ($key.dattec | Format-Table -AutoSize | Out-String -Width 4096).Trim()

    $report = @"
$($key.acc) // $($key.suc) // SD // SEGMENTO I // $($key.nomclient) // $($key.falla) // FOLIODASHBOARD //

Cuenta: $($key.acc)
Cliente: $($key.nomclient)

Segmento: I

Sucursal: $($key.suc)
Dirección: $($key.dir)

Falla: $($key.falla)
Descripción: $($key.desc)

PERMISOS ESPECIALES: $($key.perm)

Datos de Contacto:
Persona que reporta: $($key.conseg)
Teléfono: $($key.consegnum)
Correo: $($key.consegmail)
Horario de llamada: $($key.conseghor)

Contacto en sitio: $($key.con)
Teléfono: $($key.connum)
Horario de visita: $($key.convis)

==============
Datos Técnicos:
PLAN: $($key.plan)
$($key.entrega)

$dattecFormatted

╠═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬═╬
"@

    return $report
}
#main program

function Run {
    $Host.UI.RawUI.WindowTitle = "Bashdoard"
    Ascii-Art
    $key.acc = Input-Account
    Load-Url
    Clean-DashData
    Scan-Medium
    Probe
    $plantilla = Create-Template
    set-clipboard $plantilla
    Add-Content -path $config.pathwork -value $plantilla
    Write-Host "`n `n Copiado al portapapeles.`n `n"
    Finish
    $Host.UI.RawUI.WindowTitle = ""
}
