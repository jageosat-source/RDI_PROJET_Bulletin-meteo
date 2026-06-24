# test_smtp_relay.ps1 - Test controle d'un relais SMTP
#
# Par defaut, ce script teste uniquement la connectivite TCP.
# L'envoi reel exige le parametre -Send afin d'eviter tout mail accidentel.
#
# Exemple sans envoi :
# & ".\test_smtp_relay.ps1" -SmtpServer "smtp.interne.local" -To "prenom.nom@geo-sat.com"
#
# Exemple avec envoi :
# & ".\test_smtp_relay.ps1" -SmtpServer "smtp.interne.local" -Port 25 -From "meteo-bot@geo-sat.com" -To "prenom.nom@geo-sat.com" -Send

param(
    [Parameter(Mandatory = $true)]
    [string]$SmtpServer,

    [int]$Port = 25,

    [string]$From = "meteo-bot@geo-sat.com",

    [Parameter(Mandatory = $true)]
    [string]$To,

    [switch]$UseSsl,

    [switch]$Send,

    [string]$PdfPath = "",

    [int]$TimeoutSeconds = 30
)

$ErrorActionPreference = "Stop"

$versionRoot = Split-Path $PSScriptRoot -Parent
$logDir = Join-Path $versionRoot "Archives\Logs\smtp_tests\$(Get-Date -Format 'yyyyMMdd')"
New-Item -ItemType Directory -Force -Path $logDir | Out-Null
$logStamp = Get-Date -Format 'yyyyMMdd-HHmmss-fff'
$logFile = Join-Path $logDir "smtp_test_${logStamp}_pid$PID.log"

function Log {
    param([string]$Message)

    $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')  $Message"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Test-TcpPort {
    param(
        [string]$HostName,
        [int]$HostPort,
        [int]$TimeoutMs
    )

    $client = New-Object System.Net.Sockets.TcpClient
    try {
        $async = $client.BeginConnect($HostName, $HostPort, $null, $null)
        if (-not $async.AsyncWaitHandle.WaitOne($TimeoutMs, $false)) {
            throw "Timeout TCP apres $TimeoutMs ms"
        }
        $client.EndConnect($async)
    }
    finally {
        $client.Close()
    }
}

Log "=== Test relais SMTP ==="
Log "Machine : $env:COMPUTERNAME"
Log "Utilisateur : $env:USERDOMAIN\$env:USERNAME"
Log "Serveur : $SmtpServer"
Log "Port : $Port"
Log "SSL : $([bool]$UseSsl)"
Log "From : $From"
Log "To : $To"
Log "Mode envoi reel : $([bool]$Send)"

try {
    Test-TcpPort -HostName $SmtpServer -HostPort $Port -TimeoutMs ($TimeoutSeconds * 1000)
    Log "Connectivite TCP : OK"
}
catch {
    Log "Connectivite TCP : ERREUR - $_"
    exit 1
}

if (-not $Send) {
    Log "Aucun mail envoye. Ajouter -Send pour tester un envoi reel."
    Log "Log : $logFile"
    exit 0
}

if ($PdfPath) {
    if (-not [System.IO.Path]::IsPathRooted($PdfPath)) {
        $bulletinDir = Join-Path $versionRoot "Bulletins"
        $PdfPath = Join-Path $bulletinDir $PdfPath
    }
    if (-not (Test-Path -LiteralPath $PdfPath)) {
        Log "Piece jointe introuvable : $PdfPath"
        exit 1
    }
}

$mail = New-Object System.Net.Mail.MailMessage
$smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $Port)

try {
    $mail.From = $From
    $mail.To.Add($To)
    $mail.Subject = "Test SMTP RDI MeteoBot - $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    $mail.Body = "Test d'envoi SMTP depuis $env:COMPUTERNAME sans Outlook."

    if ($PdfPath) {
        $attachment = New-Object System.Net.Mail.Attachment($PdfPath)
        $mail.Attachments.Add($attachment)
        Log "Piece jointe : $(Split-Path $PdfPath -Leaf)"
    }

    $smtp.EnableSsl = [bool]$UseSsl
    $smtp.Timeout = $TimeoutSeconds * 1000
    $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
    $smtp.UseDefaultCredentials = $false

    Log "Envoi SMTP..."
    $smtp.Send($mail)
    Log "Envoi SMTP : OK"
    Log "Log : $logFile"
}
catch {
    Log "Envoi SMTP : ERREUR - $_"
    exit 1
}
finally {
    $mail.Dispose()
    $smtp.Dispose()
}
