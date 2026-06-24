# send_bulletin.ps1 — Envoi du bulletin météo par e-mail (Outlook)
# UTF-8 BOM requis — ne pas modifier l'encodage de ce fichier

param(
    [string]$PdfPath = "",
    [string]$DiffusionListPath = (Join-Path $PSScriptRoot "Liste-diffusion.csv"),
    [ValidateSet("Outlook", "Smtp")]
    [string]$DeliveryMethod = "Outlook",
    [string]$SmtpServer = "mx-mibc-fr-08.mailinblack.com",
    [int]$SmtpPort = 25,
    [switch]$SmtpUseSsl,
    [string]$From = "j.augeraud@geo-sat.com",
    [switch]$DryRun
)

$ErrorActionPreference = "Stop"

function Get-LatestPdf {
    $bulletinDir = Join-Path (Split-Path $PSScriptRoot -Parent) "Bulletins"
    $latest = Get-ChildItem -LiteralPath $bulletinDir -Filter "Meteo_*.pdf" -File |
              Sort-Object LastWriteTime -Descending |
              Select-Object -First 1
    if (-not $latest) {
        throw "Aucun PDF trouve dans $bulletinDir"
    }
    return $latest.FullName
}

function Get-CsvDelimiter {
    param([string]$Path)

    $firstLine = Get-Content -LiteralPath $Path -TotalCount 1 -Encoding UTF8
    $semicolonCount = ($firstLine.ToCharArray() | Where-Object { $_ -eq ';' }).Count
    $commaCount = ($firstLine.ToCharArray() | Where-Object { $_ -eq ',' }).Count

    if ($semicolonCount -gt $commaCount) {
        return ';'
    }
    return ','
}

function Get-DiffusionRecipients {
    param([string]$Path)

    if (-not (Test-Path -LiteralPath $Path)) {
        throw "Liste de diffusion introuvable : $Path"
    }

    $delimiter = Get-CsvDelimiter -Path $Path
    $rows = @(Import-Csv -LiteralPath $Path -Delimiter $delimiter -Encoding UTF8)
    if ($rows.Count -eq 0) {
        throw "Liste de diffusion vide : $Path"
    }

    $preferredColumns = @("adresse de courrier", "adresse", "email", "mail", "courriel", "e-mail")
    $addresses = @()
    $invalid = @()

    foreach ($row in $rows) {
        $address = $null
        foreach ($column in $preferredColumns) {
            $property = $row.PSObject.Properties | Where-Object { $_.Name -ieq $column } | Select-Object -First 1
            if ($property) {
                $address = [string]$property.Value
                break
            }
        }

        if (-not $address) {
            $property = $row.PSObject.Properties |
                        Where-Object { $_.Name -match '(?i)(mail|courriel|adresse)' } |
                        Select-Object -First 1
            if ($property) {
                $address = [string]$property.Value
            }
        }

        if ($address) {
            $address = $address.Trim()
        }
        if (-not $address) {
            continue
        }

        if ($address -notmatch '^[^@\s]+@[^@\s]+\.[^@\s]+$') {
            $invalid += $address
            continue
        }

        $addresses += $address
    }

    if ($invalid.Count -gt 0) {
        throw "Adresse(s) invalide(s) dans la liste de diffusion : $($invalid -join ', ')"
    }
    if ($addresses.Count -eq 0) {
        throw "Aucune adresse valide trouvee dans la liste de diffusion : $Path"
    }

    $unique = @()
    $seen = @{}
    foreach ($address in $addresses) {
        $key = $address.ToLowerInvariant()
        if (-not $seen.ContainsKey($key)) {
            $seen[$key] = $true
            $unique += $address
        }
    }

    return $unique
}

if (-not $PdfPath) {
    $PdfPath = Get-LatestPdf
}
elseif (-not [System.IO.Path]::IsPathRooted($PdfPath)) {
    $bulletinDir = Join-Path (Split-Path $PSScriptRoot -Parent) "Bulletins"
    $PdfPath = Join-Path $bulletinDir $PdfPath
}

if (-not (Test-Path -LiteralPath $PdfPath)) {
    Write-Error "PDF introuvable : $PdfPath"
    exit 1
}
if ([System.IO.Path]::GetExtension($PdfPath) -ine ".pdf") {
    Write-Error "Le fichier a envoyer doit etre un PDF : $PdfPath"
    exit 1
}

$pdfName = Split-Path $PdfPath -Leaf
Write-Host "Bulletin PDF : $pdfName"

$recipients = @(Get-DiffusionRecipients -Path $DiffusionListPath)
$bccLine = $recipients -join ";"

# Sujet avec date française
$now     = Get-Date
$joursFR = @("Dimanche","Lundi","Mardi","Mercredi","Jeudi","Vendredi","Samedi")
$moisFR  = @("","Janvier","Février","Mars","Avril","Mai","Juin","Juillet","Août","Septembre","Octobre","Novembre","Décembre")
$dateLabel = "$($joursFR[$now.DayOfWeek.value__]) $($now.Day) $($moisFR[$now.Month]) $($now.Year)"
$subject   = "Bulletin Météo Bordeaux — $dateLabel"
$body      = "Bonjour,$([Environment]::NewLine * 2)Veuillez trouver ci-joint le bulletin météo de l'agglomération Bordelaise.$([Environment]::NewLine * 2)$dateLabel"

if ($DryRun) {
    Write-Host ""
    Write-Host "=== Controle e-mail (sans envoi) ==="
    Write-Host "Mode envoi : $DeliveryMethod"
    Write-Host "Liste diffusion : $DiffusionListPath"
    Write-Host "Destinataires CCI : $($recipients.Count)"
    Write-Host "Piece jointe : PDF"
    if ($DeliveryMethod -eq "Smtp") {
        Write-Host "SMTP : $SmtpServer`:$SmtpPort SSL=$([bool]$SmtpUseSsl)"
        Write-Host "Expediteur : $From"
    }
    Write-Host "[OK] Controle sans envoi termine — $subject"
    exit 0
}

if ($DeliveryMethod -eq "Smtp") {
    Write-Host ""
    Write-Host "=== Envoi e-mail (SMTP) ==="
    Write-Host "SMTP : $SmtpServer`:$SmtpPort SSL=$([bool]$SmtpUseSsl)"
    Write-Host "Expediteur : $From"
    Write-Host "Liste diffusion : $DiffusionListPath"
    Write-Host "Destinataires CCI : $($recipients.Count)"
    Write-Host "Piece jointe : PDF"

    $mail = New-Object System.Net.Mail.MailMessage
    $smtp = New-Object System.Net.Mail.SmtpClient($SmtpServer, $SmtpPort)
    try {
        $mail.From = $From
        foreach ($recipient in $recipients) {
            $mail.Bcc.Add($recipient)
        }
        $mail.Subject = $subject
        $mail.Body = $body
        $mail.Attachments.Add($PdfPath) | Out-Null

        $smtp.EnableSsl = [bool]$SmtpUseSsl
        $smtp.Timeout = 60000
        $smtp.DeliveryMethod = [System.Net.Mail.SmtpDeliveryMethod]::Network
        $smtp.UseDefaultCredentials = $false
        $smtp.Send($mail)
        Write-Host "[OK] E-mail envoye SMTP — $subject"
    } catch {
        Write-Error "[ERREUR] SMTP : $_"
        exit 1
    } finally {
        $mail.Dispose()
        $smtp.Dispose()
    }
}
else {
    # Envoi via Outlook
    Write-Host ""
    Write-Host "=== Envoi e-mail (Outlook) ==="
    try {
        # S'attacher a l'Outlook deja ouvert (sinon en lancer un) — evite l'erreur
        # CO_E_SERVER_EXEC_FAILURE quand Outlook tourne deja dans la session.
        $outlook = $null
        try { $outlook = [Runtime.InteropServices.Marshal]::GetActiveObject("Outlook.Application") } catch {}
        if (-not $outlook) { $outlook = New-Object -ComObject Outlook.Application }
        try { $outlook.GetNamespace("MAPI").Logon($null,$null,$false,$false) } catch {}

        $mail         = $outlook.CreateItem(0)
        $mail.BCC     = $bccLine
        $mail.Subject = $subject
        $mail.Body    = $body

        $mail.Attachments.Add($PdfPath) | Out-Null
        Write-Host "Liste diffusion : $DiffusionListPath"
        Write-Host "Destinataires CCI : $($recipients.Count)"
        Write-Host "Piece jointe : PDF"

        $mail.Send()
        Write-Host "[OK] E-mail envoye Outlook — $subject"
    } catch {
        # exit 1 : l'appelant (run_bulletin_complet.ps1 / tache planifiee) doit voir l'echec
        Write-Error "[ERREUR] Outlook : $_"
        exit 1
    }
}

Write-Host ""
Write-Host "Termine."
exit 0
