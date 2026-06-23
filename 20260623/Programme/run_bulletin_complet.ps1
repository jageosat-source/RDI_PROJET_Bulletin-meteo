# run_bulletin_complet.ps1 - Pipeline complet bulletin meteo
# 1. Genere le bulletin HTML via Python (generate_bulletin.py)
# 2. Convertit en PDF via Playwright (html_to_pdf.py)
# 3. Envoie l'e-mail via Outlook avec le PDF en piece jointe

param(
    [switch]$DryRun,
    [int]$MailTimeoutSeconds = 120
)

$ErrorActionPreference = "Stop"
$scriptDir   = $PSScriptRoot
$logFile     = Join-Path $scriptDir "run_log.txt"
$bulletinDir = Join-Path (Split-Path $scriptDir -Parent) "Bulletins"

function Log($msg) {
    $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "$ts  $msg"
    Write-Host $line
    Add-Content -Path $logFile -Value $line -Encoding UTF8
}

function Find-Python {
    $versionRoot = Split-Path $scriptDir -Parent
    $projectRoot = Split-Path $versionRoot -Parent
    $jauRoot = Split-Path $projectRoot -Parent

    # 1) Environnements portables controles, avant PATH/systeme.
    #    Le serveur 10.2.2.12 peut utiliser D:\JAU\.venv cree via uv sans droits admin.
    $pythonPaths = @(
        (Join-Path $versionRoot ".venv\Scripts\python.exe"),
        (Join-Path $projectRoot ".venv\Scripts\python.exe"),
        (Join-Path $jauRoot ".venv\Scripts\python.exe"),
        "$env:LOCALAPPDATA\miniconda3\python.exe",
        "$env:USERPROFILE\miniconda3\python.exe",
        "$env:LOCALAPPDATA\miniconda3\envs\base\python.exe",
        "C:\ProgramData\miniconda3\python.exe",
        "C:\ProgramData\Anaconda3\python.exe",
        "C:\ProgramData\anaconda3\python.exe",
        "$env:USERPROFILE\Anaconda3\python.exe",
        "$env:USERPROFILE\anaconda3\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
        "C:\Python312\python.exe",
        "C:\Python311\python.exe"
        # NOTE : WindowsApps/python3.exe deliberement absent - pas de Playwright
    )

    foreach ($path in $pythonPaths) {
        if (Test-Path -LiteralPath $path) {
            return @{ Exe = $path; Pre = @() }
        }
    }

    # 2) Lanceur officiel Windows 'py' (repli)
    $candidate = Get-Command py -ErrorAction SilentlyContinue
    if ($candidate) {
        return @{ Exe = "py"; Pre = @("-3") }
    }

    # 3) python / python3 sur le PATH (dernier recours)
    foreach ($name in @("python", "python3")) {
        $candidate = Get-Command $name -ErrorAction SilentlyContinue
        if ($candidate -and ($candidate.Source -notlike "*\Microsoft\WindowsApps\*")) {
            return @{ Exe = $candidate.Source; Pre = @() }
        }
    }

    return $null
}

Log "=== Debut pipeline bulletin ==="

$python = Find-Python
if (-not $python) {
    Log "Etape 1 : ERREUR - Python introuvable (attendu : D:\JAU\.venv\Scripts\python.exe ou Python avec Playwright, hors alias Windows Store)."
    exit 1
}
$pythonExe = $python.Exe
$pythonPre = $python.Pre
Log "Python : $pythonExe $($pythonPre -join ' ')"

# Etape 1 : generation du bulletin HTML
Log "Etape 1 : generation bulletin Python..."
$pyScript = Join-Path $scriptDir "generate_bulletin.py"
if (-not (Test-Path $pyScript)) {
    Log "Etape 1 : ERREUR - generate_bulletin.py introuvable."
    exit 1
}

$result = & $pythonExe @pythonPre $pyScript 2>&1
$code = $LASTEXITCODE
$result | ForEach-Object { Log "  [py] $_" }
if ($code -ne 0) {
    Log "Etape 1 : ERREUR - Python a quitte avec le code $code"
    exit 1
}
Log "Etape 1 : OK"

# Trouver le bulletin le plus recent genere dans le dossier commun du projet
$latest = Get-ChildItem $bulletinDir -Filter "Meteo_*.html" |
          Sort-Object LastWriteTime -Descending |
          Select-Object -First 1
if (-not $latest) {
    Log "ERREUR : aucun bulletin trouve dans $bulletinDir"
    exit 1
}
$htmlPath = $latest.FullName
$pdfPath  = [System.IO.Path]::ChangeExtension($htmlPath, "pdf")
Log "Bulletin HTML : $($latest.Name)"

# Etape 2 : conversion PDF via html_to_pdf.py (Playwright)
Log "Etape 2 : conversion PDF..."
$pdfScript = Join-Path $scriptDir "html_to_pdf.py"
if (-not (Test-Path $pdfScript)) {
    Log "Etape 2 : ERREUR - html_to_pdf.py introuvable."
    exit 1
}

if (Test-Path $pdfPath) {
    Remove-Item -LiteralPath $pdfPath -Force
}

$pdfResult = & $pythonExe @pythonPre $pdfScript --input $htmlPath --output $pdfPath --scale 0.9 --margin 6mm 2>&1
$pdfCode = $LASTEXITCODE
$pdfResult | ForEach-Object { Log "  [pdf] $_" }
if ($pdfCode -ne 0) {
    Log "Etape 2 : ERREUR - conversion PDF echouee avec le code $pdfCode"
    exit 1
}
if (-not (Test-Path $pdfPath)) {
    Log "Etape 2 : ERREUR - PDF non genere : $pdfPath"
    exit 1
}
Log "Etape 2 : OK -> $([System.IO.Path]::GetFileName($pdfPath))"

# Etape 3 : envoi e-mail
if ($DryRun) {
    Log "Etape 3 : controle e-mail sans envoi..."
} else {
    Log "Etape 3 : envoi e-mail..."
}
$sendScript = Join-Path $scriptDir "send_bulletin.ps1"
if (-not (Test-Path $sendScript)) {
    Log "Etape 3 : ERREUR - send_bulletin.ps1 introuvable."
    exit 1
}

if ($DryRun) {
    $sendOutput = & $sendScript -PdfPath $pdfPath -DryRun 2>&1
    $sendCode = $LASTEXITCODE
} else {
    $mailStamp = Get-Date -Format "yyyyMMdd-HHmmss-ffff"
    $mailOut = Join-Path $env:TEMP "bulletin_mail_$mailStamp.out.log"
    $mailErr = Join-Path $env:TEMP "bulletin_mail_$mailStamp.err.log"
    $powershellExe = Join-Path $PSHOME "powershell.exe"
    $mailArgs = @(
        "-ExecutionPolicy", "Bypass",
        "-NoProfile",
        "-File", $sendScript,
        "-PdfPath", $pdfPath
    )

    $mailProcess = Start-Process -FilePath $powershellExe -ArgumentList $mailArgs `
        -PassThru -WindowStyle Hidden `
        -RedirectStandardOutput $mailOut -RedirectStandardError $mailErr

    if (-not $mailProcess.WaitForExit($MailTimeoutSeconds * 1000)) {
        try { $mailProcess.Kill() } catch {}
        $sendCode = 124
        $sendOutput = @("Timeout envoi e-mail apres $MailTimeoutSeconds secondes. Processus PowerShell mail arrete.")
    } else {
        $sendCode = $mailProcess.ExitCode
        $sendOutput = @()
    }

    foreach ($mailLog in @($mailOut, $mailErr)) {
        if (Test-Path -LiteralPath $mailLog) {
            $sendOutput += Get-Content -LiteralPath $mailLog -ErrorAction SilentlyContinue
            Remove-Item -LiteralPath $mailLog -Force -ErrorAction SilentlyContinue
        }
    }
}
$sendOutput | Where-Object { "$($_)".Trim() } | ForEach-Object { Log "  [mail] $_" }
if ($sendCode -ne 0) {
    Log "Etape 3 : ERREUR - envoi non confirme (code $sendCode)"
    exit 1
}
Log "Etape 3 : OK"

Log "=== Pipeline termine ==="
exit 0