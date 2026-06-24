# setup_tache_planifiee.ps1
# Cree / met a jour la tache planifiee : Bulletin Meteo, Lundi->Vendredi a 08h00,
# avec envoi SMTP, sans dependance Outlook.
#
# Robustesse :
#   - StartWhenAvailable : rattrape un declenchement manque (PC eteint/en veille a 08h00)
#                          -> la tache part des que possible apres disponibilite
#   - batterie : demarre et continue meme sur batterie (portable)
#   - WakeToRun : sort le PC de veille pour executer (sans effet si totalement eteint)
#   - LogonType Interactive : valide en session ouverte ou deconnectee.
#     Pour fonctionner apres redemarrage sans session, utiliser un compte/service dedie.

$taskName = "BulletinMeteoLunVen"
$scriptDir = $PSScriptRoot
$ps1      = Join-Path $scriptDir "run_bulletin_complet.ps1"
$logOut   = Join-Path $scriptDir "setup_result.txt"
$smtpServer = "mx-mibc-fr-08.mailinblack.com"
$smtpPort = 25
$mailFrom = "j.augeraud@geo-sat.com"

$out = @()
$out += "=== Configuration tache : $taskName ==="
$out += "Date : $(Get-Date)"

# Supprimer d'eventuelles anciennes taches
foreach ($old in @("BulletinMeteoLundi")) {
    try { Unregister-ScheduledTask -TaskName $old -Confirm:$false -ErrorAction Stop; $out += "Ancienne tache supprimee : $old" } catch {}
}

$actionArgs = @(
    "-ExecutionPolicy", "Bypass",
    "-NoProfile",
    "-WindowStyle", "Hidden",
    "-File", "`"$ps1`"",
    "-DeliveryMethod", "Smtp",
    "-SmtpServer", $smtpServer,
    "-SmtpPort", $smtpPort,
    "-From", $mailFrom,
    "-MailTimeoutSeconds", "120"
) -join " "

$action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument $actionArgs

$trigger = New-ScheduledTaskTrigger -Weekly `
           -DaysOfWeek Monday,Tuesday,Wednesday,Thursday,Friday -At 8:00AM

# LogonType Interactive => tourne dans la session utilisateur ouverte ou deconnectee,
# sans mot de passe stocke. Suffisant pour SMTP tant que la session reste active.
$principal = New-ScheduledTaskPrincipal -UserId "$env:USERDOMAIN\$env:USERNAME" `
             -LogonType Interactive -RunLevel Limited

$settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -WakeToRun `
            -ExecutionTimeLimit (New-TimeSpan -Hours 1) -MultipleInstances IgnoreNew `
            -RestartCount 2 -RestartInterval (New-TimeSpan -Minutes 5)
$settings.DisallowStartIfOnBatteries = $false
$settings.StopIfGoingOnBatteries     = $false

Register-ScheduledTask -TaskName $taskName -Action $action -Trigger $trigger `
    -Principal $principal -Settings $settings -Force | Out-Null

# Verification
$info = Get-ScheduledTaskInfo -TaskName $taskName
$t    = Get-ScheduledTask -TaskName $taskName
$out += ""
$out += "=== Verification ==="
$out += "Etat               : $($t.State)"
$out += "Prochaine execution: $($info.NextRunTime)"
$out += "Mode envoi         : SMTP"
$out += "SMTP               : $smtpServer`:$smtpPort"
$out += "Expediteur         : $mailFrom"
$out += "StartWhenAvailable : $($t.Settings.StartWhenAvailable)"
$out += "Sur batterie OK    : $(-not $t.Settings.DisallowStartIfOnBatteries)"
$out += "LogonType          : $($t.Principal.LogonType)"
$out += "RunLevel           : $($t.Principal.RunLevel)"

$out | Out-File $logOut -Encoding UTF8
$out | ForEach-Object { Write-Host $_ }
