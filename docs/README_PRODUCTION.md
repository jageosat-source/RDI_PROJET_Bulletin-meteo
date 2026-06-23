# Exploitation production

## Environnement

Machine cible : `10.2.2.12`
Utilisateur : `j.augeraud`
Python portable : `D:\JAU\.venv\Scripts\python.exe`

Le pipeline ne depend pas du `PATH` Windows pour Python. Il recherche d'abord les environnements `.venv`, dont `D:\JAU\.venv`.

## Flux quotidien

1. Generation du bulletin HTML.
2. Conversion HTML vers PDF avec Playwright/Chromium.
3. Envoi Outlook avec le PDF en piece jointe.
4. Destinataires places en CCI depuis `20260623/Programme/Liste-diffusion.csv`.

Statut de recette : envoi reel confirme recu apres relance PowerShell le 2026-06-23 a 11:10.

## Tache planifiee

Nom : `BulletinMeteoLunVen`
Horaire : lundi-vendredi a 08:00
Script : `20260623/Programme/run_bulletin_complet.ps1`

## Tests

Test sans envoi :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -DryRun
```

Envoi reel :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1"
```

## Logs

Log courant : `20260623/Programme/run_log.txt`

Les logs de preuve ou de recette doivent etre archives dans `20260623/Archives/Preuves` ou `20260623/Archives/Logs`.

## Protection anti-blocage Outlook

`run_bulletin_complet.ps1` accepte `-MailTimeoutSeconds`.
Par defaut, l'etape mail est arretee apres 120 secondes si Outlook COM ne repond pas.

Exemple de test court :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -MailTimeoutSeconds 20
```

## Point de vigilance Outlook

Les etapes Python, Playwright et PDF sont validees.
L'envoi reel depend d'Outlook COM dans la session `j.augeraud`.
Si l'erreur `0x80080005 CO_E_SERVER_EXEC_FAILURE` apparait, ouvrir Outlook manuellement dans la session serveur, terminer toute configuration de profil/compte, puis relancer un test d'envoi.
