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
