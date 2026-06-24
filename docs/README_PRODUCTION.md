# Exploitation production

## Environnement

Machine cible : `10.2.2.12`
Utilisateur : `j.augeraud`
Python portable : `D:\JAU\.venv\Scripts\python.exe`

Le pipeline ne depend pas du `PATH` Windows pour Python. Il recherche d'abord les environnements `.venv`, dont `D:\JAU\.venv`.

## Flux quotidien

1. Generation du bulletin HTML.
2. Conversion HTML vers PDF avec Playwright/Chromium.
3. Envoi SMTP avec le PDF en piece jointe.
4. Destinataires places en CCI depuis `20260623/Programme/Liste-diffusion.csv`.

Statut de recette : envoi SMTP reel confirme recu le 2026-06-24 a 10:28,
session utilisateur deconnectee cote poste client.

## Tache planifiee

Nom : `BulletinMeteoLunVen`
Horaire : lundi-vendredi a 08:00
Script : `20260623/Programme/run_bulletin_complet.ps1`
Mode : SMTP via `mx-mibc-fr-08.mailinblack.com:25`

## Tests

Test sans envoi :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -DryRun
```

Envoi reel :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1"
```

Test SMTP explicite :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -DeliveryMethod Smtp
```

## Logs

Log courant : `20260623/Programme/run_log.txt`

Les logs de preuve ou de recette doivent etre archives dans `20260623/Archives/Preuves` ou `20260623/Archives/Logs`.

## Protection mail

`run_bulletin_complet.ps1` accepte `-MailTimeoutSeconds`.
Par defaut, l'etape mail est arretee apres 120 secondes si le sous-processus mail ne repond pas.

Exemple de test court :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -MailTimeoutSeconds 20
```

## Point de vigilance session

Le SMTP ne depend pas d'Outlook.
La tache planifiee actuelle utilise encore `LogonType Interactive` : elle fonctionne en session ouverte ou deconnectee.
Pour garantir l'execution apres redemarrage sans aucune session utilisateur, creer une tache avec mot de passe stocke ou compte de service dedie.
