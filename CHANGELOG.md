# Changelog

## RDI_MeteoBot_20260623-v1.0.1 - 2026-06-23

- Ajout d'une protection timeout autour de l'envoi Outlook pour eviter qu'une tache reste bloquee indefiniment.
- Validation DryRun apres correction : OK.
- Tentative d'envoi reel : generation HTML/PDF OK, blocage Outlook COM confirme.
- Preuves d'echec archivees dans `20260623/Archives/Preuves/20260623/outlook-com-fail`.
- Envoi reel confirme recu apres relance PowerShell le 2026-06-23 a 11:10.
- Correction du nettoyage des logs temporaires mail pour eviter l'erreur non bloquante `Remove-Item`.

## 20260623-prod-v1.0.0 - 2026-06-23

- Mise en production cible sur le serveur 10.2.2.12.
- Abandon de l'envoi Teams.
- Envoi e-mail Outlook avec destinataires en CCI.
- Lecture de la liste de diffusion depuis `Programme/Liste-diffusion.csv`.
- Piece jointe PDF uniquement.
- Horaire de tache planifiee ajuste a 08:00, lundi-vendredi.
- Support du Python portable `D:\JAU\.venv\Scripts\python.exe`.
- Ajout du mode `-DryRun` pour `send_bulletin.ps1` et `run_bulletin_complet.ps1`.
- Validation complete sans envoi realisee le 2026-06-23 a 10:11.
