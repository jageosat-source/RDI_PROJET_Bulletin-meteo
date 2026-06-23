# RDI_PROJET_Bulletin-meteo

Production du bulletin meteo Bordeaux.

Version de production active : `20260623-prod-v1.0.1`.

Le code source de production se trouve dans `20260623/Programme`.
Les bulletins, logs et preuves de recette sont conserves localement et exclus du versionnage Git courant.

## Commandes utiles

Controle complet sans envoi :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1" -DryRun
```

Execution production avec envoi :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\run_bulletin_complet.ps1"
```

Configuration de la tache planifiee :

```powershell
& "D:\JAU\RDI_PROJET_Bulletin-meteo\20260623\Programme\setup_tache_planifiee.ps1"
```