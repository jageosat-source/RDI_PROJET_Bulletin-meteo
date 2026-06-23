# Politique d'archivage

## Code et releases

Les releases validees sont conservees localement dans `20260623/Archives/Releases` avec un manifeste SHA256.
Git conserve l'historique source ; les archives ZIP servent de preuve restaurable.

## Bulletins

Le dossier `20260623/Bulletins` reste le dossier d'exploitation courant.
Les bulletins de test ou de recette sont deplaces dans `20260623/Archives/Preuves`.
Les bulletins anciens peuvent etre classes mensuellement dans `20260623/Archives/Bulletins/YYYY/MM`.

## Logs

`run_log.txt` reste le log courant.
Les copies de logs associees a une recette ou une release sont archivees avec la preuve correspondante.

## Suppression

Aucune suppression definitive sans archive ZIP ou dossier de preuve accompagne d'un manifeste SHA256.
