# CBC - Synchronisation Google Apps Script vers GitHub

Source officielle :

- Google Sheets : `CBC-2025 - Google xlsm`
- Dossier Drive : `AS Construct / AS - ADMIN / AS - BANQUE / CBC`
- Projet Apps Script : `1ecJiQ8w5UD7rNEd7k840P77Q0QruEBQvCSpjwXafgWlPReGiabQEwH6f`
- Dépôt GitHub : `CERNIGLIARO/C-GITHUB`
- Dossier GitHub : `CBC`

## Principe

La synchronisation est volontairement à sens unique :

`Google Apps Script -> PC -> GitHub`

Le script automatique exécute :

1. `git pull --ff-only origin main`
2. `clasp pull`
3. `git add`
4. `git commit` seulement s'il y a des changements
5. `git push origin main`

Il n'exécute jamais `clasp push`.

## Installation initiale sur Windows

Dans PowerShell ou Invite de commandes :

```text
node --version
npm --version
git --version
```

Installer clasp :

```text
npm install -g @google/clasp
```

Puis connecter le compte Google propriétaire du script :

```text
clasp login
```

Dans le dossier local :

```text
C:\GITHUB\C-GITHUB\CBC
```

tester :

```text
clasp pull
```

Ensuite lancer :

```text
SYNC_GOOGLE_VERS_GITHUB.bat
```

## Automatisation Windows

Créer une tâche dans le Planificateur de tâches Windows qui lance :

```text
C:\GITHUB\C-GITHUB\CBC\SYNC_GOOGLE_VERS_GITHUB.bat
```

Par exemple toutes les heures.

## Sécurité

- Le script s'arrête s'il détecte des modifications Git locales non enregistrées.
- Il utilise `git pull --ff-only` pour éviter une fusion automatique inattendue.
- Git garde l'historique de chaque version synchronisée.
- Ne pas ajouter de token GitHub ou de mot de passe dans les fichiers du dépôt.
