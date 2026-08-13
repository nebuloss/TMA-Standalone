# Dépendances Microsoft locales

Ce dossier reçoit `MicrosoftTeamsMeetingAddinInstaller.msi` au moyen de :

```powershell
.\scripts\dependencies\Get-TmaInstaller.ps1
```

Le MSI et les DLL qui en sont extraites ne sont pas distribués sous la licence
MIT du projet et restent ignorés par Git. Leur signature Microsoft et leur
empreinte sont contrôlées avant le build.
