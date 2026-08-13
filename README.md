# TMA autonome

Complément COM x64 pour Outlook classique permettant de créer des réunions Microsoft Teams avec une identité COM indépendante du complément officiel.

## Fonctionnalités

- réunion Teams classique et réunion instantanée ;
- invitation HTML en français ou en anglais ;
- rejoindre, ouvrir les options ou retirer la réunion en ligne ;
- ouverture des flux Teams officiels pour les webinaires, assemblées et rendez-vous virtuels ;
- coexistence avec `TeamsAddin.FastConnect` sans réutiliser son CLSID ;
- scheduler et authentification OneAuth fournis par le payload TMA Microsoft extrait au build.

## Organisation du dépôt

```text
src/TmaCleanRoom/       Sources C# et modèles d'invitation
installer/              Définition WiX du MSI
scripts/build/          Builds Windows et Linux
scripts/dependencies/   Acquisition vérifiée du payload Microsoft
scripts/dev/            Enregistrement COM de développement
scripts/test/           Diagnostics et audit de publication
scripts/release/        Création d'un export source-only
tools/                  Générateurs multiplateformes
vendor/                 Dépendances locales ignorées par Git
```

## Prérequis de compilation Windows

- Windows x64 ;
- Outlook classique x64 et ses Primary Interop Assemblies ;
- .NET Framework 4.8.1 Developer Pack ;
- Windows 10/11 SDK ;
- WiX Toolset 4 accessible via `wix.exe` ;
- le MSI x64 officiel « Microsoft Teams Meeting Add-in » obtenu par vos propres moyens autorisés.

Le dépôt ne contient et ne redistribue aucun binaire Microsoft.

Le MSI généré embarque un sous-ensemble validé du payload officiel : le
scheduler TMA, son wrapper OneAuth et leurs dépendances directes. Le loader COM
Microsoft, WebView2, les vues stock et les ressources inutilisées ne sont pas
inclus. L'identité du compte reste fournie par Office et Windows WAM.

## Compiler

Si New Teams est déjà installé :

```powershell
.\scripts\dependencies\Get-TmaInstaller.ps1
.\scripts\build\Build-Windows.ps1
```

La version utilisée pour les releases est verrouillée dans
`dependencies.lock.json`. Si la version installée de Teams est plus récente,
`Get-TmaInstaller.ps1 -AcceptNewVersion` permet de l'utiliser explicitement ;
le build reste signé et vérifié, mais ne sera plus identique à la release
verrouillée.

Si Teams n'est pas installé, le script peut utiliser le bootstrapper officiel
Microsoft et demander une élévation UAC :

```powershell
.\scripts\dependencies\Get-TmaInstaller.ps1 -ProvisionTeamsIfMissing
.\scripts\build\Build-Windows.ps1
```

Il est aussi possible de fournir directement un MSI existant avec
`.\scripts\build\Build-Windows.ps1 -MsiPath C:\chemin\MicrosoftTeamsMeetingAddinInstaller.msi`.

Le résultat est `TMA-Standalone.msi`. Le build extrait le payload x64 dans `_work`, compile le complément, génère les composants WiX puis vérifie que le MSI possède uniquement l’identité COM clean-room.

### CI Linux

Le workflow `linux-build.yml` construit le MSI sur Ubuntu sans installer ni
exécuter Teams. Il télécharge le MSIX x64 officiel Microsoft, extrait le MSI
avec `msitools`, compile avec le SDK .NET puis génère le MSI Windows avec WiX :

```powershell
./scripts/build/Build-Linux.ps1
```

Les tests d'installation, COM et Outlook restent nécessairement exécutés sur
Windows.

## Identité COM

- ProgID : `TmaCleanRoom.Connect`
- CLSID : `{8F5373B8-4973-4E58-A69E-CB57AA22691C}`
- installation : `C:\Program Files\TMA-Standalone`

Le MSI ne modifie ni le CLSID, ni la TypeLib, ni le `LoadBehavior` du complément Teams officiel.

## Diagnostic

```powershell
.\scripts\test\Test-TmaInstallation.ps1
```

Les journaux d’exécution sont conservés dans `%LOCALAPPDATA%\TMA-CleanRoom\teams-bridge.log`. Ils ne doivent contenir aucun jeton OAuth.

## Limites

- Outlook classique x64 uniquement ;
- les flux webinaire, assemblée et rendez-vous virtuel ouvrent Teams, comme le complément officiel ;
- le comportement dépend des API privées du payload TMA et doit être revalidé après une mise à jour Microsoft importante ;
- le MSI per-machine enregistre le complément Outlook pour l’utilisateur qui effectue l’installation.

## Développement

`scripts\dev\Register-Dev.ps1` et `Unregister-Dev.ps1` inscrivent uniquement
la build locale pour l’utilisateur courant. Fermez Outlook avant de remplacer
la DLL chargée.

Ce projet n’est ni affilié, ni approuvé, ni pris en charge par Microsoft.
