<# .SYNOPSIS
     Synchronisation aller Dateien im lokalen Workspace(C:\Work)
.DESCRIPTION
     Alle heruntergeladen Dateien im Arbeitsbereich werden aktualisiert bzw. gelöscht 
.NOTES
     File Name : WorkSpaceSync.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version 1.00 – Vault 2023 support
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > WorkSpaceSync.ps1
.INPUTTYPE
     none
.RETURNVALUE
     void (Es werden Logdateien im Ordner C:\Work\Administration\WorkspaceSyncBericht\ erstellt )
.COMPONENT
     Vault Server
#>
$administrationFolderPath = "C:\Work\Administration"
$vaultExplorer = '"C:\Program Files\Autodesk\Vault Client 2023\Explorer\Connectivity.WorkspaceSync.exe"'
$server = "192.168.0.1"
$vaultName = "Vault"
$vaultUserName = "BE-Automation"
$vaultPasswort = "BE-Automation"
$workspaceSyncSettings = "C:\Work\Administration\Standardeinstellungen\Vault\Einstellungen_Vault\Work_Sync_Settings.xml"
$workspaceSyncBerichtPath = "C:\Work\Administration\WorkspaceSyncBericht"
$logFileName = "SyncBericht-" + (Get-Date -Format yyyy) + "-" + (Get-Date -Format MM) + ".csv"
$logFile = $workspaceSyncBerichtPath + "\" + $logFileName

#Ordner und Dateiüberpüfung
if (!(Test-Path ($administrationFolderPath))) {
     Write-Host "Ordner ("+ $administrationFolderPath +") wurde nicht gefunden!"
     exit  
}

if (!(Test-Path ($workspaceSyncSettings))) {
     Write-Host "Work_Sync_Settings.xml wurde nicht gefunden!"
     exit  
}

if (!(Test-Path ($workspaceSyncBerichtPath))) {
     New-Item -Path $workspaceSyncBerichtPath -ItemType Directory
}
if (!(Test-Path ($workspaceSyncBerichtPath + "\SyncBericht-JJJJ-MM.csv"))) {
     New-Item -Path ($workspaceSyncBerichtPath + "\SyncBericht-JJJJ-MM.csv") -ItemType File
}

$startArgsServer = "-N" + $server + "\" + $vaultName
$startArgsUser = "-VU" + $vaultUserName
$startArgsPasswort = "-VP" + $vaultPasswort
$startArgsSettings = "-S" + $workspaceSyncSettings
$startArgslogFile = "-F" + $logFile

$startArgs = @(
     '/c',
     $vaultExplorer,
     $startArgsServer,
     $startArgsUser,
     $startArgsPasswort
     $startArgsSettings,
     $startArgslogFile  
)

start-process "cmd" -Argumentlist $startArgs -WindowStyle Hidden