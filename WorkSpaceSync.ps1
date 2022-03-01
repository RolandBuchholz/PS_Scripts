<# .SYNOPSIS
     Synchronisation aller Dateien im lokalen Workspace(C:\Work)
.DESCRIPTION
     Alle heruntergeladen Dateien im Arbeitsbereich werden aktualiesiert bzw. gelöscht 
.NOTES
     File Name : WorkSpaceSync.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version Version 0.1 – Setup Startparameter
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > WorkSpaceSync.ps1
.INPUTTYPE
     none
.RETURNVALUE
     void (Es wird eine log Datei erstellt C:\Work\Administration\SyncBericht.csv)
.COMPONENT
     Vault Server
#>

$vaultExplorer = '"C:\Program Files\Autodesk\Vault Client 2022\Explorer\Connectivity.WorkspaceSync.exe"'
$server = "192.168.0.1"
$vaultName = "Vault"
$vaultUserName = "BE-Automation"
$vaultPasswort = "BE-Automation"
$workspaceSyncSettings = "C:\Work\Administration\Standardeinstellungen\Vault\Einstellungen_Vault\Work_Sync_Settings.xml"
$logFile = "C:\Work\Administration\SyncBericht.csv"

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

start-process "cmd" -Argumentlist $startArgs -NoNewWindow -wait
