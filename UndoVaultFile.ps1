<# .SYNOPSIS
     Auschecken Rückgängig von Dateien auf den VaultServer
.DESCRIPTION
     Reservierungen von Dateien werden enfernt und der alte Stand wird hergestellt  
.NOTES
     File Name : UndoVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version 0.25 – bugfix => undo SpezifikationsPdf
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > UndoVaultFile.ps1 -Auftragsnummer „8951234“
.INPUTTYPE
     Auftragsnummer 
.RETURNVALUE
     $downloadresult
     $errCode
.COMPONENT
     Vault Server
#>

       
Param(
    [Parameter(Mandatory = $true)]          
    [String]$Auftragsnummer
)

class DownloadInfo {
    [bool]$Success = $null
    [string]$FileName
    [string]$FullFileName
    [string]$CheckOutState
    [bool]$IsCheckOut = $null
    [string]$CheckOutPC
    [string]$EditedBy
    [string]$ErrorState
}

try {
    Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
    Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
    [System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2022\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")
}
catch {
    Write-Host "Vault Client 2022 oder DataStandard wurde nicht gefunden!"
    $errCode = 9 #Vault Client 2022 oder DataStandard wurde nicht gefunden
    $downloadresult.Success = $false
    LogOut($downloadresult)
}

$downloadresult = [DownloadInfo]::new()
function LogOut {
    param (
        [DownloadInfo]$downloadinfo
    )
    Write-Host "---DownloadInfo---" 
    Write-Output (ConvertTo-Json $downloadinfo)
    Write-Host "---DownloadInfo---"
    if ($null -ne $connection) {
        # $vault.Dispose()
        $logOff = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($connection) #Vault Connection schließen
        Write-Host "Vaultverbindungsstatus:" + $logOff
    }
    $Host.SetShouldExit([int]$errCode)
    exit
}
# Auftragsnummervalidierung
if (($Auftragsnummer.Length -eq 6 -or $Auftragsnummer.Length -eq 7) -and $Auftragsnummer -match '^\d+$') {
    $AuftragsTyp = "Auftrag"
}
elseif ($Auftragsnummer -match '[0-9]{2}[-]0[1-9]|1[0-2][-][0-9]{4}') {
    $AuftragsTyp = "Angebot"
}
else {
    $errCode = 6 #Invalide Auftrags bzw. Angebotsnummer
    $downloadresult.Success = $false
    LogOut($downloadresult)
}
# Vault Login

try {

    $AdskLicensing = "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_5.dll"
    if (!(Test-Path $AdskLicensing -PathType leaf)) {
        try {
            Copy-Item -Path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\AdskLicensingSDK_5.dll" -Destination "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_5.dll"
        }
        catch {
            Write-Host "AdskLicensingSDK_5.dll wurde nicht gefunden!"
            $errCode = 8 #Fehlende AdskLicensingSDK_5.dll
            LogOut($downloadresult)
        } 
    }

    $settings = New-Object Autodesk.DataManagement.Client.Framework.Vault.Forms.Settings.LoginSettings
    $settings.ServerName = "192.168.0.1"
    $settings.VaultName = "vault"
    $settings.AutoLoginMode = 3
    $connection = [Autodesk.DataManagement.Client.Framework.Vault.Forms.Library]::Login($settings)

    if ($null -eq $connection) {
        $settings.AutoLoginMode = 1
        $connection = [Autodesk.DataManagement.Client.Framework.Vault.Forms.Library]::Login($settings)
    }
}
catch {
    $errCode = 2 #Login fehlgeschlagen
    $downloadresult.Success = $false
    LogOut($downloadresult)
} 

try {

    #Quellpfad ermitteln
    $seachFile = $Auftragsnummer + "-AutoDeskTransfer.xml"

    if ($AuftragsTyp -eq "Auftrag") {
        $seachPath = "C:\Work\AUFTRÄGE NEU\\Konstruktion"
    }
    elseif ($AuftragsTyp -eq "Angebot") {
        $seachPath = "C:\Work\AUFTRÄGE NEU\Angebote"
    }
    else {
        $seachPath = "C:\Work\AUFTRÄGE NEU\"
    }
    
    $sourceFile = Get-ChildItem -Path $seachPath -Recurse -Include $seachFile
    if ($null -eq $sourceFile) {
        Write-Host "AutoDeskTransferXml im Arbeitsbereich nicht gefunden."-ForegroundColor DarkRed
        $errCode = "7" # Datei im Arbeitsbereich nicht gefunden
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }
    if ($sourceFile.Count -gt 1) {
        Write-Host "AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden."-ForegroundColor DarkRed
        $errCode = "5"# AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden.
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }

    $VltHelpers = New-Object VdsSampleUtilities.VltHelpers

    #FileStatus auslesen 
    $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
    $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $sourceFile.FullName) 

    if ($FileStatus["ErrorState"] -eq "VaultFileNotFound") {
        $downloadresult.Success = $false
        $downloadresult.ErrorState = $FileStatus["ErrorState"]
        $errCode = "7" # Datei im Vault nicht gefunden
        LogOut($downloadresult)
    }

    if ($FileStatus["CheckOutState"] -ne "CheckedOutByCurrentUser") {
        $downloadresult.Success = $false
        $downloadresult.FileName = $FileStatus["FileName"]
        $downloadresult.FullFileName = $FileStatus["FullFileName"]
        $downloadresult.CheckOutState = $FileStatus["CheckOutState"]
        $downloadresult.IsCheckOut = [System.Convert]::ToBoolean($FileStatus["CheckOut"])
        $downloadresult.CheckOutPC = $FileStatus["CheckOutPC"]
        $downloadresult.EditedBy = $FileStatus["EditedBy"]
        $downloadresult.ErrorState = $FileStatus["ErrorState"]
        $errCode = "1" # Datei Reservierung enfernen fehlgeschlagen
        LogOut($downloadresult)
    }

    # Auschecken Rückgängig - Reservierung enfernen
    try {

        $vault = $connection.WebServiceManager
        $vaultPathAutodesktransferXml = $VltHelpers.ConvertLocalPathToVaultPath($connection, $FileStatus["FullFileName"])
        $BerechnungenPath = $vaultPathAutodesktransferXml + "/Berechnungen/"

        #Dateinamen der benötigten Dateien
        $undoFiles = @()
        $undoFiles += $vaultPathAutodesktransferXml + "/" + $FileStatus["FileName"]
        $undoFiles += $vaultPathAutodesktransferXml + "/" + $Auftragsnummer + "-Spezifikation.pdf"
        $undoFiles += $BerechnungenPath + $Auftragsnummer + ".html"
        $undoFiles += $BerechnungenPath + $Auftragsnummer + ".aus"
        $undoFiles += $BerechnungenPath + $Auftragsnummer + ".dat"
        $undoFiles += $BerechnungenPath + $Auftragsnummer + ".LILO"

        $vaultFoundUndoFiles = $vault.DocumentService.FindLatestFilesByPaths($undoFiles)

        $downloadTicket = New-Object Autodesk.Connectivity.WebServices.ByteArray
        
        foreach ($vaultFoundUndoFile in $vaultFoundUndoFiles) {
            if ($vaultFoundUndoFile.Id -gt 0 -and $vaultFoundUndoFile.CheckedOut) {
                $vault.DocumentService.UndoCheckoutFile($vaultFoundUndoFile.MasterId, [ref]$downloadTicket)
            } 
        }
        

        #FileStatus auslesen 
        $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
        $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $sourceFile.FullName) 
        $downloadresult.Success = $true
        $downloadresult.FileName = $FileStatus["FileName"]
        $downloadresult.FullFileName = $FileStatus["FullFileName"]
        $downloadresult.CheckOutState = $FileStatus["CheckOutState"]
        $downloadresult.IsCheckOut = [System.Convert]::ToBoolean($FileStatus["CheckOut"])
        $downloadresult.CheckOutPC = $FileStatus["CheckOutPC"]
        $downloadresult.EditedBy = $FileStatus["EditedBy"]
        $downloadresult.ErrorState = $FileStatus["ErrorState"]
        $errCode = 0
        LogOut($downloadresult)   
    }
    catch {
        $downloadresult.Success = $false
        $errCode = 1 # Datei Reservierung enfernen fehlgeschlagen
        LogOut($downloadresult)
    }
}
catch {
    $downloadresult.Success = $false
    $errCode = 1 # Datei Reservierung enfernen fehlgeschlagen
    LogOut($downloadresult)
}
