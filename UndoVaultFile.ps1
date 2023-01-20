<# .SYNOPSIS
     Auschecken Rückgängig von Dateien auf den VaultServer
.DESCRIPTION
     Reservierungen von Dateien werden enfernt und der alte Stand wird hergestellt  
.NOTES
     File Name : UndoVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version 1.10 – add custom filedownload
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > UndoVaultFile.ps1 -Auftragsnummer 8951234 $true
                                                                    (Auftragsnummer)(CustomFile optional)  
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
    [String]$Auftragsnummer,
    [bool]$CustomFile = $false
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
    $clientFrameworkVaultPath = "C:\Program Files\Autodesk\Vault Client 2023\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
    if (Test-Path $clientFrameworkVaultPath) {
        Add-Type -path "C:\Program Files\Autodesk\Vault Client 2023\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
        Add-Type -path "C:\Program Files\Autodesk\Vault Client 2023\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
    }
    else {
        Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
        Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
    }
    $vdsSampleUtilitiesPath = ($Env:ProgramData + "\Autodesk\Vault 2023\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")
    if (Test-Path $vdsSampleUtilitiesPath) {
        [System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2023\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")
    }
    else {
        [System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2022\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")
    }  
}
catch {
    Write-Host "Vault Client oder DataStandard wurde nicht gefunden!"
    $errCode = 9 #Vault Client oder DataStandard wurde nicht gefunden
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
if (!$CustomFile) {
    if (($Auftragsnummer.Length -eq 6 -or $Auftragsnummer.Length -eq 7) -and $Auftragsnummer -match '^\d+$') {
        $AuftragsTyp = "Auftrag"
    }
    elseif ($Auftragsnummer -match '[0-9]{2}[-]0[1-9]|1[0-2][-][0-9]{4}') {
        $AuftragsTyp = "Angebot"
    }
    elseif ($Auftragsnummer -match 'VP[-][0-9]{2}[-][0-9]{4}') {
        $AuftragsTyp = "Vorplanung"
    }
    else {
        $errCode = 6 #Invalide Auftrags bzw. Angebotsnummer
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }
}
# Vault Login

try {

    $AdskLicensing = "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_6.dll"
    if (!(Test-Path $AdskLicensing -PathType leaf)) {
        try {
            Copy-Item -Path "C:\Program Files\Autodesk\Vault Client 2023\Explorer\AdskLicensingSDK_6.dll" -Destination "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_6.dll"
        }
        catch {
            Write-Host "AdskLicensingSDK_6.dll wurde nicht gefunden!"
            $errCode = 8 #Fehlende AdskLicensingSDK_6.dll
            $downloadresult.Success = $false
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
    If (!$CustomFile) {
        $seachFile = $Auftragsnummer + "-AutoDeskTransfer.xml"

        if ($AuftragsTyp -eq "Auftrag") {
            $seachPath = "C:\Work\AUFTRÄGE NEU\\Konstruktion"
        }
        elseif ($AuftragsTyp -eq "Angebot" -or $AuftragsTyp -eq "Vorplanung" ) {
            $seachPath = "C:\Work\AUFTRÄGE NEU\Angebote"
        }
        else {
            $seachPath = "C:\Work\AUFTRÄGE NEU\"
        }
    }
    else {
        $seachFile = $Auftragsnummer
        $seachPath = "C:\Work\"
    }
    

    $sourceFile = Get-ChildItem -Path $seachPath -Recurse -Include $seachFile -Attributes a
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


        #Dateinamen der benötigten Dateien
        $undoFiles = @()
        $undoFiles += $vaultPathAutodesktransferXml + "/" + $FileStatus["FileName"]
        if (!$CustomFile) {
            $BerechnungenPath = $vaultPathAutodesktransferXml + "/Berechnungen/"
            $undoFiles += $vaultPathAutodesktransferXml + "/" + $Auftragsnummer + "-Spezifikation.pdf"
            $undoFiles += $vaultPathAutodesktransferXml + "/" + $Auftragsnummer + "-LiftHistory.json"
            $undoFiles += $BerechnungenPath + $Auftragsnummer + ".html"
            $undoFiles += $BerechnungenPath + $Auftragsnummer + ".aus"
            $undoFiles += $BerechnungenPath + $Auftragsnummer + ".dat"
            $undoFiles += $BerechnungenPath + $Auftragsnummer + ".LILO"
        }

        $vaultFoundUndoFiles = $vault.DocumentService.FindLatestFilesByPaths($undoFiles)

        $downloadTicket = New-Object Autodesk.Connectivity.WebServices.ByteArray
        
        foreach ($vaultFoundUndoFile in $vaultFoundUndoFiles) {
            
            if ($vaultFoundUndoFile.Id -gt 0 -and $vaultFoundUndoFile.CheckedOut) {
                #FileStatus auslesen 
                $undoFileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
                $undoFileStatus = $VltHelpers.GetVaultFileStatus($connection, $vaultFoundUndoFile.CkOutSpec)

                if ($undoFileStatus["CheckOutState"] -eq "CheckedOutByCurrentUser" ) {
                    $vault.DocumentService.UndoCheckoutFile($vaultFoundUndoFile.MasterId, [ref]$downloadTicket)
                }
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