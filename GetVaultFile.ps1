<# .SYNOPSIS
     Download von Dateien in den Workspace 
.DESCRIPTION
     Download der AutoDeskTransfer.Xml in den lokalen Workspace  
.NOTES
     File Name : GetVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version Version 0.88 – new ErrorCode 11 CheckedOutLinkedFilesByOtherUser
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > GetVaultFile.ps1 8951234 $true
                                                        (Auftragsnummer)(ReadOnly)
.INPUTTYPE
     [String]Auftragsnummer
     [bool]ReadOnly  
.RETURNVALUE
     $downloadresult
     $errCode
.COMPONENT
     Vault Server
#>
      
Param(
    [Parameter(Mandatory = $true)]          
    [String]$Auftragsnummer,
    [bool]$ReadOnly = $false
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
        Write-Host "LogOut successful:" $logOff
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
if ($ReadOnly) {
    try {
        $serverName = "192.168.0.1"
        $vaultName = "vault"
        $vaultUser = "BE-Automation"
        $vaultPw = "BE-Automation"
        $authenticationFlags = [Autodesk.DataManagement.Client.Framework.Vault.Currency.Connections.AuthenticationFlags]::ReadOnly
        $vdfResults = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogIn($serverName, $vaultName, $vaultUser, $vaultPw, $authenticationFlags, $null)
        
        if ($vdfResults.Success) {
            $connection = $vdfResults.Connection;
        }
    }
    catch {
        $errCode = 2 #Login fehlgeschlagen
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }
}
else {
    try {

        $AdskLicensing = "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_5.dll"
        if (!(Test-Path $AdskLicensing -PathType leaf)) {
            try {
                Copy-Item -Path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\AdskLicensingSDK_5.dll" -Destination "C:\Windows\System32\WindowsPowerShell\v1.0\AdskLicensingSDK_5.dll"
            }
            catch {
                Write-Host "AdskLicensingSDK_5.dll wurde nicht gefunden!"
                $errCode = 8 #Fehlende AdskLicensingSDK_5.dll
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
}

try {
    
    $VltHelpers = New-Object VdsSampleUtilities.VltHelpers

    #Dateinamen der benötigten Dateien
    $downloadFiles = @()
    $downloadFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
    $downloadFiles += $Auftragsnummer + "-Spezifikation.pdf"
    $downloadFiles += $Auftragsnummer + ".html"
    $downloadFiles += $Auftragsnummer + ".aus"
    $downloadFiles += $Auftragsnummer + ".dat"
    $downloadFiles += $Auftragsnummer + ".LILO"

    #Quellpfad ermitteln
    if ($ReadOnly) {
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

        $seachFile = $Auftragsnummer + "-AutoDeskTransfer.xml"
    
        $foundFiles = Get-ChildItem -Path $seachPath -Recurse -Include $seachFile

        if ($foundFiles.Count -eq 1) {
            
            If ($foundFiles[0].IsReadOnly -eq $false) {
                #FileStatus auslesen 
                $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
                $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $foundFiles[0].FullName) 
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
        }
        elseif ($foundFiles.Count -gt 1) {
            foreach ($item in $foundFiles) {

                If (!$item.IsReadOnly) {
                    Write-Host "AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden."-ForegroundColor DarkRed
                    $errCode = "5"# AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden
                    $downloadresult.Success = $false
                    $downloadresult.IsCheckOut = $null
                    LogOut($downloadresult)
                }
            }
        }
    }
    else {
        $SearchCriteria = New-Object 'system.collections.generic.dictionary[string,string]'
        $SearchCriteria.Add("Name", "")
        $SearchCriteria["Name"] = $downloadFiles[0]
        $ADTFile = $VltHelpers.GetFileBySearchCriteria($connection, $SearchCriteria, $true, $false)

        #FileStatus auslesen 
        $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
        $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $ADTFile)

        if ($FileStatus["CheckOutState"] -eq "CheckedOutByOtherUser") {
            $ReadOnly = $true
            $errCode = 10
        }
    }

    #Dateien im Vault suchen (auschecken) und den Arbeitsbereich ermitteln
    $SearchCriteria = New-Object 'system.collections.generic.dictionary[string,string]'
    $SearchCriteria.Add("Name", "")

    $vaultFiles = @()
    for ($i = 0; $i -le $downloadFiles.Count - 1; $i++) {
        $SearchCriteria["Name"] = $downloadFiles[$i]
        $CheckOutFiles = !($ReadOnly)
        $vaultFile = $VltHelpers.GetFileBySearchCriteria($connection, $SearchCriteria, $true, $CheckOutFiles)
        
        switch ( $vaultFile ) {
            $null {
                Write-Host "Datei wurde im Vault nicht gefunden. Überprüfen Sie Ihre Eingabe!"-ForegroundColor DarkRed
                $downloadresult.Success = $false
                $errCode = 7 # Datei in Vault nicht gefunden
                LogOut($downloadresult) 
            }
            "File not found" {
                if ($downloadFiles[$i] -match "-AutoDeskTransfer.xml") {
                    Write-Host "Datei wurde im Vault nicht gefunden. Überprüfen Sie Ihre Eingabe!"-ForegroundColor DarkRed
                    $downloadresult.Success = $false
                    $errCode = 7 # Datei in Vault nicht gefunden
                    LogOut($downloadresult) 
                }
            }
            "CheckOut failed" {
                if ($downloadFiles[$i] -match "-AutoDeskTransfer.xml") {
                    $vaultFile = $VltHelpers.GetFileBySearchCriteria($connection, $SearchCriteria, $true, $false)
                    if (($null -ne $vaultFile) -and ($vaultFile -ne "File not found")) {
                        $vaultFiles += $vaultFile
                    }
                    else {
                        Write-Host "Datei wurde im Vault nicht gefunden. Überprüfen Sie Ihre Eingabe!"-ForegroundColor DarkRed
                        $downloadresult.Success = $false
                        $errCode = 7 # Datei in Vault nicht gefunden
                        LogOut($downloadresult)
                    }
                }
                elseif ($ReadOnly -eq $false) {
                    $SearchCriteriaLinkFile = New-Object 'system.collections.generic.dictionary[string,string]'
                    $SearchCriteriaLinkFile.Add("Name", "")
                    $SearchCriteriaLinkFile["Name"] = $downloadFiles[$i]
                    $linkedVaultFile = $VltHelpers.GetFileBySearchCriteria($connection, $SearchCriteriaLinkFile, $true, $false)
                    $LinkedFileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
                    $LinkedFileStatus = $VltHelpers.GetVaultFileStatus($connection, $linkedVaultFile)

                    if ($LinkedFileStatus["CheckOutState"] -eq "CheckedOutByOtherUser") {
                        Write-Host "AutoDeskTransferXml verbundene Dateien durch anderen Benutzer ausgechecked."-ForegroundColor DarkRed
                        $downloadresult.FileName = $LinkedFileStatus["FileName"]
                        $downloadresult.FullFileName = $LinkedFileStatus["FullFileName"]
                        $downloadresult.CheckOutState = $LinkedFileStatus["CheckOutState"]
                        $downloadresult.IsCheckOut = [System.Convert]::ToBoolean($LinkedFileStatus["CheckOut"])
                        $downloadresult.CheckOutPC = $LinkedFileStatus["CheckOutPC"]
                        $downloadresult.EditedBy = $LinkedFileStatus["EditedBy"]
                        $downloadresult.ErrorState = $LinkedFileStatus["ErrorState"]
                        $errCode = 11 # Xml verbundene Dateien durch anderen Benutzer ausgechecked.
                        $downloadresult.Success = $false
                        LogOut($downloadresult)
                    }
                }
            }
            default {
                $vaultFiles += $vaultFile 
            }          
        }
    }

    if ($vaultFiles.count -gt 0) {
        $WorkFolderPath = $vaultFiles[0] -replace $downloadFiles[0], ""
    }
    #Verzeichnissstruktur anlegen
    if ($WorkFolderPath.StartsWith("C:\Work\AUFTRÄGE NEU")) {
        if (!(Test-Path $WorkFolderPath"Berechnungen")) { New-Item -Path $WorkFolderPath"Berechnungen" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Berechnungen/PDF")) { New-Item -Path $WorkFolderPath"Berechnungen/PDF" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Bestellungen")) { New-Item -Path $WorkFolderPath"Bestellungen" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Bgr00")) { New-Item -Path $WorkFolderPath"Bgr00" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Bgr00/CAD-CFP")) { New-Item -Path $WorkFolderPath"Bgr00/CAD-CFP" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Fotos")) { New-Item -Path $WorkFolderPath"Fotos" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"SV")) { New-Item -Path $WorkFolderPath"SV" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Montage-TÜV-Dokumentation")) { New-Item -Path $WorkFolderPath"Montage-TÜV-Dokumentation" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Montage-TÜV-Dokumentation/TÜV")) { New-Item -Path $WorkFolderPath"Montage-TÜV-Dokumentation/TÜV" -ItemType Directory }
        if (!(Test-Path $WorkFolderPath"Montage-TÜV-Dokumentation/TÜV/Zertifikate")) { New-Item -Path $WorkFolderPath"Montage-TÜV-Dokumentation/TÜV/Zertifikate" -ItemType Directory }
    }
    #FileStatus auslesen 
    $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
    $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $vaultFiles[0]) 
    
    $downloadresult.Success = $true
    $downloadresult.FileName = $FileStatus["FileName"]
    $downloadresult.FullFileName = $FileStatus["FullFileName"]
    $downloadresult.CheckOutState = $FileStatus["CheckOutState"]
    $downloadresult.IsCheckOut = [System.Convert]::ToBoolean($FileStatus["CheckOut"])
    $downloadresult.CheckOutPC = $FileStatus["CheckOutPC"]
    $downloadresult.EditedBy = $FileStatus["EditedBy"]
    $downloadresult.ErrorState = $FileStatus["ErrorState"]

    if (((!$ReadOnly) -and ($downloadresult.CheckOutState -eq "CheckedOutByOtherUser")) -or ($errCode -eq 10)) {
        $errCode = 10 # Datei wurde von anderem User ausgechecked
        LogOut($downloadresult)
    }

    $errCode = 0
    LogOut($downloadresult)
}
catch {
    $downloadresult.Success = $false
    $errCode = 1 # Datei download ist fehlgeschlagen
    LogOut($downloadresult)
}