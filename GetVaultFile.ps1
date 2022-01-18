<# .SYNOPSIS
     Download von Dateien in den Workspace 
.DESCRIPTION
     Download der AutoDeskTransfer.Xml in den lokalen Workspace  
.NOTES
     File Name : GetVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version Version 0.8 – ReadOnly und VDF Login
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > GetVaultFile.ps1 8951234 $true
                                                        (Auftragsnummer)(ReadOnly)
.INPUTTYPE
     [String]Auftragsnummer
     [bool]ReadOnly  
.RETURNVALUE
     $errCode
.COMPONENT
     Vault Server
#>
      
Param(
    [Parameter(Mandatory = $true)]          
    [String]$Auftragsnummer,
    [bool]$ReadOnly = $false       
)
   
# Vault Login

try {
    Import-Module powerVault
    Initialize-VDF
    #Create-LogRepository
    #Get-LogRepository
    Get-VaultInstallationDirectory

    $vaultUser = "BE-Automation"
    $vaultPw = "BE-Automation"

    Open-VaultConnection -Password $vaultpW -Server 192.168.0.1:8080 -User $vaultUser -Vault vault
}
catch {
    $errCode = 2 #Login fehlgeschlagen
    $Host.SetShouldExit([int]$errCode)
    exit
}

try {
    #Dateinamen der benötigten Dateien
    $downloadFiles = @()
    $downloadFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
    $downloadFiles += $Auftragsnummer + ".html"
    $downloadFiles += $Auftragsnummer + ".aus"
    $downloadFiles += $Auftragsnummer + ".dat"
    $downloadFiles += $Auftragsnummer + ".LILO"

    #Dateien im Vault suchen und den Arbeitsbereich ermitteln 
    $vaultFiles = @()
    for ($i = 0; $i -le $downloadFiles.Count - 1; $i++) {

        $vaultFiles += Get-VaultFiles -FileName $downloadFiles[$i]
    }


    if ($vaultFiles.count -gt 0) {
        $WorkFolderPath = $vaultFiles[0].Path.TrimStart("$") -replace "/AUFTRÄGE", "C:/Work/AUFTRÄGE"
    }
    #Verzeichnissstruktur anlegen
    if (!(Test-Path $WorkFolderPath"/Berechnungen")) { New-Item -Path $WorkFolderPath"/Berechnungen" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Berechnungen/PDF")) { New-Item -Path $WorkFolderPath"/Berechnungen/PDF" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Bestellungen")) { New-Item -Path $WorkFolderPath"/Bestellungen" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Bgr00")) { New-Item -Path $WorkFolderPath"/Bgr00" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Bgr00/CAD-CFP")) { New-Item -Path $WorkFolderPath"/Bgr00/CAD-CFP" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Fotos")) { New-Item -Path $WorkFolderPath"/Fotos" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/SV")) { New-Item -Path $WorkFolderPath"/SV" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Montage-TÜV-Dokumentation")) { New-Item -Path $WorkFolderPath"/Montage-TÜV-Dokumentation" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Montage-TÜV-Dokumentation/TÜV")) { New-Item -Path $WorkFolderPath"/Montage-TÜV-Dokumentation/TÜV" -ItemType Directory }
    if (!(Test-Path $WorkFolderPath"/Montage-TÜV-Dokumentation/TÜV/Zertifikate")) { New-Item -Path $WorkFolderPath"/Montage-TÜV-Dokumentation/TÜV/Zertifikate" -ItemType Directory }

    #Dateien werden ausgechecked
    $downloadTicket = New-Object Autodesk.Connectivity.WebServices.ByteArray

    for ($i = 0; $i -le $vaultFiles.Count - 1; $i++) {

        If ($vaultFiles[$i].IsCheckedOut) {
            Write-Host "Datei"$vaultFiles[$i]._Name"ist bereits ausgechecked!"  
        }
        Else {
            $checkedOutFile = $vault.DocumentService.CheckoutFile($vaultFiles[$i].Id, "Master", $env:COMPUTERNAME, $WorkFolderPath, "Für automatsische Bearbeitung abgerufen.", [ref]$downloadTicket)
            Write-Host "Datei"$vaultFiles[$i]._Name"wurde ausgechecked!"   
        }
    }

    #Dateien werden in den Arbeitsbereich geladen
    if ($vaultFiles.count -gt 0) {
        for ($i = 0; $i -le $vaultFiles.Count - 1; $i++) {
            $downloadFile = Save-VaultFile -File $vaultFiles[$i].'Full Path'
            $downloadFullFileName = $vaultFiles[$i]._FullPath.TrimStart("$") -replace "/AUFTRÄGE", "C:/Work/AUFTRÄGE"
            Set-ItemProperty $downloadFullFileName -Name IsReadOnly -Value $false
            Write-Host "Datei"$vaultFiles[$i]._Name"wurde heruntergeladen!"
        }
    }
    Else {
        Write-Host "Auftrag $Auftragsnummer nicht gefunden! Prüfen Sie ob im Vault bereits ein Auftrag angelegt worden ist!!"
        $vault.Dispose() #Vault Connection schließen

        $errCode = 1 # Datei download ist fehlgeschlagen

        $Host.SetShouldExit([int]$errCode)
        exit
    }

    #Read-Host Debug

    $vault.Dispose() #Vault Connection schließen

    $errCode = 0

    $Host.SetShouldExit([int]$errCode)
    exit
}
catch {
    $vault.Dispose() #Vault Connection schließen

    $errCode = 1 # Datei download ist fehlgeschlagen

    $Host.SetShouldExit([int]$errCode)
    exit

}