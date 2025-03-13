<# .SYNOPSIS
     Upload von Dateien aus dem Workspace zum Server
.DESCRIPTION
     Upload der AutoDeskTransfer.Xml in den lokalen Workspace  
.NOTES
     File Name : SetVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
       Version 1.25 – add ZUES-Category
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > SetVaultFile.ps1 -Auftragsnummer 8951234 $true
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
    $pdfSharpPath = "C:\Work\Administration\PowerShellScripts\PdfSharp\PdfSharp.dll"
    if (Test-Path $pdfSharpPath) {
        Add-Type -path "C:\Work\Administration\PowerShellScripts\PdfSharp\PdfSharp.dll"

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
        Write-Host "Vaultverbindungsstatus:" $logOff
    }
    $Host.SetShouldExit([int]$errCode)
    exit
}

function Remove-EmptyFolders([string]$folders) {
    Get-Childitem $folders -Recurse | Where-Object { $_.PSIsContainer -and !(Get-Childitem $_.Fullname -Recurse | 
            Where-Object { !$_.PSIsContainer }) } | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
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
        $errCode = 7 # Datei im Arbeitsbereich nicht gefunden
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }
    if ($sourceFile.Count -gt 1) {
        Write-Host "AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden."-ForegroundColor DarkRed
        $errCode = 5 # AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden.
        $downloadresult.Success = $false
        LogOut($downloadresult)
    }

    $VltHelpers = New-Object VdsSampleUtilities.VltHelpers
    $vault = $connection.WebServiceManager

    #FileStatus auslesen 
    $FileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
    $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $sourceFile) 
    
    $downloadresult.FileName = $FileStatus["FileName"]
    $downloadresult.FullFileName = $FileStatus["FullFileName"]
    $downloadresult.CheckOutState = $FileStatus["CheckOutState"]
    $downloadresult.IsCheckOut = [System.Convert]::ToBoolean($FileStatus["CheckOut"])
    $downloadresult.CheckOutPC = $FileStatus["CheckOutPC"]
    $downloadresult.EditedBy = $FileStatus["EditedBy"]
    $downloadresult.ErrorState = $FileStatus["ErrorState"]

    $sourcePath = $sourceFile.DirectoryName.Replace("\", "/") + "/"
    $targetPath = $VltHelpers.ConvertLocalPathToVaultPath($connection, $sourceFile)
    #Dateinamen der einzucheckenden

    $uploadFiles = @()
    if (!$CustomFile) {
        $uploadFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
        $pathExtBerechnungen = "Berechnungen/"
        $pathExtBerechnungenPDF = "Berechnungen/PDF/"
        $pathExtCAD = "Bgr00/CAD-CFP/"
        $pathExtTUEVZertifikate = "Montage-TÜV-Dokumentation/TÜV/Zertifikate/"
        $pathExtSV = "SV/"
        if (Test-Path ($sourcePath + $Auftragsnummer + "-Spezifikation.pdf")) { $uploadFiles += $Auftragsnummer + "-Spezifikation.pdf" }
        if (Test-Path ($sourcePath + $Auftragsnummer + "-LiftHistory.json")) { $uploadFiles += $Auftragsnummer + "-LiftHistory.json" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".html")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + ".html" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".aus")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + ".aus" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + ".dat" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".LILO")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + ".LILO" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Jupiter.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-Jupiter.txt" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Pluto.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-Pluto.txt" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Beripac.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-Beripac.txt" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Pluto-Seil.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-Pluto-Seil.txt" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-ZZE-S.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-ZZE-S.txt" }
        if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-G.txt")) { $uploadFiles += $pathExtBerechnungen + $Auftragsnummer + "-G.txt" }
    }
    else {
        $uploadFiles += $Auftragsnummer
    }

    #FileCheck CheckedOutLinkedFilesByOtherUser
    foreach ($uploadFile in $uploadFiles) {
        #FileStatus auslesen
        $LinkedFile = Get-ChildItem -Path ($sourcePath + $uploadFile)
        $LinkedFileStatus = New-Object 'system.collections.generic.dictionary[string,string]'
        $LinkedFileStatus = $VltHelpers.GetVaultFileStatus($connection, $LinkedFile )

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
    if (!$CustomFile) {
        if (Test-Path ($sourcePath + $pathExtBerechnungen)) { $berechnungenFiles = Get-ChildItem -Path ($sourcePath + $pathExtBerechnungen) -Filter  *.pdf }
        if (Test-Path ($sourcePath + $pathExtBerechnungenPDF)) { $berechnungenPDFFiles = Get-ChildItem -Path ($sourcePath + $pathExtBerechnungenPDF) -Filter  *.pdf }
        if (Test-Path ($sourcePath + $pathExtCAD)) { $cadFiles = Get-ChildItem -Path ($sourcePath + $pathExtCAD) -Filter  *.dwg }
        if (Test-Path ($sourcePath + $pathExtTUEVZertifikate)) { $zertifikateFiles = Get-ChildItem -Path ($sourcePath + $pathExtTUEVZertifikate) -Filter  *.pdf }
        if (Test-Path ($sourcePath + $pathExtSV)) { $correspondenceFiles = Get-ChildItem -Path ($sourcePath + $pathExtSV) -Filter  *.pdf }

        foreach ($berechnung in $berechnungenFiles) {
            $uploadFiles += $pathExtBerechnungen + $berechnung
        }

        foreach ($berechnungensPDF in $berechnungenPDFFiles) {
            $uploadFiles += $pathExtBerechnungenPDF + $berechnungensPDF
        }

        foreach ($cadFile in $cadFiles) {
            $uploadFiles += $pathExtCAD + $cadFile
        }

        foreach ($zertifikateFile in $zertifikateFiles) {
            $uploadFiles += $pathExtTUEVZertifikate + $zertifikateFile
        }

        foreach ($correspondenceFile in $correspondenceFiles) {

            if ($correspondenceFile.Name.StartsWith($Auftragsnummer)){
                $uploadFiles += $pathExtSV + $correspondenceFile
            }
        }

        #Prüfen ob Verzeichnisstruktur im Vault vorhanden ist
        $vaultPaths = @()
        $vaultPaths += ($targetPath + "/" + $pathExtBerechnungen).TrimEnd("/")
        if ($berechnungenPDFFiles.Count -gt 0) {
            $vaultPaths += ($targetPath + "/" + $pathExtBerechnungenPDF).TrimEnd("/")
        }
        if ($cadFiles.Count -gt 0) {
            $vaultPaths += ($targetPath + "/" + $pathExtCAD).TrimEnd("/CAD-CFP/")
            $vaultPaths += ($targetPath + "/" + $pathExtCAD).TrimEnd("/")
        }
        if ($zertifikateFiles.Count -gt 0) {
            $vaultPaths += ($targetPath + "/" + $pathExtTUEVZertifikate).TrimEnd("/TÜV/Zertifikate/")
            $vaultPaths += ($targetPath + "/" + $pathExtTUEVZertifikate).TrimEnd("/Zertifikate/")
            $vaultPaths += ($targetPath + "/" + $pathExtTUEVZertifikate).TrimEnd("/")
        }

        foreach ($vaultPath in $vaultPaths) {
            $mFolder = $vault.DocumentService.FindFoldersByPaths($vaultPath)[0]

            if ($mFolder.Id -eq -1) {
                try {
                    $mFolderName = $vaultPath.Split("/")[-1]
                    $mFolderparentId = $vault.DocumentService.FindFoldersByPaths(($vaultPath).TrimEnd("/" + $mFolderName))[0].Id
                    $vault.DocumentService.AddFolder($mFolderName, $mFolderparentId, $false)
                }
                catch {
                    Write-Host  $vaultPath " konnte nicht erstellt werden"-ForegroundColor DarkRed
                }
            }
        }

        #Prüfen ob veraltete Berechnungs Daten vorhanden sind
        if ($berechnungenPDFFiles -match 'Anlagedaten' -or 
            $berechnungenPDFFiles -match 'Lift data' -or 
            $berechnungenPDFFiles -match 'Données techniques de l´installation') {
        
            #Daten im Vault löschen
            $toDeleteVaultFiles = @()

            $propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
            $custPropDefIds = $propDefs | Where-Object { $_.IsSys -eq $false } | Select-Object -ExpandProperty Id

            if ($berechnungenPDFFiles.Count -gt 3) {
                $vaultPathBerechnungenPDF = ($targetPath + "/" + $pathExtBerechnungenPDF).TrimEnd("/")
                $vaultFolderBerechnungenPDF = $vault.DocumentService.GetFolderByPath($vaultPathBerechnungenPDF)
                $files = $vault.DocumentService.GetLatestFilesByFolderId($vaultFolderBerechnungenPDF.Id, $true)
                foreach ($file in $files) {
                    if (($file.Cat.CatName -eq "Office" -or $file.Cat.CatName -eq "ZÜS-Unterlagen") -and $file.Name.EndsWith(".pdf")) {

                        $props = $vault.PropertyService.GetPropertiesByEntityIds("FILE", @($file.Id))
                        $custProps = $props | Where-Object { $custPropDefIds -contains $_.PropDefId }

                        if ((($custProps | Where-Object { $_.PropDefId -eq 26 }).Val -eq "Berechnungen") -and (($custProps | Where-Object { $_.PropDefId -eq 104 }).Val.StartsWith("CFP"))) {
                            $toDeleteVaultFiles += $file
                        }  
                    }
                } 
            }

            if ($zertifikateFiles.Count -gt 0) {
                $vaultPathTUEVZertifikate = ($targetPath + "/" + $pathExtTUEVZertifikate).TrimEnd("/")
                $vaultFolderTUEVZertifikate = $vault.DocumentService.GetFolderByPath($vaultPathTUEVZertifikate)
                $files = $vault.DocumentService.GetLatestFilesByFolderId($vaultFolderTUEVZertifikate.Id, $true)
                foreach ($file in $files) {
                    if (($file.Cat.CatName -eq "Office" -or $file.Cat.CatName -eq "ZÜS-Unterlagen") -and $file.Name.EndsWith(".pdf")) {

                        $props = $vault.PropertyService.GetPropertiesByEntityIds("FILE", @($file.Id))
                        $custProps = $props | Where-Object { $custPropDefIds -contains $_.PropDefId }

                        if (($custProps | Where-Object { $_.PropDefId -eq 104 }).Val.StartsWith("CFP")) {
                            $toDeleteVaultFiles += $file
                        }  
                    }
                } 
            }
            foreach ($toDeleteVaultFile in $toDeleteVaultFiles) {
                try {
                    $toDeleteFolder = $vault.DocumentService.GetFoldersByFileMasterId($toDeleteVaultFile.MasterId)
                    $vault.DocumentService.DeleteFileFromFolderUnconditional( $toDeleteVaultFile.MasterId , $toDeleteFolder[0].Id)
                    Write-Host  $toDeleteVaultFile.Name  "gelöscht..."-ForegroundColor Yellow
                }
                catch { 
                    Write-Host  $toDeleteVaultFile.Name "nicht gelöscht,keine Rechte zum Löschen..."-ForegroundColor DarkRed
                }
            }
        }

        #Prüfen ob veraltete CFP-DB-Modification vorhanden sind
        if ($berechnungenFiles -match 'DB-Anpassungen'){
            
            # delete outdated CFP-DB-Modification
            if ($berechnungenFiles.Count -gt 0) {

                $vaultPathBerechnungen = ($targetPath + "/" + $pathExtBerechnungen).TrimEnd("/")
                $vaultFolderBerechnungen = $vault.DocumentService.GetFolderByPath($vaultPathBerechnungen)
                $vaultCalculationsFiles = $vault.DocumentService.GetLatestFilesByFolderId($vaultFolderBerechnungen.Id, $true)

                if ($files.Count -gt 0) {
                    $oldCFPDBModifications =  $vaultCalculationsFiles.Where{$_.Name -match 'DB-Anpassungen' } 
                    $oldCFPDBModificationsNames =  $oldCFPDBModifications.Name
                    $newCFPDBModifications = $berechnungenFiles -match 'DB-Anpassungen'
                    if($null -ne $oldCFPDBModificationsNames){
                        $toDeleteCFPDBModifications = Compare-Object $newCFPDBModifications $oldCFPDBModificationsNames -PassThru | Where-Object{$_.Sideindicator -eq '=>'}
                        if ($toDeleteCFPDBModifications.Count -gt 0){
                            foreach ($vaultCalculationsFile in $vaultCalculationsFiles) {
                                if($toDeleteCFPDBModifications.Contains($vaultCalculationsFile.Name)){
                                    try {
                                        $toDeleteFolder = $vault.DocumentService.GetFoldersByFileMasterId($vaultCalculationsFile.MasterId)
                                        $vault.DocumentService.DeleteFileFromFolderUnconditional($vaultCalculationsFile.MasterId , $toDeleteFolder[0].Id)
                                        Write-Host  $vaultCalculationsFile.Name  "gelöscht..."-ForegroundColor Yellow
                                    }
                                    catch { 
                                    Write-Host  $vaultCalculationsFile.Name "nicht gelöscht,keine Rechte zum Löschen..."-ForegroundColor DarkRed
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    #Dateien hochladen und aktualisieren
    for ($i = 0; $i -le $uploadFiles.Count - 1; $i++) {
        $verfasser = $Env:USERNAME
        $uploadSource = -join ($sourcePath, $uploadFiles[$i])
        $uploadTarget = -join ($targetPath, "/", $uploadFiles[$i])
        $uploadTargetPath = ( -join (Split-Path -Path $uploadTarget, "\")).Replace("\", "/")

        If ($null -eq $newProps) {        
            $newProps = New-Object 'system.collections.generic.dictionary[string,string]'
        }
        else {
            $newProps.Clear()
        }

        $uploadFileResult = $VltHelpers.AddFile($connection, $uploadSource, $uploadTargetPath, $true)
    
        if ($uploadFileResult -and !$CustomFile) {
            $uploadFile = ($vault.DocumentService.FindLatestFilesByPaths($uploadTarget))[0]
            $Beschreibung = $uploadFile.Name.TrimStart($Auftragsnummer + "-")

            switch ([System.IO.Path]::GetExtension($uploadTarget)) {
                ".xml" {
                    $Kategorie = "Berechnungen"
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    if ($uploadFile.Cat.CatName -ne "AnlageDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 31, $uploadFile.Comm)
                    }
                }
                ".pdf" {
                    If ($uploadTargetPath -match "Berechnungen") {
                        $Kategorie = "Berechnungen"
                        $document = [PdfSharp.Pdf.IO.PdfReader]::Open($uploadSource)
                        if ($null -ne $document) {
                            $verfasser = $document.Info.Author
                        }
                    }
                    ElseIf ($uploadTargetPath -match "Zertifikate") {
                        $Kategorie = "Baumuster-Zertifikate"
                        $document = [PdfSharp.Pdf.IO.PdfReader]::Open($uploadSource)
                        if ($null -ne $document) {
                            $verfasser = $document.Info.Author
                        }
                    }
                    Else {
                        $Kategorie = "Berechnungen"
                        $newProps.Add('Kommentare', "Von Spezifikation automatisch generierte Datei")
                    }
                    
                    $newProps.Add('Beschreibung', $Beschreibung.Replace(".pdf",""))
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)

                    if ($verfasser -eq "CFP-TÜV") {
                        if ($uploadFile.Cat.CatName -ne "ZÜS-Unterlagen") {
                            $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 40, $uploadFile.Comm)
                        }
                    }
                    else{
                        if ($uploadFile.Cat.CatName -ne "Office") {
                            $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 3, $uploadFile.Comm)
                        }
                    }
                }
                ".json" {
                    $Kategorie = "Dokumentation"
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    if ($uploadFile.Cat.CatName -ne "AnlageDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 31, $uploadFile.Comm)
                    }
                }
                { ($_ -eq ".html") -or ($_ -eq ".aus") } {
                    $html = New-Object -ComObject "HTMLFile"
                    try {
                        $html.IHTMLDocument2_write((Get-Content ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".html") -raw))
                    }
                    catch {
                        try {
                            $src = [System.Text.Encoding]::Unicode.GetBytes((Get-Content ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".html") -raw))
                            $html.write($src)
                        }
                        catch {
                            Write-Host "Html konnte nicht gelesen werden."-ForegroundColor DarkRed 
                        }
                    }

                    if ($null -ne $html) {
                        $motortyp = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Motortyp*" -or $_.innerText -like "Motor type*" }).innerText
                        $infoAufhaengung = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Aufhängung*" -or $_.innerText -like "Suspension/roping*" }).innerText
                        if ($null -ne $infoAufhaengung) {
                            $aufhaengung = $infoAufhaengung.Replace("Aufhängung is ", "").Replace("Suspension/roping is ", "") 
                        }
                        $infoTreibscheibe = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Treibscheibe *" -or $_.innerText -like "Traction sheave*" })
                        if ($null -ne $infoTreibscheibe) {
                            if ($infoTreibscheibe.Length -gt 2) {
                                $lageTreibscheibe = $infoTreibscheibe.innerText[0] 
                                $Treibscheibe = $infoTreibscheibe.innerText[2] 
                            }
                            else {
                                try {
                                    $lageTreibscheibe = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Maschine*" -or $_.innerText -like "Machine*" }).innerText
                                }
                                catch {
                                    Write-Host "Lage Treibscheibe konnte nicht gelesen werden."-ForegroundColor DarkRed 
                                }
                                $Treibscheibe = $infoTreibscheibe.innerText[1] 
                            }
                        }
                    }
                    else {
                        $motortyp = "Keine Angaben"
                        $aufhaengung = "Keine Angaben"
                        $lageTreibscheibe = "Keine Angaben"
                        $treibscheibe = "Keine Angaben"
                    }
                    $Beschreibung = "Antriebsauslegung Ziehl Abegg";
                    $Kategorie = "Berechnungen"    
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    $newProps.Add('Antriebtyp', $motortyp)
                    $newProps.Add('Aufhängung', $aufhaengung)
                    $newProps.Add('Lage Antrieb', $lageTreibscheibe)
                    $newProps.Add('Treibscheibe Zylinder', $treibscheibe )
                    if ($uploadFile.Cat.CatName -ne "AntriebsDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 35, $uploadFile.Comm)
                    }
                }
                ".dat" { 
                    $Beschreibung = "Daten Bausatzprogram CFP"
                    $Kategorie = "Berechnungen"
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    if ($uploadFile.Cat.CatName -ne "AnlageDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 31, $uploadFile.Comm)
                    }
                }
                ".LILO" {
                    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")) {
                        
                        try {
                            $hydroDat = Get-Content -path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")

                            $motortyp = ($hydroDat -match "Power_Unit_Type").Replace("[Power_Unit_Type] ", "") + ($hydroDat -match "Valve_Model").Replace("[Valve_Model] ", " - ") + ($hydroDat -match "Pumpenbezeichnung").Replace("[Pumpenbezeichnung] ", "- ")
                            $aufhaengung = ($hydroDat -match "Bauart")[0].Replace("[Bauart] ", "")
                            $lageTreibscheibe = If (($hydroDat -match "Antrieb_im_Schacht").Replace("[Antrieb_im_Schacht] ", "") -eq "0") { "Antrieb im Maschinenraum" }else { "Antrieb im Schacht" }
                            $treibscheibe = ($hydroDat -match "Zylinderbezeichnung").Replace("[Zylinderbezeichnung] ", "")
                        }
                        catch {
                            $motortyp = "Keine Angaben"
                            $aufhaengung = "Keine Angaben"
                            $lageTreibscheibe = "Keine Angaben"
                            $treibscheibe = "Keine Angaben"
                        }

                    }
                    Else {
                        $motortyp = "Keine CFP-Auslegung vorhanden"
                        $aufhaengung = "Keine CFP-Auslegung vorhanden"
                        $lageTreibscheibe = "Keine CFP-Auslegung vorhanden"
                        $treibscheibe = "Keine CFP-Auslegung vorhanden"
                    }

                    $Beschreibung = "Antriebsauslegung Lilo"
                    $Kategorie = "Berechnungen"    
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    $newProps.Add('Antriebtyp', $motortyp)
                    $newProps.Add('Aufhängung', $aufhaengung)
                    $newProps.Add('Lage Antrieb', $lageTreibscheibe)
                    $newProps.Add('Treibscheibe Zylinder', $treibscheibe )
                    if ($uploadFile.Cat.CatName -ne "AntriebsDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 35, $uploadFile.Comm)
                    }
                
                }
                ".txt" {
                    $Beschreibung = "Fertigungsunterlagen CFP"
                    $Kategorie = "Berechnungen"
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    if ($uploadFile.Cat.CatName -ne "FertigungsDaten") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 32, $uploadFile.Comm)
                    }
                }
                ".dwg" {
                    $Beschreibung = "Bausatz Zeichnungen"
                    $Kategorie = "Montagebaugruppe"
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    $newProps.Add('Verfasser', $verfasser)
                    $newProps.Add('Kategorie', $Kategorie)
                    $newProps.Add('Kommentare', "von CFP automatisch generierte Zeichnung")
                    if ($uploadFile.Cat.CatName -ne "Zeichnungsableitungen") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 24, $uploadFile.Comm)
                    }
                }
                default {
                    $newProps.Add('Beschreibung', $Beschreibung)
                    $newProps.Add('Projekt', $Auftragsnummer)
                    if ($uploadFile.Cat.CatName -ne "Basis") {
                        $vault.DocumentServiceExtensions.UpdateFileCategories($uploadFile.MasterId, 1, $uploadFile.Comm)
                    }
                }
            }

            $PropertyUpdateResult = $VltHelpers.mUpdateFileProperties2($connection, $uploadFile, $newProps)

            if (!$PropertyUpdateResult) { Write-Host "Eigenschaften"$uploadFiles[$i]"konnten nicht aktualisiert werden!"-ForegroundColor DarkRed }

            Write-Host "Datei"$uploadFiles[$i]"wurde hochgeladen und eingechecked!"-ForegroundColor Yellow
        }
        else {
            Write-Host "Datei"$uploadFiles[$i]"konnte nicht hochgeladen werden!"-ForegroundColor DarkRed
        }
    }
}
catch {
    $errCode = 1 # Datei upload ist fehlgeschlagen
    $downloadresult.Success = $false
    LogOut($downloadresult)
}

try {
    if (!$CustomFile) {
        #XML-Datei ermitteln und auslesen
        $pfadxml = $sourceFile.FullName

        $xml = [XML] (Get-Content -Path $pfadxml -Encoding UTF8)

        $parameter = $xml.selectNodes("//ParamWithValue")

        $var_FabrikNummer = $parameter | Where-Object { $_.name -eq "var_FabrikNummer" }
        $var_Kennwort = $parameter | Where-Object { $_.name -eq "var_Kennwort" }
        $var_Projekt = $parameter | Where-Object { $_.name -eq "var_Projekt" }                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                        
        $var_Betreiber = $parameter | Where-Object { $_.name -eq "var_Betreiber" }                                                                                                                                                  
        $var_Q = $parameter | Where-Object { $_.name -eq "var_Q" }
        $var_F = $parameter | Where-Object { $_.name -eq "var_F" }
        $var_Personen = $parameter | Where-Object { $_.name -eq "var_Personen" }
        $var_v = $parameter | Where-Object { $_.name -eq "var_v" } 
        $var_FH = $parameter | Where-Object { $_.name -eq "var_FH" }
        $var_SB = $parameter | Where-Object { $_.name -eq "var_SB" }
        $var_ST = $parameter | Where-Object { $_.name -eq "var_ST" }
        $var_SG = $parameter | Where-Object { $_.name -eq "var_SG" }
        $var_SK = $parameter | Where-Object { $_.name -eq "var_SK" }
        $var_KBI = $parameter | Where-Object { $_.name -eq "var_KBI" }
        $var_KTI = $parameter | Where-Object { $_.name -eq "var_KTI" }
        $var_KHLicht = $parameter | Where-Object { $_.name -eq "var_KHLicht" }
        $var_A_Kabine = $parameter | Where-Object { $_.name -eq "var_A_Kabine" }
        $var_Kommentare = $parameter | Where-Object { $_.name -eq "var_Kommentare" }

        #Ordnereigenschaften ermitteln und auslesen

        $folder = $vault.DocumentService.GetFolderByPath($targetPath)
        $propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")

        $folderProps = $vault.PropertyService.GetPropertiesByEntityIds("FLDR", @($folder.Id))

        $udpIds = $propDefs | Where-Object { $_.IsSys -eq $false } | Select-Object -ExpandProperty Id
        $folderProps = $folderProps | Where-Object { $_.Propdefid -in $udpIds }

        $fabriknummer = $folderProps | Where-Object { $_.PropDefId -eq "124" }
        $projektTitel = $folderProps | Where-Object { $_.PropDefId -eq "27" }
        $aufstellungsort = $folderProps | Where-Object { $_.PropDefId -eq "145" }
        $betreiber = $folderProps | Where-Object { $_.PropDefId -eq "144" }
        $nutzlast = $folderProps | Where-Object { $_.PropDefId -eq "132" }
        $fahrkorbgewicht = $folderProps | Where-Object { $_.PropDefId -eq "133" }
        $nenngeschwingigkeit = $folderProps | Where-Object { $_.PropDefId -eq "131" }
        $personenAnzahl = $folderProps | Where-Object { $_.PropDefId -eq "139" }
        $schachtbreite = $folderProps | Where-Object { $_.PropDefId -eq "137" }
        $schachttiefe = $folderProps | Where-Object { $_.PropDefId -eq "138" }
        $schachtgrube = $folderProps | Where-Object { $_.PropDefId -eq "135" }
        $schachtkopf = $folderProps | Where-Object { $_.PropDefId -eq "136" }
        $foerderhoehe = $folderProps | Where-Object { $_.PropDefId -eq "134" }
        $kabinenbreite = $folderProps | Where-Object { $_.PropDefId -eq "140" }
        $kabinentiefe = $folderProps | Where-Object { $_.PropDefId -eq "141" }
        $kabinenhoehe = $folderProps | Where-Object { $_.PropDefId -eq "142" }
        $kabinenflaeche = $folderProps | Where-Object { $_.PropDefId -eq "143" }
        $kommentare = $folderProps | Where-Object { $_.PropDefId -eq "24" }

        #Föderhöhe in Millimeter umwandeln
        if (-Not [string]::IsNullOrWhiteSpace($var_FH.value)) {
            $FHmm = [System.Convert]::ToDecimal($var_FH.value, [cultureinfo]::GetCultureInfo('de-DE')) * 1000
            $var_FH.value = $FHmm.tostring()
        }

        #Ordnereigenschaften schreiben und übermitteln

        $fabriknummer.Val = $var_FabrikNummer.value
        $projektTitel.Val = $var_Kennwort.value
        $aufstellungsort.Val = $var_Projekt.value
        $betreiber.Val = $var_Betreiber.value
        $nutzlast.Val = $var_Q.value
        $fahrkorbgewicht.Val = $var_F.value
        $nenngeschwingigkeit.Val = $var_v.value
        $personenAnzahl.Val = $var_Personen.value
        $schachtbreite.Val = $var_SB.value
        $schachttiefe.Val = $var_ST.value
        $schachtgrube.Val = $var_SG.value
        $schachtkopf.Val = $var_SK.value
        $foerderhoehe.Val = $var_FH.value
        $kabinenbreite.Val = $var_KBI.value
        $kabinentiefe.Val = $var_KTI.value
        $kabinenhoehe.Val = $var_KHLicht.value
        $kommentare.Val = $var_Kommentare.value
        if (-Not [string]::IsNullOrWhitespace($var_A_Kabine.value)) {
            $kabinenflaeche.Val = $var_A_Kabine.value
        }

        $propValues = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
        $propValues.Items = New-Object Autodesk.Connectivity.WebServices.PropInstParam[] $folderProps.Count
        $i = 0
        foreach ($d in $folderProps.GetEnumerator()) {
            if ($d.ValTyp.ToString() -eq "Numeric" -and [string]::IsNullOrWhiteSpace($d.Val)){
                $d.Val = $null   
            }
            $propValues.Items[$i] = New-Object Autodesk.Connectivity.WebServices.PropInstParam -Property @{PropDefId = $d.PropDefId; Val = $d.Val }
            $i++
        }

        $vault.DocumentServiceExtensions.UpdateFolderProperties(@($folder.Id), @($propValues))
  
        #Housekeeping

        $workPathBerechnungenPDF = $sourcePath + $pathExtBerechnungenPDF
        $workPathTUEVZertifikate = $sourcePath + $pathExtTUEVZertifikate
        $workPathBerechnungen = $sourcePath + $pathExtBerechnungen
        $workPathCAD = $sourcePath + $pathExtCAD
        $workPathSV = $sourcePath + $pathExtSV

        If ($sourcePath -match "C:/Work/AUFTRÄGE NEU"){
            if (Test-Path ($workPathBerechnungenPDF)) { Remove-Item -Path $workPathBerechnungenPDF -Recurse -Force }
            if (Test-Path ($workPathTUEVZertifikate)) { Remove-Item -Path $workPathTUEVZertifikate -Recurse -Force }
            if (Test-Path ($workPathBerechnungen)) { Remove-Item -Path $workPathBerechnungen  -Recurse -Force }
            if (Test-Path ($workPathCAD)) { Remove-Item -Path $workPathCAD -Recurse -Force }
            if (Test-Path ($workPathSV)) { Remove-Item -Path $workPathSV  -Recurse -Force }
        }

        $deleteFiles = @()
        $deleteFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
        if (Test-Path ($sourcePath + $Auftragsnummer + "-Spezifikation.pdf")) { $deleteFiles += $Auftragsnummer + "-Spezifikation.pdf" }
        if (Test-Path ($sourcePath + $Auftragsnummer + "-LiftHistory.json")) { $deleteFiles += $Auftragsnummer + "-LiftHistory.json" }
        foreach ($deleteFile in $deleteFiles) {
            try {
                $pathDeleteFile = $sourcePath + $deleteFile
                Remove-Item $pathDeleteFile -Force
            }
            catch {
                # TODO Ausgabe Fehlermeldung
            }
        }
    }
    # leere Ordner löschen
    If (($sourcePath -match "C:/Work/AUFTRÄGE NEU")) {
        Remove-EmptyFolders $sourcePath.Replace("/" + $Auftragsnummer + "/","")
    }
    #FileStatus auslesen 
    $FileStatus = $VltHelpers.GetVaultFileStatus($connection, $sourceFile) 
    
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
    $errCode = 3 # Eigenschaftenabgleich Vault ist fehlgeschlagen
    $downloadresult.Success = $false
    LogOut($downloadresult)
}