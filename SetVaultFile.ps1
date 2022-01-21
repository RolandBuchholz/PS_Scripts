<# .SYNOPSIS
     Upload von Dateien aus dem Workspace zum Server
.DESCRIPTION
     Upload der AutoDeskTransfer.Xml in den lokalen Workspace  
.NOTES
     File Name : SetVaultFile.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version 0.8 – Dateibereinigung angepasst
.EXAMPLE
     Beispiel wie das Script aufgerufen wird > SetVaultFile.ps1 -Auftragsnummer „8951234“
.INPUTTYPE
     Auftragsnummer 
.RETURNVALUE
     Void
.COMPONENT
     Vault Server
#>

       
Param(            
    [String]$Auftragsnummer           
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
    $errCode = "2" #Login fehlgeschlagen
    $Host.SetShouldExit($errCode -as [int])
    exit
}

try {

    #$Auftragsnummer = "1001042"

    #Quellpfad ermitteln
    $seachFile = $Auftragsnummer + "-AutoDeskTransfer.xml"
    $sourceFile = Get-ChildItem -Path "C:\Work\AUFTRÄGE NEU\" -Recurse -Include $seachFile
    if ($sourceFile.Count -gt 1) {
        $vault.Dispose() #Vault Connection schließen
        $errCode = "5"# AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden.
        Write-Host "AutoDeskTransferXml mehrfach im Arbeitsbereich vorhanden."-ForegroundColor DarkRed
        $Host.SetShouldExit($errCode -as [int])
    }
    $sourcePath = $sourceFile.DirectoryName.Replace("\", "/") + "/"
    $targetPath = $sourcePath.Replace("C:/Work", "$")

    #Dateinamen der einzucheckenden

    $pathExtBerechnungen = "Berechnungen/"
    $pathExtBerechnungenPDF = "Berechnungen/PDF/"
    $pathExtCAD = "Bgr00/CAD-CFP/"
    $pathExtTUEVZertifikate = "Montage-TÜV-Dokumentation/TÜV/Zertifikate/"

    $uploadFiles = @()
    $uploadFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
    if (Test-Path ($sourcePath + $Auftragsnummer + "-Spezifikation.pdf")) { $uploadFiles += $Auftragsnummer + "-Spezifikation.pdf" }
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
    
    if (Test-Path ($sourcePath + $pathExtBerechnungenPDF)) { $berechnungenPDFFiles = Get-ChildItem -Path ($sourcePath + $pathExtBerechnungenPDF) -Filter  *.pdf }
    if (Test-Path ($sourcePath + $pathExtCAD)) { $cadFiles = Get-ChildItem -Path ($sourcePath + $pathExtCAD) -Filter  *.dwg }
    if (Test-Path ($sourcePath + $pathExtTUEVZertifikate)) { $zertifikateFiles = Get-ChildItem -Path ($sourcePath + $pathExtTUEVZertifikate) -Filter  *.pdf }
 

    foreach ($berechnungensPDF in $berechnungenPDFFiles) {
        $uploadFiles += $pathExtBerechnungenPDF + $berechnungensPDF
    }

    foreach ($cadFile in $cadFiles) {
        $uploadFiles += $pathExtCAD + $cadFile
    }

    foreach ($zertifikateFile in $zertifikateFiles) {
        $uploadFiles += $pathExtTUEVZertifikate + $zertifikateFile
    }
        
    
    #Prüfen ob Daten zum Upload vorhanden sind 
    if ($berechnungenPDFFiles -match 'Anlagedaten' -or 
        $berechnungenPDFFiles -match 'Lift data' -or 
        $berechnungenPDFFiles -match 'Données techniques de l´installation' ) {
        #Daten im Vault löschen
        $toDeleteVaultFiles = @()
        $vaultPathBerechnungen = $targetPath + $pathExtBerechnungenPDF
        $vaultPathTUEVZertifikate = $targetPath + $pathExtTUEVZertifikate

        if ($berechnungenPDFFiles.Count -ge 1) {
            $files = Get-VaultFiles -Folder $vaultPathBerechnungen
            foreach ($file in $files) {
                if (($file._Author -eq "CFP") -and ($file._CategoryName -eq "Office") -and ($file._Extension -eq "pdf") -and ($file.Kategorie -eq "Berechnungen")) {
                    $toDeleteVaultFiles += $file
                }
            } 
        }

        if ($zertifikateFiles.Count -ge 1) {
            $files = Get-VaultFiles -Folder $vaultPathTUEVZertifikate
            foreach ($file in $files) {
                if (($file._Author -eq "CFP") -and ($file._CategoryName -eq "Office") -and ($file._Extension -eq "pdf") -and ($file.Kategorie -eq "Baumuster-Zertifikate")) {
                    $toDeleteVaultFiles += $file
                }
            } 
        }

        foreach ($toDeleteVaultFile in $toDeleteVaultFiles) {
            try {
                $toDeleteFolder = $vault.DocumentService.GetFolderByPath($toDeleteVaultFile._EntityPath)
                $vault.DocumentService.DeleteFileFromFolderUnconditional( $toDeleteVaultFile.MasterId , $toDeleteFolder.Id)
                Write-Host  $toDeleteVaultFile.Name  "gelöscht..."-ForegroundColor Yellow
            }
            catch { 
                Write-Host  $toDeleteVaultFile.Name "nicht gelöscht,keine Rechte zum Löschen..."-ForegroundColor DarkRed
            }
        }
    }

    $verfasser = $Env:USERNAME

    #Dateien hochladen und aktualisieren
    for ($i = 0; $i -le $uploadFiles.Count - 1; $i++) {
        $uploadSource = -join ($sourcePath, $uploadFiles[$i])
        $uploadTarget = -join ($targetPath, $uploadFiles[$i])
        $uploadFile = Add-VaultFile -From $uploadSource -To $uploadTarget
        $Beschreibung = (($uploadFile._Name.TrimStart($Auftragsnummer + "-")).TrimEnd("." + $uploadFile._Extension))
        switch ($uploadFile._Extension) {
            "xml" { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = $Beschreibung; 'Projekt' = $Auftragsnummer; 'Verfasser' = $verfasser; 'Kategorie' = "Berechnungen" } -Category "AnlageDaten" }
            "pdf" {
                If ($updateFile.Path -match "Berechnungen") {
                    $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = $Beschreibung; 'Projekt' = $Auftragsnummer; 'Verfasser' = "CFP"; 'Kategorie' = "Berechnungen" } -Category "Office" 
                }
                ElseIf ($updateFile.Path -match "Zertifikate") {
                    $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = $Beschreibung; 'Projekt' = $Auftragsnummer; 'Verfasser' = "CFP"; 'Kategorie' = "Baumuster-Zertifikate" } -Category "Office"
                }
                Else {
                    $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = $Beschreibung; 'Projekt' = $Auftragsnummer; 'Verfasser' = $verfasser; 'Kategorie' = "Berechnungen"; 'Kommentare' = "Von Spezifikation automatisch generierte Datei" }-Category "Office"
                }
          
            }
            "html" {

                $html = New-Object -ComObject "HTMLFile"
                $html.IHTMLDocument2_write($(Get-Content ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".html") -raw))
                $motortyp = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Motortyp*" }).innerText
                $aufhaengung = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Aufhängung*" }).innerText.Replace("Aufhängung is ", "")
                $lageTreibscheibe = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Treibscheibe *" }).innerText[0]
                $treibscheibe = ($HTML.body.getElementsByTagName('tr') | Where-Object { $_.innerText -like "Treibscheibe *" }).innerText[2]
          
                $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Antriebsauslegung Ziehl Abegg"; 'Projekt' = $Auftragsnummer; 'Kategorie' = "Berechnungen"; 'Antriebtyp' = $motortyp; 'Aufhängung' = $aufhaengung; 'Lage Antrieb' = $lageTreibscheibe; 'Treibscheibe Zylinder' = $treibscheibe } -Category "AntriebsDaten"
            }
            "aus" { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Antriebsauslegung Ziehl Abegg"; 'Projekt' = $Auftragsnummer; 'Kategorie' = "Berechnungen" } -Category "AntriebsDaten" }
            "dat" { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Daten Bausatzprogram CFP"; 'Projekt' = $Auftragsnummer; 'Verfasser' = $verfasser; 'Kategorie' = "Berechnungen" } -Category "AnlageDaten" }
            "LILO" {
                if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")) {
                    $hydroDat = Get-Content -path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")

                    $motortyp = ($hydroDat -match "Power_Unit_Type").Replace("[Power_Unit_Type] ", "") + ($hydroDat -match "Valve_Model").Replace("[Valve_Model] ", " - ") + ($hydroDat -match "Pumpenbezeichnung").Replace("[Pumpenbezeichnung] ", "- ")
                    $aufhaengung = ($hydroDat -match "Bauart")[0].Replace("[Bauart] ", "")
                    $lageTreibscheibe = If (($hydroDat -match "Antrieb_im_Schacht").Replace("[Antrieb_im_Schacht] ", "") -eq "0") { "Antrieb im Maschinenraum" }else { "Antrieb im Schacht" }
                    $treibscheibe = ($hydroDat -match "Zylinderbezeichnung").Replace("[Zylinderbezeichnung] ", "")
                }
                Else {
                    $motortyp = "Keine CFP-Auslegung vorhanden"
                    $aufhaengung = "Keine CFP-Auslegung vorhanden"
                    $lageTreibscheibe = "Keine CFP-Auslegung vorhanden"
                    $treibscheibe = "Keine CFP-Auslegung vorhanden"
                }
                $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Antriebsauslegung Ziehl Abegg"; 'Projekt' = $Auftragsnummer; 'Kategorie' = "Berechnungen"; 'Antriebtyp' = $motortyp; 'Aufhängung' = $aufhaengung; 'Lage Antrieb' = $lageTreibscheibe; 'Treibscheibe Zylinder' = $treibscheibe } -Category "AntriebsDaten"
            }
            "txt" { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Fertigungsunterlagen CFP"; 'Projekt' = $Auftragsnummer; 'Kategorie' = "Berechnungen" } -Category "FertigungsDaten" }
            "dwg" { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = "Bausatz Zeichnungen"; 'Projekt' = $Auftragsnummer; 'Kategorie' = "Montagebaugruppe"; 'Kommentare' = "von CFP automatisch generierte Zeichnung" } -Category "Zeichnungsableitungen" }
            default { $updateFile = Update-VaultFile -File $uploadFile.'Full Path' -Properties @{'Beschreibung' = $Beschreibung; 'Projekt' = $Auftragsnummer } -Category "Basis" }
        }

        Write-Host "Datei"$uploadFile._Name"wurde hochgeladen und eingechecked!"-ForegroundColor Yellow
    }

}
catch {
    $vault.Dispose() #Vault Connection schließen

    $errCode = "1" # Datei upload ist fehlgeschlagen

    $Host.SetShouldExit($errCode -as [int])
    exit

}

try {

    #XML-Datei ermitteln und auslesen
    $pfadxml = $sourceFile.FullName

    $xml = [XML] (Get-Content -Path $pfadxml -Encoding UTF8)

    $parameter = $xml.selectNodes("//ParamWithValue")
    #$parameter | select name, value, typeCode


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

    #Föderhöhe in Meter umwandeln
    if ($null -ne $var_FH.value) {
        $var_FH.value = $var_FH.value * 1000
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
    $kabinenflaeche.Val = $var_A_Kabine.value
    $kommentare.Val = $var_Kommentare.value

    $propValues = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
    $propValues.Items = New-Object Autodesk.Connectivity.WebServices.PropInstParam[] $folderProps.Count
    $i = 0
    foreach ($d in $folderProps.GetEnumerator()) {
        $propValues.Items[$i] = New-Object Autodesk.Connectivity.WebServices.PropInstParam -Property @{PropDefId = $d.PropDefId; Val = $d.Val }
        $i++
    }

        
    #Housekeeping

    $workPathBerechnungenPDF = $sourcePath + $pathExtBerechnungenPDF
    $workPathTUEVZertifikate = $sourcePath + $pathExtTUEVZertifikate


    If (($workPathBerechnungenPDF -match "C:/Work/AUFTRÄGE NEU") -and ($workPathTUEVZertifikate -match "C:/Work/AUFTRÄGE NEU")) {
    
        if (Test-Path ($workPathBerechnungenPDF)) { Remove-Item -Path $workPathBerechnungenPDF -Recurse -Force }
        if (Test-Path ($workPathTUEVZertifikate)) { Remove-Item -Path $workPathTUEVZertifikate -Recurse -Force }
    }

    $deleteFiles = @()
    $deleteFiles += $Auftragsnummer + "-AutoDeskTransfer.xml"
    if (Test-Path ($sourcePath + $Auftragsnummer + "-Spezifikation.pdf")) { $deleteFiles += $Auftragsnummer + "-Spezifikation.pdf" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".html")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + ".html" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".aus")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + ".aus" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".dat")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + ".dat" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + ".LILO")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + ".LILO" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Jupiter.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-Jupiter.txt" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Pluto.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-Pluto.txt" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Beripac.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-Beripac.txt" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-Pluto-Seil.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-Pluto-Seil.txt" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-ZZE-S.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-ZZE-S.txt" }
    if (Test-Path ($sourcePath + $pathExtBerechnungen + $Auftragsnummer + "-G.txt")) { $deleteFiles += $pathExtBerechnungen + $Auftragsnummer + "-G.txt" }


    foreach ($deleteFile in $deleteFiles) {
        try {
            $pathDeleteFile = $sourcePath + $deleteFile
            Remove-Item $pathDeleteFile -Force
        }
        catch
        {}
    }







    #Read-Host Debug:

    $vault.DocumentServiceExtensions.UpdateFolderProperties(@($folder.Id), @($propValues))

    $vault.Dispose() #Vault Connection schließen

    $errCode = "0"

    $Host.SetShouldExit($errCode -as [int])
    exit
}
catch {
    $vault.Dispose() #Vault Connection schließen

    $errCode = "3"# Eigenschaftenabgleich Vault ist fehlgeschlagen

    $Host.SetShouldExit($errCode -as [int])
    exit
}
