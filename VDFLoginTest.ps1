Add-Type -path "C:\Program Files\Autodesk\Vault Client 2023\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
Add-Type -path "C:\Program Files\Autodesk\Vault Client 2023\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2023\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")

$ReadOnly = $false
#$ReadOnly = $true     

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
        $Host.SetShouldExit([int]$errCode)
        exit
    }
}
else {
    try {
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
        $Host.SetShouldExit([int]$errCode)
        exit
    }   
}

$vault = $connection.WebServiceManager
$VltHelpers = New-Object VdsSampleUtilities.VltHelpers


$targetPath = "$/AUFTRÄGE NEU/Konstruktion/895/8951475"

$folder = $vault.DocumentService.GetFolderByPath($targetPath)
$propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")

$folderProps = $vault.PropertyService.GetPropertiesByEntityIds("FLDR", @($folder.Id))

$udpIds = $propDefs | Where-Object { $_.IsSys -eq $false } | Select-Object -ExpandProperty Id
$folderProps = $folderProps | Where-Object { $_.Propdefid -in $udpIds }

$kommentare = $folderProps | Where-Object { $_.PropDefId -eq "24" }
$kabinenflaeche = $folderProps | Where-Object { $_.PropDefId -eq "143" }

$kommentare.Val = "Hallllo"
$kabinenflaeche.Val = 2.26


$propValues = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
$propValues.Items = New-Object Autodesk.Connectivity.WebServices.PropInstParam[] $folderProps.Count
$i = 0
foreach ($d in $folderProps.GetEnumerator()) {
    $propValues.Items[$i] = New-Object Autodesk.Connectivity.WebServices.PropInstParam -Property @{PropDefId = $d.PropDefId; Val = $d.Val }
    $i++
}



try {
    $vault.DocumentServiceExtensions.UpdateFolderProperties(@($folder.Id), @($propValues))
}
catch {
    $test = "Hallo"
}




# $vault.Dispose()
$logOff = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($connection)
$logOff

