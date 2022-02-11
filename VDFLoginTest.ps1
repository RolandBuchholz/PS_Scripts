Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2022\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")

$ReadOnly = $false
$ReadOnly = $true     

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


function FindFile($fileName) {
    $filePropDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FILE")
    $fileNamePropDef = $filePropDefs | Where-Object { $_.SysName -eq "Name" }
    $srchCond = New-Object 'Autodesk.Connectivity.WebServices.SrchCond'
    $srchCond.PropDefId = $fileNamePropDef.Id
    $srchCond.PropTyp = "SingleProperty"
    $srchCond.SrchOper = 3 #is equal
    $srchCond.SrchRule = "Must"
    $srchCond.SrchTxt = $fileName

    $bookmark = ""
    $status = $null
    $totalResults = @()
    while ($null -eq $status -or $totalResults.Count -lt $status.TotalHits) {
        $results = $vault.DocumentService.FindFilesBySearchConditions(@($srchCond), $null, $null, $false, $true, [ref]$bookmark, [ref]$status)
        if ($null -ne $results) {
            $totalResults += $results
        }
        else { break }
    }
    return $totalResults;
}

$vault = $connection.WebServiceManager

# $downloadFullFileName = FindFile("1001042-AutoDeskTransfer.xml")


$user = $connection.UserName
$VltHelpers = New-Object VdsSampleUtilities.VltHelpers

$Path = "$/Administration/Vault.ico"
$Test = $VltHelpers.mGetFileByFullFileName($connection, $Path)

$SearchCriteria = New-Object 'system.collections.generic.dictionary[string,string]'
$SearchCriteria.Add("Name", "1001042-AutoDeskTransfer.xml")

$Test2 = $VltHelpers.GetFileBySearchCriteria($connection, $SearchCriteria)
$Test2
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$window = New-Object System.Windows.Forms.Form
$window.Width = 400
$window.Height = 100
 
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(10, 10)
$Label.Text = $user
$Label.AutoSize = $True

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Size(10, 30)
$Label2.Text = $downloadFullFileName.Name
$Label2.AutoSize = $True

$window.Controls.Add($Label)
$window.Controls.Add($Label2)
[void]$window.ShowDialog()

# $vault.Dispose()
$logOff = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($connection)
$logOff

