# Import-Module powerVault
# Initialize-VDF
# # Create-LogRepository
# # Get-LogRepository
# Get-VaultInstallationDirectory
Add-Type -path "E:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
Add-Type -path "E:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2022\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")

$settings = New-Object Autodesk.DataManagement.Client.Framework.Vault.Forms.Settings.LoginSettings
$settings.ServerName = "localhost"
$settings.VaultName = "Samples"
$settings.AutoLoginMode = 3
$connection = [Autodesk.DataManagement.Client.Framework.Vault.Forms.Library]::Login($settings)
$vault = $connection.WebServiceManager


$user = $connection.UserName

$InvHelpers = New-Object VdsSampleUtilities.InvHelpers
$AcadHelpers = New-Object VdsSampleUtilities.AcadHelpers 
$VltHelpers = New-Object VdsSampleUtilities.VltHelpers

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
# $Label2.Text = $file.'Full Path'.ToString()
$Label2.Text = "Halllllllllllllllllllllllooooooo"
$Label2.AutoSize = $True

$window.Controls.Add($Label)
$window.Controls.Add($Label2)
[void]$window.ShowDialog()

$vault.Dispose()
$logOff = [Autodesk.DataManagement.Client.Framework.Vault.Library]::ConnectionManager.LogOut($connection)

