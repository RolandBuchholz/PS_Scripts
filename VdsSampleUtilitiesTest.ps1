Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.Forms.dll"
Add-Type -path "C:\Program Files\Autodesk\Vault Client 2022\Explorer\Autodesk.DataManagement.Client.Framework.Vault.dll"
[System.Reflection.Assembly]::LoadFrom($Env:ProgramData + "\Autodesk\Vault 2022\Extensions\DataStandard\Vault.Custom\addinVault\VdsSampleUtilities.dll")


$VltHelpers = New-Object VdsSampleUtilities.VltHelpers

$VltHelpers.GetType()
