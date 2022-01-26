Param(            
    [String]$Auftragsnummer           
)
class DownloadInfo {
    [bool]$Success
    [string]$FileName
    [string]$FullFileName
    [bool]$IsCheckOut
    [string]$CheckOutBy
}

$result = [DownloadInfo]::new()
$result.Success = $true
$result.FileName = $Auftragsnummer
$result.FullFileName = "C:/........../" + $Auftragsnummer
$result.IsCheckOut = $null
$result.CheckOutBy = "Der Checker" 

# $Test = ConvertTo-Json $result


Write-Output (ConvertTo-Json $result)

$errCode = "2" #Login fehlgeschlagen
$Host.SetShouldExit($errCode -as [int])
exit
