Param(
    [Parameter(Mandatory = $true)]          
    [String]$Auftragsnummer,
    [bool]$ReadOnly = $false       
)


If ($ReadOnly) {
    $DateiStatus = "Datei ist schreibgeschützt"
}
else {
    $DateiStatus = "Datei ist kann bearbeitet werden"
}


Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$window = New-Object System.Windows.Forms.Form
$window.Width = 400
$window.Height = 100
 
$Label = New-Object System.Windows.Forms.Label
$Label.Location = New-Object System.Drawing.Size(10, 10)
$Label.Text = $Auftragsnummer
$Label.AutoSize = $True

$Label2 = New-Object System.Windows.Forms.Label
$Label2.Location = New-Object System.Drawing.Size(10, 30)
$Label2.Text = $DateiStatus
$Label2.AutoSize = $True

$window.Controls.Add($Label)
$window.Controls.Add($Label2)
[void]$window.ShowDialog()