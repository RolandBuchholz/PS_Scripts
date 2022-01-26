$sourcePath = 'C:\Work\Test\'


if (Test-Path ($sourcePath)) { $berechnungenPDFFiles = Get-ChildItem -Path ($sourcePath) -Filter  *.pdf }

foreach ($berechnungensPDF in $berechnungenPDFFiles) {
    $uploadFiles += $pathExtBerechnungenPDF + $berechnungensPDF
    Write-Host $berechnungensPDF
}

if ($berechnungenPDFFiles -match 'Anlagedaten' -or 
    $berechnungenPDFFiles -match 'Lift data' -or 
    $berechnungenPDFFiles -match 'Données techniques de l´installation' ) {
    Write-Host "Datei found"
}
else {
    Write-Host "Datei not found" 
}