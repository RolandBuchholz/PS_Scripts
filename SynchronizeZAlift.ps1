<# .SYNOPSIS
     Update Registry ZAlift
.DESCRIPTION
     Registryeinträge des ZAlift Programmes werden mit der AutoDeskTransfer.xml synchronisiert
.NOTES
     File Name : SynchronizeZAlift.ps1
     Author : Buchholz Roland – roland.buchholz@berchtenbreiter-gmbh.de
.VERSION
     Version 0.27 – update compensationRopeWeight and tractionSheavePos
     Beispiel wie das Script aufgerufen wird > SynchronizeZAlift.ps1 get "C:\Work\AUFTRÄGE NEU\Konstruktion\100\1001042-1048\1001042\Save-1001042-AutoDeskTransfer.xml"
                                                                 (set or reset)(FullPath)                                            
.INPUTTYPE
     [String]$SynchronizeDirection
     [string]$FullPathXml
.RETURNVALUE
     $errCode
.COMPONENT
     ZAliftProgram
#>

Param(    
     [String]$SynchronizeDirection,
     [string]$FullPathXml
)

class ZALiftKey {
     [string]$RegistrySubPath
     [string]$Key
     [System.Xml.XmlElement]$Value
     [string]$PropertyType

     ZALiftKey(
          [string]$s,
          [string]$k,
          [System.Xml.XmlElement]$v,
          [string]$p
     ) {
          $this.RegistrySubPath = $s
          $this.Key = $k
          $this.Value = $v
          $this.PropertyType = $p
     }
}

#$SynchronizeDirection = "set"
#$FullPathXml = 'C:\Work\AUFTRÄGE NEU\Konstruktion\100\1003000\1003000-AutoDeskTransfer.xml'

try {

     $RegistryPathLast = 'HKCU:\SOFTWARE\VB and VBA Program Settings\ZETALIFT\LAST'
     $RegistryPathAll = 'HKCU:\SOFTWARE\VB and VBA Program Settings\ZETALIFT\All'
     $RegistryPathZAL = 'HKCU:\SOFTWARE\VB and VBA Program Settings\ZETALIFT\ZAL'

     if ($SynchronizeDirection -eq "set") {

          $xml = [XML] (Get-Content -Path $FullPathXml -Encoding UTF8)

          $parameter = $xml.selectNodes("//ParamWithValue")

          $var_Kunde = $parameter | Where-Object { $_.name -eq "var_AnPersonZ1" }
          $var_AufhaengungsartRope = $parameter | Where-Object { $_.name -eq "var_AufhaengungsartRope" }
          $var_Bausatz = $parameter | Where-Object { $_.name -eq "var_Bausatz" }
          $var_Umschlingungswinkel = $parameter | Where-Object { $_.name -eq "var_Umschlingungswinkel" }
          $var_Tragseiltyp = $parameter | Where-Object { $_.name -eq "var_Tragseiltyp" }
          $var_Normen = $parameter | Where-Object { $_.name -eq "var_Normen" }
          $var_FH = $parameter | Where-Object { $_.name -eq "var_FH" }
          $var_AuftragsNummer = $parameter | Where-Object { $_.name -eq "var_AuftragsNummer" }
          $var_F = $parameter | Where-Object { $_.name -eq "var_F" }
          $var_GegenGewicht_Masse = $parameter | Where-Object { $_.name -eq "var_Gegengewichtsmasse" }
          $var_Q = $parameter | Where-Object { $_.name -eq "var_Q" }
          $var_v = $parameter | Where-Object { $_.name -eq "var_v" } 
          $var_NumberOfRopes = $parameter | Where-Object { $_.name -eq "var_NumberOfRopes" }
          $var_AnzahlUmlenkrollen = $parameter | Where-Object { $_.name -eq "var_AnzahlUmlenkrollen" }
          $var_AnzahlUmlenkrollenFk = $parameter | Where-Object { $_.name -eq "var_AnzahlUmlenkrollenFk" }
          $var_AnzahlUmlenkrollenGgw = $parameter | Where-Object { $_.name -eq "var_AnzahlUmlenkrollenGgw" }
          $var_ZA_IMP_Regler_Typ = $parameter | Where-Object { $_.name -eq "var_ZA_IMP_Regler_Typ" }
          $var_Erkennungsweg = $parameter | Where-Object { $_.name -eq "var_Erkennungsweg" }   
          $var_Totzeit = $parameter | Where-Object { $_.name -eq "var_Totzeit" }                                                                                                                    
          $var_Vdetektor = $parameter | Where-Object { $_.name -eq "var_Vdetektor" }
          $var_KHLicht = $parameter | Where-Object { $_.name -eq "var_KHLicht" }
          $var_Umlenkrollendurchmesser = $parameter | Where-Object { $_.name -eq "var_Umlenkrollendurchmesser" }
          $var_ZA_IMP_Treibscheibe_RIA = $parameter | Where-Object { $_.name -eq "var_ZA_IMP_Treibscheibe_RIA" }
          $var_Fremdbelueftung = $parameter | Where-Object { $_.name -eq "var_Fremdbelueftung" }
          $var_ElektrBremsenansteuerung = $parameter | Where-Object { $_.name -eq "var_ElektrBremsenansteuerung" }
          $var_Treibscheibegehaertet = $parameter | Where-Object { $_.name -eq "var_Treibscheibegehaertet" }
          $var_Handlueftung = $parameter | Where-Object { $_.name -eq "var_Handlueftung" }
          $var_MotorGeber = $parameter | Where-Object { $_.name -eq "var_MotorGeber" }
          $var_UnterseilGewicht = $parameter | Where-Object { $_.name -eq "var_UnterseilGewicht" }
          $var_LageAntrieb = $parameter | Where-Object { $_.name -eq "var_LageAntrieb" }
          #$var_SeilabgangAntrieb = $parameter | Where-Object { $_.name -eq "var_SeilabgangAntrieb" }

          # default values
          if ($var_Kunde.value -eq ""){$var_Kunde.value = "Berchtenbreiter GmbH"}
          if ($var_Q.value -eq ""){$var_Q.value = "100"}
          if (($var_F.value -eq "") -or ($var_F.value -eq "0")){$var_F.value = "150"}
          if ($var_FH.value -eq ""){$var_FH.value = "1"}
          if ($var_V.value -eq ""){$var_V.value = "1"}
          if (($var_GegenGewicht_Masse.value -eq "") -or ($var_GegenGewicht_Masse.value -eq "0")){$var_GegenGewicht_Masse.value = "200"}
          if ($var_AufhaengungsartRope.value -eq ""){$var_AufhaengungsartRope.value = "1"}
          if ($var_Tragseiltyp.value -eq ""){$var_Tragseiltyp.value = "D 8mm WOLF PAWO F7"}
          if ($var_NumberOfRopes.value -eq ""){$var_NumberOfRopes.value = "2"}
          if ($var_MotorGeber.value -eq ""){$var_MotorGeber.value = "ECN 1313-2048Endat"}
          if ($var_AnzahlUmlenkrollen.value -eq ""){$var_AnzahlUmlenkrollen.value = "0"}
          if ($var_AnzahlUmlenkrollenFk.value -eq ""){$var_AnzahlUmlenkrollenFk.value = "0"}
          if ($var_AnzahlUmlenkrollenGgw.value -eq ""){$var_AnzahlUmlenkrollenGgw.value = "0"}
          if ($var_Umschlingungswinkel.value -eq ""){$var_Umschlingungswinkel.value = "180"}
          if ($var_Erkennungsweg.value -eq ""){$var_Erkennungsweg.value = "0"}
          if ($var_Totzeit.value -eq ""){$var_Totzeit.value = "0"}
          if ($var_Vdetektor.value -eq ""){$var_Vdetektor.value = "0"}
          if ($var_KHLicht.value -eq ""){$var_KHLicht.value = "2100"}
          if ($var_UnterseilGewicht.value -eq ""){$var_UnterseilGewicht.value = "0 kg"}
          if ($var_LageAntrieb.value -eq ""){$var_LageAntrieb.value = "oben, im Schacht"}
          #if ($var_SeilabgangAntrieb.value -eq ""){$var_SeilabgangAntrieb.value = "Motorfuss unten"}


          New-ItemProperty -Path $RegistryPathAll -Name "HtmlFormat" -Value "0" -PropertyType "String" -Force
     
          $ListZALiftKeys = New-Object 'system.collections.generic.dictionary[string,ZALiftKey]'

          $ListZALiftKeys.Add("A", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "A", $var_AufhaengungsartRope, "String")))
          $ListZALiftKeys.Add("Aufzugsbauart", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Aufzugsbauart", $var_Bausatz, "String")))
          $ListZALiftKeys.Add("BETA", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "BETA", $var_Umschlingungswinkel, "String")))
          $ListZALiftKeys.Add("D", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "D", $var_Tragseiltyp , "String")))
          $ListZALiftKeys.Add("EN81-20", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "EN81-20", $var_Normen , "String")))
          $ListZALiftKeys.Add("FH", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "FH", $var_FH, "String")))
          $ListZALiftKeys.Add("Filename1", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Filename1", $var_AuftragsNummer, "String")))
          $ListZALiftKeys.Add("Fkg", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Fkg", $var_F, "String")))
          $ListZALiftKeys.Add("Gkg", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Gkg", $var_GegenGewicht_Masse, "String")))
          $ListZALiftKeys.Add("Kunde", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Kunde", $var_Kunde, "String")))
          $ListZALiftKeys.Add("Projektb1", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Projektb1", $var_AuftragsNummer, "String")))
          $ListZALiftKeys.Add("Qkg", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Qkg", $var_Q, "String")))
          $ListZALiftKeys.Add("V", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "V", $var_v, "String")))
          $ListZALiftKeys.Add("Seilalt", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Seilalt", $var_Tragseiltyp , "String")))
          $ListZALiftKeys.Add("Z", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "Z", $var_NumberOfRopes, "String")))
          $ListZALiftKeys.Add("ZUM", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "ZUM", $var_AnzahlUmlenkrollen, "String")))
          $ListZALiftKeys.Add("ZUMF", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "ZUMF", $var_AnzahlUmlenkrollenFk, "String")))
          $ListZALiftKeys.Add("ZUMG", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "ZUMG", $var_AnzahlUmlenkrollenGgw, "String")))
          $ListZALiftKeys.Add("A3_Ausloeseweg", ([ZALiftKey]$key = New-Object ZALiftKey("All", "A3_Ausloeseweg", $var_Erkennungsweg, "String")))
          $ListZALiftKeys.Add("A3_Auslöseweg", ([ZALiftKey]$key = New-Object ZALiftKey("All", "A3_Auslöseweg", $var_Erkennungsweg, "String")))
          $ListZALiftKeys.Add("A3_Ausloesetotzeit", ([ZALiftKey]$key = New-Object ZALiftKey("All", "A3_Ausloesetotzeit", $var_Totzeit, "String")))
          $ListZALiftKeys.Add("A3_Ausloesegeschwindigkeit", ([ZALiftKey]$key = New-Object ZALiftKey("All", "A3_Ausloesegeschwindigkeit", $var_Vdetektor, "String")))
          $ListZALiftKeys.Add("A3_Kabinenhoehe", ([ZALiftKey]$key = New-Object ZALiftKey("All", "A3_Kabinenhoehe", $var_KHLicht, "String")))
          $ListZALiftKeys.Add("Filename_next", ([ZALiftKey]$key = New-Object ZALiftKey("All", "Filename_next", $var_AuftragsNummer, "String")))
          $ListZALiftKeys.Add("Anlage-A", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-A", $var_AufhaengungsartRope, "String")))
          $ListZALiftKeys.Add("Anlage-Art", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-Art", $var_Bausatz, "String")))
          $ListZALiftKeys.Add("Anlage-FH", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-FH", $var_FH, "String")))
          $ListZALiftKeys.Add("Anlage-F", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-F", $var_F, "String")))
          $ListZALiftKeys.Add("Anlage-G", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-G", $var_GegenGewicht_Masse, "String")))
          $ListZALiftKeys.Add("Anlage-Q", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-Q", $var_Q, "String")))
          $ListZALiftKeys.Add("Anlage-URF", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-URF", $var_AnzahlUmlenkrollenFk, "String")))
          $ListZALiftKeys.Add("Anlage-URG", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-URG", $var_AnzahlUmlenkrollenGgw, "String")))
          $ListZALiftKeys.Add("Anlage-V", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-V", $var_v, "String")))
          $ListZALiftKeys.Add("Konfignummer", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Konfignummer", $var_AuftragsNummer, "String")))
          $ListZALiftKeys.Add("Treibscheibe-N", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-N", $var_NumberOfRopes, "String")))
          $ListZALiftKeys.Add("Treibscheibe-SZ", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-SZ", $var_NumberOfRopes, "String")))
          $ListZALiftKeys.Add("Treibscheibe-RIA", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-RIA", $var_ZA_IMP_Treibscheibe_RIA, "String")))
          $ListZALiftKeys.Add("Treibscheibe-RIA_", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-RIA_", $var_ZA_IMP_Treibscheibe_RIA, "String")))
          $ListZALiftKeys.Add("Treibscheibe-SD", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-SD", $var_Tragseiltyp, "String")))
          $ListZALiftKeys.Add("Treibfaehigkeit-Seilrollendurchmesser", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibfaehigkeit-Seilrollendurchmesser", $var_Umlenkrollendurchmesser, "String")))
          $ListZALiftKeys.Add("Treibscheibe-Umschlingung", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-Umschlingung", $var_Umschlingungswinkel, "String")))
          $ListZALiftKeys.Add("UCM-Erkennungsweg", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "UCM-Erkennungsweg", $var_Erkennungsweg, "String")))
          $ListZALiftKeys.Add("UCM-Geschwindigkeitsdetektor", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "UCM-Geschwindigkeitsdetektor", $var_Vdetektor, "String")))
          $ListZALiftKeys.Add("UCM-Totzeit", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "UCM-Totzeit", $var_Totzeit, "String")))
          $ListZALiftKeys.Add("Motor-FAN", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Motor-FAN", $var_Fremdbelueftung, "String")))
          $ListZALiftKeys.Add("Bremsmodul-Typ", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Bremsmodul-Typ", $var_ElektrBremsenansteuerung, "String")))
          $ListZALiftKeys.Add("Treibscheibe-H", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Treibscheibe-H", $var_Treibscheibegehaertet, "String")))
          $ListZALiftKeys.Add("Bremse-Handlueftung", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Bremse-Handlueftung", $var_Handlueftung, "String")))
          $ListZALiftKeys.Add("Bremse-Handlüftung", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Bremse-Handlüftung", $var_Handlueftung, "String")))
          $ListZALiftKeys.Add("Bremse-Lueftueberwachung", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Bremse-Lueftueberwachung", $var_Handlueftung, "String")))
          $ListZALiftKeys.Add("Regler-Typ", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Regler-Typ", $var_ZA_IMP_Regler_Typ, "String")))
          $ListZALiftKeys.Add("Geber-Typ", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Geber-Typ", $var_MotorGeber, "String")))
          $ListZALiftKeys.Add("bfsu", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "bfsu", $var_UnterseilGewicht, "String")))
          $ListZALiftKeys.Add("FSU", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "FSU", $var_UnterseilGewicht, "String")))
          $ListZALiftKeys.Add("sukg", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "sukg", $var_UnterseilGewicht, "String")))
          $ListZALiftKeys.Add("Anlage-bfsu", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-bfsu", $var_UnterseilGewicht, "String")))
          $ListZALiftKeys.Add("MLAGE", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "MLAGE", $var_LageAntrieb, "String")))
          $ListZALiftKeys.Add("NLAGE", ([ZALiftKey]$key = New-Object ZALiftKey("LAST", "NLAGE", $var_LageAntrieb, "String")))
          $ListZALiftKeys.Add("Anlage-MLAGE", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-MLAGE", $var_LageAntrieb, "String")))
          $ListZALiftKeys.Add("Anlage-NLAGE", ([ZALiftKey]$key = New-Object ZALiftKey("ZAL", "Anlage-NLAGE", $var_LageAntrieb, "String")))

          foreach ($par in $ListZALiftKeys.Values) {
               switch ($par.RegistrySubPath) {
                    "LAST" {
                         $RegistryPath = $RegistryPathLast
                         break
                    }
                    "All" {
                         $RegistryPath = $RegistryPathAll
                         break
                    }
                    "ZAL" {
                         $RegistryPath = $RegistryPathZAL
                         break
                    }
                    Default {
                         $RegistryPath = $null 
                    }
               }

               #Validate NewValue
               switch ($par.Key) {
                    { ($_ -eq "Aufzugsbauart") -or ($_ -eq "Anlage-Art") } {
                         $isRucksack = ($var_Bausatz.value.StartsWith("BRR") -or $var_Bausatz.value.StartsWith("EZE-SR") -or $var_Bausatz.value.StartsWith("Sonderbausatz Seil Rucksack"))
                         If ($par.Key -eq "Aufzugsbauart") {
                              if ($isRucksack) {
                                   $newValue = "rucksack"
                              }
                              else {
                                   $newValue = "standard"
                              }
                         }
                         elseif ($par.Key -eq "Anlage-Art") {
                              if ($isRucksack) {
                                   $newValue = "1"
                              }
                              else {
                                   $newValue = "0"
                              }
                         }
                    }
                    { ($_ -eq "D") -or ($_ -eq "Seilalt") -or ($_ -eq "Treibscheibe-SD") } {
                         if ($par.Value.value -ne ""){
                              $ropeSplit = $par.Value.value -split "mm"
                              if ($par.Key -eq "Seilalt" -and $ropeSplit.Count -ge 2) {
                                   $newValue = $ropeSplit[1].Trim()
                              }
                              elseif ($ropeSplit.Count -ge 2) {
                                   $newValue = $ropeSplit[0].Replace("D", "").Trim()
                              }
                         }
                    }
                    "EN81-20" {
                         if ($par.Value.value.StartsWith("EN81-20")) {
                              $newValue = "1"
                         }
                         else {
                              $newValue = "0"
                         }
                    }
                    { ($_ -eq "Gkg") -or ($_ -eq "Anlage-G") } {

                         if ($var_GegenGewicht_Masse.value -ne "") {
                              $newValue = $var_GegenGewicht_Masse.value
                         }
                    }
                    "A3_Ausloesegeschwindigkeit" {
                         if ($par.Value.value -ne "") {
                              $Vdetektor = [System.Convert]::ToDecimal($par.Value.value, [cultureinfo]::GetCultureInfo('de-DE'))
                              $newValue = ($Vdetektor * 1000).ToString()
                         }
                    }
                    "A3_Kabinenhoehe" {
                         if ($par.Value.value -ne "") {
                              $newValue = $par.Value.value
                         }
                    }
                    "UCM-Erkennungsweg" {
                         if ($par.Value.value -ne "") {
                              $Erkennungsweg = [System.Convert]::ToDecimal($par.Value.value, [cultureinfo]::GetCultureInfo('de-DE'))
                              $newValue = ($Erkennungsweg / 1000).ToString()
                         }
                    }
                    "Filename_next" {
                         $newValue = (Split-Path -Path $FullPathXml) + "\Berechnungen\" + $var_AuftragsNummer.value
                    }
                    "Motor-FAN" {
                         if ($par.Value.value -eq "True") {
                              #regedit not possible
                         }
                    }
                    "Bremsmodul-Typ" {
                         if ($par.Value.value -eq "True") {
                              $newValue = "ZAsbc4C 230"
                         }
                         else {
                              $newValue = "ohne"
                         }
                    }
                    "Treibscheibe-H" {
                         if ($par.Value.value -eq "True") {
                              $newValue = "1"
                         }
                         else {
                              $newValue = "0"
                         }
                    }
                    { ($_ -eq "Bremse-Handlueftung" ) -or ($_ -eq "Bremse-Handlüftung") } {
                         if ($par.Value.value -match "mit Hand") {
                              $newValue = "mit Handlueftung"
                         }
                         elseif ($par.Value.value -match "Bowden") {
                              $newValue = "fuer Bowdenzug"
                         }
                         else {
                              $newValue = "ohne Handlueftung"
                         }
                    }
                    "Bremse-Lueftueberwachung" {
                         if ($par.Value.value -match "Mikrosch") {
                              $newValue = "Mikroschalter"
                         }
                         else {
                              $newValue = "Naeherungsschalter"
                         }
                    }
                    "Regler-Typ" {
                         if ($par.Value.value -eq "") {
                              $newValue = ""
                         }
                         else {
                              $newValue = $par.Value.value.Replace("ZAdyn4CS", "ZAdyn4CS ")
                         }
                    }
                    { ($_ -eq "bfsu" ) -or ($_ -eq "FSU") -or ($_ -eq "Anlage-bfsu")-or ($_ -eq "sukg")} {
                             if (($_ -eq "bfsu") -or ($_ -eq "Anlage-bfsu")){
                                   if ($par.Value.value.EndsWith("kg")){
                                        $newValue = "kg"
                                   }
                                   else{
                                        $newValue = "%"
                                   }
                              }
                              else {
                                   $newValue = $par.Value.value.Split("")[0]
                              }
                    }
                    { ($_ -eq "MLAGE" ) -or ($_ -eq "Anlage-MLAGE")} {
                         if ($par.Value.value.StartsWith("unten")) {
                              $newValue = "1"
                         }
                         else {
                              $newValue = "0"
                         }
                    }
                    { ($_ -eq "NLAGE" ) -or ($_ -eq "Anlage-NLAGE")} {
                         if ($par.Value.value.EndsWith("über")) {
                              $newValue = "0"
                         }
                         elseif ($par.Value.value.EndsWith("neben")) {
                              $newValue = "1"
                         }
                         else{
                              $newValue = "2"
                         }
                    }
                    Default {
                         $newValue = $par.Value.value
                    }
               }
               if ($null -ne $RegistryPath) {
                    if (-NOT (Test-Path $RegistryPath)) {
                         New-Item -Path $RegistryPath -Force | Out-Null
                    }  
                    New-ItemProperty -Path $RegistryPath -Name $par.Key -Value $newValue -PropertyType $par.PropertyType -Force
               }
          }
     }
     elseif ($SynchronizeDirection -eq "reset") {

          if (-NOT (Test-Path $RegistryPathAll)) {
               New-Item -Path $RegistryPathAll -Force | Out-Null
          }  
          New-ItemProperty -Path $RegistryPathAll -Name "Filename_next" -Value "" -PropertyType "String" -Force

          if (-NOT (Test-Path $RegistryPathLast)) {
               New-Item -Path $RegistryPathLast -Force | Out-Null
          }  
          New-ItemProperty -Path $RegistryPathLast -Name "Filename1" -Value "" -PropertyType "String" -Force

          if (-NOT (Test-Path $RegistryPathZAL)) {
               New-Item -Path $RegistryPathZAL -Force | Out-Null
          }  
          New-ItemProperty -Path $RegistryPathZAL -Name "Konfignummer" -Value "" -PropertyType "String" -Force    
     }
     $errCode = 0
}
catch {
     $errCode = 1
}

$Host.SetShouldExit([int]$errCode)
exit