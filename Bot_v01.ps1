while(0 -eq 0){
Stop-Process -Name iexplore
Start-Sleep -s 3
function Get-RandomCharacters($length, $characters) {
    $random = 1..$length | ForEach-Object { Get-Random -Maximum $characters.length }
    $private:ofs=""
    return [String]$characters[$random]
}
 
function Scramble-String([string]$inputString){     
    $characterArray = $inputString.ToCharArray()   
    $scrambledStringArray = $characterArray | Get-Random -Count $characterArray.Length     
    $outputString = -join $scrambledStringArray
    return $outputString 
}
 
$password = Get-RandomCharacters -length 5 -characters 'ABCDEFGHKLMNOPRSTUVWXYZ'
$password = Scramble-String $password

$ie = New-Object -ComObject 'internetExplorer.Application' -ErrorAction Ignore -ErrorVariable global:Fehler
$ie.Navigate("https://techinthebasketde.zendesk.com/hc/de/requests/new") #Webformular
$ie.Visible= $true
While($ie.Busy -eq $true) {Start-Sleep -Seconds 3;if($exit -eq 1){$global:Main_Tool_Icon.Visible      = $false;exit}}
$mail = $ie.document.getElementByID('request_anonymous_requester_email')
$mail.click()
$mail.value = "$password@$password.de"
$betreff = $ie.document.getElementByID('request_subject')
$betreff.click()
$betreff.value = "#Betreff"
$beschreibung = $ie.document.getElementByID('request_description')
$beschreibung.click()
$beschreibung.value = "Bitte leiten sie die Rückzahlung für die Bestellung #Bestellnummer in die Wege"
$frage = $ie.document.getElementByID('request_custom_fields_360000226209')
$frage.click()
$frage.value = "fragen_zu_ihrer_bestellung" #Thema
$ID = $ie.document.getElementByID('request_custom_fields_45382549')
$ID.click()
$ID.value = "#Bestellnummer"
$send = $ie.Document.links | where type -match "submit"
#$send.click()
write-host "klick"
start-sleep -s 3
While($ie.Busy -eq $true) {Start-Sleep -Seconds 3;if($exit -eq 1){$global:Main_Tool_Icon.Visible      = $false;exit}}
}
