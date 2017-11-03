$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath

Get-Content "$dir\..\REGLAGES.txt" | foreach-object -begin {$settings=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $settings.Add($k[0], $k[1]) } }


$directories = Get-ChildItem "$($settings.Get_Item("EmplacementPourEleves"))" | where {$_.Attributes -match'Directory'}
$word = New-Object -ComObject word.application
$word.Visible = $true
foreach ($d in $directories){
$filepathyolo = "$($settings.Get_Item("EmplacementPerso"))\$d.docx"
$pathalreadyexists = Test-Path $filepathyolo
If ($pathalreadyexists -eq 0 ){
    "Création du fichier de: $d"
$doc = $word.documents.add()
$selection = $word.Selection
$doc.Saveas([REF]$filepathyolo)
$doc.close()
}
Else{
"$d a déjà un fichier"
}
}
$word.Quit()