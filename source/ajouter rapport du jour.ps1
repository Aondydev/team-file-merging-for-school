$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath
#from http://tlingenf.spaces.live.com/blog/cns!B1B09F516B5BAEBF!213.entry
#
Get-Content "$dir\..\REGLAGES.txt" | foreach-object -begin {$settings=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $settings.Add($k[0], $k[1]) } }

Function SearchAWord($Document,$findtext,$replacewithtext)
{
  $FindReplace=$Document.ActiveWindow.Selection.Find
  $matchCase = $false;
  $matchWholeWord = $true;
  $matchWildCards = $false;
  $matchSoundsLike = $false;
  $matchAllWordForms = $false;
  $forward = $true;
  $format = $false;
  $matchKashida = $false;
  $matchDiacritics = $false;
  $matchAlefHamza = $false;
  $matchControl = $false;
  $read_only = $false;
  $visible = $true;
  $replace = 2;
  $wrap = 1;
  $FindReplace.Execute($findText, $matchCase, $matchWholeWord, $matchWildCards, $matchSoundsLike, $matchAllWordForms, $forward, $wrap, $format, $replaceWithText, $replace, $matchKashida ,$matchDiacritics, $matchAlefHamza, $matchControl)

}

$word = New-Object -ComObject word.application
$word.Visible = $true

$outputfiles = Get-ChildItem "$($settings.Get_Item("EmplacementPerso"))\*.docx"
foreach($f in $outputfiles){
$docentete = $word.documents.open("$dir\..\en tete.docx")
SearchAWord -Document $docentete -findtext '***nom***' -replacewithtext $f.BaseName
SearchAWord -Document $docentete -findtext '***date***' -replacewithtext $(get-date -f dd/MM/yyyy)
$docentete.SaveAs([ref]"$dir\entetetemp.docx")
$docentete.close()
$doc = $word.documents.open("$($settings.Get_Item("EmplacementPerso"))\$($f.Name)")
$selection = $word.Selection
    $inputfiles = Get-ChildItem "$($settings.Get_Item("EmplacementPourEleves"))\$($f.BaseName)\*.docx"
    foreach($indoc in $inputfiles){
    $selection.EndKey(6,0) #https://technet.microsoft.com/en-us/library/ee692877.aspx 6 means going to the end of doc
    $selection.InsertFile("$dir\entetetemp.docx")
    $selection.TypeParagraph
    $selection.InsertFile("$($settings.Get_Item("EmplacementPourEleves"))\$($f.BaseName)\$($indoc.Name)")
    $selection.InsertBreak(7) #https://technet.microsoft.com/en-us/library/ee692855.aspx 7 inserts a page break
    Move-Item "$($indoc.FullName)" "$($indoc.DirectoryName)\traité\$($indoc.BaseName) $(get-date -f dd-MM-yyyy)$($indoc.Extension)"
    }
$doc.Fields.Update()
$doc.Save()
$doc.close()
Copy-Item -Path "$($f.FullName)" -Destination "$($settings.Get_Item("EmplacementPourEleves"))\$($f.BaseName)\traité\$($f.Name)" -Force
}

$word.Quit()