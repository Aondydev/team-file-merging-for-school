$scriptpath = $MyInvocation.MyCommand.Path
$dir = Split-Path $scriptpath

Get-Content "$dir\..\REGLAGES.txt" | foreach-object -begin {$settings=@{}} -process { $k = [regex]::split($_,'='); if(($k[0].CompareTo("") -ne 0) -and ($k[0].StartsWith("[") -ne $True)) { $settings.Add($k[0], $k[1]) } }


Function OpenExcelBook($FileName)

{

$Excel=new-object -ComObject Excel.Application

Return $Excel.workbooks.open($Filename)

}
Function SaveExcelBook($Workbook)

{

$Workbook.save()

$Workbook.close()

}
Function ReadCellData($Workbook,$Cell)

{

$Worksheet=$Workbook.Activesheet

Return $Worksheet.Range($Cell).text

}


$Workbook=OpenExcelBook -FileName "$($dir)\..\liste.xlsx"

$Row=1

Do
{

$Data=ReadCellData -Workbook $Workbook -Cell "A$Row"

$pathtocreate = "$($settings.Get_Item("EmplacementPourEleves"))\$($Data)\traité"

If ($Data.length –ne '0')
{
"Création du dossier de $Data"
New-Item -Path "$pathtocreate" -ItemType "directory"

$Row++

}

}until($Data.length -eq 0)

SaveExcelBook –workbook $Workbook
New-Item -Path "$($settings.Get_Item("EmplacementPerso"))" -ItemType "directory"