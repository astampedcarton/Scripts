#List all the file extention to be search
$extlst = @("rtf","doc","docx","xlsx","xls")
$delfold = "c:\temp\"

#Get all the files for a particular extention
$lst = Get-ChildItem -LiteralPath "$delfold" -Recurse | Where-Object {$_.Extension -match '\.(rtf|doc|docx|xlsx|xls)$'}
clear

# $rtfs = Get-ChildItem -LiteralPath "$delfold" -Recurse -Filter *.rtf
$rtfs = Get-ChildItem -LiteralPath "$delfold" -Filter *.rtf

$word = New-Object -ComObject Word.Application

$word.Visible = $false

if (-not $word){
  Write-Host "Word did not start"
}

foreach($rtf in $rtfs){
  #$rtfcont = Get-Content $rtf.FullName -RAw
  #Write-Host($rtf.FullName)
  $doc = $word.Documents.Open($rtf.FullName)
  $fname = "{0, -30}" -f $rtf 
  foreach($section in $doc.Sections) {
     
    $header = $section.Headers.Item(1).Range.Text
    Write-Output "$Fname : $header"
  }

  $doc.Close()
}
