#Small script to check for Todo and FixMe's in a file.
#Refer to the readme for setup in VScodium/Vscode or NPP 

param([Parameter(Mandatory)]
      [string]$file    
) 

Get-Content $file | 
  Select-String -pattern 'TODO|FIXME|TO DO' |
  ForEach-Object{
#    Write-host "{0}:{1}  {2}" -f $_.path, $_.LineNumber, $_.Line.Trim()
    if (Select-String -InputObject $_ -pattern 'FIXME'){
      Write-host "Line #: $($_.LineNumber): $($_.Line.Trim())" -ForeGroundColor "Red"
    } else {
      Write-host "Line #: $($_.LineNumber): $($_.Line.Trim())" -ForeGroundColor "Green"
    }
  }

#Ensure that the terminal window is kept active. 
#this can be removed without an effect to the script
Read-Host "Press Enter to close"

