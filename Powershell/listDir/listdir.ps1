#Script to act as simple ls command, plus color coding the file by type or flagging files that 
#contains the word copy. Specifically created for the command line as it overs quicker access 
# Todo: Add support for wild cards
# Todo: Need to think about better classification of file types
# Todo: Add support to open a file that was found
#
#Make sure that the terminal is cleanr everytime the script runs.

clear
$dir = ""

#Retrieve all the files in a directory and display in a list
$histFile = "C:\Users\henti\OneDrive\Programming\Powershell\hist.txt"
$maxln = 120
$line = "-"*($maxln-2)
$blline = " "*($maxln-2)
$blkst = "+"

$head  = "Directories"
$head1 = " "

$msg   = " Enter the path to display the content for."

function CreateUI{
    param(
        [string]$path
    )
   clear
   Write-Host "$blkst$line$blkst" -ForegroundColor Cyan
   $dirpath = Read-Host "Dir Path " 
   Write-Host "$blkst$blline$blkst" -ForegroundColor Cyan
   Write-Host "$blkst$line$blkst"  -ForegroundColor Cyan
}


#Main loop.
while ($true){
#  CreateUI
  Write-Host " "
  Write-Host "$line"
  $dirpath = Read-Host "Dir Path " 

  switch($dirpath){
    "exit" {
      Write-Host "Exit "
      exit
    }
    default {
      Write-Host "$line"
      $dircont = Get-ChildItem -path $dirpath
      foreach ($cont in $dircont){
        $ext = [System.IO.path]::GetExtension($cont)
        #Write-Host "File attri: $cont.PSIsContainer"
        # switch($ext){
        #   ".ps1" {$color = "green"}
        # }
        if ($ext -eq ".ps1") {
          $color = "green"
        } else {
          if ($ext -eq ".sas" -or $ext -eq ".r") {
            $color = "Cyan"
          } else {
            if ($ext -eq ".xlsx" -or $ext -eq ".xls" -or $ext -eq ".xlsm") {
              $color = "DarkBlue"
            } else {
              if ($ext -eq ".doc" -or $ext -eq ".docx" -or $ext -eq ".docm"-or $ext -eq ".rtf") {
                $color = "DarkBlue"
              } else {
                if ($cont.PSIsContainer){
                  # Write-Host "$cont" -ForegroundColor Yellow
                  $color = "Yellow"
                } else {
                  # Write-Host "$cont" -ForegroundColor White
                  $color = "White"
                }
              }
            }
          }
        }
        
        # this is an overall classification. Regardless what the file type if it's a copy or delete
        # It's to be flagged in Red.
        if ($cont -match '(?i)(- copy|copy of|delete)') {
          $color = "Red"
        } 
        
        #Universal output. It's afterall just the colors that differ.
        Write-Host "$($cont.Name.PadRight(40)) $($cont.LastWritetime.Tostring().Padleft(24)) $($cont.Length.Tostring().PadLeft(20))" -ForegroundColor $color
      }
    }
   
  }  
}
