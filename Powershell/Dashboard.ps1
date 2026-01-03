# Dashboard.ps1 - Multi-Column Display
# Lightweight PowerShell Dashboard to open links from a text file
#--------------------------------------------------------------------------
# Current expected layout of the links.txt file
#--------------------------------------------------------------------------
# Name,ShortName,Path,Color,readonly
# Example:
# Google,Google, https://www.google.com,Magenta
# Dispaly Name for link, Group Name,Path to file or folder or URL,Color,readonly
# SomeDoc,SomeDoc, c:\temp\doc.doc,Magenta,readonly
# Valid colors are: Black, DarkBlue, DarkGreen, DarkCyan, DarkRed, DarkMagenta, DarkYellow, 
# Gray, DarkGray, Blue, Green, Cyan, Red, Magenta, Yellow, White
#

# Author: (henti) (https://github.com/astampedcarton/GeneralTools/blob/main/Powershell/Dashboard.ps1)
# Date: 20SEP2025
# Notes: Dashboad is provided as is.

# --- Configuration ---
$linksFile = ".\links.txt"
#$linksFile = "C:\Users\userid\OneDrive\Programming\Powershell\links.txt"
$maxln = 120
$line = "-"*($maxln-2)
$blkst = "+"

$head  = "Links Dashboard"
$head1 = "Entries in RED indicate Folder links that's not working."
$head2 = "* Indicates that a file is to be opened as Read-Only. URL's are not checked if they exist."
$head3 = "r: Reload the Links file. p: Launch Powershell. b: Launch Browser. e: Launch Explorer."
$head4 = " "

$msg   = " Enter the number of the link to open, or 'Press Ctrl+C' to quit."

function CreateBlocks {
    param (
        [string]$text,
        [bool]$inclline = $false
    )

    if ($inclline -eq $true) {
        Write-Host "$blkst$line$blkst" -ForegroundColor Cyan
    } 

    if ($text.Length -gt 0) {
        Write-Host "|" -ForegroundColor Cyan -NoNewline
        Write-Host "$text"  -ForegroundColor White -NoNewline
        Write-Host (" " * ($maxln - $text.Length-2)) -ForegroundColor White -NoNewline
        Write-Host "|" -ForegroundColor Cyan
    }
}

function WriteMsgs{
    param (
        [string]$text,
        [string]$color = "Yellow"
    )
    Write-Host $text -ForegroundColor Yellow
    Start-Sleep -Seconds 2
}

function LaunchOfficeProc{
    param(
      [string]$app, 
      [string]$targetpath,
      [bool]$readonly = $true
    )
    switch ($app) {
        "word" {
            $office = New-Object -ComObject Word.Application
            $office.Visible = $true
            $doc = $office.Documents.Open($targetpath, $false, $true) # Open as Read-Only
        }
        "excel" {
            $office = New-Object -ComObject Excel.Application
            $office.Visible = $true
            $doc = $office.workbooks.Open($targetpath, [ref]$false, [ref]$readonly) # Open as Read-Only
        }
        "powerpoint" {
            $office = New-Object -ComObject PowerPoint.Application
            $office.Visible = $true
            $doc = $office.Presentation.Open($targetpath, [ref]$false, [ref]$readonly) # Open as Read-Only
        }
    }
}

# --- Load Links from File ---
function Load-Links{
  # Check for file existence
  if (-not (Test-Path -Path $linksFile)) {
      Write-Host "Error: The '`links.txt' file was not found." -ForegroundColor Red
      Write-Host "Please create it and add your links." -ForegroundColor Yellow
      Read-Host "Press Enter to exit..."
      exit
  }
  Write-Host "Loading links from $linksFile" -ForegroundColor Yellow
  Start-Sleep -Seconds 1
  
  # Read the links and colors from the file
  $links = Get-Content -Path $linksFile | ForEach-Object {
      $parts = $_ -split ',', 5      
      [PSCustomObject]@{
          Name    = $parts[0].Trim()
          Group   = $parts[1].Trim()
          Path    = $parts[2].Trim().Trim('"')
          Color   = $parts[3].Trim()
          mode    = if($parts.Count -ge 5) {$parts[4].Trim().ToLower()} else {""}
      }
  }

  return $links
}  

function Show-Groups($link) {
  Clear-Host
  # Header in Cyan
  CreateBlocks -Text $head -inclline $true
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $head1 -inclline $false
  CreateBlocks -Text $head2 -inclline $false   
  CreateBlocks -Text $head3 -inclline $false   
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $head4 -inclline $false   
  CreateBlocks -Text "" -inclline $true
  
  $groups =$link.Group | Sort-Object -Unique
  for ($i=0; $i -lt $groups.Count; $i++) {
      Write-Host "$($i+1). $($groups[$i])" -ForegroundColor Green
  }

  Write-Host " "
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $msg -inclline $false
  CreateBlocks -Text "" -inclline $true
  Write-Host " "
  
  return $groups
}

function Show-LinksInGroup($link, $group) {
  Clear-Host
  # Header in Cyan
  $head = $head + " - Group: $group"
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $head -inclline $true
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $head1 -inclline $false
  CreateBlocks -Text $head2 -inclline $false   
  CreateBlocks -Text $head3 -inclline $false   
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $head4 -inclline $false   
  CreateBlocks -Text "" -inclline $true 

  $grouplinks = $link | Where-Object { $_.Group -eq $group }
  $columns = 3 # You can change this to 4 for a 4-column layout
  $linksPerColumn = [math]::Ceiling($grouplinks.Count / $columns)     
  # Write-Host "$($i+1). $($grouplinks[$i].Name)" -ForegroundColor $grouplinks[$i].Color
  for ($i=0; $i -lt $linksPerColumn; $i++) {
    # Display the menu in columns

    for ($j = 0; $j -lt $columns; $j++) {
        $index = $i + ($j * $linksPerColumn)
        
        if ($index -lt $grouplinks.Count) {
            $link = $grouplinks[$index]
            $displayText = "$($index + 1). $($link.Name)"
        
            # Format for consistent spacing
            $format = "{0,-" + ([math]::Ceiling($maxln / $columns)) + "}"
            #Test if the link exists or not and color it red if it doesn't
            if (-not (Test-Path -Path $link.Path)) {
                if ($link.Path -match "^https.*" -or $link.Path -match "^www.*") {
                    $displayText = "$($index + 1). $($link.Name) (URL)"
                    Write-Host ($format -f $displayText) -ForegroundColor $link.Color -NoNewline
                } else {
                    Write-Host ($format -f $displayText)  -ForegroundColor Red -NoNewline   
                }                    
            } else {
                if ($link.mode -eq "readonly") {
                    $displayText = "$($index + 1). $($link.Name) (*)"
                }
                Write-Host ($format -f $displayText) -ForegroundColor $link.Color -NoNewline
            }
        }
  
    }
   }
  Write-Host " "
  CreateBlocks -Text "" -inclline $true
  CreateBlocks -Text $msg -inclline $false
  CreateBlocks -Text "" -inclline $true
  Write-Host " "

  return $grouplinks
}

$allLinks = Load-Links
# --- Main loop ---
while ($true) {
    $group = Show-Groups $allLinks
    $grpChoice = Read-Host "Select a group by number (or 'exit' to quit)"
    $grpChoice = $grpChoice.Trim()

          #To reload the links file
    if ($grpChoice -eq "r") {
          Write-Host "Reloading the links file..." -ForegroundColor Yellow
          $allLinks = Load-Links
          continue
      }
    elseif ($grpChoice -eq "P") {
          Write-Host "Launching Powershell" -ForegroundColor Yellow
          Start-Process "powershell.exe"
          continue
      }
    elseif ($grpChoice -eq "b") {      
          Write-Host "Launching Browser" -ForegroundColor Yellow
          Start-Process "https://"
          continue
      }
    elseif ($grpChoice -eq "e") {      
          Write-Host "Launching Explorer" -ForegroundColor Yellow
          Start-Sleep -Seconds 1
          Start-Process "explorer.exe"
          continue
      }       
    elseif ($grpChoice -match '^\d+$'){
        $selGrp = $group[[int]$grpChoice - 1] 
    } else {
        $selGrp = -1
        Write-Host "Non-numeric input detected." -ForegroundColor Yellow
    }

  # inner loop to select links in the group
    while ($true) {
      $grplnks = Show-LinksInGroup $allLinks $selGrp
      $linkChoice = Read-Host "Enter your chouce (or 'back' to select another group, 'exit' to quit)"

      switch ($linkChoice) {
          "exit" {
              exit
          }
          "back" {
            $back = $true
            break
          }         
          default {
              if ($linkChoice -ne -1 -and $linkChoice -match '^\d+$' -and 
                  $linkChoice -gt 0 -and 
                  $linkChoice -le $grplnks.Count) {
                  try {
                       $targetPath = $grplnks[$linkChoice - 1].Path
                       if ($targetPath -like "*.lnk") {
                           $shell = New-Object -ComObject WScript.Shell
                           $shortcut = $shell.CreateShortcut($targetPath)
                           $targetPath = $shortcut.TargetPath
                       } else {
                           if ($targetPath -match "^https.*" -or $targetPath -match "^www.*") {
                              # It's a URL, open with default browser
                              Start-Process $targetPath
                              continue
                           } else {
                              if ($grplnks[$choicenum - 1].mode -eq "readonly") {                            
                                  if ($targetpath -match ".*\.docx$" -or $targetpath -match ".*\.doc$" -or $targetpath -match ".*\.rtf$") {
                                      # Open Word document in read-only mode
                                      WriteMsgs -text "Opening $targetpath in Word." -Color Yellow   
                                      LaunchOfficeProc("word", $targetpath)
                                      #Start-Process "winword.exe" -ArgumentList "/r", "`"$targetpath`""
                                      continue
                                  } elseif ($targetpath -match ".*\.xlsx$" -or $targetpath -match ".*\.xls$") {
                                      # Open Excel document in read-only mode
                                      WriteMsgs -text "Opening $targetpath in Excel." -Color Yellow                                   
                                      LaunchOfficeProc("excel", $targetpath)
  
                                      continue
                                  } elseif ($targetpath -match ".*\.pptx$" -or $targetpath -match ".*\.ppt$") {
                                      # Open PowerPoint document in read-only mode
                                      WriteMsgs -text "Opening $targetpath om Power Point." -Color Yellow                                     
                                      LaunchOfficeProc("excel", $targetpath)
                                      #Start-Process -FilePath "powerpnt.exe" -ArgumentList "/r", "`"$targetpath`""
                                      continue
                                  }
                              } else {
                                WriteMsgs -text "Opening $targetpath." -Color Yellow                               
                                Start-Process -FilePath "$targetpath"
                                continue
                              }
                           }
                       }
#                      Start-Process -FilePath "$selectLink.Path"
                  } catch {
                      WriteMsgs -text "An error occurred while trying to open the link. $($_.Exception.Message)" -Color Red
                  }
              } else {
                  WriteMsgs -text "Invalid choice. Please enter a valid number or 'Press Ctrl+C' to quit." 
  
              }
          }
      }
      if ($back -eq $true) {
          $back  = $false 
          break
      }      
    }
}
