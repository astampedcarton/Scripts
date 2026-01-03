Clear-Host

$tmploc = "c:\temp\tmpimg\"
$outfolder = "c:\temp\results"
$rundt = Get-Date -Format "yyyymmdd"
$logfile = "$outfolder" + "\log_$rundt" + ".txt"

#.Net assemblies for the image processing portion
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Windows.Forms

#general setup of what will be used to annotate that the figure has changed
$font = New-Object System.Drawing.Font("Arial", 8)
#$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::FromName($color))
#$pen = New-Object System.Drawing.Pen ([System.Drawing.Color]::Red, 1)
$brush = New-Object System.Drawing.SolidBrush ([System.Drawing.Color]::Red)
$markerSize = 20
$bdebug = $true

#Create the temp folder
if (Test-path -path "$tmploc") {
  write-output "Temp Folder already exists "
} else {
  mkdir -Path "$tmploc"
}

if (Test-path -path "$outfolder") {
  write-output "Results Folder already exists"
} else {
  mkdir -Path "$outfolder"
}

write-output "Results will be saved too $outfolder"
#PNG Header identifier DO NOT CHANGE ANYTHING HERE
$pnghead = @(137, 80, 78, 71, 13, 10, 26, 10)
#$pngstr = "137 80 78 71 13 10 26 10"
#$pnghex = @(89, 50, 4E, 47, 0D, 0A, 1A, 0A)

#HEX Codes for the different image types these must be the first few characters
#==============================================================================
#PNG  : 89504e470d0a
#JPEG : FFD8FF
#GIF  : 47494638
#BMP  : 424D

function Debuglog{
     param(
     [string]$text
     )
#  if ($bdebug) {
     write-host "Debugging INFO: $text"
#  }
}


function Create-ImageFiles{
        param ([int]$icnt, [string]$tag, [string]$hex)
#    write-output "Fig seq: $icnt"
    #Create image files;
    $str = $hex.Substring(0,$icnt) + $hex.Substring($icnt)
    $slen = $str.Length
    $str = $str -replace '\s', ''
  
    if ($slen % 2 -ne 0) {
        Debuglog -text "ERROR: Hex string has an odd number of characters. Cannot convert safely."
        return        
    }

    #Convert the data to bytes
    $bytes = for ($j = 0; $j -lt $slen; $j += 2) {
      [Convert]::ToByte($str.Substring($j, 2), 16)
    }

    #For Debugging;
    $fcnt = $bytes.Count
    #Debuglog -text "F Cont: $fcnt"

    #Extract the the location of the potential corrupt text.
    #The string must start with what png header
    for ($i = 0; $i -le $bytes.Count - $pnghead.Count; $i++) {
        $match = $true
        for ($j = 0; $j -lt $pnghead.Count; $j++) {
            if ($bytes[$i + $j] -ne $pnghead[$j]) {
                $match = $false
                break
            }
        }
        if ($match) {
            $istr = $i
            break
        }
    }

#    write-output "Byte St: $istr"        
    $bytepng = $bytes[$istr..($bytes.Count - 1)]
    $outpng = $tmploc + "$tag" + "image" + $icnt + ".png"

    #Debuglog -text "Outpng file: $outpng"

    [IO.File]::WriteAllBytes("$outpng", $bytepng)

  #Return the file name. We need to load these later to draw the marker
  return $outpng
}

#Extract the figure information
function Get-RtfPictHex {
    param (
        [string]$Path,
        [string]$tag
    )

    $content = Get-Content -Path $Path -Raw
    $pattern = '{\\pict[\s\S]*?}'  # non-greedy pict block
    $pattern = '\\shppict{\\pict\\pngblip\\picwgoal\d+\\pichgoal\d+\s+([\da-fA-F\s]+?)}'

    $patmatches = [regex]::Matches($content, $pattern)

    #Because more than 1 figure could be present in a file we need to 
    #create the individual images.
    $figcnt = 0
    foreach ($match in $patmatches) {
        $figcnt++
        #$filename = $tmploc + "$tag_image_hex$figcnt.png";
        $hex = $match.Groups[1].Value -replace '\s+', ''
#        $hex = ($match.Value -replace '[^\da-fA-F]', '').ToLower()
        $hex = ($hex  -replace '[^\da-fA-F]', '').ToLower()
        #Create the images from the files thats in the RTF
        $null = Create-ImageFiles -icnt $figcnt -tag $tag -hex $hex        
    }
    Debuglog -text "Get-RTFPicHext: Figure count from : $figcnt"
    return $figcnt
}

function Compare-RtfImages {
    param (
        [string]$oldfile,
        [string]$newfile
    )
    #TODO: Need to add a check that the tmploc is not part of any of the current locations 
    #Because we can debug we need to make sure that the tmpimg folder is clean
    $tmploccl = "$tmploc" + "\" + "*.png"
    Remove-Item -Path "$tmploccl"
    
    #Extract the picture portion from the RTF.
    #Because we only want to check if there's a differences or not
    #we compare the whole results.
    if ((Test-Path -path "$oldfile") -and (Test-Path -path "$newfile")) {
      #Return the nr of images found in the different files
      $oldcnt = Get-RtfPictHex -Path $oldfile -tag "old"
      $newcnt = Get-RtfPictHex -Path $newfile -tag "new"
      Debuglog -text "Old Figure count: $oldcnt"
      Debuglog -text "New Figure count: $newcnt"
    } else {
      Debuglog -text "One or both of the provided files was not found."
      exit
    }

    $newfname = Split-Path -Path $newfile -Leaf
    $rtfName = $newfname

    if ($oldcnt -eq $newcnt){   
      #Load the image files
      $outdiff = 0
      $rtfbox = New-Object System.Windows.Forms.RichTextBox
      for ($img=1; $img -lt $newcnt; $img++){
        #Load the images one by one to check what compares and what not.
        #Even if the images files were rearranges it will be picked up as a
        #non-compare. The images will not be compared against each other 
        #TODO: Consider if this might be something needed or not.
        $imgofl = "$tmploc" + "oldimage"+$img+".png"   
        $imgnfl = "$tmploc" + "newimage"+$img+".png"

        Debuglog -text "Starting on: $imgofl"

        $oldbmp = [System.Drawing.Bitmap]::FromFile($imgofl)
        $newbmp = [System.Drawing.Bitmap]::FromFile($imgnfl)

        #Make sure that the dimensions between the two figures are the same            
        if ($oldbmp.Width -ne $newbmp.Width -or $oldbmp.Height -ne $newbmp.Height) {
            Write-Error "Images have different dimensions"
            return
        }
        #Flag the difference on the 2nd image this would be the new one
        $graphics = [System.Drawing.Graphics]::FromImage($newbmp)
        # Scan pixels
        $diffnd = $false
        for ($x = 0; $x -lt $oldbmp.Width; $x++) {
            for ($y = 0; $y -lt $oldbmp.Height; $y++) {
                $oldcolr= $oldbmp.GetPixel($x, $y)
                $newcolr = $newbmp.GetPixel($x, $y)
                if ($oldcolr.ToArgb() -ne $newcolr.ToArgb()) {
                    Debuglog -text "Difference found at ($x, $y)"
                    $diffnd = $true
                    $outdiff++
                    #Draw a yellow marker (square)
                    #$graphics.DrawRectangle($pen, $x, $y, $markerSize, $markerSize)
                    $graphics.FillRectangle($brush, $x, $y, $markerSize, $markerSize)
                    $graphics.DrawString("There are diffs to checks", $font, $brush, 5, 1)
                    break
                    #Break once the first difference is encountered. Because we are evaulating images 
                    #there's no sense in continuing as from that point all pixels will be different.
                }
            }
            if ($diffnd) { 
              write-output "Differences found on image $img."
              $result = $result + "Differences found on image $img".Padright(30)
              break 
            }
        }
        
        #paste to the RTF box 
        # for ($nimg=1; $nimg -lt $newcnt; $nimg++){
        $imgnfl = "$tmploc" + "newimage"+$nimg+".png"
        [System.Windows.Forms.Clipboard]::SetImage($newbmp)
        $rtfbox.Paste()
        # }

      } #For close
      
      if ($outdiff -eq 0) {
          write-output "Checked on a pixel level. No differences found."
          $result = $result + "No diff on pixel level".Padright(30)
      } else {
          # Save the result                    
          $outresrtf = $outfolder + "\$rtfName"+ ".rtf"
          debuglog -text "Output to rtf: $outresrtf"

          #Create a richtext box to copy to and save as an rtf
          # $rtfbox = New-Object System.Windows.Forms.RichTextBox
          # for ($nimg=1; $nimg -lt $newcnt; $nimg++){
          #    $imgnfl = "$tmploc" + "newimage"+$nimg+".png"
          #    [System.Windows.Forms.Clipboard]::SetImage($newbmp)
          #    $rtfbox.Paste()
          # }

          $rtfbox.SaveFile($outresrtf, [System.Windows.Forms.RichTextBoxStreamType]::RichText)
          $result = $result + ". Re-compare figure".Padright(20)
          Debuglog -text "Marker image saved to $outfolder"
      }
      # Clear all handle to the images
      if ($graphics) {$graphics.Dispose()}
      if ($oldbmp) {$oldbmp.Dispose()}
      if ($newbmp) {$newbmp.Dispose()}

      #Clean up the temp directory
      if (-not $bdebug) {
        Remove-Item -path "$oldimgpath"
        Remove-Item -path "$newimgpath"
        Debuglog -text "Removed: $oldimgpath, $newimgpath "
      }
    } else {
       write-output "Old File: $oldfname New File: $newfname : Different nr of images found on RTF. Recompare is needed."
       write-output "Old Count: $oldcnt New Count: $newcnt"
    }
    write-output "Old File: $oldfname New File: $newfname : $result"
}


# Example usage:
#Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple.rtf" -newfile "C:\Temp\figures\figure2\fig-simple.rtf" | Out-File -FilePath $logfile  
# Compare-RtfImages -oldfile "C:\Temp\figures\figures1 (2).rtf" -newfile "C:\Temp\figures\figures2 (2).rtf" | Out-File -FilePath $logfile  
Compare-RtfImages -oldfile "C:\Temp\figures\figures2 (3).rtf" -newfile "C:\Temp\figures\figures3 (2).rtf" | Out-File -FilePath $logfile  
# Compare-RtfImages -oldfile "C:\Temp\figures\figures2 (2).rtf" -newfile "C:\Temp\figures\figures3 (2).rtf" | Out-File -FilePath $logfile  

#Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple2.rtf" -newfile "C:\Temp\figures\figure2\fig-simple2.rtf"
#Compare-RtfImages -oldfile "C:\Temp\figures\figure1\fig-simple3.rtf" -newfile "C:\Temp\figures\figure2\fig-simple3.rtf"
