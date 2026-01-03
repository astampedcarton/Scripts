# Verify the file before uploading
clear

$zippath = "c:\temp\p21.zip"

$reqfiles = @(
  "ae.xpt",
  "ex.xpt",
  "dm.xpt",
  "adrg.pdf",
  "define.xml",
  "define.xsl"
)

Add-Type -AssemblyName System.IO.Compression.FileSystem
$zip = [System.IO.Compression.ZipFile]::OpenRead($zipPath)
$zipentry1 = $zip.Entries | Select-Object -ExpandProperty FullName

$zipentry = $zip.Entries | 
    ForEach-Object {
        $entry = [System.IO.Path]::GetFileName($_.FullName)
        if ($entry -in $reqfiles){
          [PSCustomObject]@{
            Name         = $_.FullName
            FileName     = [System.IO.Path]::GetFileName($_.FullName)
            SizeBytes    = $_.Length
            LastModified = $_.LastWriteTime
          }
        }
} 

$zip.Dispose()

$missingfiles = $reqfiles | Where-Object { $_ -notin $zipentry1 }
$fndfiles = $reqfiles | Where-Object { $_ -in $zipentry1}

if ($missingfiles.Count -eq 0) {
  Write-Host "Required files are present:  $($fndfiles -join ' ,')"
} else {
  Write-Host "Required files are missing: $($missingfiles -join ' ,')"
  Write-Host "Required files are present:  $($fndfiles -join ' ,')"
  $zipentry | Format-Table -Autosize
}

