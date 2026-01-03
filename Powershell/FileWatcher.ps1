$filewatch = New-Object System.IO.fileSystemWatcher

$filewatch.Path = "c:\temp\"
$filewatch.IncludeSubdirectories = $true

$filewatch.EnableRaisingEvents = $true

$writeaction = {$path=$Event.SourceEventArgs.FullPath
                $changeType=$Event.SourceEventArgs.ChangeType
                $logline = "$(Get-Date),$changeType, $path" 
                Add-content "C:\Users\userid\Documents\FileWatcherlog.log"  value $logline
                }

Register-ObjectEvent $filewatch "Created" -Action $writeaction
Register-ObjectEvent $filewatch "Changed" -Action $writeaction
Register-ObjectEvent $filewatch "Deleted" -Action $writeaction
Register-ObjectEvent $filewatch "Renamed" -Action $writeaction

while ($true) {sleep(5)}
