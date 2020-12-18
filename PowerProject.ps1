Set-ExecutionPolicy -ExecutionPolicy Unrestricted

#Unregister any current events
Get-EventSubscriber | Unregister-Event

$watcher = New-Object System.IO.FileSystemWatcher
$watcher.IncludeSubdirectories = $false
$watcher.Path = 'C:\PowerPoint'
$watcher.EnableRaisingEvents = $true
$LogPath = 'C:\PowerPoint\Logs'

$action =
{
    Start-Transcript -OutputDirectory $LogPath -NoClobber
    $changetype = $event.SourceEventArgs.ChangeType
    Write-Host "$path was $changetype at $(get-date)"
    $path = $event.SourceEventArgs.FullPath
    Start-Sleep -s 3
    Write-Host "Stopping PowerPoint"
    Stop-Process -name POWERPNT -Force
    Start-Sleep -s 3
    Copy-Item $path -Destination 'C:\PowerPoint\Stage\current.pps' -Force -Verbose
    Write-Host "File ($path) Copied to Stage"
    Start-Sleep -s 3
    $stage = gci 'C:\Powerpoint\Stage'  | where { ! $_.PSIsContainer } | sort LastWriteTime | select -last 1
    $newpath = "C:\PowerPoint\Stage\current.pps" #+ $stage.Name
    Write-Host "Starting Powerpoint ($newpath)"
    Invoke-Item $newpath
}

Register-ObjectEvent $watcher 'Created' -Action $action

##Start-Sleep -s 3600
##.\PowerPoint.ps1