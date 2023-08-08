# Set EnvPath: OneNoteBackup=C:\Users\Name\AppData\Local\Microsoft\OneNote\16.0\Sicherung
# Set OneNote backup interval to 1 day.

$targetDiskName='Levin_Joller'
$targetDiskPath='D:\'
Write-host "[1] ---Health status---"
Write-host "[1] Check disk connection"
try{
    Get-PnpDevice -FriendlyName $targetDiskName -Class 'WPD' -Status OK -erroraction Stop
}catch{
    Write-Error $Error[0].Exception.Message
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[1] OK" -f Green

$oneNoteTarget=$targetDiskPath+"$((Get-Date).ToString('yyMMdd'))_oneNote"
Write-host "[2] ---OneNote backup---"
Write-host "[2] Check target folder"
if(Test-Path $oneNoteTarget){
    Write-host "Folder Exists! Terminate." -f Red
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[2] OK" -f Green

Write-host "[2] Create folder"
try{
    New-Item -ItemType Directory -Path $oneNoteTarget -erroraction  stop
}catch{
    Write-Error $Error[0].Exception.Message
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[2] OK" -f Green

Write-host "[3] Copy data"
try{
    Copy-Item -Path $env:OneNoteBackup -Destination $oneNoteTarget -Recurse -erroraction  stop
}catch{
    Write-Error $Error[0].Exception.Message
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[3] OK" -f Green

Write-host "[4] ---OneDrive backup---"
$oneDriveTarget=$targetDiskPath+"$((Get-Date).ToString('yyMMdd'))_oneDrive"
Write-host "[4] Check target folder"
if(Test-Path $oneDriveTarget){
    Write-host "Folder Exists! Terminate." -f Red
    Read-Host -Prompt "Press Enter to exit"
    exit
}

Write-host "[4] Create folder"
$oneDriveSource=$env:OneDrive+'\*'
try{
    New-Item -ItemType Directory -Path $oneDriveTarget -erroraction  stop
}catch{
    Write-Error $Error[0].Exception.Message
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[4] OK" -f Green

Write-host "[4] Copy data"
$excludes='BMBT22A-sharePoint','Anlagen','Aufnahmen','Microsoft Teams-Chatdateien','Levin @ sluz'
try{
Copy-Item -Path $oneDriveSource -Destination $oneDriveTarget -Recurse -Exclude $excludes -erroraction  stop
}catch{
    Write-Error $Error[0].Exception.Message
    Read-Host -Prompt "Press Enter to exit"
    exit
}
Write-host "[4] OK" -f Green
Write-host "Backup successfully created!" -f Green
Read-Host -Prompt "Press Enter to exit"