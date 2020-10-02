$script:manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json
$script:driversPath = ".\drivers"
$script:wimPath = ".\wim"
$script:unattendPath = ".\unattend"
$script:machineModel = ""

function main {
  Show-ApplicationTitle
  Get-MachineSerialNumber
  Get-MachineModel
  Get-TasksMenuForDevice
  Get-FinishOptions
}

function Get-TasksMenuForDevice {
  [bool]$isDeviceInManifest = Test-ManifestForModel
 
  $counter = 0
  [System.Collections.ArrayList]$taskCollection = @()

  Write-Host "`nTask Menu" -ForegroundColor Magenta
  Write-Host "+++++++++" -ForegroundColor Gray
  
  if ($isDeviceInManifest) {
    
    foreach ($property in $manifest.'tasks'.PSObject.Properties) { 
    
      $taskCollection.Add($property.Name) | Out-Null
  
      $taskName = $property.Name

      if ($manifest.'models'.$script:machineModel -contains $taskName) {
        Write-Host "[" -ForegroundColor DarkGray -NoNewline; Write-Host "$counter" -ForegroundColor Yellow -NoNewline; Write-Host "] " -ForegroundColor DarkGray -NoNewline; Write-Host $taskName -ForegroundColor Yellow
      } else {
        Write-Host "[$counter] $taskName" -ForegroundColor DarkGray
      }
      
      $counter++
    }
  } else {
    foreach ($property in $manifest.'tasks'.PSObject.Properties) { 
    
      $taskCollection.Add($property.Name) | Out-Null
  
      $taskName = $property.Name
      Write-Host "[$counter] $taskName" -ForegroundColor DarkGray
  
      $counter++
    }
  }
  Write-Host ""

  $taskCollection.Add("Capture Image") | Out-Null
  Write-Host "[$counter] Capture Image" -ForegroundColor Cyan

  $counter++
  $taskCollection.Add("Shutdown Windows PE Gracefully") | Out-Null
  Write-Host "[$counter] Shutdown Windows PE Gracefully" -ForegroundColor Green

  $selectedOption = Read-Host "`nSelect an option and hit enter or hit CTRL-C to exit`n"
  New-ImageTask($selectedOption)
}

function New-ImageTask($selectedOption) {

  if($taskCollection[[int]$selectedOption] -eq "Capture Image") {
    New-CaptureImageTask
    exit
  }

  if($selectedOption -match '\d' -ne 1) {
    Write-Host "Invalid option. Invalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-TasksMenuForDevice
    return
  }

  $taskName = $taskCollection[[int]$selectedOption]

  if([string]::IsNullOrWhiteSpace($taskName)){
    Write-Host "Invalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-TasksMenuForDevice
    return
  }

  if ($taskName -eq "Shutdown Windows PE Gracefully") {
    wpeutil shutdown
  }

  Test-TaskName($taskName)

  $wimFile = Test-TaskForWimFile($taskName)

  [bool]$isDrivers = Test-TaskForDrivers($taskName) 
  [bool]$isUnattendFile = Test-UnattendFile($taskName)
  
  Set-InternalDrivePartitions
  Set-PowerSchemeToHigh

  Write-Host "`n[ Beginning image task... ]" -ForegroundColor Cyan
  Expand-WindowsImage -ImagePath "$script:wimPath\$wimFile" -index 1 -ApplyPath "w:\"
  Write-Host "`n[ >> Image task complete << ]" -ForegroundColor Yellow

  if([bool]$isDrivers -eq 1){
    $version = $script:manifest.'tasks'.$taskName.'drivers'
    Set-DriversOnImagedPartition($version)
  }

  if([bool]$isUnattendFile) {
    $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'
    Set-UnattendFile($unattendFile)
  }
  Set-BootLoader
  Get-MachineSerialNumber
  Write-Host "`n[ >> All Good <<]" -ForegroundColor Magenta
}

function New-CaptureImageTask {
  $counter = 0
  [System.Collections.ArrayList]$taskCollection = @()
  [System.Collections.ArrayList]$selectableOptions = @()
  
  Write-Host "`n[ Capture Image Selected... ]`n" -ForegroundColor Cyan
  Write-Host "Volumes" -ForegroundColor Magenta
  Write-Host "-------"

  Get-Volume | ForEach-Object {
    if($_.DriveLetter.length -gt 0) {

      $taskCollection.Add($_.DriveLetter) | Out-Null
      $selectableOptions.Add($counter) | Out-Null

      $size = [math]::Round($_.Size/1GB)
      $readableSize = "$size GB" 
      Write-Host "[$counter] " $_.DriveLetter $_.FileSystem $readableSize $_.HealthStatus $_.FileSystemLabel -ForegroundColor DarkGray
      $counter++  
    }
    
  }

  do {
    $driveOption = Read-Host "`nVolume to capture`n"
  } while ($selectableOptions -notcontains $driveOption)
  
  do {
    Write-Host "Enter a file name for the image. Allowed characeters [a-z,0-9,-]`n" -ForegroundColor DarkYellow
    $fileName = Read-Host "Filename for image"
  } while ($fileName -notmatch '^[a-z0-9.-]{1,200}$')

  
  
  $vol = $taskCollection[[int]$driveOption]
  $z = ":\"
  $capturePath = ($vol)+($z)

  $imgPath = "$script:wimPath\$fileName"

  if(Test-Path $imgPath -PathType Leaf) {
    Write-Host "`n[ Error: File with name '$imgPath' already exists. Try again. ]" -ForegroundColor Red
    New-CaptureImageTask
    return
  }

  Write-Host "[ Beginning Capture Task... ]" -ForegroundColor Cyan 
  New-WindowsImage -ImagePath $imgPath -CapturePath $capturePath -Name "Windows"
  Write-Host "`n[ >> Capture task completed << ]" -ForegroundColor Yellow
  
  Get-FinishOptions
}

function Get-FinishOptions {
  Write-Host "`nPress any key to shutdown gracefully or ESC key to exit" -ForegroundColor Green
  
  $key = [console]::ReadKey()
  if ($key.Key -ne '27') {
    wpeutil Shutdown
  } else {
    exit
  }
}

function Test-TaskName($taskName){
  if(![bool]($script:manifest.'tasks'.PSobject.Properties.name -Match $taskName)){
    Write-Host "[ Error: Selected taskname: '$taskName' not found in manifest 'tasks' object. Check manifest tasks under '$script:machineModel'. Script will now exit. ]`n" -ForegroundColor DarkRed
    exit
  }
}

function Set-UnattendFile ($unattendFile){
  Copy-Item "$script:unattendPath\$unattendFile" -Destination "w:\windows\panther\unattend.xml"
  Write-Host "[ >> Copied $unattendFile to windows\panther... << ]" -ForegroundColor Yellow
}

function Set-DriversOnImagedPartition($version) {
  Write-Host "`n[ Injecting drivers... ]" -ForegroundColor Cyan
  Add-WindowsDriver -Path "w:\" -Driver "$script:driversPath\$version\$script:machineModel\" -Recurse -ForceUnsigned | Out-Null
  Write-Host "[ >> Finished injecting drivers << ]" -ForegroundColor Yellow
}

function Set-BootLoader{
  Write-Host "`n[ Setting Bootloader... ]`n" -ForegroundColor Cyan

  if($(Get-ComputerInfo).BiosFirmwareType -eq "Uefi"){
    $efiCommand = "bcdboot w:\windows"
    Invoke-Expression $efiCommand
    Invoke-Expression $efiCommand
    Invoke-Expression $efiCommand
  } else {
    try {
      Write-Host "`n[ Checking Windows BootLoader for stale entries... ]" -ForegroundColor Cyan
      Clear-BootLoaderEntries
    } catch {
      Write-Host "[ >> No entries found << ]" -ForegroundColor Yellow
    }
    
    $biosCommand = "bcdboot w:\windows /s w: /f BIOS"
    Invoke-Expression $biosCommand
  }
}

function Get-MachineModel{
	Write-Host "[ Retrieving computer model... ]`n" -ForegroundColor Cyan
	$script:machineModel = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty "Model"
	Write-Host "[ >> Found machine model: $script:machineModel << ]" -ForegroundColor Yellow
}

function Get-MachineSerialNumber{
	$serialNumber = get-ciminstance win32_bios | Select-Object -ExpandProperty serialnumber
	Write-Host "`n[ >> Machine Serial Number: $serialNumber << ]`n" -ForegroundColor Yellow
}
function Test-ManifestForModel{
  if(![bool]($script:manifest.'models'.PSobject.Properties.name -Match $script:machineModel)){
    Write-Host "`n[ Warning: Machine model not found in manifest. ]" -ForegroundColor DarkYellow
    0
  } else {
    Write-Host "`n[ Found Machine model in manifest. Tasks in " -ForegroundColor DarkYellow -NoNewline; Write-Host "yellow" -ForegroundColor Yellow -NoNewline; Write-Host " are recommended for this device. ]`n" -ForegroundColor DarkYellow
    1
  }
}


function Test-TaskForWimFile($taskName){
  if(![bool]($script:manifest.'tasks'.$taskName.PSobject.Properties.name -Match "wim")){
    Write-Host "[ Error: Manifest task named '$taskName' missing key named 'wim' type[string]. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  $wimFile = $script:manifest.'tasks'.$taskName.'wim'

  if ([string]::IsNullOrWhiteSpace($wimFile)){
    Write-Host  "[ Error: No WIM file defined for '$imageName' in manifest. Script will now exit. ]`n"  -ForegroundColor DarkRed
    exit
  }

  $wimFile
}

function Test-TaskForWindowsVersionKey($taskName){
  if(![bool]($script:manifest.'tasks'.$taskName.PSobject.Properties.name -Match "drivers")){
    Write-Host "[ Error: Manifest task named '$taskName' missing key named 'drivers' type[string]. Example 1709, 1803, 1909 etc. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }
}


function Test-TaskForDrivers($taskName){
  $version = $script:manifest.'tasks'.$taskName.'drivers'

  if([string]::IsNullOrWhiteSpace($version)){
    0
    return
  }

  Write-Host "`n[ Checking for drivers in .\drivers\$version folder... ]`n" -ForegroundColor Cyan

  [bool]$isDriversFolder = Test-Path -Path "$script:driversPath\$version"
  if($isDriversFolder -ne 1){
    New-Item -Path "$script:driversPath\$version" -ItemType Directory | OUT-NULL
    Write-Host "[ >> Created missing root drivers folder '$script:driversPath\$version' << ]" -ForegroundColor Yellow
  }

  [bool]$isDriversForModel = Test-Path -Path "$script:driversPath\$version\$script:machineModel"
  if($isDriversForModel -ne 1){
    New-Item -Path "$script:driversPath\$version\$script:machineModel" -ItemType Directory | OUT-NULL
    Write-Host "[ >> Created drivers folder for device model: '$script:machineModel' << ]" -ForegroundColor Yellow
  }

  $driversDirectoryInfo = Get-ChildItem -Path "$script:driversPath\$version\$script:machineModel" | Measure-Object
  if($driversDirectoryInfo.count -eq 0) {
    Write-Host "`n[ WARNING: There were no drivers detected for this device ]" -ForegroundColor DarkYellow

    $selectedOption = Read-Host "`nWould you like to continue anyway? (y/n)`n"
    if ($selectedOption.ToLower() -eq 'y'){
      Write-Host "`n[ >> Continuing without drivers << ]`n" -ForegroundColor Yellow
      0
    } else {
      Write-Host "`n[ Script Exiting... ]`n" -ForegroundColor Cyan
      exit
    }
  } else {
    Write-Host "`n[ >> Found drivers folder for $script:machineModel << ]" -ForegroundColor Yellow 
    1
  }

}

function Test-UnattendFile ($taskName){
  $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'

  if(![string]::IsNullOrWhiteSpace($unattendFile)){
    Test-UnattendFolder

    Write-Host "`n[ Checking for specified unattend file: '$unattendFile'... ]" -ForegroundColor Cyan

    if(Test-Path -Path "$script:unattendPath\$unattendFile" -PathType leaf){
      Write-Host "`n[ >> Found $unattendFile << ]" -ForegroundColor Yellow
      1
    } else {
      Write-Host "`n[ Warning: Could not locate unattend file: $unattendFile ]`n" -ForegroundColor DarkYellow
      $selectedOption = Read-Host "Would you like to continue anyway? (y/n)"
      if($selectedOption.ToLower() -eq 'y') {
        0
      } else {
        Write-Host "`n[ Script Exiting... ]`n" -ForegroundColor Cyan
        exit
      }
    }
  }
}

function Test-UnattendFolder {
  [bool] $isUnattendFolder = Test-Path -Path "$script:unattendPath"
  if($isUnattendFolder -ne 1){
    New-Item -Path $script:unattendPath -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing unattend folder << ]" -ForegroundColor Yellow
  }
}

function Test-WimFolderForImageFile($wimFile){
  Write-Host "[ Checking WIM exists... ]" -ForegroundColor Cyan
  if (Test-Path -Path "$script:wimPath\$wimFile" -PathType leaf) {
    Write-Host "[ >> Found $wimFile << ]" -ForegroundColor DarkYellow
  } else {
    Write-Host "[ Error: Could not locate '$wimFile' in wim folder. Compare manifest and wim folder. Script will now exit ]`n" -ForegroundColor DarkRed
    exit
  }
}
function Set-InternalDrivePartitions{
  $internalDisk = Get-InternalDiskNumber

  $efiLayout = @"
select disk $internalDisk
clean
convert gpt
create partition efi size=260
format fs=fat32 quick
assign letter=s
create partition msr size=128
create partition primary
format fs=ntfs quick label=WINDOWS
assign letter=W
"@

  $mbrLayout = @"
select disk $internalDisk
clean
convert mbr
create partition primary
format fs=ntfs quick label=WINDOWS
active
assign letter=W
"@

  New-Item -Name disklayout.txt -ItemType File -Force | OUT-NULL
  $message = ""

  if($(Get-ComputerInfo).BiosFirmwareType -eq "Uefi"){
    Add-Content -Path disklayout.txt $efiLayout
    $message = "`n[ Writing efi disk layout... ]"
  } else {
    $message = "`n[ Writing mbr/bios disk layout... ]"
    Add-Content -Path disklayout.txt $mbrLayout
  }

  Write-Host $message -ForegroundColor Cyan

  $command = "diskpart /s disklayout.txt"
  Invoke-Expression $command

  Write-Host "`n[ >> Disk layout complete << ]" -ForegroundColor Yellow
}

function Get-InternalDiskNumber {
  #Get-WmiObject win32_diskdrive | Where-Object {Write-Host $_.MediaType}
  $internalDisks = Get-WmiObject win32_diskdrive | Where-Object {$_.MediaType -ne "External hard disk media" -and $_.MediaType -ne "Removable Media" -and $_.InterfaceType -ne "USB"}
  
  $InternalDiskNumbers = @()

  $internalDisks | ForEach-Object {
    $InternalDiskNumbers += $_.DeviceID.Substring($_.DeviceID.Length -1)
  }
  if ($InternalDiskNumbers.Count -gt 1){
    Write-Host "[ Error: Found more than one internal disk. Script must exit ]`n" -ForegroundColor DarkRed
    exit
  }

  if ($InternalDiskNumbers.Count -eq 0) {
    Write-Host "[ Error: No internal disk found for imaging. Script must exit ]`n" -ForegroundColor DarkRed
    exit
  }
  return $InternalDiskNumbers[0]
}

function Clear-BootLoaderEntries {
  $bcdOutput = (bcdedit /v) -join "`n"
  $entries = New-Object System.Collections.Generic.List[pscustomobject]]
  ($bcdOutput -split '(?m)^(.+\n-)-+\n' -ne '').ForEach({
    if ($_.EndsWith("`n-")) { # entry header 
      $entries.Add([pscustomobject] @{ Name = ($_ -split '\n')[0]; Properties = [ordered] @{} })
    } else {  # block of property-value lines
      ($_ -split '\n' -ne '').ForEach({
        $propAndVal = $_ -split '\s+', 2 # split line into property name and value
        if ($propAndVal[0] -ne '') { # [start of] new property; initialize list of values
          $currProp = $propAndVal[0]
          $entries[-1].Properties[$currProp] = New-Object Collections.Generic.List[string]
        }
        $entries[-1].Properties[$currProp].Add($propAndVal[1]) # add the value
      })
    }
  })

  $entries | ForEach-Object {
    if ($_.Name -ne "Windows Boot Manager") {
      $bootLoaderIdentifier = $_.Properties['identifier']
      Write-Host "[ >> Removed entry: $bootLoaderIdentifier << ]" -ForegroundColor Yellow
      $command = "bcdedit /delete '$bootLoaderIdentifier'"
      Invoke-Expression $command
    }
  }
}

function Set-PowerSchemeToHigh {
  Write-Host "`n[ Setting Power Scheme to HIGH... ]" -ForegroundColor Cyan
  powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
}

function Show-ApplicationTitle {
  $title =@"
                              __
  The                /\    .-" /
    - AG -          /  ; .'  .'
  Disk Imager      :   :/  .'
                    \  ;-.'
        .--""""--..__/     `.
      .'           .'    `o  \
    /                    `   ;
    :                  \      :
  .-;        -.         `.__.-'
:  ;          \     ,   ;
'._:           ;   :   (
    \/  .__    ;    \   `-.
  bug ;     "-,/_..--"`-..__)
      '""--.._:

[ https://github.com/telekram/ag-disk-imager ]
"@

  write-host " "
  Write-Host $title -ForegroundColor DarkYellow
}

main