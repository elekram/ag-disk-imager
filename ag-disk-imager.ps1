$script:manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json
$script:driversPath = ".\drivers"
$script:wimPath = ".\wim"
$script:unattendPath = ".\unattend"
$script:machineModel = ""

function main {
  Show-ApplicationTitle
  Get-MachineSerialNumber
  Get-MachineModel
  Test-ManifestForModel
  Get-TasksMenuForDevice
  Get-MachineSerialNumber
  Write-Host "[ >> All Good <<]" -ForegroundColor Green
}

function Get-TasksMenuForDevice {
  $counter = 0
  [System.Collections.ArrayList]$taskCollection = @('0')

  foreach ($item in $manifest.'models'.$script:machineModel) {
    $taskCollection.Add($item) | Out-Null
  }

  Write-Host "`nMenu" -ForegroundColor Magenta
  Write-Host "++++" -ForegroundColor Gray

  Foreach ($task in $taskCollection) {
    if ($task -ne '0'){
      Write-Host "[$counter] Task: $task" -ForegroundColor DarkYellow
    }
    $counter++
  }

  $selectedOption = Read-Host "`nSelect image option and hit enter or punch CTRL-C to exit`n"
  New-ImageTask($selectedOption)
}

function New-ImageTask($selectedOption) {
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

  Test-TaskName($taskName)

  $wimFile = Test-TaskForWimFile($taskName)

  [bool]$isDrivers = Test-TaskForDrivers($taskName) 

  [bool]$isUnattendFile = Test-UnattendFile($taskName)

  Set-InternalDrivePartitions
  Set-PowerSchemeToHigh

  Write-Host "[ Beginning image task... ]" -ForegroundColor Cyan
  Expand-WindowsImage -ImagePath "$script:wimPath\$wimFile" -index 1 -ApplyPath "w:\"
  Write-Host "[ >> Completed image task << ]" -ForegroundColor DarkYellow

  if([bool]$isDrivers -eq 1){
    $version = $script:manifest.'tasks'.$taskName.'wim'
    Set-DriversOnImagedPartition($version)
  }

  if([bool]$isUnattendFile) {
    $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'
    Set-UnattendFile($unattendFile)
  }
  Set-BootLoader
}

function Test-TaskName($taskName){
  if(![bool]($script:manifest.'tasks'.PSobject.Properties.name -Match $taskName)){
    Write-Host "[ Error: Selected taskname: '$taskName' not found in manifest 'tasks' object. Script will now exit. ]`n" -ForegroundColor DarkRed
    exit
  }
}

function Set-UnattendFile ($unattendFile){
  Copy-Item "$script:unattendPath\$unattendFile" -Destination "w:\windows\panther\unattend.xml"
  Write-Host "[ >> Copied $unattendFile to windows\panther... << ]" -ForegroundColor DarkYellow
}

function Set-DriversOnImagedPartition($version) {
  Write-Host "`n[ Injecting drivers... ]" -ForegroundColor Cyan
  Add-WindowsDriver -Path "w:\" -Driver "$script:driversPath\$version\$script:machineModel" -Recurse -ForceUnsigned | Out-Null
  Write-Host "[ >> Finished injecting drivers << ]" -ForegroundColor DarkYellow
}

function Set-BootLoader{
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
      Write-Host "[ >> No entries found << ]" -ForegroundColor DarkYellow
    }
    
    $biosCommand = "bcdboot w:\windows /s w: /f BIOS"
    Invoke-Expression $biosCommand
  }
}

function Get-MachineModel{
	Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
	$script:machineModel = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty "Model"
	Write-Host "[ >> Found machine model $script:machineModel << ]" -ForegroundColor DarkYellow
}

function Get-MachineSerialNumber{
	$serialNumber = get-ciminstance win32_bios | Select-Object -ExpandProperty serialnumber
	Write-Host "`n[ Machine Serial Number: $serialNumber ]`n" -ForegroundColor Yellow
}
function Test-ManifestForModel{
  if(![bool]($script:manifest.'models'.PSobject.Properties.name -Match $script:machineModel)){
    Write-Host "[ Error: Machine model not found in manifest. Script will now exit. ]`n" -ForegroundColor DarkRed
    exit
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

function Test-TaskForDrivers($taskName){
  $version = $script:manifest.'tasks'.$taskName.'version'
  $drivers = $script:manifest.'tasks'.$taskName.'drivers'

  if([bool]$drivers -eq 1){
    Write-Host "[ Warning: Inject drivers flag set to 'true' ]" -ForegroundColor Yellow
    Write-Host "`n[ Checking for drivers... ]" -ForegroundColor Cyan

    if([string]::IsNullOrWhiteSpace($version)){
      Write-Host "[ ERROR: Manifest missing 'version' key for selected task. Version key required when 'driver' key set to true. Script will now exit ]" -ForegroundColor DarkRed
      exit
    }

    [bool] $isDriversFolder = Test-Path -Path "$script:driversPath\$version"
    if($isDriversFolder -ne 1){
      New-Item -Path "$script:driversPath\$version" -ItemType Directory | OUT-NULL
      Write-Host "[ >> Created missing root drivers folder '$script:driversPath\$version' << ]" -ForegroundColor DarkYellow
    }

    [bool] $isDriversForModel = Test-Path -Path "$script:driversPath\$version\$script:machineModel"
    if($isDriversForModel -ne 1){
      New-Item -Path "$script:driversPath\$version\$script:machineModel" -ItemType Directory | OUT-NULL
      Write-Host "[ >> Created drivers folder for device model: '$script:machineModel' << ]" -ForegroundColor DarkYellow
    }

    $driversDirectoryInfo = Get-ChildItem -Path "$script:driversPath\$version\$script:machineModel" | Measure-Object
    if($driversDirectoryInfo.count -eq 0) {
      Write-Host "[ Error: Please add drivers for model: $script:machineModel. Script will now exit ]`n" -ForegroundColor DarkRed
      exit
    }

    Write-Host "[ >> Found drivers folder for $script:machineModel << ]" -ForegroundColor DarkYellow
    1
  }
}

function Test-UnattendFile ($taskName){
  $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'

  if(![string]::IsNullOrWhiteSpace($unattendFile)){
    Test-UnattendFolder

    Write-Host "[ Checking for specified unattend file: '$unattendFile'...]" -ForegroundColor Cyan

    if(Test-Path -Path "$script:unattendPath\$unattendFile" -PathType leaf){
      Write-Host "[ >> Found $unattendFile << ]" -ForegroundColor DarkYellow
      1
    } else {
      Write-Host "[ Error: Could not locate $unattendFile in unattend folder. Compare manifest and unattend folder. Script will now exit ]`n" -ForegroundColor DarkRed
      exit
    }
  }
}

function Test-UnattendFolder {
  [bool] $isUnattendFolder = Test-Path -Path "$script:unattendPath"
  if($isUnattendFolder -ne 1){
    New-Item -Path $script:unattendPath -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing unattend folder << ]" -ForegroundColor DarkYellow
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

  Write-Host "[ >> Disk layout complete << ]" -ForegroundColor DarkYellow
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
      Write-Host "[ >> Removed entry: $bootLoaderIdentifier << ]" -ForegroundColor DarkYellow
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