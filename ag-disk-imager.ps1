$script:manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json
$script:driversPath = '.\drivers'
$script:wimPath = '.\wim'
$script:unattendPath = '.\unattend'
$script:diskLayoutPath = '.\custom-disk-layouts'
$script:machineModel = ''
$script:machineSerialNumber = ''

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
  Write-Host '+++++++++' -ForegroundColor Gray
  
  if ($isDeviceInManifest) {
    $defaultTask = ''
    $defaultTaskOption = ''

    foreach ($property in $manifest.'tasks'.PSObject.Properties) { 
    
      $taskCollection.Add($property.Name) | Out-Null
      $taskName = $property.Name

      if ($manifest.'models'.$script:machineModel -contains $taskName) {
        Write-Host '[' -ForegroundColor DarkGray -NoNewline; 
        Write-Host "$counter" -ForegroundColor Yellow -NoNewline; 
        Write-Host '] ' -ForegroundColor DarkGray -NoNewline; 
        Write-Host $taskName -ForegroundColor Yellow
      
        if ([string]::IsNullOrWhiteSpace($defaultTask)) {
          $defaultTaskOption = $counter
          $defaultTask = $taskName
        }
        $counter++
        continue
      }
      
      Write-Host "[$counter] $taskName" -ForegroundColor DarkGray
      $counter++
    }
  }

  if (!$isDeviceInManifest) {
    foreach ($property in $manifest.'tasks'.PSObject.Properties) { 
    
      $taskCollection.Add($property.Name) | Out-Null
  
      $taskName = $property.Name
      Write-Host "[$counter] $taskName" -ForegroundColor DarkGray
  
      $counter++
    }
  }

  Write-Host ''

  $taskCollection.Add('Capture Image') | Out-Null
  Write-Host "[$counter] Capture Image" -ForegroundColor Cyan

  $counter++
  $taskCollection.Add('Shutdown Windows PE Gracefully') | Out-Null
  Write-Host "[$counter] Shutdown Windows PE Gracefully" -ForegroundColor Green

  if (![string]::IsNullOrWhiteSpace($defaultTask)) {
    
    Write-Host "`n`nDefault task " -ForegroundColor DarkYellow -NoNewline; 
    Write-Host $defaultTask -ForegroundColor Blue -NoNewline; 
    Write-Host ' found for machine model ' -ForegroundColor DarkYellow -NoNewline; 
    Write-Host $script:machineModel -ForegroundColor Blue -NoNewline; 
    Write-Host '' -ForegroundColor DarkYellow
    
    $menuInput = Read-Host "Hit enter to select the default task or select a number (0-$counter)"

    if ($menuInput.Trim() -eq [string]::empty) {
      New-ImageTask($defaultTaskOption)
      return
    }

    if (!$menuInput -match '^\d{1,2}$') {
      Write-Host "`nInvalid option. Choose a number from the menu." -ForegroundColor DarkRed
      Get-TasksMenuForDevice
      return
    }

    if ($menuInput -match '^\d{1,2}$') {
      if ($menuInput -le $counter) {
        New-ImageTask($menuInput)
        return
      }

      Get-TasksMenuForDevice
      return
    }

    Write-Host "`nInvalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-TasksMenuForDevice
    return
  }

  $menuInput = Read-Host "`nSelect a number from the menu (0-$counter)`n"

  if ($menuInput -match '^\d{1,2}$') {
    if ($menuInput -le $counter) {
      New-ImageTask($menuInput)
      return
    }
  }

  Write-Host "`nInvalid option. Choose a number from the menu." -ForegroundColor DarkRed
  Get-TasksMenuForDevice
  return
}

function New-ImageTask($selectedOption) {

  if ($taskCollection[[int]$selectedOption] -eq 'Capture Image') {
    New-CaptureImageTask
    exit
  }

  $taskName = $taskCollection[[int]$selectedOption]

  if ([string]::IsNullOrWhiteSpace($taskName)) {
    Write-Host "`nInvalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-TasksMenuForDevice
    return
  }

  if ($taskName -eq 'Shutdown Windows PE Gracefully') {
    wpeutil shutdown
  }

  if (![bool]($script:manifest.'tasks'.PSobject.Properties.name -Match $taskName)) {
    Write-Host "`n[ Error: Selected taskname: '$taskName' not found in manifest 'tasks' object. Check manifest tasks under '$script:machineModel'. Script will now exit. ]" -ForegroundColor DarkRed
    exit
  }

  $wimFile = Test-TaskForWimFile($taskName)

  [bool]$isDrivers = Test-TaskForDrivers($taskName) 
  [bool]$isUnattendFile = Test-UnattendFile($taskName)
  
  Set-InternalDrivePartitions($taskName)
  Set-PowerSchemeToHigh

  Write-Host "`n[ Beginning image task... ]" -ForegroundColor Cyan
  try {
    Expand-WindowsImage -ImagePath "$script:wimPath\$wimFile" -index 1 -ApplyPath 'w:\'
  }
  catch {
    Write-Host "`n[ Error: something went wrong when trying to write file: $wimFile. See log.txt for details. Script will now exit ]" -ForegroundColor DarkRed
    New-LogEntry($_)
    exit
  }
  
  Write-Host "`n[ >> Image task complete << ]" -ForegroundColor Yellow

  if ([bool]$isDrivers -eq $true) {
    $version = $script:manifest.'tasks'.$taskName.'drivers'
    Set-DriversOnImagedPartition($version)
  }

  if ([bool]$isUnattendFile) {
    $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'
    Set-UnattendFile($unattendFile)
  }

  Set-BootLoader
  Get-MachineModel
  Get-MachineSerialNumber
  Write-Host "`n[ >> Task: $taskName completed <<]" -ForegroundColor Magenta
}

function New-CaptureImageTask {
  $counter = 0
  [System.Collections.ArrayList]$taskCollection = @()
  [System.Collections.ArrayList]$selectableOptions = @()
  
  Write-Host "`n[ Capture Image Selected... ]`n" -ForegroundColor Cyan
  Write-Host 'Volumes' -ForegroundColor Magenta
  Write-Host '-------'

  Get-Volume | ForEach-Object {
    if ($_.DriveLetter.length -gt 0) {

      $taskCollection.Add($_.DriveLetter) | Out-Null
      $selectableOptions.Add($counter) | Out-Null

      $size = [math]::Round($_.Size / 1GB)
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
    $fileName = Read-Host 'Filename for image'
  } while ($fileName -notmatch '^[a-z0-9.-]{1,200}$')
  
  $vol = $taskCollection[[int]$driveOption]
  $z = ':\'
  $capturePath = ($vol) + ($z)

  $imgPath = "$script:wimPath\$fileName"

  if (Test-Path $imgPath -PathType Leaf) {
    Write-Host "`n[ Error: File with name '$imgPath' already exists. Try again. ]" -ForegroundColor Red
    New-CaptureImageTask
    return
  }

  Write-Host "`n[ Beginning Capture Task... ]" -ForegroundColor Cyan 
  New-WindowsImage -ImagePath $imgPath -CapturePath $capturePath -Name 'Windows'
  Write-Host "`n[ >> Capture task completed << ]" -ForegroundColor Yellow
  
  Get-FinishOptions
}

function Get-FinishOptions {
  Write-Host "`nPress any key to shutdown gracefully or ESC key to exit" -ForegroundColor Green
  
  $key = [console]::ReadKey()
  if ($key.Key -ne '27') {
    wpeutil Shutdown
    return
  }
  exit
}

function Set-UnattendFile ($unattendFile) {
  Copy-Item "$script:unattendPath\$unattendFile" -Destination 'w:\windows\panther\unattend.xml'
  Write-Host "`n[ >> Copied $unattendFile to windows\panther... << ]" -ForegroundColor Yellow
}

function Set-DriversOnImagedPartition($version) {
  Write-Host "`n[ Injecting drivers... ]" -ForegroundColor Cyan
  
  try {
    Add-WindowsDriver -Path 'w:\' -Driver "$script:driversPath\$version\$script:machineModel\" -Recurse -ForceUnsigned | Out-Null
  }
  catch {
    Write-Host "`n[ Error: Something went wrong when attempting to inject drivers. See log.txt. Script will now exit ]" -ForegroundColor DarkRed
    New-LogEntry($_)
    return
  }

  Write-Host "`n[ >> Finished injecting drivers << ]" -ForegroundColor Yellow
}

function Set-BootLoader {
  Write-Host "`n[ Setting Bootloader... ]`n" -ForegroundColor Cyan

  $firmwareType = (Get-ComputerInfo).BiosFirmwareType

  if ($firmwareType -eq 'Uefi') {
    $efiCommand = 'bcdboot w:\windows'
    Invoke-Expression $efiCommand
    Invoke-Expression $efiCommand | Out-Null
    Invoke-Expression $efiCommand | Out-Null
    return
  }

  try {
    Write-Host "`n[ Checking Windows BootLoader for stale entries... ]" -ForegroundColor Cyan
    Clear-BootLoaderEntries
  }
  catch {
    Write-Host "`n[ >> No entries found << ]" -ForegroundColor Yellow
  }
    
  $biosCommand = 'bcdboot w:\windows /s w: /f BIOS'
  Invoke-Expression $biosCommand
}

function Get-MachineModel {
  Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
  $script:machineModel = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty 'Model'
  Write-Host "`n[ >> Found machine model: $script:machineModel << ]" -ForegroundColor Yellow
}

function Get-MachineSerialNumber {
  $serialNumber = get-ciminstance win32_bios | Select-Object -ExpandProperty serialnumber
  $script:machineSerialNumber = $serialNumber
  Write-Host "`n[ >> Machine Serial Number: $serialNumber << ]" -ForegroundColor Yellow
}
function Test-ManifestForModel {
  if (![bool]($script:manifest.'models'.PSobject.Properties.name -Match $script:machineModel)) {
    Write-Host "`n[ Warning: Machine model not found in manifest. ]" -ForegroundColor DarkYellow
    return $false
  }
  
  Write-Host "`n[ Found Machine model in manifest. Tasks in " -ForegroundColor DarkYellow -NoNewline; 
  Write-Host 'yellow' -ForegroundColor Yellow -NoNewline; 
  Write-Host ' are default/recommended for this device. ]' -ForegroundColor DarkYellow

  return $true
}


function Test-TaskForWimFile($taskName) {
  if (![bool]($script:manifest.'tasks'.$taskName.PSobject.Properties.name -Match 'wim')) {
    Write-Host "`n[ Error: Manifest task named '$taskName' missing key named 'wim' type[string]. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  $wimFile = $script:manifest.'tasks'.$taskName.'wim'

  if ([string]::IsNullOrWhiteSpace($wimFile)) {
    Write-Host  "`n[ Error: No WIM file defined for '$imageName' in manifest. Script will now exit. ]"  -ForegroundColor DarkRed
    exit
  }

  Test-WimFolderForImageFile($wimFile)
  $wimFile
}

function Test-TaskForWindowsVersionKey($taskName) {
  if (![bool]($script:manifest.'tasks'.$taskName.PSobject.Properties.name -Match 'drivers')) {
    Write-Host "`n[ Error: Manifest task named '$taskName' missing key named 'drivers' type[string]. Example 1709, 1803, 1909 etc. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }
}

function Test-TaskForDrivers($taskName) {
  $version = $script:manifest.'tasks'.$taskName.'drivers'

  if ([string]::IsNullOrWhiteSpace($version)) {
    return $false
  }

  Write-Host "`n[ Checking for drivers in .\drivers\$version folder... ]" -ForegroundColor Cyan

  [bool]$isDriversFolder = Test-Path -Path "$script:driversPath\$version"
  if ($isDriversFolder -ne $true) {
    New-Item -Path "$script:driversPath\$version" -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing root drivers folder '$script:driversPath\$version' << ]" -ForegroundColor Yellow
  }

  [bool]$isDriversForModel = Test-Path -Path "$script:driversPath\$version\$script:machineModel"
  if ($isDriversForModel -ne $true) {
    New-Item -Path "$script:driversPath\$version\$script:machineModel" -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created drivers folder for device model: '$script:machineModel' << ]" -ForegroundColor Yellow
  }

  $driversDirectoryInfo = Get-ChildItem -Path "$script:driversPath\$version\$script:machineModel" | Measure-Object
  if ($driversDirectoryInfo.count -eq 0) {
    Write-Host "`n[ WARNING: There were no drivers detected for this device ]" -ForegroundColor DarkYellow

    $selectedOption = Read-Host "`nWould you like to continue anyway? (y/n)`n"
    if ($selectedOption.ToLower() -eq 'y' -or $selectedOption -eq [string]::empty) {
      Write-Host "`n[ >> Continuing without drivers << ]`n" -ForegroundColor Yellow
      return $false
    }
    
    Write-Host "`n[ Script Exiting... ]`n" -ForegroundColor Cyan
    exit
  }
  
  Write-Host "`n[ >> Found drivers folder for $script:machineModel << ]" -ForegroundColor Yellow 
  return $true
}

function Test-UnattendFile ($taskName) {
  $unattendFile = $script:manifest.'tasks'.$taskName.'unattend'

  if (![string]::IsNullOrWhiteSpace($unattendFile)) {
    Test-UnattendFolder

    Write-Host "`n[ Checking for specified unattend file: '$unattendFile'... ]" -ForegroundColor Cyan

    if (Test-Path -Path "$script:unattendPath\$unattendFile" -PathType leaf) {
      Write-Host "`n[ >> Found $unattendFile << ]" -ForegroundColor Yellow
      return $true
    }
    
    Write-Host "`n[ Warning: Could not locate unattend file: $unattendFile ]`n" -ForegroundColor DarkYellow
    $selectedOption = Read-Host 'Would you like to continue anyway? (y/n)'
    if ($selectedOption.ToLower() -eq 'y' -or $selectedOption -eq [string]::empty) {
      return $true
    }
    
    Write-Host "`n[ Script Exiting... ]`n" -ForegroundColor Cyan
    exit
  }
}

function Test-UnattendFolder {
  [bool] $isUnattendFolder = Test-Path -Path "$script:unattendPath"
  if ($isUnattendFolder -ne $true) {
    New-Item -Path $script:unattendPath -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing unattend folder << ]" -ForegroundColor Yellow
  }
}

function Test-WimFolderForImageFile($wimFile) {
  Write-Host "`n[ Checking WIM exists... ]" -ForegroundColor Cyan
  if (Test-Path -Path "$script:wimPath\$wimFile" -PathType leaf) {
    Write-Host "`n[ >> Found $wimFile << ]" -ForegroundColor Yellow
    return
  }
  
  Write-Host "`n[ Error: Could not locate '$wimFile' in wim folder. Compare manifest and wim folder. Script will now exit ]" -ForegroundColor DarkRed
  exit
}
function Set-InternalDrivePartitions($taskName) {
  $internalDisk = Get-InternalDiskNumber

  $customDiskLayout = $script:manifest.'tasks'.$taskName.'disklayout'

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
  $message = ''
  $command = ''
  $firmwareType = (Get-ComputerInfo).BiosFirmwareType 

  ## This block is for custom disk layouts if specified for the task in the manifest
  if (![string]::IsNullOrWhiteSpace($customDiskLayout)) {

    Write-Host "`n[ >> Found custom disk layout reference '$customDiskLayout' << ]" -ForegroundColor Yellow

    Write-Host "`n[ Checking '$customDiskLayout' file exists... ]" -ForegroundColor Cyan
    if (Test-Path -Path "$script:diskLayoutPath\$customDiskLayout" -PathType leaf) {
      
      Write-Host "`n[ >> Found $customDiskLayout << ]" -ForegroundColor Yellow
      Test-LayoutFileForExceptedDiskNumber($customDiskLayout)

      $command = "diskpart /s $script:diskLayoutPath\$customDiskLayout"
      $message = "`n[ Writing custom disk layout: $customDiskLayout ... ]"

      Write-Host $message -ForegroundColor Cyan
      Invoke-Expression $command

      Write-Host "`n[ >> Disk layout complete << ]" -ForegroundColor Yellow
      return
    }
    
    Write-Host "`n[ Error: Could not locate '$customDiskLayout' in 'custom-disk-layouts' folder. Compare manifest and 'custom-disk-layouts' folder. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  if ($firmwareType -eq 'Uefi') {
    Add-Content -Path disklayout.txt $efiLayout
    $message = "`n[ Writing efi disk layout... ]"
    $command = 'diskpart /s disklayout.txt'
    Invoke-Expression $command
    return
  }
  
  if ($firmwareType -ne 'Uefi') {
    $message = "`n[ Writing mbr/bios disk layout... ]"
    Add-Content -Path disklayout.txt $mbrLayout
    $command = 'diskpart /s disklayout.txt'
    Invoke-Expression $command
    return
  }
}

function Test-LayoutFileForExceptedDiskNumber($customDiskLayout) {
  # ensures that the diskpart text file doesn't try to wipe the external usb drive
  foreach ($line in Get-Content "$script:diskLayoutPath\$customDiskLayout") {
    if ($line -match 'select' -and $line -match 'disk') {
      $lineElements = $line.split(' ')
    }
  }

  foreach ($e in $lineElements) {
    if ($e -match '^\d+$') {
      $targetDiskNumber = [int]$e
    }
  }

  $internalDisk = Get-InternalDiskNumber
  if ($targetDiskNumber -ne $internalDisk) {
    Write-Host "`n[ Error: Custom layout file: '$customDiskLayout' is attempting to write to a excepted disk. Script will now exit ]" -ForegroundColor Red
    exit
  }
}

function Get-InternalDiskNumber {
  #Get-WmiObject win32_diskdrive | Where-Object {Write-Host $_.MediaType}
  $internalDisks = Get-WmiObject win32_diskdrive | Where-Object { $_.MediaType -ne 'External hard disk media' -and $_.MediaType -ne 'Removable Media' -and $_.InterfaceType -ne 'USB' }
  
  $InternalDiskNumbers = @()

  $internalDisks | ForEach-Object {
    $InternalDiskNumbers += $_.DeviceID.Substring($_.DeviceID.Length - 1)
  }
  if ($InternalDiskNumbers.Count -gt 1) {
    Write-Host "`n[ Error: Found more than one internal disk. Script must exit ]" -ForegroundColor DarkRed
    exit
  }

  if ($InternalDiskNumbers.Count -eq 0) {
    Write-Host "`n[ Error: No internal disk found for imaging. Script must exit ]" -ForegroundColor DarkRed
    exit
  }
  return $InternalDiskNumbers[0]
}

function Clear-BootLoaderEntries {
  $bcdOutput = (bcdedit /v) -join "`n"
  $entries = New-Object System.Collections.Generic.List[pscustomobject]]
  ($bcdOutput -split '(?m)^(.+\n-)-+\n' -ne '').ForEach({
      if ($_.EndsWith("`n-")) {
        # entry header 
        $entries.Add([pscustomobject] @{ Name = ($_ -split '\n')[0]; Properties = [ordered] @{} })
      }
      else {
        # block of property-value lines
        ($_ -split '\n' -ne '').ForEach({
            $propAndVal = $_ -split '\s+', 2 # split line into property name and value
            if ($propAndVal[0] -ne '') {
              # [start of] new property; initialize list of values
              $currProp = $propAndVal[0]
              $entries[-1].Properties[$currProp] = New-Object Collections.Generic.List[string]
            }
            $entries[-1].Properties[$currProp].Add($propAndVal[1]) # add the value
          })
      }
    })

  $entries | ForEach-Object {
    if ($_.Name -ne 'Windows Boot Manager') {
      $bootLoaderIdentifier = $_.Properties['identifier']
      Write-Host "`n[ >> Removed entry: $bootLoaderIdentifier << ]" -ForegroundColor Yellow
      $command = "bcdedit /delete '$bootLoaderIdentifier'"
      Invoke-Expression $command
    }
  }
}

function New-LogEntry($entry) {
  $currentDateAndTime = Get-Date
  $divider = '------------'

  if (!(Test-Path .\log.txt -PathType Leaf)) {
    New-Item -Name .\log.txt -ItemType File -Force | OUT-NULL
    $logTitle = 'AG ERROR LOG'
    Add-Content -Path .\log.txt -Value $logTitle
    Add-Content -Path .\log.txt -Value $divider
  }

  Add-Content -Path .\log.txt -Value "Logged: $currentDateAndTime"
  Add-Content -Path .\log.txt -Value $entry
  Add-Content -Path .\log.txt -Value $divider
}

function Set-PowerSchemeToHigh {
  Write-Host "`n[ Setting Power Scheme to HIGH... ]" -ForegroundColor Cyan
  powercfg /s 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
}

function Show-ApplicationTitle {
  $title = @"
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

  write-host ' '
  Write-Host $title -ForegroundColor DarkYellow
}

main