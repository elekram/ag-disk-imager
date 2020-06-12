$manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json
$script:driversPath = ".\drivers"
$script:wimPath = ".\wim"
$script:unattendPath = ".\unattend"
$script:machineModel = ""
[System.Collections.ArrayList]$script:imageNames = @('0')

function main {
  Show-ApplicationTitle
  Get-MachineSerialNumber
  Get-MachineModel
  Test-ManifestForModel
  Test-ManifestForDefaults
  Test-ManifestForMachineRequiredProperties
  Test-ManifestForDuplicateNames
  Test-WimFolder
  Get-ImageMenuForDevice
  Get-MachineSerialNumber
  Write-Host "[ >> All Good <<]" -ForegroundColor Green
}

function Test-ManifestForDefaults{
  if(![bool]($manifest.PSobject.Properties.name -Match "defaults")){
    Write-Host "[ Error: defaults key found not found in manifest. At least one default task required. Script will now exit. ]`n" -ForegroundColor DarkRed
    exit
  }
}
function Test-ManifestForModel{
  if(![bool]($manifest.PSobject.Properties.name -Match $script:machineModel)){
    Write-Host "[ Error: Machine model not found in manifest. Script will now exit. ]`n" -ForegroundColor DarkRed
    exit
  }
}

function Test-ManifestForMachineRequiredProperties{
  if(![bool]($manifest.$script:machineModel.PSobject.Properties.name -Match "config")){
    Write-Host "[ Error: Manifest missing 'config' key for $script:machineModel. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  if(![bool]($manifest.$script:machineModel.PSobject.Properties.name -Match "name")){
    Write-Host "[ Error: Manifest missing 'name' key for $script:machineModel. Machines have names to you know. Script will now exit ]" -ForegroundColor DarkRed
    exit
  } else {
    $machineLongName = $manifest.$script:machineModel.'name'
    Write-Host "[ $machineLongName ]" -ForegroundColor Yellow
  }

  if($manifest.$script:machineModel.'config' -ne 'defaults'){
    if(![bool]($manifest.$script:machineModel.PSobject.Properties.name -Match "tasks")){
      Write-Host "[ Error: Manifest missing 'tasks' key for $script:machineModel. Script will now exit ]" -ForegroundColor DarkRed
      exit
    }
  }
}

function Test-ManifestTaskProperties($task) {
  if(![bool]($task.PSobject.Properties.name -Match "wim")){
    Write-Host "[ Error: Selected task missing 'wim' key in manifest. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  if(![bool]($task.PSobject.Properties.name -Match "drivers")){
    Write-Host "`n[ Warning: Selected task missing 'drivers' key in manifest. No drivers will be injected ]`n" -ForegroundColor Yellow
    Write-Host -NoNewLine 'Press any key to continue image task or CTRL-C to exit';
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    Write-Host ""
  }

  if(![bool]($task.PSobject.Properties.name -Match "unattend")){
    Write-Host "`n[ Warning: Selected task missing 'unattend' key in manifest. Unattend file will not be injected ]`n" -ForegroundColor Yellow
    Write-Host -NoNewLine 'Press any key to continue image task or CTRL-C to exit';
    $null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown');
    Write-Host ""
  }
}

function Test-ManifestForDuplicateNames {
  $manifest."defaults".PSObject.Properties | ForEach-Object {
    $defName = $_.Name
    $manifest.$script:machineModel.'tasks'.PSobject.Properties | ForEach-Object {
      if ($defName -eq $_.Name){
        Write-Host "[ Error: Duplicate task '$defName' found in manifest. Compare defaults and $script:machineModel." -ForegroundColor DarkRed
        Write-Host "[ Duplicate entries are not allowed. Script will now exit. ]`n" -ForegroundColor DarkRed
        exit
      }
    }
  }
}

function Get-ImageMenuForDevice {
  $counter = 0

  $machineConfig = $manifest.$script:machineModel."config"

  if([string]::IsNullOrWhiteSpace($machineConfig)){
    Write-Host "[ WARNING: Config key missing from manifest for $script:machineModel. Using 'defaults' ]" -ForegroundColor Yellow
    $machineConfig = "defaults"
  }

  [System.Collections.ArrayList]$script:imageNames = @('0')

  switch ($machineConfig) {
    "defaults" {
      $manifest."defaults".PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }
    }
    "append" {
      $manifest."defaults".PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }

      $manifest.$script:machineModel.'tasks'.PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }
    }
    "replace" {
      $manifest.$script:machineModel.'tasks'.PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }
    }
  }

  Write-Host "`nMenu" -ForegroundColor Magenta
  Write-Host "++++" -ForegroundColor Gray

  Foreach ($img in $script:imageNames) {
    if ($img -ne 'config' -and $img -ne '0'){
      Write-Host "[$counter] $img" -ForegroundColor DarkYellow
    }
    $counter++
  }

  $selectedOption = Read-Host "`nSelect image option and hit enter or punch CTRL-C to exit`n"
  New-ImageJob($machineConfig, $selectedOption)
}

function New-ImageJob ($_args){

  $machineConfig = $_args[0]
  $menuOption = $_args[1]

  $imageFile = ""
  $unattendFile = ""
  [bool]$drivers = 0

  # check if selected option is a number character
  if($menuOption -match '\d' -ne 1) {
    Write-Host "Invalid option. Invalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-ImageMenuForDevice
    return
  }

  $imageName = $script:imageNames[[int]$menuOption]

  if([string]::IsNullOrWhiteSpace($imageName)){
    Write-Host "Invalid option. Choose a number from the menu." -ForegroundColor DarkRed
    Get-ImageMenuForDevice
    return
  }

  if($machineConfig.Trim().ToLower() -eq 'defaults' -or $machineConfig.Trim().ToLower() -eq 'append'){

    if($null -eq $manifest.'defaults'.$imageName){
      Test-ManifestTaskProperties($manifest.$script:machineModel.'tasks'.$imageName)

      $imageFile = $manifest.$script:machineModel.'tasks'.$imageName.'wim'
      $unattendFile = $manifest.$script:machineModel.'tasks'.$imageName.'unattend'
      $drivers = $manifest.$script:machineModel.'tasks'.$imageName.'drivers'
    } else {
      Test-ManifestTaskProperties($manifest.'defaults'.$imageName)

      $imageFile = $manifest.'defaults'.$imageName.'wim'
      $unattendFile = $manifest.'defaults'.$imageName.'unattend'
      $drivers = $manifest.'defaults'.$imageName.'drivers'
    }

  } else {
    $imageFile = $manifest.$script:machineModel.'tasks'.$imageName.'wim'
    $unattendFile = $manifest.$script:machineModel.'tasks'.$imageName.'unattend'
    $drivers = $manifest.$script:machineModel.'tasks'.$imageName.'drivers'
  }


  if ([string]::IsNullOrWhiteSpace($imageFile)){
    Write-Host  "[ Error: No WIM file defined for '$imageName' in manifest. Script will now exit. ]`n"  -ForegroundColor DarkRed
    exit
  }

  Test-WimFolderForImageFile($imageFile)
  Test-DriversForMachineModelExist($drivers)
  [bool]$isUnattendFile = Test-UnattendFile($unattendFile)

  Set-InternalDrivePartitions
  Set-PowerSchemeToHigh
  Write-Host "[ Beginning image task... ]" -ForegroundColor Cyan
  Expand-WindowsImage -ImagePath "$script:wimPath\$imageFile" -index 1 -ApplyPath "w:\"
  Write-Host "[ >> Completed image task << ]" -ForegroundColor DarkYellow

  if([bool]$drivers -eq 1){
    Set-DriversOnImagedPartition
  }

  if([bool]$isUnattendFile) {
    Set-UnattendFile($unattendFile)
  }
  Set-BootLoader
}

function Set-UnattendFile ($unattendFile){
  Copy-Item "$script:unattendPath\$unattendFile" -Destination "w:\windows\panther\unattend.xml"
  Write-Host "[ >> Copied $unattendFile to windows\panther... << ]" -ForegroundColor DarkYellow
}
function Test-UnattendFile ($unattendFile){
  if(![string]::IsNullOrWhiteSpace($unattendFile)){
    Test-UnattendFolder

    Write-Host "[ Checking for unattend file: '$unattendFile'...]" -ForegroundColor Cyan

    if(Test-Path -Path "$script:unattendPath\$unattendFile" -PathType leaf){
      Write-Host "[ >> Found $unattendFile << ]" -ForegroundColor DarkYellow
      1
    } else {
      Write-Host "[ Error: Could not locate $unattendFile in unattend folder. Compare manifest and unattend folder. Script will now exit ]`n" -ForegroundColor DarkRed
      exit
    }
  }
}

function Test-DriversForMachineModelExist($drivers){
  if([bool]$drivers -eq 1){
    Write-Host "[ Warning: Inject drivers flag set to 'true' ]" -ForegroundColor Yellow
    Write-Host "`n[ Checking for drivers... ]" -ForegroundColor Cyan

    [bool] $isDriversFolder = Test-Path -Path "$script:driversPath"
    if($isDriversFolder -ne 1){
      New-Item -Path $script:driversPath -ItemType Directory | OUT-NULL
      Write-Host "[ >> Created missing drivers root folder << ]" -ForegroundColor DarkYellow
    }

    [bool] $isDriversForModel = Test-Path -Path "$script:driversPath\$script:machineModel"
    if($isDriversForModel -ne 1){
      New-Item -Path "$script:driversPath\$script:machineModel" -ItemType Directory | OUT-NULL
      Write-Host "[ >> Created drivers folder for $script:machineModel << ]" -ForegroundColor DarkYellow
    }

    $driversDirectoryInfo = Get-ChildItem -Path "$script:driversPath\$script:machineModel" | Measure-Object
    if($driversDirectoryInfo.count -eq 0) {
      Write-Host "[ Error: Please add drivers for model: $script:machineModel. Script will now exit ]`n" -ForegroundColor DarkRed
      exit
    }

    Write-Host "[ >> Found drivers folder for $script:machineModel << ]" -ForegroundColor DarkYellow
  }
}

function Test-UnattendFolder {
  [bool] $isUnattendFolder = Test-Path -Path "$script:unattendPath"
  if($isUnattendFolder -ne 1){
    New-Item -Path $script:unattendPath -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing unattend folder << ]" -ForegroundColor DarkYellow
  }
}



function Test-WimFolderForImageFile($imageFile){
    Write-Host "[ Checking WIM exists... ]" -ForegroundColor Cyan
  if (Test-Path -Path "$script:wimPath\$imageFile" -PathType leaf) {
    Write-Host "[ >> Found $imageFile << ]" -ForegroundColor DarkYellow
  } else {
    Write-Host "[ Error: Could not locate $imageFile in wim folder. Compare manifest and wim folder. Script will now exit ]`n" -ForegroundColor DarkRed
    exit
  }
}

function Set-DriversOnImagedPartition {
  Write-Host "`n[ Injecting drivers... ]" -ForegroundColor Cyan
  Add-WindowsDriver -Path "w:\" -Driver "$script:driversPath\$script:machineModel" -Recurse -ForceUnsigned | Out-Null
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

function Get-MachineSerialNumber{
  $serialNumber = get-ciminstance win32_bios | Select-Object -ExpandProperty serialnumber
  Write-Host "`n[ Machine Serial Number: $serialNumber ]`n" -ForegroundColor Yellow
}

function Get-MachineModel{
  Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
  $script:machineModel = Get-CimInstance -ClassName Win32_ComputerSystem | Select-Object -ExpandProperty "Model"
  Write-Host "[ >> Found machine model $script:machineModel << ]" -ForegroundColor DarkYellow
}

function Test-WimFolder {
  [bool] $isWimFolder = Test-Path -Path "$script:wimPath"
  if($isWimFolder -ne 1){
    New-Item -Path $script:wimPath -ItemType Directory | OUT-NULL
    Write-Host "[ >> Created missing WIM folder << ]" -ForegroundColor DarkYellow
  }
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