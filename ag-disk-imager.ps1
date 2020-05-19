$manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json

$script:driversPath = ".\drivers"
$script:wimPath = ".\wim"
$script:machineModel = ""
$script:selectedOption = ""
$script:manifestConfig = ""
[System.Collections.ArrayList]$script:imageNames = @('0')
function main {
  Show-ApplicationTitle
  Set-PowerSchemeToHigh
  Get-MachineModel
  Test-ManifestForModel
  Test-WimFolder
  Test-DriversForMachineModelExist
  Get-ImageMenuForDevice
  Write-Host "[ >> Finished! <<]" -ForegroundColor Green
}

function Test-ManifestForModel{
  if(![bool]($manifest.PSobject.Properties.name -Match $script:machineModel)){
    Write-Host "[ Error: Machine model not found in manifest. Script will now exit. ]" -ForegroundColor DarkRed
    exit
  }
}
function Get-ImageMenuForDevice {
  $counter = 0
  $manifestConfig = $manifest.$script:machineModel."config"

  [System.Collections.ArrayList]$script:imageNames = @('0')
  
  switch ($manifestConfig) {
    "defaults" {
      $manifest."defaults".PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }
    }
    "append" {
      $manifest."defaults".PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }

      $manifest.$script:machineModel.PSObject.Properties | ForEach-Object {
        $script:imageNames.Add($_.Name) | Out-Null
      }
    }
    "replace" {
      $manifest.$script:machineModel.PSObject.Properties | ForEach-Object {
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

  $script:selectedOption = Read-Host "`nSelect image option and hit enter or punch CTRL-C to exit`n" 
  New-ImageJob
}

function New-ImageJob{
  $imageFile = ""
  # check if selected option is a number character
  if($script:selectedOption -match '\d' -ne 1) {
    Write-Host "Invalid option." -ForegroundColor DarkRed
    Get-ImageMenuForDevice
    return
  }

  $imgName = $script:imageNames[[int]$script:selectedOption]

  if([string]::IsNullOrWhiteSpace($manifest.$script:machineModel.$imgName)){
    if([string]::IsNullOrWhiteSpace($manifest."defaults".$imgName)){
      Write-Host "Invalid option." -ForegroundColor DarkRed
      Get-ImageMenuForDevice
      return
    }
    $imageFile = $manifest."defaults".$imgName
  } else {
    $imageFile = $manifest.$script:machineModel.$imgName
  }
  
  #check that wim exits before proceeding
  Write-Host "$script:wimPath\$imageFile"
  Write-Host "[ Checking WIM exists... ]" -ForegroundColor Cyan
  if (Test-Path -Path "$script:wimPath\$imageFile" -PathType leaf) {
    Write-Host "[ >> Found $imageFile << ]" -ForegroundColor DarkYellow 
  } else {
    Write-Host "[ Error: Could not locate $imageFile. Compare manifest and WIM folder. Script will not exit ]" -ForegroundColor DarkRed 
    exit
  }

  return
  Set-InternalDrivePartitions
  
  Write-Host "[ Beginning image task... ]" -ForegroundColor Cyan
  Expand-WindowsImage -ImagePath "$script:wimPath\$imageFile" -index 1 -ApplyPath "w:\" 
  Write-Host "[ >> Completed image task << ]" -ForegroundColor DarkYellow 
  Set-DriversOnImagedPartition
  Set-BootLoader

}

function Test-WimExists {
  Test-Path
}

function Set-DriversOnImagedPartition {
  Write-Host "`n[ Injecting drivers... ]" -ForegroundColor Cyan
  Add-WindowsDriver -Path "w:\" -Driver "$script:driversPath\$script:machineModel" -Recurse -ForceUnsigned
  Write-Host "[ >> Finished injecting drivers << ]" -ForegroundColor DarkYellow
}

function Set-BootLoader{
  if($(Get-ComputerInfo).BiosFirmwareType -eq "Uefi"){
    $efiCommand = "bcdboot w:\windows"
    Invoke-Expression $efiCommand
    Invoke-Expression $efiCommand
    Invoke-Expression $efiCommand
  } else {
    $biosCommand = "bcdboot w:\windows /s w: /f BIOS"
    Invoke-Expression $biosCommand
  }
}


function Get-MachineModel{
  Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
  $script:machineModel = Get-ComputerInfo | Select-Object -ExpandProperty "csmodel*"
  #$script:machineModel = "20G80001AU"
  Write-Host "[ >> Found machine model $script:machineModel << ]" -ForegroundColor DarkYellow
}
function Test-DriversForMachineModelExist{
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
    Write-Host "[ Error: Please add drivers for model: $script:machineModel. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

  Write-Host "[ >> Found drivers folder for $script:machineModel << ]" -ForegroundColor DarkYellow
}

function Test-WimFolder {
  [bool] $isWimFolder = Test-Path -Path "$script:wimPath"
  if($isWimFolder -ne 1){
    New-Item -Path $script:wimPath -ItemType Directory | OUT-NULL
    Write-Host "`n[ >> Created missing WIM folder << ]" -ForegroundColor DarkYellow
  }

  $wimDirectoryInfo = Get-ChildItem -Path "$script:wimPath" | Measure-Object
  if($wimDirectoryInfo.count -eq 0) {
    Write-Host "[ Error: No files in WIM directory. Script will now exit ]" -ForegroundColor DarkRed
    exit
  }

}

function Get-InternalDiskNumber {
  $internalDisks = Get-WmiObject win32_diskdrive | Where-Object {$_.interfacetype -ne "USB"}
  $InternalDiskNumbers = @()

  $internalDisks | ForEach-Object {
    $InternalDiskNumbers += $_.DeviceID.Substring($_.DeviceID.Length -1)
  }
  if ($InternalDiskNumbers.Count -gt 1){
    Write-Host "[ Error: Found more than one internal disk. Script must exit ]" -ForegroundColor DarkRed
    exit
  }

  if ($InternalDiskNumbers.Count -eq 0) {
    Write-Host "[ Error: No internal disk found for imaging. Script must exit ]" -ForegroundColor DarkRed
    exit
  }
  return $InternalDiskNumbers[0]
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

  if($(Get-ComputerInfo).BiosFirmwareType -eq "Uefi"){
    Add-Content -Path disklayout.txt $efiLayout
  } else {
    Add-Content -Path disklayout.txt $mbrLayout
  }

  Write-Host "`n[ Writing disk layout... ]" -ForegroundColor Cyan

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
"@
  
    write-host " "
    Write-Host $title -ForegroundColor DarkYellow
}

main