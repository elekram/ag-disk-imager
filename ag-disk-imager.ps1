$manifest = Get-Content -Raw -Path .\manifest.json | ConvertFrom-Json
$settings = Get-Content -Raw -Path .\settings.json | ConvertFrom-Json
$script:machineModel = ""
$script:selectedOption = ""
[System.Collections.ArrayList]$script:imageNames = @('0')
function main {
  #Show-ApplicationTitle
  #Set-PowerSchemeToHigh
  Get-MachineModel
  #Test-DriversForMachineModelExist
  #Write-Host $manifest."20G80001AU"."home"
  Get-ImageMenuForDevice
  #New-ImageJob
  #$manifest."20G80001AU"

}

function Get-ImageMenuForDevice {
  $counter = 0

  $manifest.$script:machineModel.PSObject.Properties | ForEach-Object {
    $script:imageNames.Add($_.Name) | Out-Null
  }
  #$imageNames.Add('Exit') | Out-Null

  Write-Host "`nMenu" -ForegroundColor Magenta
  Write-Host "++++" -ForegroundColor Gray
  
  Foreach ($img in $script:imageNames) {

    if ($img -ne '0'){
      Write-Host "[$counter] $img" -ForegroundColor DarkYellow
    }
    $counter++
  }
  #$script:imageNames
  $script:selectedOption = Read-Host "`nSelect image option and hit enter or punch CTRL-C to exit`n" 
  New-ImageJob
}

function New-ImageJob{

  $imgName = $script:imageNames[[int]$script:selectedOption]
  Write-Host $manifest.$script:machineModel.$imgName

  return
  try {  
    $manifest.$script:machineModel.$imgName
  } catch {
    Write-Host "bad option"
  }
}

function Get-MachineModel{
  Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
  $script:machineModel = Get-ComputerInfo | Select-Object -ExpandProperty "csmodel*"
  $script:machineModel = "20G80001AU"
  Write-Host "[ >> Found $script:machineModel << ]" -ForegroundColor DarkYellow
}
function Test-DriversForMachineModelExist{
  Get-Location | Write-Host
  Write-Host "`n[ Checking for drivers... ]" -ForegroundColor Cyan
  [bool] $isDriversForModel = Test-Path -Path "$settings.'driversPath'\$script:machineModel"
  if($isDriversForModel -ne 1){
    throw "ERROR: No drivers for $script:machineModel found. Script must exit."
  }
  Write-Host "`n[ >> Found drivers for $script:machineModel << ]" -ForegroundColor DarkYellow
  Write-Host ""
}

function Get-InternalDiskNumber {
  $internalDisks = Get-WmiObject win32_diskdrive | Where-Object {$_.interfacetype -ne "USB"}
  $InternalDiskNumbers = @()

  $internalDisks | ForEach-Object {
    $InternalDiskNumbers += $_.DeviceID.Substring($_.DeviceID.Length -1)
  }
  if ($InternalDiskNumbers.Count -gt 1){
    throw "ERROR: Found more than one internal disk. Script must exit."
  }

  if ($InternalDiskNumbers.Count -eq 0) {
    throw "ERROR: No internal disk found for imaging. Script must exit."
  }
  return $InternalDiskNumbers[0]
}

function New-DiskPartitions{

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
    Add-Content –Path disklayout.txt $efiLayout
  } else {
    Add-Content –path disklayout.txt $mbrLayout
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
      ;     "-,/_..--"`-..__)    
      '""--.._:
"@
  
    write-host " "
    Write-Host $title -ForegroundColor DarkYellow
}

main