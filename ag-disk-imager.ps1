function main {
  Get-MachineModel
  Test-DriversForMachineModelExist
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
  [bool] $isDriversForModel = Test-Path -Path "$script:driversPath\$script:machineModel"
  if($isDriversForModel -ne 1){
    throw "ERROR: No drivers for $script:machineModel found. Script must exit."
  }
  Write-Host "`n[ >> Found drivers for $script:machineModel << ]" -ForegroundColor DarkYellow
  Write-Host ""
}

main