function Get-MachineModel{
  Write-Host "`n[ Retrieving computer model... ]" -ForegroundColor Cyan
  $script:machineModel = Get-ComputerInfo | Select-Object -ExpandProperty "csmodel*"
  $script:machineModel = "20G80001AU"
  Write-Host "[ >> Found $script:machineModel << ]" -ForegroundColor DarkYellow
}