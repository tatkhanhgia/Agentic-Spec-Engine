$cfg = [xml](Get-Content "$env:LOCALAPPDATA\Syncthing\config.xml")
$key = $cfg.configuration.gui.apikey
$h = @{ 'X-API-Key' = $key }

$status = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/system/status' -Headers $h
$devices = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/config/devices' -Headers $h
$folders = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/config/folders' -Headers $h
$connections = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/system/connections' -Headers $h

Write-Output "MY_ID=$($status.myID)"
Write-Output "DEVICE_COUNT=$($devices.Count)"
foreach ($d in $devices) {
    Write-Output "PEER=$($d.deviceID)|$($d.name)"
}
foreach ($f in $folders) {
    Write-Output "FOLDER=$($f.id)|$($f.path)|devices=$($f.devices.Count)"
}
Write-Output "CONNECTIONS=$($connections.connections | ConvertTo-Json -Compress)"
