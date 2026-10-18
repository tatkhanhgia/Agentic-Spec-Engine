$cfg = [xml](Get-Content "$env:LOCALAPPDATA\Syncthing\config.xml")
$key = $cfg.configuration.gui.apikey
$h = @{ 'X-API-Key' = $key }

try {
    $pending = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/cluster/pending/devices' -Headers $h
    Write-Output "PENDING=$($pending | ConvertTo-Json -Compress)"
} catch {
    Write-Output "PENDING_ERROR=$($_.Exception.Message)"
}

try {
    $pendingFolders = Invoke-RestMethod -Uri 'http://127.0.0.1:8384/rest/cluster/pending/folders' -Headers $h
    Write-Output "PENDING_FOLDERS=$($pendingFolders | ConvertTo-Json -Compress)"
} catch {
    Write-Output "PENDING_FOLDERS_ERROR=$($_.Exception.Message)"
}
