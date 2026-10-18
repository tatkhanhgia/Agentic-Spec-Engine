# GoPaperless Sync Pack - shared settings
#
# May ban (may dang chay script nay):
#   LAPTOP-3KA0F0DN  - 100.81.203.97
#
# Cac may o xa ket noi qua Tailscale: xem $Peers ben duoi.
#
# Syncthing Device ID moi may lay tai: http://127.0.0.1:8384 -> Action -> Show ID

function ConvertTo-WslPath {
    param([string]$WindowsPath)
    $normalized = $WindowsPath -replace '\\', '/'
    if ($normalized -match '^([A-Za-z]):(.*)$') {
        $drive = $Matches[1].ToLower()
        $rest = $Matches[2]
        return "/mnt/$drive$rest"
    }
    return $normalized
}

# May dang chay script nay.
$LocalMachineIp   = '100.81.203.97'
$LocalMachineName = 'LAPTOP-3KA0F0DN'

# Cau hinh Syncthing.
$PackConfig = @{
    SyncFolder           = 'C:\Users\Admin\Documents\Sync'
    LocalMachineIp       = $LocalMachineIp
    LocalMachineName     = $LocalMachineName
    SyncthingFolderId    = 'documents-sync'
    SyncthingFolderLabel = 'Documents Sync'
    SyncthingPort        = 8384
    InstallDir           = "$env:LOCALAPPDATA\GoPaperlessSync"
}

# Danh sach may dong bo (peer). Moi peer: Name / TailscaleIp / SyncthingDeviceId.
# Muon them may: them 1 dong o day, roi chay lai Cai-Dat-Sync.cmd.
$Peers = @(
    @{ Name = 'Cursor';                 TailscaleIp = '100.108.172.26'; SyncthingDeviceId = 'RJWC5MH-PSNA2FD-ZYRSJYV-JR6ZMHZ-P3TFCMF-576XLSI-2VQOULU-D4WI3QT' },
    @{ Name = 'May thu 3 (placeholder)'; TailscaleIp = 'CHANGE-ME';     SyncthingDeviceId = '' }
)

# Che do nang cao: SSH + Unison (WSL). Chi ho tro 2 may, dung peer dau tien lam may remote.
$SyncConfig = @{
    RemoteHost    = $Peers[0].TailscaleIp
    RemoteUser    = 'admin'
    LocalMachine  = $LocalMachineIp
    SyncFolder    = $PackConfig.SyncFolder
    WslSyncPath   = ConvertTo-WslPath $PackConfig.SyncFolder
    UnisonProfile = 'sync'
}

function Get-SyncthingConfigPath {
    return Join-Path $env:LOCALAPPDATA 'Syncthing\config.xml'
}

function Get-SyncthingGuiUrl {
    return "http://127.0.0.1:$($PackConfig.SyncthingPort)"
}

function Get-PeerList {
    param($Peers = $script:Peers)
    return $Peers
}

function Get-PeerDeviceIds {
    param($Peers = $script:Peers)
    return @($Peers | ForEach-Object { $_.SyncthingDeviceId.Trim() } | Where-Object { $_ })
}

function Get-SyncRemoteUri {
    param($Config = $script:SyncConfig)
    $winPath = $Config.SyncFolder -replace '\\', '/'
    return "ssh://$($Config.RemoteUser)@$($Config.RemoteHost)//$winPath"
}
