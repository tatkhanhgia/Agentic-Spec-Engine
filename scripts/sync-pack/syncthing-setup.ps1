# Syncthing setup helpers for GoPaperless Sync Pack

function Write-Info([string]$Message) {
    Write-Host $Message -ForegroundColor Cyan
}

function Write-Ok([string]$Message) {
    Write-Host $Message -ForegroundColor Green
}

function Write-Warn([string]$Message) {
    Write-Host $Message -ForegroundColor Yellow
}

function Write-Err([string]$Message) {
    Write-Host $Message -ForegroundColor Red
}

function Ensure-SyncFolder {
    param($Config)
    New-Item -ItemType Directory -Force -Path $Config.SyncFolder | Out-Null
}

function Test-SyncthingInstalled {
    param($Config)
    if (Get-Command syncthing -ErrorAction SilentlyContinue) {
        return $true
    }
    $portable = Join-Path $Config.InstallDir 'syncthing\syncthing.exe'
    return Test-Path $portable
}

function Get-SyncthingExe {
    param($Config)

    if (Get-Command syncthing -ErrorAction SilentlyContinue) {
        return (Get-Command syncthing).Source
    }

    $commonPaths = @(
        (Join-Path $Config.InstallDir 'syncthing\syncthing.exe'),
        "$env:ProgramFiles\Syncthing\syncthing.exe",
        "$env:LocalAppData\Programs\Syncthing\syncthing.exe"
    )
    foreach ($path in $commonPaths) {
        if (Test-Path $path) {
            return $path
        }
    }
    return $null
}

function Install-SyncthingPackage {
    param($Config)

    Write-Info 'Dang cai Syncthing...'

    if (Get-Command winget -ErrorAction SilentlyContinue) {
        Write-Host 'Thu cai bang winget...'
        cmd /c "winget install --id Syncthing.Syncthing -e --accept-source-agreements --accept-package-agreements --silent" | Out-Null
        Start-Sleep -Seconds 3
        if (Get-Command syncthing -ErrorAction SilentlyContinue) {
            Write-Ok 'Da cai Syncthing bang winget.'
            return
        }
    }

    $installRoot = Join-Path $Config.InstallDir 'syncthing'
    $exePath = Join-Path $installRoot 'syncthing.exe'
    if (Test-Path $exePath) {
        Write-Ok 'Syncthing portable da co san.'
        return
    }

    Write-Host 'Tai Syncthing portable...'
    New-Item -ItemType Directory -Force -Path $installRoot | Out-Null
    $zipPath = Join-Path $env:TEMP 'syncthing-windows.zip'
    $url = 'https://github.com/syncthing/syncthing/releases/download/v1.29.3/syncthing-windows-amd64-v1.29.3.zip'

    Invoke-WebRequest -Uri $url -OutFile $zipPath -UseBasicParsing
    Expand-Archive -Path $zipPath -DestinationPath $installRoot -Force

    $extracted = Get-ChildItem -Path $installRoot -Recurse -Filter 'syncthing.exe' | Select-Object -First 1
    if (-not $extracted) {
        throw 'Khong tim thay syncthing.exe sau khi giai nen.'
    }

    Copy-Item $extracted.FullName $exePath -Force
    Write-Ok 'Da tai Syncthing portable.'
}

function Start-SyncthingProcess {
    param($Config)

    $exe = Get-SyncthingExe -Config $Config
    if (-not $exe) {
        throw 'Syncthing chua duoc cai.'
    }

    $running = Get-Process syncthing -ErrorAction SilentlyContinue
    if ($running) {
        Write-Ok 'Syncthing dang chay.'
        return
    }

    Write-Info 'Khoi dong Syncthing...'
    Start-Process -FilePath $exe -WindowStyle Hidden
    Start-Sleep -Seconds 4
}

function Wait-SyncthingConfig {
    param([int]$TimeoutSec = 60)

    $configPath = Get-SyncthingConfigPath
    $deadline = (Get-Date).AddSeconds($TimeoutSec)
    while ((Get-Date) -lt $deadline) {
        if (Test-Path $configPath) {
            return $configPath
        }
        Start-Sleep -Seconds 2
    }
    throw 'Khong tim thay file cau hinh Syncthing.'
}

function Get-SyncthingApiKey {
    param([string]$ConfigPath)

    [xml]$xml = Get-Content $ConfigPath
    return $xml.configuration.gui.apikey
}

function Invoke-SyncthingApi {
    param(
        [string]$ApiKey,
        [string]$Method,
        [string]$Path,
        [object]$Body = $null
    )

    $uri = "http://127.0.0.1:8384$Path"
    $params = @{
        Uri         = $uri
        Method      = $Method
        Headers     = @{ 'X-API-Key' = $ApiKey }
        ContentType = 'application/json'
    }
    if ($null -ne $Body) {
        $params.Body = ($Body | ConvertTo-Json -Depth 10)
    }
    return Invoke-RestMethod @params
}

function Ensure-SyncthingFolder {
    param(
        $Config,
        [string]$ApiKey,
        [string[]]$DeviceIds = @()
    )

    $folders = Invoke-SyncthingApi -ApiKey $ApiKey -Method GET -Path '/rest/config/folders'
    $existing = $folders | Where-Object { $_.id -eq $Config.SyncthingFolderId }

    $deviceEntries = @()
    foreach ($id in $DeviceIds) {
        if ($id) {
            $deviceEntries += @{ deviceID = $id.Trim() }
        }
    }

    $folder = @{
        id                 = $Config.SyncthingFolderId
        label              = $Config.SyncthingFolderLabel
        path               = $Config.SyncFolder
        type               = 'sendreceive'
        devices            = $deviceEntries
        rescanIntervalS    = 3600
        fsWatcherEnabled   = $true
        fsWatcherDelayS    = 10
        versioning         = @{ type = '' }
        ignorePatterns     = @()
    }

    if ($existing) {
        Invoke-SyncthingApi -ApiKey $ApiKey -Method PUT -Path "/rest/config/folders/$($Config.SyncthingFolderId)" -Body $folder | Out-Null
    } else {
        Invoke-SyncthingApi -ApiKey $ApiKey -Method POST -Path '/rest/config/folders' -Body $folder | Out-Null
    }
}

function Add-SyncthingRemoteDevice {
    param(
        [string]$ApiKey,
        [string]$DeviceId,
        [string]$Name = 'May remote'
    )

    if (-not $DeviceId) {
        return
    }

    $devices = Invoke-SyncthingApi -ApiKey $ApiKey -Method GET -Path '/rest/config/devices'
    $exists = $devices | Where-Object { $_.deviceID -eq $DeviceId.Trim() }
    if ($exists) {
        return
    }

    $device = @{
        deviceID            = $DeviceId.Trim()
        name                = $Name
        addresses           = @('dynamic')
        compression         = 'metadata'
        introducer          = $false
        skipIntroductionRemovals = $false
        paused              = $false
    }
    Invoke-SyncthingApi -ApiKey $ApiKey -Method POST -Path '/rest/config/devices' -Body $device | Out-Null
}

function Get-LocalSyncthingDeviceId {
    param([string]$ApiKey)
    $status = Invoke-SyncthingApi -ApiKey $ApiKey -Method GET -Path '/rest/system/status'
    return $status.myID
}

function Register-SyncthingStartup {
    param($Config)

    $exe = Get-SyncthingExe -Config $Config
    if (-not $exe) {
        return
    }

    $startup = [Environment]::GetFolderPath('Startup')
    $shortcutPath = Join-Path $startup 'GoPaperless Syncthing.lnk'
    if (Test-Path $shortcutPath) {
        return
    }

    $wsh = New-Object -ComObject WScript.Shell
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $exe
    $shortcut.WindowStyle = 7
    $shortcut.Description = 'GoPaperless folder sync'
    $shortcut.Save()
    Write-Ok 'Da them Syncthing vao khoi dong Windows.'
}

function New-SyncFolderShortcut {
    param($Config)

    $desktop = [Environment]::GetFolderPath('Desktop')
    $shortcutPath = Join-Path $desktop 'Sync - GoPaperless.lnk'
    $wsh = New-Object -ComObject WScript.Shell
    $shortcut = $wsh.CreateShortcut($shortcutPath)
    $shortcut.TargetPath = $Config.SyncFolder
    $shortcut.Description = 'Thu muc dong bo 2 chieu'
    $shortcut.Save()
}

function Install-SyncthingMode {
    param(
        $Config,
        [string[]]$RemoteDeviceIds = @()
    )

    Ensure-SyncFolder -Config $Config
    Install-SyncthingPackage -Config $Config
    Start-SyncthingProcess -Config $Config

    $configPath = Wait-SyncthingConfig
    $apiKey = Get-SyncthingApiKey -ConfigPath $configPath

    $sharedDevices = @()
    foreach ($id in $RemoteDeviceIds) {
        if ($id) {
            Add-SyncthingRemoteDevice -ApiKey $apiKey -DeviceId $id
            $sharedDevices += $id.Trim()
        }
    }

    Ensure-SyncthingFolder -Config $Config -ApiKey $apiKey -DeviceIds $sharedDevices
    Register-SyncthingStartup -Config $Config
    New-SyncFolderShortcut -Config $Config

    $myId = Get-LocalSyncthingDeviceId -ApiKey $apiKey
    return @{
        DeviceId = $myId
        GuiUrl   = Get-SyncthingGuiUrl
        SyncFolder = $Config.SyncFolder
    }
}
