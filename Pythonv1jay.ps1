# Load servers from text file
$servers = Get-Content -Path "C:\Scripts\servers.txt" | Where-Object { $_ -and $_.Trim() -ne "" } | Sort-Object -Unique

# Central log path
$logRoot = "C:\Temp"
$timestamp = Get-Date -Format "yyyyMMdd_HHmmss"

# Script block to run remotely
$cleanupScript = {
    param($logPath)

    # --- Static Python MSI Product Codes (from your list) ---
    $pythonGUIDs = @(
        "{88AF4D20-BE9D-4CA6-8BD4-5DB380A41CC8}",
        "{AD923240-0ACE-45C9-8749-05BF77AAE101}",
        "{BDFB7011-0AB2-440F-8F00-32AF7A9ED1ED}",
        "{65B0F976-5151-427E-95B4-2320DC64F91E}",
        "{A36C1168-60E6-42E4-93DB-6BE8C6DD9DD6}",
        "{52AB506A-EC3C-4060-9EBF-6A975994CB35}",
        "{8EEE042B-6EAF-4171-BA6E-01319ED99DA8}",
        "{33F9B46C-EB19-4BB7-ABFA-F8C71B73E9A4}",
        "{FCA1EB7D-2F62-4659-AA5F-42C37CE5D3CB}",
        "{F6DA05CF-67B5-47D0-ABD4-371C80BA0717}",
        "{9BE9E7F0-F9F1-487B-A2FC-790CD2898388}",
        "{D6580352-5B95-49A9-B2F3-313D12D13968}",
        "{511119D2-41C4-48E1-A3DA-0A6A1E68AC76}",
        "{EC27BF73-AB7E-4867-9EEC-3AD456006835}",
        "{5C5B7907-C4E8-4E09-8CD6-3E844C7D65E2}",
        "{7C56D977-225C-4EBA-8308-E47DF9FA867F}",
        "{4DD10049-CC97-48AE-BE76-4CB6E3111F7B}",
        "{C4B7FF79-1195-436F-AA85-28EE995151B7}",
        "{69BCB7EC-54AF-47F2-A891-D335CE44A530}",
        "{2994270E-FE74-49E5-98BB-E65F5F0EC304}",
        "{1F5C7063-8305-4755-A643-32DE2BE966F9}",
        "{A15F08D3-26E4-4F0B-BA8B-ED59A52D6A02}",
        "{6F13A394-E3EA-4585-9ADE-046B69F1F902}",
        "{48E8B3E4-EEE2-4DB3-A518-C2B8A3075B5A}",
        "{83C32D05-F3C4-4D61-877E-0A4C6717E7DC}",
        "{6CE85987-8440-409D-BE75-F5128943F67B}",
        "{6E84DCAA-19DD-4560-AAE7-043EADF5C1F8}",
        "{6C19B2EE-FA34-4270-A87F-1FF008C1AC6E}"
    )

    # --- Static registry keys you provided ---
    $registryKeys = @(
        "HKLM:\Software\Classes\Installer\Products\B240EEE8FAE61714ABE61013E99DD98A",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\B240EEE8FAE61714ABE61013E99DD98A",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{8EEE042B-6EAF-4171-BA6E-01319ED99DA8}",
        "HKLM:\Software\Classes\Installer\Products\D7BE1ACF26F29564AAF5243CC75E3DBC",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\D7BE1ACF26F29564AAF5243CC75E3DBC",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{FCA1EB7D-2F62-4659-AA5F-42C37CE5D3CB}",
        "HKLM:\Software\Classes\Installer\Products\779D65C7C522ABE438804ED79FAF68F7",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\779D65C7C522ABE438804ED79FAF68F7",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{7C56D977-225C-4EBA-8308-E47DF9FA867F}",
        "HKLM:\Software\Classes\Installer\Products\CE7BCB96FA452F748A193D53EC445A03",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\CE7BCB96FA452F748A193D53EC445A03",
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\{69BCB7EC-54AF-47F2-A891-D335CE44A530}"
    )

    # --- Prepare logging ---
    $log = @()
    function Write-Log { param([string]$Message) $script:log += ("{0:u} {1}" -f (Get-Date), $Message) }

    Write-Log "----- Python cleanup starting on $env:COMPUTERNAME -----"

    # --- Uninstall static GUIDs (MSI) ---
    foreach ($guid in $pythonGUIDs) {
        Write-Log "Uninstalling static GUID $guid..."
        try {
            Start-Process "msiexec.exe" -ArgumentList "/x $guid /quiet /norestart" -Wait -ErrorAction Stop
            Write-Log "Success: $guid"
        } catch {
            Write-Log "Failed: $guid - $($_.Exception.Message)"
        }
    }

    # --- Uninstall dynamically discovered Python packages via registry (avoid Win32_Product) ---
    $uninstallRoots = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    $pythonPackages = foreach ($path in $uninstallRoots) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -like "*Python*"
        }
    }

    foreach ($pkg in $pythonPackages) {
        $name   = $pkg.DisplayName
        $quiet  = $pkg.QuietUninstallString
        $uninst = if ($quiet) { $quiet } else { $pkg.UninstallString }

        if (-not $uninst) {
            Write-Log "Skipping (no uninstall string): $name"
            continue
        }

        Write-Log "Uninstalling dynamic package $name ..."
        try {
            if ($uninst -match 'msiexec(\.exe)?') {
                # Try to normalize to msiexec with GUID
                $guidMatch = [regex]::Match($uninst, '\{[0-9A-Fa-f\-]{36}\}')
                if ($guidMatch.Success) {
                    Start-Process "msiexec.exe" -ArgumentList "/x $($guidMatch.Value) /quiet /norestart" -Wait -ErrorAction Stop
                } else {
                    # Fallback: pass through to msiexec, ensure quiet flags
                    if ($uninst -notmatch '/quiet')    { $uninst += ' /quiet' }
                    if ($uninst -notmatch '/norestart') { $uninst += ' /norestart' }
                    $tokens = $uninst -split '\s+',2
                    Start-Process $tokens[0].Trim('"') -ArgumentList ($tokens[1]) -Wait -ErrorAction Stop
                }
            } else {
                # Generic EXE uninstallers; try to enforce silence
                if ($uninst -notmatch '(/quiet|/silent|/s|/S|--silent)') { $uninst += ' /quiet' }
                if ($uninst -notmatch '/norestart') { $uninst += ' /norestart' }
                $tokens = $uninst -split '\s+',2
                Start-Process $tokens[0].Trim('"') -ArgumentList ($tokens[1]) -Wait -ErrorAction Stop
            }
            Write-Log "Success: $name"
        } catch {
            Write-Log "Failed: $name - $($_.Exception.Message)"
        }
    }

    # --- Remove leftover Python folders ---
    $paths = @(
        "$env:ProgramFiles\Python*",
        "$env:ProgramFiles(x86)\Python*",
        "$env:LocalAppData\Programs\Python",
        "$env:AppData\Python",
        "$env:USERPROFILE\AppData\Local\Programs\Python"
    )

    foreach ($path in $paths) {
        Get-ChildItem -Path $path -ErrorAction SilentlyContinue | ForEach-Object {
            Write-Log "Removing folder: $($_.FullName)"
            Remove-Item -Path $_.FullName -Recurse -Force -ErrorAction SilentlyContinue
        }
    }

    # --- Remove Python segments from system PATH (Machine) ---
    try {
        $envRegPath = "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\Environment"
        $envPath = (Get-ItemProperty -Path $envRegPath -Name Path -ErrorAction SilentlyContinue).Path
        if ($envPath) {
            # Conservative filtering for typical Python paths, including Scripts subfolders
            $updatedParts = ($envPath -split ";") | Where-Object {
                $_ -and ($_ -notmatch '(?i)\\Python(?:\d+(?:\.\d+)*)?(?:\\|$)') -and ($_ -notmatch '(?i)\\Python\\Scripts(?:\\|$)')
            }
            if ($updatedParts.Count -ne ($envPath -split ";").Count) {
                Set-ItemProperty -Path $envRegPath -Name Path -Value ($updatedParts -join ";")
                Write-Log "Updated system PATH"
            } else {
                Write-Log "No Python segments found in PATH"
            }
        }
    } catch {
        Write-Log "Failed updating PATH: $($_.Exception.Message)"
    }

    # --- Remove static registry keys ---
    foreach ($key in $registryKeys) {
        Write-Log "Removing registry key: ${key}"
        try {
            if (Test-Path $key) { Remove-Item -Path $key -Recurse -Force -ErrorAction Stop }
        } catch {
            Write-Log "Failed to remove registry key ${key}: $($_.Exception.Message)"
        }
    }

    # --- Remove dynamic registry keys (for uninstall entries) ---
    $uninstallPaths = @(
        "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*",
        "HKLM:\Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*"
    )

    foreach ($path in $uninstallPaths) {
        Get-ItemProperty -Path $path -ErrorAction SilentlyContinue | Where-Object {
            $_.DisplayName -like "*Python*"
        } | ForEach-Object {
            $regPath = $_.PSPath
            Write-Log "Removing dynamic registry key: ${regPath}"
            try {
                if (Test-Path $regPath) {
                    Remove-Item -Path $regPath -Recurse -Force -ErrorAction Stop
                }
            } catch {
                Write-Log "Failed to remove dynamic registry key ${regPath}: $($_.Exception.Message)"
            }
        }
    }

    Write-Log "----- Python cleanup completed on $env:COMPUTERNAME -----"

    # Emit log back to caller
    return ,$log
} # <-- closes $cleanupScript scriptblock

# Ensure central log root exists on the caller
New-Item -Path $logRoot -ItemType Directory -Force | Out-Null

# Run remotely & save per-server logs
foreach ($server in $servers) {
    try {
        $log = Invoke-Command -ComputerName $server -ScriptBlock $cleanupScript -ArgumentList $logRoot -ErrorAction Stop
        $nodeDir = Join-Path $logRoot $server
        New-Item -Path $nodeDir -ItemType Directory -Force | Out-Null
        $outFile = Join-Path $nodeDir ("PythonCleanup_{0}.log" -f $timestamp)
        $log | Out-File -FilePath $outFile -Encoding UTF8
        Write-Host "[$server] Log saved to $outFile"
    } catch {
        Write-Warning "[$server] Failed: $($_.Exception.Message)"
    }
}