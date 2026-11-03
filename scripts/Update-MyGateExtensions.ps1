#requires -Version 5.1
#requires -RunAsAdministrator

[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

$owner = 'Raghav-45'
$repository = 'mygate-tools'
$installRoot = Join-Path $env:USERPROFILE 'Downloads'
$tools = @(
    'mygate-dump-tool',
    'mygate-report-tool',
    'mygate-summary-tool',
    'mygate-analysis-tool',
    'mygate-handoff-tool'
)
$apiHeaders = @{
    Accept = 'application/vnd.github+json'
    'User-Agent' = 'MyGate-Extension-Updater'
    'X-GitHub-Api-Version' = '2022-11-28'
}

if (-not (Test-Path -LiteralPath $installRoot -PathType Container)) {
    throw "Downloads folder not found: $installRoot"
}

$installedTools = @($tools | Where-Object {
    Test-Path -LiteralPath (Join-Path $installRoot $_) -PathType Container
})

if ($installedTools.Count -eq 0) {
    Write-Host "No MyGate extension folders found directly inside $installRoot."
    Write-Host 'Install an extension with Chrome Load unpacked first, then run this updater again.'
    exit 0
}

$temporaryRoot = Join-Path $env:TEMP ("mygate-extension-update-{0}" -f [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $temporaryRoot | Out-Null

try {
    $releaseUri = "https://api.github.com/repos/$owner/$repository/releases/latest"
    Write-Host "Checking the latest MyGate release..."
    $release = Invoke-RestMethod -Uri $releaseUri -Headers $apiHeaders -Method Get

    if (-not $release -or -not $release.tag_name) {
        throw 'GitHub did not return a published release.'
    }

    $assetsByName = @{}
    foreach ($asset in $release.assets) {
        $assetsByName[$asset.name] = $asset.browser_download_url
    }

    foreach ($tool in $installedTools) {
        $assetName = "$tool.zip"
        if (-not $assetsByName.ContainsKey($assetName)) {
            throw "Release $($release.tag_name) is missing required asset $assetName. No folders have been replaced."
        }
    }

    Write-Host "Latest release: $($release.tag_name)"

    $manifestsByTool = @{}
    foreach ($tool in $installedTools) {
        $archivePath = Join-Path $temporaryRoot "$tool.zip"
        $stagingPath = Join-Path $temporaryRoot "$tool-new"

        Write-Host "Downloading $tool..."
        Invoke-WebRequest -Uri $assetsByName["$tool.zip"] -Headers $apiHeaders -OutFile $archivePath -UseBasicParsing
        Expand-Archive -LiteralPath $archivePath -DestinationPath $stagingPath -Force

        $manifestPath = Join-Path $stagingPath 'manifest.json'
        if (-not (Test-Path -LiteralPath $manifestPath -PathType Leaf)) {
            throw "The $tool archive has no root manifest.json; its installed folder was left untouched."
        }

        $manifest = Get-Content -LiteralPath $manifestPath -Raw | ConvertFrom-Json
        if ($manifest.manifest_version -ne 3 -or -not $manifest.version) {
            throw "The $tool archive has an invalid Chrome extension manifest; its installed folder was left untouched."
        }

        $manifestsByTool[$tool] = $manifest
    }

    foreach ($tool in $installedTools) {
        $destination = Join-Path $installRoot $tool
        $stagingPath = Join-Path $temporaryRoot "$tool-new"
        $backupPath = "${destination}.update-backup-$([guid]::NewGuid().ToString('N'))"

        # Keep a rollback copy until the freshly extracted extension is in place.
        Move-Item -LiteralPath $destination -Destination $backupPath
        try {
            Move-Item -LiteralPath $stagingPath -Destination $destination
        }
        catch {
            $replacementError = $_
            try {
                Move-Item -LiteralPath $backupPath -Destination $destination
            }
            catch {
                throw "Update failed and rollback could not restore the folder. The original files are preserved at '$backupPath'. Update error: $($replacementError.Exception.Message) Rollback error: $($_.Exception.Message)"
            }
            throw $replacementError
        }

        Remove-Item -LiteralPath $backupPath -Recurse -Force
        Write-Host ("Updated {0} to version {1}." -f $tool, $manifestsByTool[$tool].version)
    }

    Write-Host ''
    Write-Host 'Update complete. Open chrome://extensions and click Reload for each updated unpacked extension.'
}
finally {
    if (Test-Path -LiteralPath $temporaryRoot) {
        Remove-Item -LiteralPath $temporaryRoot -Recurse -Force
    }
}
