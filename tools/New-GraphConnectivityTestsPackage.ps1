<#
.SYNOPSIS
    Copies the files needed to run Graph Connectivity Tests into a separate folder and zips it.

.DESCRIPTION
    The package contains only what Invoke-GraphConnectivityTests.ps1 needs at run time, plus the HTML guides:
        package\Invoke-GraphConnectivityTests.ps1, package\GraphConnectivityTests.psd1, package\GraphConnectivityTests.psm1,
        package\src\, package\config\, package\templates\, package\docs\GraphConnectivityTests-UserGuide.html,
        package\docs\GraphConnectivityTests-Guide.html, package\README.md, CHANGELOG.md, package\LICENSE, package\THIRD-PARTY-NOTICES.md
    The HTML guides (user guide, developer guide) are rebuilt first from their Markdown sources (tools\Build-Documentation.ps1):
    they are self-contained (images inline), so the Markdown sources and the images are not copied.
    It never copies reports\, logs\, artifacts\, tests\ or tools\.

    The configuration is copied as delivered (contoso.com example values). The script checks the content of the
    package and that the module loads from it, then writes <Destination>.zip next to the folder (the content of the
    folder at the root of the zip).

.PARAMETER Destination
    Package folder. Default: artifacts\GraphConnectivityTests-<version> in the repository (artifacts\ is not tracked).

.PARAMETER Force
    Replace the destination folder (and its zip) if it already contains a package. A folder that contains reports\
    or logs\ (a package that has been run) is never replaced.

.PARAMETER NoZip
    Only the folder, no zip.

.EXAMPLE
    .\tools\New-GraphConnectivityTestsPackage.ps1
    Creates artifacts\GraphConnectivityTests-1.0.0 and artifacts\GraphConnectivityTests-1.0.0.zip.

.NOTES
    Author  : Nicolas Fabert
    Version : 1.0.0
    Part of : Graph Connectivity Tests (repository tool, not in the package)
#>
#Requires -Version 7.4
[CmdletBinding()]
param(
    [string]$Destination,
    [switch]$Force,
    [switch]$NoZip
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$packageRoot = Join-Path $root 'package'
$version = (Import-PowerShellDataFile -LiteralPath (Join-Path $packageRoot 'GraphConnectivityTests.psd1')).ModuleVersion
if (-not $Destination) { $Destination = Join-Path $root "artifacts\GraphConnectivityTests-$version" }
$Destination = [IO.Path]::GetFullPath($Destination, (Get-Location).Path).TrimEnd('\')
$rootPrefix = [IO.Path]::GetFullPath($root).TrimEnd('\') + '\'
$artifacts = $rootPrefix + 'artifacts\'
$inside = ($Destination + '\').StartsWith($rootPrefix, [StringComparison]::OrdinalIgnoreCase)
if (($inside -and -not $Destination.StartsWith($artifacts, [StringComparison]::OrdinalIgnoreCase)) -or $rootPrefix.StartsWith($Destination + '\', [StringComparison]::OrdinalIgnoreCase)) {
    throw "The destination must be under artifacts\ or outside the tool folder: $Destination"
}
$zip = "$Destination.zip"
if (Test-Path -LiteralPath $Destination) {
    if (-not $Force) { throw "The destination already exists: $Destination. Use -Force to replace it." }
    if (-not (Test-Path -LiteralPath (Join-Path $Destination 'Invoke-GraphConnectivityTests.ps1'))) { throw "The destination is not a Graph Connectivity Tests package, it is not replaced: $Destination" }
    foreach ($used in 'reports', 'logs') {
        if (Test-Path -LiteralPath (Join-Path $Destination $used)) { throw "The destination contains a $used folder (a package that has been run), it is not replaced: $Destination" }
    }
    Remove-Item -LiteralPath $Destination -Recurse -Force
}
if (-not $NoZip -and (Test-Path -LiteralPath $zip)) {
    if (-not $Force) { throw "The zip already exists: $zip. Use -Force to replace it." }
    Remove-Item -LiteralPath $zip -Force
}

# ---- HTML guides, rebuilt from the Markdown sources -----------------------------------------------
& (Join-Path $PSScriptRoot 'Build-Documentation.ps1') | Out-Null

# ---- Files needed at run time -------------------------------------------------------------------
$files = [Collections.Generic.List[string]]::new()
foreach ($f in 'Invoke-GraphConnectivityTests.ps1', 'GraphConnectivityTests.psd1', 'GraphConnectivityTests.psm1', 'README.md', 'CHANGELOG.md', 'LICENSE', 'THIRD-PARTY-NOTICES.md',
    'config\GraphConnectivityTests.config.psd1', 'templates\Report.template.html', 'docs\GraphConnectivityTests-UserGuide.html', 'docs\GraphConnectivityTests-Guide.html') { $files.Add($f) }
Get-ChildItem -LiteralPath (Join-Path $packageRoot 'src') -Filter '*.ps1' -File | ForEach-Object { $files.Add("src\$($_.Name)") }
foreach ($f in $files) {
    $source = Join-Path $packageRoot $f
    if ($f -in 'CHANGELOG.md', 'LICENSE' -and -not (Test-Path -LiteralPath $source)) { $source = Join-Path $root $f }
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) { throw "Missing file in the tool folder: $f" }
    $target = Join-Path $Destination $f
    [void][IO.Directory]::CreateDirectory((Split-Path $target -Parent))
    Copy-Item -LiteralPath $source -Destination $target
}

# ---- Checks ---------------------------------------------------------------------------------------
$problems = [Collections.Generic.List[string]]::new()
foreach ($name in 'reports', 'logs', 'tests', 'artifacts', 'tools', 'docs\images') {
    if (Test-Path -LiteralPath (Join-Path $Destination $name)) { $problems.Add("Folder $name\ must not be in the package.") }
}
Get-ChildItem -LiteralPath $Destination -Recurse -File -Include '*.log', '*.csv', '*.json', '*.png', '*.md.bak', '*.Tests.ps1', '*Simulator*' |
    ForEach-Object { $problems.Add("Not a run-time file: $($_.Name)") }
$manifest = [IO.File]::ReadAllText((Join-Path $packageRoot 'GraphConnectivityTests.psm1'))
$parts = @([regex]::Match($manifest, "foreach \(\`$part in ([^)]+)\)").Groups[1].Value -split ',' | ForEach-Object { $_.Trim(" '") } | Where-Object { $_ })
if (-not $parts.Count) { $problems.Add('The list of the parts of the module was not found in GraphConnectivityTests.psm1.') }
foreach ($part in $parts) {
    if (-not (Test-Path -LiteralPath (Join-Path $Destination "src\GraphConnectivityTests.$part.ps1"))) { $problems.Add("Missing in the package: src\GraphConnectivityTests.$part.ps1") }
}
$config = [IO.File]::ReadAllText((Join-Path $Destination 'config\GraphConnectivityTests.config.psd1'))
if ($config -notmatch 'contoso\.com') { $problems.Add('The configuration of the package must keep the contoso.com example values.') }
if ($problems.Count) { throw ("Package not valid ($Destination):`n - " + ($problems -join "`n - ")) }

# Same tests as the source code, and reports written under the package folder.
$expected = [string](& pwsh -NoProfile -Command "Import-Module '$packageRoot\GraphConnectivityTests.psd1'; (Get-GctTestCatalog).Count")
$loaded = & pwsh -NoProfile -Command "Import-Module '$Destination\GraphConnectivityTests.psd1'; (Get-GctTestCatalog).Count; (Import-GctConfiguration).OutputPath"
if ($LASTEXITCODE -ne 0 -or $loaded[0] -ne $expected -or -not ([string]$loaded[1]).StartsWith($Destination)) { throw "The module does not load correctly from the package (expected $expected tests): $loaded" }

# ---- Zip --------------------------------------------------------------------------------------------
if (-not $NoZip) { Compress-Archive -Path (Join-Path $Destination '*') -DestinationPath $zip -CompressionLevel Optimal }

$all = Get-ChildItem -LiteralPath $Destination -Recurse -File
Write-Host ''
Write-Host "  Graph Connectivity Tests $version - package ready" -ForegroundColor Green
Write-Host "  Folder   : $Destination"
if (-not $NoZip) { Write-Host ("  Zip      : {0} ({1:N1} MB)" -f $zip, ((Get-Item -LiteralPath $zip).Length / 1MB)) }
Write-Host ("  Content  : {0} files, {1:N1} MB" -f $all.Count, (($all | Measure-Object Length -Sum).Sum / 1MB))
Write-Host "  Check    : module loads, $expected tests, reports written under the package folder"
Write-Host "  Config   : contoso.com example values - fill in Target before the first run (user guide, chapter 1.1)"
Write-Host ''
$all | Sort-Object FullName | ForEach-Object { '    {0,10:N0}  {1}' -f $_.Length, $_.FullName.Substring($Destination.Length + 1) }
