$ErrorActionPreference = 'Stop'
$projectRoot = Split-Path -Parent $PSScriptRoot
$hook = Join-Path $projectRoot '.githooks/pre-commit'
if (-not (Test-Path -LiteralPath $hook)) {
    throw 'The pre-commit hook is missing from this checkout.'
}
$existing = git -C $projectRoot config --local --get core.hooksPath
if ($existing -and $existing -ne '.githooks') {
    throw "A different hooks directory is already configured: $existing. Review it before replacing it."
}
git -C $projectRoot config --local core.hooksPath .githooks
if ($LASTEXITCODE -ne 0) { throw 'Could not configure Git hooks.' }
Write-Host 'Pre-commit hook enabled for this checkout.'
Write-Host 'Before the first commit, restore dependencies:'
Write-Host '  dotnet restore ReportExport.csproj'
Write-Host '  dotnet restore tests/ReportExport.Tests.csproj'
