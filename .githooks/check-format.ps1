$ErrorActionPreference = 'Stop'
$PSNativeCommandUseErrorActionPreference = $false
$projectRoot = Split-Path -Parent $PSScriptRoot
$dotnetCommand = Get-Command dotnet -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if (-not $dotnetCommand) {
    throw 'Install .NET 8 SDK and make dotnet available on PATH.'
}
$dotnetPath = $dotnetCommand.Source
$env:DOTNET_HOST_PATH = $dotnetPath
$env:DOTNET_ROOT = Split-Path -Parent $dotnetPath

Push-Location $projectRoot
try {
    foreach ($project in @('ReportExport.csproj', 'tests/ReportExport.Tests.csproj')) {
        if (-not (Test-Path -LiteralPath $project)) {
            throw "Missing $project. Use a branch that contains both application and tests."
        }
        $assets = Join-Path (Split-Path -Parent (Join-Path $projectRoot $project)) 'obj/project.assets.json'
        if (-not (Test-Path -LiteralPath $assets)) {
            throw "Restore dependencies first: dotnet restore $project"
        }
        Write-Host "Checking formatting: $project"
        for ($attempt = 1; $attempt -le 2; $attempt++) {
            $output = & $dotnetPath format whitespace $project --verify-no-changes --no-restore 2>&1
            $code = $LASTEXITCODE
            $output | ForEach-Object { Write-Host $_ }
            if ($code -eq 0) { break }
            if ($attempt -eq 2 -or ($output -join "`n") -notmatch 'Unable to locate dotnet CLI') {
                Write-Host "Commit blocked. Fix formatting with: dotnet format whitespace $project --no-restore"
                exit $code
            }
            Write-Warning 'Temporary .NET CLI discovery failure; retrying once.'
            Start-Sleep -Seconds 2
        }
    }
    Write-Host 'Formatting checks passed.'
}
finally {
    Pop-Location
}
