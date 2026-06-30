$ErrorActionPreference = "Stop"
$project = Join-Path $PSScriptRoot "XlsxToXlsBatch\XlsxToXlsBatch.csproj"
$output = Join-Path $PSScriptRoot "publish"

dotnet restore $project --ignore-failed-sources
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

dotnet publish $project -c Release --no-restore -o $output
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

Write-Host ""
Write-Host "Build completed:" -ForegroundColor Green
Get-ChildItem $output -Filter *.exe | Select-Object -ExpandProperty FullName