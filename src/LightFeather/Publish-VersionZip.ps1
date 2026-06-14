param(
    [string]$ProjectPath = (Join-Path $PSScriptRoot 'LightFeather.csproj'),
    [string]$PublishFolder = (Join-Path $PSScriptRoot 'publish'),
    [string]$OutputFolder = $PSScriptRoot
)

$ErrorActionPreference = 'Stop'

if (-not (Test-Path -LiteralPath $ProjectPath)) {
    throw "Nie znaleziono pliku projektu: $ProjectPath"
}

if (-not (Test-Path -LiteralPath $PublishFolder)) {
    throw "Nie znaleziono folderu publish: $PublishFolder"
}

[xml]$projectXml = Get-Content -LiteralPath $ProjectPath
$ns = New-Object System.Xml.XmlNamespaceManager($projectXml.NameTable)
$ns.AddNamespace('msb', 'http://schemas.microsoft.com/developer/msbuild/2003')

$applicationVersionNode = $projectXml.SelectSingleNode('//msb:ApplicationVersion', $ns)
if ($null -eq $applicationVersionNode -or [string]::IsNullOrWhiteSpace($applicationVersionNode.InnerText)) {
    throw "Nie znaleziono ApplicationVersion w pliku: $ProjectPath"
}

$version = $applicationVersionNode.InnerText.Trim()
$zipName = "LightFeather-$version.zip"
$zipPath = Join-Path $OutputFolder $zipName

if (-not (Test-Path -LiteralPath $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder | Out-Null
}

if (Test-Path -LiteralPath $zipPath) {
    Remove-Item -LiteralPath $zipPath -Force
}

# Pakujemy zawartosc folderu publish (bez nadrzednego katalogu publish).
Compress-Archive -Path (Join-Path $PublishFolder '*') -DestinationPath $zipPath -CompressionLevel Optimal

Write-Host "Gotowe: $zipPath"
