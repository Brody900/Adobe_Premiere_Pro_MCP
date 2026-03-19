$ErrorActionPreference = "Stop"

if ($IsLinux -or $IsMacOS) {
    Write-Host "This installer currently supports Windows only."
    Exit
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
$CepExtensionsDir = "$env:APPDATA\Adobe\CEP\extensions"
$CepTargetDir = "$CepExtensionsDir\MCPBridgeCEP"
$ClaudeConfigPath = "$env:APPDATA\Claude\claude_desktop_config.json"
$TempDir = "$env:TEMP\premiere-mcp-bridge"
$DistEntry = "$RepoRoot\dist\index.js"

if (-not (Get-Command node -ErrorAction SilentlyContinue)) {
    Write-Host "Node.js 18+ is required but 'node' was not found."
    Exit
}

$NodeVersionOutput = node -v
$NodeMajor = [int]($NodeVersionOutput -replace '^v', '').Split('.')[0]
if ($NodeMajor -lt 18) {
    Write-Host "Node.js 18+ is required. Found: $NodeVersionOutput"
    Exit
}

Write-Host "Installing npm dependencies..."
Set-Location -Path $RepoRoot
npm install

Write-Host "Building MCP server..."
npm run build

if (-not (Test-Path $DistEntry)) {
    Write-Host "Build completed but dist\index.js was not created."
    Exit
}

Write-Host "Enabling Adobe CEP debug mode..."
foreach ($version in "12", "11", "10") {
    $RegistryPath = "HKCU:\Software\Adobe\CSXS.$version"
    if (-not (Test-Path $RegistryPath)) {
        New-Item -Path $RegistryPath -Force | Out-Null
    }
    Set-ItemProperty -Path $RegistryPath -Name "PlayerDebugMode" -Value "1" -Type String
}

Write-Host "Installing Premiere CEP extension..."
if (-not (Test-Path $CepExtensionsDir)) {
    New-Item -ItemType Directory -Force -Path $CepExtensionsDir | Out-Null
}
if (Test-Path $CepTargetDir) {
    Remove-Item -Recurse -Force -Path $CepTargetDir
}
Copy-Item -Recurse -Path "$RepoRoot\cep-plugin" -Destination $CepTargetDir

Write-Host "Preparing bridge temp directory..."
if (-not (Test-Path $TempDir)) {
    New-Item -ItemType Directory -Force -Path $TempDir | Out-Null
}

Write-Host "Updating Claude Desktop config..."
$ClaudeConfigDir = Split-Path -Parent $ClaudeConfigPath
if (-not (Test-Path $ClaudeConfigDir)) {
    New-Item -ItemType Directory -Force -Path $ClaudeConfigDir | Out-Null
}

$ConfigData = @{}
if (Test-Path $ClaudeConfigPath) {
    try {
        $ConfigRaw = Get-Content -Path $ClaudeConfigPath -Raw
        if (-not [string]::IsNullOrWhiteSpace($ConfigRaw)) {
            $ConfigData = $ConfigRaw | ConvertFrom-Json -AsHashtable
        }
    } catch {
        Write-Error "Claude Desktop config is not valid JSON: $ClaudeConfigPath"
        Exit
    }
}

if (-not $ConfigData.mcpServers) {
    $ConfigData.mcpServers = @{}
}

$ConfigData.mcpServers["premiere-pro"] = @{
    command = "node"
    args = @($DistEntry.Replace('\', '\\'))
    env = @{
        PREMIERE_TEMP_DIR = $TempDir.Replace('\', '\\')
    }
}

$ConfigData | ConvertTo-Json -Depth 10 | Set-Content -Path $ClaudeConfigPath

Write-Host ""
Write-Host "Install complete."
Write-Host "Next:"
Write-Host "1. Restart Claude Desktop."
Write-Host "2. Restart Premiere Pro."
Write-Host "3. Open Window > Extensions > MCP Bridge (CEP)."
Write-Host "4. Set Temp Directory to $TempDir."
Write-Host "5. Click Save Configuration, then Start Bridge, then Test Connection."
