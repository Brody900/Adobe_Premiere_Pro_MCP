$ErrorActionPreference = "Stop"

if ($IsLinux -or $IsMacOS) {
    Write-Host "This uninstaller currently supports Windows only."
    Exit
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
$CepTargetDir = "$env:APPDATA\Adobe\CEP\extensions\MCPBridgeCEP"
$ClaudeConfigPath = "$env:APPDATA\Claude\claude_desktop_config.json"
$TempDir = "$env:TEMP\premiere-mcp-bridge"

Write-Host "Removing Premiere CEP extension..."
if (Test-Path $CepTargetDir) {
    Remove-Item -Recurse -Force -Path $CepTargetDir
    Write-Host "Removed $CepTargetDir"
} else {
    Write-Host "No CEP extension found at $CepTargetDir"
}

Write-Host "Removing bridge temp directory..."
if (Test-Path $TempDir) {
    Remove-Item -Recurse -Force -Path $TempDir
    Write-Host "Removed $TempDir"
} else {
    Write-Host "No temp directory found at $TempDir"
}

Write-Host "Removing Claude Desktop config entry..."
if (Test-Path $ClaudeConfigPath) {
    try {
        $ConfigRaw = Get-Content -Path $ClaudeConfigPath -Raw
        $ConfigData = $ConfigRaw | ConvertFrom-Json -AsHashtable

        if ($ConfigData.mcpServers -and $ConfigData.mcpServers["premiere-pro"]) {
            $ConfigData.mcpServers.Remove("premiere-pro")
            $ConfigData | ConvertTo-Json -Depth 10 | Set-Content -Path $ClaudeConfigPath
            Write-Host "Removed premiere-pro entry from $ClaudeConfigPath"
        } else {
            Write-Host "No premiere-pro entry found in $ClaudeConfigPath"
        }
    } catch {
        Write-Error "Failed to update $ClaudeConfigPath. It may not be valid JSON."
    }
} else {
    Write-Host "No Claude Desktop config found at $ClaudeConfigPath"
}

Write-Host ""
Write-Host "Uninstall complete."
Write-Host "Note: Adobe CEP debug mode was left enabled as it may be used by other extensions."
