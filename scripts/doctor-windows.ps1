$ErrorActionPreference = "Stop"

if ($IsLinux -or $IsMacOS) {
    Write-Host "This doctor command currently supports Windows only."
    Exit
}

$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$RepoRoot = Split-Path -Parent $ScriptDir
$DistEntry = "$RepoRoot\dist\index.js"
$CepTargetDir = "$env:APPDATA\Adobe\CEP\extensions\MCPBridgeCEP"
$ClaudeConfigPath = "$env:APPDATA\Claude\claude_desktop_config.json"
$TempDir = "$env:TEMP\premiere-mcp-bridge"
$Failures = 0

function Pass-Check ($Message) {
    Write-Host "[ok] $Message" -ForegroundColor Green
}

function Fail-Check ($Message) {
    Write-Host "[missing] $Message" -ForegroundColor Red
    $script:Failures++
}

function Info-Check ($Message) {
    Write-Host "[info] $Message" -ForegroundColor Cyan
}

if (Get-Command node -ErrorAction SilentlyContinue) {
    $NodeVersion = node -v
    $NodeMajor = [int]($NodeVersion -replace '^v', '').Split('.')[0]
    if ($NodeMajor -ge 18) {
        Pass-Check "Node.js available ($NodeVersion)"
    } else {
        Fail-Check "Node.js 18+ required (found $NodeVersion)"
    }
} else {
    Fail-Check "Node.js not found in PATH"
}

if (Test-Path $DistEntry) {
    Pass-Check "Built MCP server found at $DistEntry"
} else {
    Fail-Check "Build output missing at $DistEntry (run npm run build)"
}

if (Test-Path $CepTargetDir) {
    if ((Test-Path "$CepTargetDir\CSXS\manifest.xml") -and (Test-Path "$CepTargetDir\index.html")) {
        Pass-Check "Premiere CEP extension installed at $CepTargetDir"
    } else {
        Fail-Check "CEP extension folder exists but is incomplete at $CepTargetDir"
    }
} else {
    Fail-Check "Premiere CEP extension not installed at $CepTargetDir"
}

if (Test-Path $TempDir) {
    Pass-Check "Bridge temp directory exists at $TempDir"
} else {
    Fail-Check "Bridge temp directory missing at $TempDir"
}

foreach ($version in "12", "11", "10") {
    $RegistryPath = "HKCU:\Software\Adobe\CSXS.$version"
    try {
        $Value = Get-ItemProperty -Path $RegistryPath -Name "PlayerDebugMode" -ErrorAction Stop
        if ($Value.PlayerDebugMode -eq "1") {
            Pass-Check "Adobe CEP debug mode enabled for CSXS.$version"
        } else {
            Fail-Check "Adobe CEP debug mode not enabled for CSXS.$version (is $($Value.PlayerDebugMode))"
        }
    } catch {
        Fail-Check "Adobe CEP debug mode not enabled for CSXS.$version (Registry key not found)"
    }
}

if (Test-Path $ClaudeConfigPath) {
    try {
        $ConfigRaw = Get-Content -Path $ClaudeConfigPath -Raw
        $ConfigData = $ConfigRaw | ConvertFrom-Json
        $Server = $ConfigData.mcpServers."premiere-pro"

        if (-not $Server) {
            Fail-Check "Claude Desktop config is present but missing the premiere-pro entry"
        } else {
            $Arg0 = if ($Server.args) { $Server.args[0] } else { "" }
            $Temp = if ($Server.env) { $Server.env.PREMIERE_TEMP_DIR } else { "" }

            if ($Server.command -ne "node") {
                Fail-Check "Claude Desktop config has a premiere-pro entry with the wrong command ($($Server.command))"
            } elseif ($Arg0 -ne ($DistEntry -replace '\\', '\\\\')) {
                # Note: powershell might handle this differently, so just check string equality mostly
                $ExpectedDist = $DistEntry.Replace('\', '\\')
                if ($Arg0 -eq $ExpectedDist -or $Arg0 -eq $DistEntry) {
                     # good
                } else {
                    Fail-Check "Claude Desktop config points to the wrong dist path ($Arg0)"
                }
            } elseif ($Temp -ne ($TempDir -replace '\\', '\\\\') -and $Temp -ne $TempDir) {
                Fail-Check "Claude Desktop config points to the wrong temp dir ($Temp)"
            } else {
                Pass-Check "Claude Desktop config contains a valid premiere-pro entry"
            }
        }
    } catch {
        Fail-Check "Claude Desktop config is not valid JSON"
    }
} else {
    Fail-Check "Claude Desktop config not found at $ClaudeConfigPath"
}

Info-Check "Premiere panel check must still be done manually inside Premiere Pro."
Info-Check "Open Window > Extensions > MCP Bridge (CEP), then click Test Connection."

if ($Failures -gt 0) {
    Write-Host ""
    Write-Host "Doctor found $Failures issue(s)."
    Exit
}

Write-Host ""
Write-Host "Doctor check passed."
