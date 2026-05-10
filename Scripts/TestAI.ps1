# Import the PSLLM module if not already imported
Import-Module PSLLM

# Set up configuration for local Llama model
$configPath = "$env:USERPROFILE\.psllm\config.json"
$configDir = [System.IO.Path]::GetDirectoryName($configPath)

if (-not (Test-Path $configDir)) {
    New-Item -Path $configDir -ItemType Directory -Force | Out-Null
}

$config = @{
    DefaultProvider = "Llama"
    Providers = @{
        Llama = @{
            ModelPath = "C:\Users\deep1ne\Mentor\models\llama-2-7b-chat.gguf"
            # Optional parameters you might want to adjust
            ContextSize = 2048
            GpuLayers = 0  # Set to higher number if you have GPU support
        }
    }
} | ConvertTo-Json -Depth 5

Set-Content -Path $configPath -Value $config

# Verify the configuration
Get-Content -Path $configPath

# Test the local Llama model
try {
    $response = Invoke-PSLLM -Prompt "Hello, please provide a simple PowerShell tip" -Provider Llama
    $response
} catch {
    Write-Error "Error using local Llama model: $_"
    Write-Host "Make sure llama.cpp is properly installed and accessible by PSLLM" -ForegroundColor Yellow
}