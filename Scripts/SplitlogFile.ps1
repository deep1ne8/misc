<#
.SYNOPSIS
    Log File Splitter
.DESCRIPTION
    Splits a large log file into smaller parts based on user-defined size.
#>

param()

# Prompt user for inputs
$LogFilePath = Read-Host "Enter full path of the log file"
$Destination = Read-Host "Enter destination folder path"
$Unit = Read-Host "Enter size unit (MB, KB, Lines)"
$Size = Read-Host "Enter maximum size per split (integer only)"

# Validate file
if (-not (Test-Path $LogFilePath)) {
    Write-Error "Log file not found at: $LogFilePath"
    exit
}

# Validate destination
if (-not (Test-Path $Destination)) {
    try {
        New-Item -ItemType Directory -Path $Destination -Force | Out-Null
    } catch {
        Write-Error "Destination not writable: $Destination"
        exit
    }
}

# Convert size input
switch ($Unit.ToUpper()) {
    "MB"    { $Bytes = [int64]$Size * 1MB }
    "KB"    { $Bytes = [int64]$Size * 1KB }
    "LINES" { $Bytes = $null } # handled separately
    default { Write-Error "Invalid unit. Use MB, KB, or Lines"; exit }
}

$BaseName = [System.IO.Path]::GetFileNameWithoutExtension($LogFilePath)
$Extension = [System.IO.Path]::GetExtension($LogFilePath)
$Counter = 1

if ($Unit.ToUpper() -eq "LINES") {
    # Split by lines
    $LineBuffer = @()
    $LineCount = 0

    Get-Content $LogFilePath | ForEach-Object {
        $LineBuffer += $_
        $LineCount++

        if ($LineCount -ge $Size) {
            $OutFile = Join-Path $Destination ("{0}_part{1}{2}" -f $BaseName, $Counter, $Extension)
            $LineBuffer | Out-File -FilePath $OutFile -Encoding utf8
            Write-Host "Created $OutFile"
            $Counter++
            $LineBuffer = @()
            $LineCount = 0
        }
    }

    # Write remaining lines
    if ($LineBuffer.Count -gt 0) {
        $OutFile = Join-Path $Destination ("{0}_part{1}{2}" -f $BaseName, $Counter, $Extension)
        $LineBuffer | Out-File -FilePath $OutFile -Encoding utf8
        Write-Host "Created $OutFile"
    }

} else {
    # Split by size
    $reader = [System.IO.File]::OpenRead($LogFilePath)
    $buffer = New-Object byte[] $Bytes

    while ($true) {
        $bytesRead = $reader.Read($buffer, 0, $Bytes)
        if ($bytesRead -le 0) { break }

        $OutFile = Join-Path $Destination ("{0}_part{1}{2}" -f $BaseName, $Counter, $Extension)
        $fs = [System.IO.File]::Create($OutFile)
        $fs.Write($buffer, 0, $bytesRead)
        $fs.Close()

        Write-Host "Created $OutFile"
        $Counter++
    }

    $reader.Close()
}

Write-Host "Splitting complete."
