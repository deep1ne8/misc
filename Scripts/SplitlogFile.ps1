# PowerShell script to split a file into smaller chunks

# Prompt for input file path
$InputFile = Read-Host "Enter the full path of the file to split"

# Prompt for chunk size (in lines or MB)
$Choice = Read-Host "Do you want to split by (1) number of lines per chunk OR (2) size in MB? Enter 1 or 2"

if ($Choice -eq 1) {
    $LinesPerChunk = Read-Host "Enter the number of lines per chunk"
    $LinesPerChunk = [int]$LinesPerChunk

    $Counter = 1
    Get-Content $InputFile -ReadCount $LinesPerChunk | ForEach-Object {
        $OutputFile = "{0}_part{1}.txt" -f $InputFile, $Counter
        $_ | Set-Content $OutputFile
        Write-Host "Created $OutputFile"
        $Counter++
    }

} elseif ($Choice -eq 2) {
    $SizeMB = Read-Host "Enter chunk size in MB"
    $SizeBytes = $SizeMB * 1MB

    $InputStream = [System.IO.File]::OpenRead($InputFile)
    $Buffer = New-Object byte[] $SizeBytes
    $Counter = 1

    while (($BytesRead = $InputStream.Read($Buffer, 0, $Buffer.Length)) -gt 0) {
        $OutputFile = "{0}_part{1}.bin" -f $InputFile, $Counter
        $OutputStream = [System.IO.File]::Create($OutputFile)
        $OutputStream.Write($Buffer, 0, $BytesRead)
        $OutputStream.Close()
        Write-Host "Created $OutputFile"
        $Counter++
    }
    $InputStream.Close()
} else {
    Write-Host "Invalid choice. Please run again and choose 1 or 2."
}

 
