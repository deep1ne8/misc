# Path to the Liongard agent log file
$LogPath = "C:\Program Files (x86)\LiongardInc\LiongardAgent\logs\agent.log"

# Validate file exists
if (-not (Test-Path $LogPath)) {
    Write-Host "Log file not found: $LogPath"
    
}

# Backup ACL
$OriginalACL = Get-Acl $LogPath

# Grant SYSTEM read access temporarily
$Rule = New-Object System.Security.AccessControl.FileSystemAccessRule("SYSTEM","Read","Allow")
$ACL = Get-Acl $LogPath
$ACL.AddAccessRule($Rule)
Set-Acl -Path $LogPath -AclObject $ACL

# Upload endpoint
$UploadUrl = "https://tmpfiles.org/api/v1/upload"

# Create form-data boundary
$Boundary = [System.Guid]::NewGuid().ToString()
$LF = "`r`n"

# Prepare header for multipart upload
$FileName = Split-Path $LogPath -Leaf
$Header = "--$Boundary$LF" +
          "Content-Disposition: form-data; name=`"file`"; filename=`"$FileName`"$LF" +
          "Content-Type: application/octet-stream$LF$LF"

$Footer = "$LF--$Boundary--$LF"

# Convert header/footer to bytes
$HeaderBytes = [System.Text.Encoding]::UTF8.GetBytes($Header)
$FooterBytes = [System.Text.Encoding]::UTF8.GetBytes($Footer)

# Create memory stream for full body
$Stream = New-Object System.IO.MemoryStream
$Stream.Write($HeaderBytes, 0, $HeaderBytes.Length)

# Stream file in chunks with progress bar
$ChunkSize = 4MB
$FileStream = [System.IO.File]::OpenRead($LogPath)
$Buffer = New-Object byte[] $ChunkSize

$TotalSize = $FileStream.Length
$Uploaded = 0

Write-Output "Uploading $FileName ($TotalSize bytes)..."

while (($Read = $FileStream.Read($Buffer, 0, $ChunkSize)) -gt 0) {

    # Write chunk
    $Stream.Write($Buffer, 0, $Read)
    $Uploaded += $Read

    # Progress bar
    $Percent = [math]::Round(($Uploaded / $TotalSize) * 100, 2)
    Write-Progress -Activity "Uploading log file" -Status "$Percent% complete" -PercentComplete $Percent
}

$FileStream.Close()
$Stream.Write($FooterBytes, 0, $FooterBytes.Length)
$Stream.Position = 0

try {
    $Response = Invoke-RestMethod -Uri $UploadUrl -Method Post -ContentType "multipart/form-data; boundary=$Boundary" -Body $Stream
    $FileUrl = $Response.data.url

    Write-Output ""
    Write-Output "==============================================="
    Write-Output "UPLOAD COMPLETE"
    Write-Output "Log file is available at:"
    Write-Output "$FileUrl"
    Write-Output "==============================================="
    Write-Output ""

    # Optional: save URL to a file for Datto RMM retrieval
    $OutFile = "C:\ProgramData\LiongardAgentUploadURL.txt"
    Set-Content -Path $OutFile -Value $FileUrl
    Write-Output "URL saved to: $OutFile"
}
catch {
    Write-Output "Upload failed: $($_.Exception.Message)"
}
finally {
    # Restore original ACL
    Set-Acl -Path $LogPath -AclObject $OriginalACL
}
