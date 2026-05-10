# CertKeyVerifier.ps1 (Updated with Native + OpenSSL Hybrid Verification)
# GUI tool to verify a private key matches a certificate and export to various formats

Add-Type -AssemblyName PresentationFramework

function Compare-KeyAndCert {
    param(
        [string]$KeyPath,
        [string]$CertPath,
        [ref]$Result
    )
    try {
        $pem = Get-Content -Path $KeyPath -Raw
        if ($pem -notmatch '-----BEGIN [A-Z ]*PRIVATE KEY-----') {
            throw "Invalid or unsupported private key format."
        }
        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportFromPem($pem)
        $keyMod = [Convert]::ToBase64String($rsa.ExportParameters($false).Modulus)

        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertPath)
        $pubRsa = $cert.GetRSAPublicKey()
        $certMod = [Convert]::ToBase64String($pubRsa.ExportParameters($false).Modulus)

        $Result.Value = ($keyMod -eq $certMod)
        return $Result.Value
    } catch {
        $Result.Value = $false
        return $false
    }
}

# Load GUI (previously loaded XAML string code remains unchanged)
# ... [XAML loading block unchanged above] ...

# Hook Verify Button
$VerifyBtn.Add_Click({
    $keyPath  = $KeyPathBox.Text
    $certPath = $CertPathBox.Text

    if (-not (Test-Path $keyPath))  { [System.Windows.MessageBox]::Show('Please select a private key file.') ; return }
    if (-not (Test-Path $certPath)) { [System.Windows.MessageBox]::Show('Please select a certificate file.') ; return }

    $matched = $false

    # Try native PowerShell first
    Compare-KeyAndCert -KeyPath $keyPath -CertPath $certPath -Result ([ref]$matched) | Out-Null

    # Fallback to OpenSSL if native failed
    if (-not $matched) {
        $openssl = Get-Command openssl -ErrorAction SilentlyContinue
        if ($openssl) {
            try {
                $certOutput = & $openssl.Source x509 -noout -modulus -in "$certPath" 2>&1
                $certModulus = ($certOutput -join "") -replace 'Modulus=', '' -replace '\s',''

                $keyOutput = & $openssl.Source rsa -noout -modulus -in "$keyPath" 2>&1
                $keyModulus = ($keyOutput -join "") -replace 'Modulus=', '' -replace '\s',''

                if ($certModulus -and $keyModulus -and ($certModulus -eq $keyModulus)) {
                    $matched = $true
                }
            } catch {
                Write-Warning "OpenSSL verification failed: $_"
            }
        }
    }

    if ($matched) {
        $ResultText.Text = '✅ Key MATCHES certificate.'
        $ResultText.Foreground = 'Green'
    } else {
        $ResultText.Text = '❌ Key does NOT match certificate.'
        $ResultText.Foreground = 'Red'
    }
})

# Show GUI
$Form.ShowDialog() | Out-Null