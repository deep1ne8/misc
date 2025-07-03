# CertKeyVerifier.ps1
# GUI tool to verify a private key matches a certificate and export to various formats
# Requires PowerShell 7+ (.NET 5+) for ImportFromPem / CopyWithPrivateKey

<#+
.SYNOPSIS
    Graphical utility to compare a private key file with an X.509 certificate and export the certificate
    in PEM, DER, PKCS#7 or PFX format.
.DESCRIPTION
    * Upload or browse to a private key file (.key/.pem) and certificate file (.cer/.crt/.pem).
    * Click **Verify** to confirm the modulus of the private key matches that of the certificate.
    * Choose an export format, output path and (for PFX) a password, then click **Export**.
.NOTES
    Author : ChatGPT (OpenAI) – generated for Earl
    Version: 1.0 – 2025‑07‑03
    Requires: Windows PowerShell 7+, .NET 5+, RSA keys in PEM (PKCS#1 or PKCS#8) or PKCS#12 (PFX) with matching certificate.
#>

# Ensure PresentationFramework for WPF GUI
Add-Type -AssemblyName PresentationFramework

# --------------- Helper Functions ---------------
function Show-Error ($msg) { [System.Windows.MessageBox]::Show($msg,'Error',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Error) }
function Show-Info  ($msg) { [System.Windows.MessageBox]::Show($msg,'Information',[System.Windows.MessageBoxButton]::OK,[System.Windows.MessageBoxImage]::Information) }

function Get-ModulusFromPrivateKey {
    param([string]$Path)
    $pem = Get-Content $Path -Raw
    try {
        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportFromPem($pem)
        $params = $rsa.ExportParameters($false)
        return [Convert]::ToBase64String($params.Modulus)
    } catch {
        throw "Unsupported or invalid private key format: $Path"
    }
}

function Get-ModulusFromCertificate {
    param([string]$Path)
    try {
        $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($Path)
        $rsaPublic = $cert.GetRSAPublicKey()
        $params = $rsaPublic.ExportParameters($false)
        return [Convert]::ToBase64String($params.Modulus)
    } catch {
        throw "Unable to read certificate: $Path"
    }
}

function Compare-KeyAndCert {
    param([string]$KeyPath,[string]$CertPath,[ref]$Result)
    try {
        $keyMod  = Get-ModulusFromPrivateKey $KeyPath
        $certMod = Get-ModulusFromCertificate $CertPath
        if ($keyMod -eq $certMod) { $Result.Value = $true; return $true }
        else { $Result.Value = $false; return $false }
    } catch {
        Show-Error $_.Exception.Message
        return $false
    }
}

function Export-Certificate {
    param(
        [string]$CertPath,
        [string]$KeyPath,
        [ValidateSet('PEM','DER','P7B','PFX')][string]$Format,
        [string]$OutFile,
        [System.Security.SecureString]$Password
    )

    switch ($Format) {
        'PEM' {
            $bytes = [IO.File]::ReadAllBytes($CertPath)
            $pemBody = [Convert]::ToBase64String($bytes,[System.Base64FormattingOptions]::InsertLineBreaks)
            $pem = "-----BEGIN CERTIFICATE-----`n$pemBody`n-----END CERTIFICATE-----"
            Set-Content -Path $OutFile -Value $pem -NoNewline
        }
        'DER' {
            Copy-Item -Path $CertPath -Destination $OutFile -Force
        }
        'P7B' {
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertPath)
            $bytes = $cert.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs7)
            [IO.File]::WriteAllBytes($OutFile,$bytes)
        }
        'PFX' {
            if (-not (Test-Path $KeyPath)) { throw 'Private key path is required for PFX export.' }
            if (-not $Password) { throw 'A password is required for PFX export.' }
            $cert = [System.Security.Cryptography.X509Certificates.X509Certificate2]::new($CertPath)
            $pemKey = Get-Content $KeyPath -Raw
            $rsa = [System.Security.Cryptography.RSA]::Create()
            $rsa.ImportFromPem($pemKey)
            $certWithKey = $cert.CopyWithPrivateKey($rsa)
            $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($Password)
            try {
                $plainPwd = [System.Runtime.InteropServices.Marshal]::PtrToStringBSTR($BSTR)
                $bytes = $certWithKey.Export([System.Security.Cryptography.X509Certificates.X509ContentType]::Pkcs12, $plainPwd)
            } finally {
                [System.Runtime.InteropServices.Marshal]::ZeroFreeBSTR($BSTR)
            }
            [IO.File]::WriteAllBytes($OutFile,$bytes)
        }
    }
}

function Show-ExportHelp {
    $msg = @"
Export formats:
  • PEM  (.pem/.crt) – Base‑64 text; widely used on Linux/Apache/Nginx.
  • DER  (.cer/.der) – Binary; preferred by Java keystores & Windows MMC.
  • P7B  (.p7b/.spc) – PKCS#7, certificate chain only; IIS import, Java trust stores.
  • PFX  (.pfx/.p12) – PKCS#12, includes private key; Windows, Azure, code‑signing.
"@
    Show-Info $msg
}

# --------------- XAML UI ---------------
[xml]$Xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation" Title="Certificate ↔ Key Verifier" Height="380" Width="640" WindowStartupLocation="CenterScreen">
  <Grid Margin="10" >
    <Grid.RowDefinitions>
      <RowDefinition Height="Auto"/>
      <RowDefinition Height="*"/>
      <RowDefinition Height="Auto"/>
    </Grid.RowDefinitions>

    <!-- File selectors -->
    <StackPanel Orientation="Vertical" Grid.Row="0" Margin="0,0,0,10">
      <StackPanel Orientation="Horizontal" Margin="0,4">
        <Label Content="Private Key:" Width="90"/>
        <TextBox x:Name="KeyPathBox" Width="400" IsReadOnly="True"/>
        <Button x:Name="BrowseKeyBtn" Content="Browse" Margin="5,0,0,0" Width="60"/>
      </StackPanel>
      <StackPanel Orientation="Horizontal" Margin="0,4">
        <Label Content="Certificate:" Width="90"/>
        <TextBox x:Name="CertPathBox" Width="400" IsReadOnly="True"/>
        <Button x:Name="BrowseCertBtn" Content="Browse" Margin="5,0,0,0" Width="60"/>
      </StackPanel>
    </StackPanel>

    <!-- Verify section -->
    <StackPanel Grid.Row="1" Orientation="Vertical" VerticalAlignment="Top">
      <Button x:Name="VerifyBtn" Content="Verify Key ↔ Cert" Height="32" Width="160" HorizontalAlignment="Left"/>
      <TextBlock x:Name="ResultText" FontSize="14" Margin="0,8,0,0"/>

      <!-- Export panel -->
      <GroupBox Header="Export Options" Margin="0,16,0,0" Width="600">
        <Grid Margin="8">
          <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
          </Grid.ColumnDefinitions>
          <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
          </Grid.RowDefinitions>

          <Label Grid.Row="0" Grid.Column="0" Content="Format:"/>
          <ComboBox x:Name="FormatBox" Grid.Row="0" Grid.Column="1" Width="120" SelectedIndex="0">
            <ComboBoxItem Content="PEM"/>
            <ComboBoxItem Content="DER"/>
            <ComboBoxItem Content="P7B"/>
            <ComboBoxItem Content="PFX"/>
          </ComboBox>
          <Button x:Name="HelpBtn" Grid.Row="0" Grid.Column="2" Content="?" Width="24"/>

          <Label Grid.Row="1" Grid.Column="0" Content="Output File:"/>
          <TextBox x:Name="OutPathBox" Grid.Row="1" Grid.Column="1" Width="320" IsReadOnly="True"/>
          <Button x:Name="BrowseOutBtn" Grid.Row="1" Grid.Column="2" Content="Browse" Width="60"/>

          <Label x:Name="PwdLbl" Grid.Row="2" Grid.Column="0" Content="PFX Password:" Visibility="Collapsed"/>
          <PasswordBox x:Name="PwdBox" Grid.Row="2" Grid.Column="1" Width="150" Visibility="Collapsed"/>
        </Grid>
      </GroupBox>
      <Button x:Name="ExportBtn" Content="Export" Height="28" Width="80" HorizontalAlignment="Left" Margin="0,6,0,0"/>
    </StackPanel>

    <!-- Footer -->
    <TextBlock Grid.Row="2" Text="© 2025 Earl Utilities – Powered by PowerShell 7 & .NET 5" HorizontalAlignment="Center" FontSize="10"/>
  </Grid>
</Window>
"@

# --------------- Load UI ---------------
$reader = New-Object System.Xml.XmlNodeReader $Xaml
try { $Window = [Windows.Markup.XamlReader]::Load($reader) } catch {
    Write-Error "Failed to load XAML: $_"
    return
}

# Access controls
$KeyPathBox    = $Window.FindName('KeyPathBox')
$CertPathBox   = $Window.FindName('CertPathBox')
$BrowseKeyBtn  = $Window.FindName('BrowseKeyBtn')
$BrowseCertBtn = $Window.FindName('BrowseCertBtn')
$VerifyBtn     = $Window.FindName('VerifyBtn')
$ResultText    = $Window.FindName('ResultText')
$FormatBox     = $Window.FindName('FormatBox')
$OutPathBox    = $Window.FindName('OutPathBox')
$BrowseOutBtn  = $Window.FindName('BrowseOutBtn')
$PwdLbl        = $Window.FindName('PwdLbl')
$PwdBox        = $Window.FindName('PwdBox')
$ExportBtn     = $Window.FindName('ExportBtn')
$HelpBtn       = $Window.FindName('HelpBtn')

# File dialog helpers
function Select-File($filter) {
    $dlg = New-Object Microsoft.Win32.OpenFileDialog
    $dlg.Filter = $filter
    if ($dlg.ShowDialog()) { return $dlg.FileName } else { return $null }
}
function Select-Save($filter,$defaultExt) {
    $dlg = New-Object Microsoft.Win32.SaveFileDialog
    $dlg.Filter = $filter
    $dlg.DefaultExt = $defaultExt
    if ($dlg.ShowDialog()) { return $dlg.FileName } else { return $null }
}

# Browse buttons
$BrowseKeyBtn.Add_Click({
    $p = Select-File "Key files (*.key;*.pem)|*.key;*.pem|All files (*.*)|*.*"
    if ($p) { $KeyPathBox.Text = $p }
})
$BrowseCertBtn.Add_Click({
    $p = Select-File "Certificate files (*.cer;*.crt;*.pem)|*.cer;*.crt;*.pem|All files (*.*)|*.*"
    if ($p) { $CertPathBox.Text = $p }
})
$BrowseOutBtn.Add_Click({
    $fmt = ($FormatBox.SelectedItem.Content.ToString())
    switch ($fmt) {
        'PEM' { $filter = 'PEM Certificate (*.pem)|*.pem' ; $ext = '.pem' }
        'DER' { $filter = 'DER Certificate (*.cer)|*.cer' ; $ext = '.cer' }
        'P7B' { $filter = 'PKCS#7 (*.p7b)|*.p7b'        ; $ext = '.p7b' }
        'PFX' { $filter = 'PFX Archive (*.pfx)|*.pfx'    ; $ext = '.pfx' }
    }
    $p = Select-Save $filter $ext
    if ($p) { $OutPathBox.Text = $p }
})

# Toggle password box for PFX
$FormatBox.Add_SelectionChanged({
    if ($FormatBox.SelectedItem.Content.ToString() -eq 'PFX') {
        $PwdLbl.Visibility = 'Visible'
        $PwdBox.Visibility = 'Visible'
    } else {
        $PwdLbl.Visibility = 'Collapsed'
        $PwdBox.Visibility = 'Collapsed'
    }
})

# Help
$HelpBtn.Add_Click({ Show-ExportHelp })

# Verify button
$VerifyBtn.Add_Click({
    $keyPath  = $KeyPathBox.Text
    $certPath = $CertPathBox.Text
    if (-not (Test-Path $keyPath))  { Show-Error 'Please select a private key file.'; return }
    if (-not (Test-Path $certPath)) { Show-Error 'Please select a certificate file.'; return }

    $matched = $false
    Compare-KeyAndCert -KeyPath $keyPath -CertPath $certPath -Result ([ref]$matched) | Out-Null
    if ($matched) {
        $ResultText.Text = '✅ Key MATCHES certificate.'
        $ResultText.Foreground = 'Green'
    } else {
        $ResultText.Text = '❌ Key does NOT match certificate.'
        $ResultText.Foreground = 'Red'
    }
})

# Export button
$ExportBtn.Add_Click({
    $certPath = $CertPathBox.Text
    $pfxPassword = $PwdBox.SecurePassword
    try {
        Export-Certificate -CertPath $certPath -KeyPath $keyPath -Format $fmt -OutFile $outPath -Password $pfxPassword
        Show-Info "Export successful → $outPath"
    } catch {
        Show-Error $_.Exception.Message
    }
    try {
        Export-Certificate -CertPath $certPath -KeyPath $keyPath -Format $fmt -OutFile $outPath -Password $pfxPassword
        Show-Info "Export successful → $outPath"
    } catch {
        Show-Error $_.Exception.Message
    }
})

# --------------- Launch GUI ---------------
$Window.ShowDialog() | Out-Null