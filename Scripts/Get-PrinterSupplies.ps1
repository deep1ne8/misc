#Requires -Version 5.1
<#
.SYNOPSIS
    Test script to verify HP printer web scraping functionality
.DESCRIPTION
    Tests web access to HP printer and shows what data can be extracted
    Helps troubleshoot scraping patterns before running the main monitor
.PARAMETER PrinterIP
    IP address of the HP printer to test
.PARAMETER SavePages
    Save downloaded pages to files for manual inspection
#>

param(
    [string]$PrinterIP = "10.14.0.99",
    [switch]$SavePages = $false
)

Write-Host "HP Printer Web Scraper Test" -ForegroundColor Cyan
Write-Host "===========================" -ForegroundColor Cyan
Write-Host "Testing printer at: $PrinterIP" -ForegroundColor Yellow
Write-Host ""

function Test-PrinterConnectivity {
    param([string]$PrinterIP)
    
    Write-Host "[1] Testing network connectivity..." -ForegroundColor Green
    
    try {
        # Simple TCP connection test to port 80 (web interface)
        $TcpClient = New-Object System.Net.Sockets.TcpClient
        $ConnectTask = $TcpClient.ConnectAsync($PrinterIP, 80)
        $Timeout = 5000  # 5 seconds
        
        if ($ConnectTask.Wait($Timeout)) {
            if ($TcpClient.Connected) {
                Write-Host "    ✓ Printer web interface is accessible on port 80" -ForegroundColor Green
                $TcpClient.Close()
                return $true
            } else {
                Write-Host "    ✗ Cannot connect to printer web interface on port 80" -ForegroundColor Yellow
            }
        } else {
            Write-Host "    ✗ Connection to printer timed out" -ForegroundColor Yellow
        }
        
        $TcpClient.Close()
        
        # Fallback: Try a simple web request as connectivity test
        Write-Host "    Trying web request as fallback test..." -ForegroundColor Gray
        
        try {
            $TestUrl = "http://$PrinterIP/"
            $Request = [System.Net.HttpWebRequest]::Create($TestUrl)
            $Request.Method = "GET"
            $Request.Timeout = 5000  # 5 seconds
            $Request.UserAgent = "ConnectivityTest"
            
            $Response = $Request.GetResponse()
            $Response.Close()
            
            Write-Host "    ✓ Web interface responds" -ForegroundColor Green
            return $true
        }
        catch {
            Write-Host "    ✗ Web interface not accessible: $($_.Exception.Message.Split('.')[0])" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "    ✗ Connectivity test failed: $($_.Exception.Message.Split('.')[0])" -ForegroundColor Red
        if ($TcpClient) { $TcpClient.Close() }
        return $false
    }
}

function Test-WebAccess {
    param([string]$PrinterIP, [bool]$SavePages)
    
    Write-Host "[2] Testing web interface access..." -ForegroundColor Green
    
    # Common HP printer web pages (with your specific supply status page first)
    $TestPages = @(
        @{Name="Supply Status (Primary)"; URL="http://$PrinterIP/hp/device/info_suppliesStatus.html?tab=Home&menu=SupplyStatus"},
        @{Name="Main Page"; URL="http://$PrinterIP/"},
        @{Name="Status Page"; URL="http://$PrinterIP/status.html"},
        @{Name="Device Status"; URL="http://$PrinterIP/devicestatus/supplies"},
        @{Name="HP Supplies"; URL="http://$PrinterIP/hp/device/supplies.htm"},
        @{Name="HP Ink/Toner"; URL="http://$PrinterIP/hp/device/InkAndToner.htm"},
        @{Name="Info Config"; URL="http://$PrinterIP/info_config.html"},
        @{Name="Device Info"; URL="http://$PrinterIP/device_info.html"}
    )
    
    $AccessiblePages = @()
    
    foreach ($Page in $TestPages) {
        try {
            Write-Host "    Testing: $($Page.Name)" -NoNewline
            
            # Use HttpWebRequest instead of WebClient for better timeout control
            $Request = [System.Net.HttpWebRequest]::Create($Page.URL)
            $Request.Method = "GET"
            $Request.Timeout = 10000  # 10 seconds
            $Request.UserAgent = "PowerShell PrinterTest/1.0"
            
            $Response = $Request.GetResponse()
            $Stream = $Response.GetResponseStream()
            $Reader = New-Object System.IO.StreamReader($Stream)
            $Content = $Reader.ReadToEnd()
            
            $Reader.Close()
            $Stream.Close()
            $Response.Close()
            
            if ($Content.Length -gt 0) {
                Write-Host " ✓" -ForegroundColor Green
                $AccessiblePages += @{
                    Name = $Page.Name
                    URL = $Page.URL
                    Content = $Content
                    Size = $Content.Length
                }
                
                if ($SavePages) {
                    $FileName = "printer_page_$($Page.Name -replace '[^a-zA-Z0-9]','_').html"
                    $Content | Out-File -FilePath $FileName -Encoding UTF8
                    Write-Host "      Saved to: $FileName" -ForegroundColor Gray
                }
            } else {
                Write-Host " ✗ (Empty response)" -ForegroundColor Yellow
            }
        }
        catch {
            Write-Host " ✗ ($($_.Exception.Message.Split('.')[0]))" -ForegroundColor Red
        }
    }
    
    Write-Host "    Found $($AccessiblePages.Count) accessible pages" -ForegroundColor Cyan
    return $AccessiblePages
}

function Test-SupplyExtraction {
    param([array]$AccessiblePages)
    
    Write-Host "[3] Testing supply data extraction..." -ForegroundColor Green
    
    $AllSupplies = @{}
    $ExtractionResults = @()
    
    foreach ($Page in $AccessiblePages) {
        Write-Host "    Analyzing: $($Page.Name)" -ForegroundColor Cyan
        
        $FoundSupplies = @{}
        $Content = $Page.Content
        
        # Pattern 1: Direct percentage matches with supply names
        Write-Host "      Pattern 1 (Name: XX%)" -NoNewline
        $Pattern1 = '(?i)(black|cyan|magenta|yellow|maintenance)[\s\w]*[:\s]+(\d+)%'
        $Matches1 = [regex]::Matches($Content, $Pattern1)
        
        foreach ($Match in $Matches1) {
            $SupplyName = $Match.Groups[1].Value.Trim().ToLower()
            $Percentage = [int]$Match.Groups[2].Value
            if ($Percentage -le 100) {
                $FoundSupplies["Pattern1_$SupplyName"] = $Percentage
            }
        }
        Write-Host " ($($Matches1.Count) matches)" -ForegroundColor Gray
        
        # Pattern 2: HP-specific data attributes
        Write-Host "      Pattern 2 (HP data attrs)" -NoNewline
        $Pattern2 = 'data-[\w-]*(?:level|percent|remaining)[^>]*>(\d+)%?<'
        $Matches2 = [regex]::Matches($Content, $Pattern2)
        
        $Index = 1
        foreach ($Match in $Matches2) {
            $Percentage = [int]$Match.Groups[1].Value
            if ($Percentage -le 100) {
                $FoundSupplies["Pattern2_supply$Index"] = $Percentage
                $Index++
            }
        }
        Write-Host " ($($Matches2.Count) matches)" -ForegroundColor Gray
        
        # Pattern 3: JavaScript variables
        Write-Host "      Pattern 3 (JS variables)" -NoNewline
        $Pattern3 = '(?i)(black|cyan|magenta|yellow|maintenance)[\w\s]*["\s]*:\s*["\s]*(\d+)["\s]*[,%\s]'
        $Matches3 = [regex]::Matches($Content, $Pattern3)
        
        foreach ($Match in $Matches3) {
            $SupplyName = $Match.Groups[1].Value.Trim().ToLower()
            $Percentage = [int]$Match.Groups[2].Value
            if ($Percentage -le 100 -and -not $FoundSupplies.ContainsKey("Pattern1_$SupplyName")) {
                $FoundSupplies["Pattern3_$SupplyName"] = $Percentage
            }
        }
        Write-Host " ($($Matches3.Count) matches)" -ForegroundColor Gray
        
        # Pattern 3: HTML table cells
        Write-Host "      Pattern 3 (HTML table)" -NoNewline
        $Pattern3 = '<td[^>]*>(?i)(black|cyan|magenta|yellow|maintenance)[\s\w]*</td>[\s\S]*?<td[^>]*>(\d+)%?</td>'
        $Matches3 = [regex]::Matches($Content, $Pattern3)
        
        foreach ($Match in $Matches3) {
            $SupplyName = $Match.Groups[1].Value.Trim().ToLower()
            $Percentage = [int]$Match.Groups[2].Value
            if ($Percentage -le 100 -and -not $FoundSupplies.ContainsKey("Pattern1_$SupplyName") -and -not $FoundSupplies.ContainsKey("Pattern2_$SupplyName")) {
                $FoundSupplies["Pattern3_$SupplyName"] = $Percentage
            }
        }
        Write-Host " ($($Matches3.Count) matches)" -ForegroundColor Gray
        
        # Pattern 4: Any percentage values (fallback)
        Write-Host "      Pattern 4 (All percentages)" -NoNewline
        $Pattern4 = '(\d+)%'
        $Matches4 = [regex]::Matches($Content, $Pattern4)
        $ValidPercentages = @()
        foreach ($Match in $Matches4) {
            $Value = [int]$Match.Groups[1].Value
            if ($Value -le 100) {
                $ValidPercentages += $Value
            }
        }
        Write-Host " ($($ValidPercentages.Count) valid percentages: $($ValidPercentages -join ', '))" -ForegroundColor Gray
        
        # Pattern 5: Look for specific HP printer supply keywords
        Write-Host "      Pattern 5 (HP-specific terms)" -NoNewline
        $HPKeywords = @('cartridge', 'toner', 'supply', 'remaining', 'level', 'status')
        $KeywordCount = 0
        foreach ($Keyword in $HPKeywords) {
            if ($Content -match "(?i)$Keyword") {
                $KeywordCount++
            }
        }
        Write-Host " ($KeywordCount HP keywords found)" -ForegroundColor Gray
        
        # Store results for this page
        $ExtractionResults += @{
            PageName = $Page.Name
            URL = $Page.URL
            FoundSupplies = $FoundSupplies
            ValidPercentages = $ValidPercentages
            HPKeywordCount = $KeywordCount
            ContentLength = $Content.Length
        }
        
        # Add to master supply list
        foreach ($Key in $FoundSupplies.Keys) {
            $AllSupplies[$Key] = $FoundSupplies[$Key]
        }
    }
    
    return @{
        AllSupplies = $AllSupplies
        DetailedResults = $ExtractionResults
    }
}

function Show-Results {
    param([hashtable]$Results)
    
    Write-Host "[4] EXTRACTION RESULTS" -ForegroundColor Green
    Write-Host "======================" -ForegroundColor Green
    
    if ($Results.AllSupplies.Count -eq 0) {
        Write-Host "    ✗ No supply data found!" -ForegroundColor Red
        Write-Host ""
        Write-Host "TROUBLESHOOTING SUGGESTIONS:" -ForegroundColor Yellow
        Write-Host "1. Check if SNMP is available instead of web scraping" -ForegroundColor Yellow
        Write-Host "2. Try accessing printer web interface manually" -ForegroundColor Yellow
        Write-Host "3. Look for different supply page URLs" -ForegroundColor Yellow
        Write-Host "4. Check if printer requires authentication" -ForegroundColor Yellow
    }
    else {
        Write-Host "    ✓ Found supply data:" -ForegroundColor Green
        foreach ($Supply in $Results.AllSupplies.Keys) {
            $CleanName = $Supply -replace '^Pattern\d+_', ''
            Write-Host "      $($CleanName.ToUpper()): $($Results.AllSupplies[$Supply])%" -ForegroundColor White
        }
    }
    
    Write-Host ""
    Write-Host "DETAILED PAGE ANALYSIS:" -ForegroundColor Cyan
    Write-Host "----------------------" -ForegroundColor Cyan
    
    foreach ($PageResult in $Results.DetailedResults) {
        Write-Host "  $($PageResult.PageName):" -ForegroundColor White
        Write-Host "    URL: $($PageResult.URL)" -ForegroundColor Gray
        Write-Host "    Content Size: $($PageResult.ContentLength) bytes" -ForegroundColor Gray
        Write-Host "    HP Keywords: $($PageResult.HPKeywordCount)" -ForegroundColor Gray
        Write-Host "    Supply Matches: $($PageResult.FoundSupplies.Count)" -ForegroundColor Gray
        Write-Host "    Valid Percentages: $($PageResult.ValidPercentages.Count)" -ForegroundColor Gray
        
        if ($PageResult.FoundSupplies.Count -gt 0) {
            Write-Host "    Supplies found:" -ForegroundColor Green
            foreach ($Supply in $PageResult.FoundSupplies.Keys) {
                Write-Host "      - $Supply = $($PageResult.FoundSupplies[$Supply])%" -ForegroundColor Green
            }
        }
        Write-Host ""
    }
}

function Show-Recommendations {
    param([hashtable]$Results, [array]$AccessiblePages)
    
    Write-Host "[5] RECOMMENDATIONS" -ForegroundColor Green
    Write-Host "==================" -ForegroundColor Green
    
    if ($AccessiblePages.Count -eq 0) {
        Write-Host "❌ No web pages accessible - try SNMP monitoring instead" -ForegroundColor Red
        Write-Host "   SNMP OIDs for HP M477fnw:" -ForegroundColor Yellow
        Write-Host "   Black Toner: 1.3.6.1.2.1.43.11.1.1.9.1.1" -ForegroundColor Yellow
        Write-Host "   Cyan Toner: 1.3.6.1.2.1.43.11.1.1.9.1.2" -ForegroundColor Yellow
        Write-Host "   Magenta Toner: 1.3.6.1.2.1.43.11.1.1.9.1.3" -ForegroundColor Yellow
        Write-Host "   Yellow Toner: 1.3.6.1.2.1.43.11.1.1.9.1.4" -ForegroundColor Yellow
    }
    elseif ($Results.AllSupplies.Count -eq 0) {
        Write-Host "⚠️ Web access works but no supply data found" -ForegroundColor Yellow
        Write-Host "   Try manual inspection of saved pages or different extraction patterns" -ForegroundColor Yellow
        Write-Host "   Consider SNMP as alternative monitoring method" -ForegroundColor Yellow
    }
    else {
        Write-Host "✅ Web scraping is working!" -ForegroundColor Green
        
        # Find the best page for monitoring
        $BestPage = $Results.DetailedResults | Sort-Object { $_.FoundSupplies.Count } -Descending | Select-Object -First 1
        Write-Host "   Best page for monitoring: $($BestPage.PageName)" -ForegroundColor Green
        Write-Host "   URL: $($BestPage.URL)" -ForegroundColor Green
        Write-Host "   Supplies detected: $($BestPage.FoundSupplies.Count)" -ForegroundColor Green
        
        Write-Host ""
        Write-Host "✅ Ready to use main monitoring script with these settings:" -ForegroundColor Green
        Write-Host "   PowerShell -File HPPrinterMonitor.ps1 -PrinterIP $PrinterIP" -ForegroundColor White
    }
}

# Main execution
try {
    # Test 1: Network connectivity
    if (-not (Test-PrinterConnectivity -PrinterIP $PrinterIP)) {
        Write-Host ""
        Write-Host "⚠️ Cannot reach printer web interface, but continuing with web tests..." -ForegroundColor Yellow
        # Don't exit here - continue with web tests anyway
    }
    
    Write-Host ""
    
    # Test 2: Web access
    $AccessiblePages = Test-WebAccess -PrinterIP $PrinterIP -SavePages $SavePages
    
    Write-Host ""
    
    # Test 3: Supply extraction
    $Results = Test-SupplyExtraction -AccessiblePages $AccessiblePages
    
    Write-Host ""
    
    # Test 4: Show results
    Show-Results -Results $Results
    
    Write-Host ""
    
    # Test 5: Recommendations
    Show-Recommendations -Results $Results -AccessiblePages $AccessiblePages
    
    Write-Host ""
    Write-Host "Test completed!" -ForegroundColor Cyan
    
    if ($SavePages) {
        Write-Host ""
        Write-Host "HTML pages saved to current directory for manual inspection." -ForegroundColor Gray
    }
}
catch {
    Write-Host ""
    Write-Host "❌ Test failed: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Stack trace:" -ForegroundColor Gray
    Write-Host $_.ScriptStackTrace -ForegroundColor Gray
}
finally {
    Write-Host ""
    Write-Host "Press any key to continue..." -ForegroundColor Gray
    $null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
}
