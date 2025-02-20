write-host "Audit Local Users + Administrators"
write-host "========================================="

#display the logged-in user
(((query user) -replace '>', '') -replace '\s{2,}', ',' | ConvertFrom-Csv).$(($((((query user) -replace '>', '') -replace '\s{2,}', ',' | ConvertFrom-Csv)) -as [string]).split('=')[0] -replace '@{','') | % {
    if ($_) {
        $varCurrentUser=$_
    }
}
write-host "- Logged-in User: $varCurrentUser"
write-host "- Reminder: This Component only audits local users, not users that are part"
write-host "  of a domain or network. The logged-on user may not appear in this list."
write-host "========================================="

#make our user-array
$arrUser=@{}

#enumerate the users known to the WMI
Get-WmiObject -Class Win32_UserAccount -Filter "LocalAccount=True" | % {

    $varSID=$_.SID #because $_ gets remapped in the device-admin query
    $varObject = New-Object PSObject

    #output some friendly information
    write-host "      SID: $varSID"
    write-host " Username: $($_.Name)"
    $varObject | Add-Member -MemberType NoteProperty -Name "Name" -Value $_.Name

    #is this user a device administrator?
    if ((Get-WMIObject -Class Win32_Group -Filter "LocalAccount=TRUE and SID='S-1-5-32-544'").GetRelated("Win32_Account","","","","PartComponent","GroupComponent",$FALSE,$NULL) | ? {$_.Domain -match $env:COMPUTERNAME -and $_.SID -match $varSID}) {
        write-host " Is Admin: YES"
        $varObject | Add-Member -MemberType NoteProperty -Name "isAdmin" -Value $true
    } else {
        write-host " Is Admin: NO"
        $varObject | Add-Member -MemberType NoteProperty -Name "isAdmin" -Value $false
    }

    #has this user ever actually logged into the device?
    if ((Get-ItemProperty "HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProfileList\$($_.SID)" -Name ProfileImagePath -ea 0).ProfileImagePath) {
        write-host " Activity: User has logged onto device at least once"
        $varObject | Add-Member -MemberType NoteProperty -Name "isActive" -Value $true
    } else {
        write-host " Activity: User has never logged onto device"
        $varObject | Add-Member -MemberType NoteProperty -Name "isActive" -Value $false
    }

    #is this account enabled?
    write-host " Disabled: $($_.Disabled)"
    $varObject | Add-Member -MemberType NoteProperty -Name "isDisabled" -Value ($_.Disabled -as [bool])

    #add this object to our array
    $arrUser+=@{$varSID=$varObject}
    write-host `r
}

switch ($env:usrAction) {
    'AdminAll' {
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (Administrator, Enabled) :: "}
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $true} | % {$varString+="$($_.Name) (Administrator, Disabled) :: "}
        write-host "- UDF Option: Write all Administrator Users to UDF $env:usrUDF"
    } 'AdminEnabled' {
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (Administrator, Enabled) :: "}
        write-host "- UDF Option: Write all Enabled Administrator Users to UDF $env:usrUDF"
    } 'AllAll' {
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (Administrator, Enabled) :: "}
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $true} | % {$varString+="$($_.Name) (Administrator, Disabled) :: "}
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -eq $false -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (StandardUser, Enabled) :: "}
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -eq $false -and $_.isDisabled -eq $true} | % {$varString+="$($_.Name) (StandardUser, Disabled) :: "}
        write-host "- UDF Option: Write all Users to UDF $env:usrUDF"
    } 'AllEnabled' {
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (Administrator, Enabled) :: "}
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin -eq $false -and $_.isDisabled -eq $false} | % {$varString+="$($_.Name) (StandardUser, Enabled) :: "}
        write-host "- UDF Option: Write all Enabled Users to UDF $env:usrUDF"
    } default {
        write-host "! ERROR: Unknown input type $env:usrAction."
        write-host "  Please report this issue to Support."
    }
}


if ($env:usrWarnOnAdmin -eq 'true') {
    if ($arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin} | ? {$_.isDisabled -eq $false}) {
        write-host `r
        write-host ": Warn-on-Admin enabled: The following active local accounts are Administrators."
        $arrUser.keys | % {$arrUser[$_]} | ? {$_.isAdmin} | ? {$_.isDisabled -eq $false} | % {
            write-host "- $($_.Name)"
        }
    }
}