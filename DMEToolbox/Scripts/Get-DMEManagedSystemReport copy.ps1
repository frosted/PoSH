Function Get-DMEManagedSystemReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]
        $MaximumAge = 30,
        [Parameter()]
        [string]
        $CSVFilePath = "$($env:TEMP)\DMEManagedSystemReport$(Get-Date -Format 'yyyyMMdd').csv"
    )

    #Checks Active Directory for objects that are Windows 10 AND the PasswordLastSet attribute is within the last 30 days and adds to variable $adResult.
    $oldestDate = (Get-Date).AddDays(-$MaximumAge)
    $adQuery = 'OperatingSystemversion -like "10.0*" -and PasswordLastSet -ge "' + $oldestDate + '" -and Enabled -eq "True"'
    $adQuery = 'OperatingSystemversion -like "10.0*"'
    $adQuery = '*'

    If (!(Get-Module -Name ActiveDirectory)) {
        try {
            #Gather Settings
            $WUServerSetting = Get-ItemPropertyValue -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer
            $RepairContentServerSource = Get-ItemPropertyValue -Path 'HKLM:Software\Microsoft\Windows\CurrentVersion\Policies\Servicing' -Name RepairContentServerSource

            #Allow WUServer
            If ($WUServerSetting -ne 0) {
                Set-ItemProperty -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer -Value 0
            }

            If ($RepairContentServerSource -ne 2) {
                Set-ItemProperty -Path 'HKLM:Software\Microsoft\Windows\CurrentVersion\Policies\Servicing' -Name RepairContentServerSource 2
            }

            Update-DMEGroupPolicy
    
            #Install RSAT AD optional feature
            Get-WindowsCapability -Online -Name Rsat.ActiveDirectory* | ForEach-Object { If ($_.State -ne 'Installed') { Add-WindowsCapability -Online -Name $_.name } }
    
            #Restore WUServer setting
            If ((Get-ItemPropertyValue -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer) -ne $WUServerSetting) {
                Set-ItemProperty -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer -Value $WUServerSetting
                Set-ItemPropertyValue -Path 'HKLM:Software\Microsoft\Windows\CurrentVersion\Policies\Servicing' -Name RepairContentServerSource -OutVariable $RepairContentServerSource
            }
        } catch {
            Write-Error 'Unable to install the Active Directory module.  Remediate, then try again.'; break
        }
    }

    If (!(Get-Module joinmodule)) {
        try {
            If ($DMEProxy) {
                Install-Module -Name joinmodule -Proxy $DMEProxy
            } else {
                Install-Module -Name joinmodule
            }

            Import-Module JoinModule -Force -DisableNameChecking
        } catch {
            Write-Error 'Unable to install the Active Directory module.  Remediate, then try again.'; break
        }
    }

    try {
        #Gather AD objects 
        $adResult = Get-ADComputer -Filter $adQuery -Properties Name, Enabled, OperatingSystem, PasswordLastSet, LastLogonDate -ErrorAction SilentlyContinue
        Write-Verbose -Message "AD query returned $($adResult.Count) results"
    } catch [System.Security.Authentication.AuthenticationException] {
        Write-Warning -Message "Access denied"; break
    } catch [System.Exception] {
        Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to AD Web Services."; break
    }

    try {
        #Gather CM objects 
        $cmResult = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery "SELECT Name0 AS Name, BuildExt AS CMOSBuild, Last_Logon_Timestamp0 AS CMLastLogonDate, Client0 AS CMClient, Client_Version0 AS CMClientVersion,Active0 AS CMActive FROM V_R_System Where BuildExt LIKE '10.%' AND Obsolete0 = 0" -ErrorAction SilentlyContinue
        Write-Verbose -Message "CM query returned $($cmResult.Count) results"

        $ADCMResult = Join-Object -LeftObject $adResult -RightObject $cmResult -On Name -JoinType Full
        
        $ADCMResult | Export-Csv -Path $CSVFilePath -NoTypeInformation

        $ADCMSummary = @{
            NotInAD       = $ADCMResult | Where-Object { $null -eq $_.DistinguishedName }
            NotInCM       = $ADCMResult | Where-Object { $null -eq $_.CMClient }
            MissingClient = $ADCMResult | Where-Object { $_.CMClient -eq 0 }
        }

        
        Write-Host -NoNewline 'CM object not in AD: '
        Write-Host $ADCMSummary.NotInAD.Count
        Write-Host -NoNewline 'AD object not in CM: '
        Write-Host $ADCMSummary.NotInCM.Count
        Write-Host -NoNewline 'CM object no client: '
        Write-Host $ADCMSummary.MissingClient.Count
        Write-Host ">>Detailed report: $($CSVFilePath)"
    } catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    } catch [System.Exception] {
        Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
    }
    Return $ADCMSummary
}