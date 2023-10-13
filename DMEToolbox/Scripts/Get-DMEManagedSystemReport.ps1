<#
.SYNOPSIS
Function to help determine which systems might be live/stale in AD/CM

.DESCRIPTION
This function leverages the ActiveDirectory module to retrieve AD objects, JoinModule to more efficiently join two
datasets and the DMEToolbox Module to retrieve CM objects.  Once run, a summary is provided and results are logged to CSV.

.PARAMETER MaximumAge
Based on the PasswordLastSet AD attribute.  Default value is 30 days

.PARAMETER CSVFilePath
Full path to CSV file where results will be written.  Default value is DMEManagedSystemReport<date>.csv" in your 
TEMP folder

.PARAMETER ReturnVariable
Allowing you to have results stored in a variable for further investigation.  

.EXAMPLE
#Report of systems no more than 90 days old, store results in CSV on root
Get-DMEManagedSystemReport -MaximumAge 90 -CSVFilePath c:\ILikeToStoreFilesOnRoot.csv

#Report of systems no more than 30 days old (default), CSV path in temp (default), have results stored 
# in variable $ResultObject
Get-DMEManagedSystemReport -ReturnVariable ResultObject
#Example of digging through results, returning only systems a) missing client, b) meet a valid naming convention 
# and c) shows activity within the last 30 days
$ResultObject.MissingClient | Select-Object *,@{Name='PlsAge';Expression={((Get-Date) - $_.PasswordLastSet).days}},@{Name='LogonAge';Expression={((Get-Date) - $_.CMLastLogonDate).days}} | Where-Object { $_.Name -match '(TB|[LD])\d{6}' -and (($_.PlsAge -le 30) -or ($_.LogonAge -le 30)) } | Measure-Object
#Last results and group by OS
$result.MissingClient | Select-Object *,@{Name='PlsAge';Expression={((Get-Date) - $_.PasswordLastSet).days}},@{Name='LogonAge';Expression={((Get-Date) - $_.CMLastLogonDate).days}} | Where-Object { $_.Name -match '(TB|[LD])\d{6}' -and (($_.PlsAge -le 14) -or ($_.LogonAge -le 14)) } | Group-Object CMClientVersion | Select-Object count,name | Sort-Object count -Descending

.NOTES
You can use import-csv and your PowerShell skills to further dig into the results from this command
#>
Function Get-DMEManagedSystemReport {
    [CmdletBinding()]
    param (
        [Parameter()]
        [int]
        $MaximumAge = 30,
        [Parameter()]
        [string]
        $CSVFilePath = "$($env:TEMP)\DMEManagedSystemReport$(Get-Date -Format 'yyyyMMdd').csv",
        [Parameter()]
        [string]
        $ReturnVariable
    )

    #Checks Active Directory for objects that are Windows 10 AND the PasswordLastSet attribute is within the last 30 days and adds to variable $adResult.
    $oldestDate = (Get-Date).AddDays(-$MaximumAge)
    $adQuery = 'OperatingSystem -like "Windows 1*" -and PasswordLastSet -ge "' + $oldestDate + '" -and Enabled -eq "True"'

    If (!(Get-Module -Name ActiveDirectory)) {
        try {
            #Gather Settings
            If (Get-ItemPropertyValue -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer -ErrorAction SilentlyContinue) {
                $WUServerSetting = Get-ItemPropertyValue -Path 'HKLM:/Software/Policies/Microsoft/Windows/WindowsUpdate/AU/' -Name UseWUServer -ErrorAction SilentlyContinue
            } Else {
                $WUServerSetting = $null
            }
            
            if (Get-ItemProperty -Path 'HKLM:/Software\Microsoft\Windows\CurrentVersion\Policies\Servicing' -Name RepairContentServerSource -ErrorAction SilentlyContinue) {
                $RepairContentServerSource = Get-ItemPropertyValue -Path 'HKLM:/Software\Microsoft\Windows\CurrentVersion\Policies\Servicing' -Name RepairContentServerSource -ErrorAction SilentlyContinue
            } Else {
                $RepairContentServerSource = $null
            }

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
                Install-Module -Name joinmodule -
            }

            Import-Module JoinModule -Force
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
        Write-Host $ADCMSummary.NotInAD.count
        Write-Host -NoNewline 'AD object not in CM: '
        Write-Host $ADCMSummary.NotInCM.count
        Write-Host -NoNewline 'CM object no client: '
        Write-Host $ADCMSummary.MissingClient.count
        Write-Host ">>Detailed report: $($CSVFilePath)"

        If ($ReturnVariable) {
            Set-Variable -Name $ReturnVariable -Value $ADCMSummary -Scope global
        }
    } catch [System.UnauthorizedAccessException] {
        Write-Warning -Message "Access denied" ; break
    } catch [System.Exception] {
        Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
    }
}