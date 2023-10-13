<#
.Synopsis
    Report on Chrome versions in the environment
.DESCRIPTION
    Leverages SCCM inventory (ARP) to report on Chrome browser versions with option to output a summary to screen and/or write a detailed/summarized report to CSV
    Also leverages a Google API to gather release dates to determine age of each chrome release it detects
.EXAMPLE
    # Will get Chrome summary report & output CSV to temp directory
    Get-DMEChromeVersionReport
.EXAMPLE
    # Will get Chrome summary report, use proxy for Google API calls, output to temp folder & verbose output to screen
    Get-DMEChromeVersionReport -Output Summary -Verbose
.EXAMPLE
    # Will get detailed Chromed report, save CSV to custom location & verbose output to screen
    Get-DMEChromeVersionReport -Output Detailed -OutputPath c:\temp -OutputFileName ChromeVersionReport.csv -Verbose
.NOTES
    Author: Ed Frost
    Date:   2023.06.06
#>
Function Get-DMEChromeVersionReport
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [ValidateSet('Summary', 'Detailed')] # Output summary = group by version, detailed = full report
        [string]$Output = 'Summary',
        [System.IO.DirectoryInfo]$OutputPath,
        [System.IO.FileInfo]$OutputFileName = 'Get-DMEChromeVersionReport.csv'
    )

    Begin
    {
        $timeStart = Get-Date

        $query = "Select 
            VRS.name0 as 'Computer Name',
            VRS.User_Name0 as 'Last Logon User',
            VRS.Last_Logon_Timestamp0 as 'Last Logon Time',
            VRS.Resource_Domain_OR_Workgr0  as 'Domain' , 
            VRS.BuildExt as 'OS Version',
            VRS.Client_Version0 as 'CM Client Version',
            VADR.Publisher0 as 'Publisher',
            VADR.DisplayName0 as 'Software Name',
            VADR.Version0 'Software version',
            VADR.InstallDate0 as 'Install Date',
            'https://versionhistory.googleapis.com/v1/chrome/platforms/all/channels/all/versions/' + VADR.Version0 + '/releases' as URL
        FROM 
            v_R_System VRS
            INNER JOIN v_Add_Remove_Programs VADR on VADR.ResourceID = VRS.ResourceID
        WHERE
            VADR.DisplayName0 IS NOT null 
            AND VADR.Publisher0 LIKE '%Google%'
            AND VADR.DisplayName0 LIKE '%Chrome%' 
            AND VADR.DisplayName0 NOT LIKE '%chat'
            AND VADR.DisplayName0 NOT LIKE '%legacy'
            AND VADR.DisplayName0 NOT LIKE '%updater'
            AND VADR.DisplayName0 NOT LIKE '%(DV)'
            AND VADR.DisplayName0 NOT LIKE '%(USB)'
            AND VADR.DisplayName0 NOT LIKE '%(SSON)'
            AND VADR.DisplayName0 NOT LIKE '%(AERO)'
            AND VADR.DisplayName0 NOT LIKE '%(HDX%)'
            AND VADR.DisplayName0 NOT LIKE '%(Redirection%)'"

        If ($null -eq $OutputPath) { [System.IO.DirectoryInfo]$OutputPath = $env:TEMP }

        $CSVPath = $OutputPath.FullName + "\" + $OutputFileName.Name


    }
    Process
    {
        # Google API return
        try
        {
            If ( (Invoke-RestMethod -Method Get -Uri 'https://versionhistory.googleapis.com/v1/chrome/platforms/win/channels/stable/versions' -Proxy $DMEProxy -TimeoutSec 10 -ErrorAction SilentlyContinue ) -or (-Not (Get-Command Invoke-DMESQLQuery)) )
            { 
                $return = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query

                $ReturnAPI = $return | Where-Object { $null -ne $_.URL } | Group-Object -Property URL | Select-Object count, name, @{ Name = 'Release Date'; Expression = { [datetime]((Invoke-RestMethod -Method Get -Uri $_.Name -Proxy $DMEProxy | Select-Object -ExpandProperty releases | Select-Object -first 1 -ExpandProperty serving).StartTime) } } , @{Name = '%'; Expression = { $($_.Count / $return.count).ToString('p') } } | Select-Object @{Name = 'Version'; Expression = { ($_.Name).Split("/")[-2] } }, @{ Name = 'Age (days)'; Expression = { (New-TimeSpan -End (Get-Date) -Start $_.'Release Date').Days } }, 'Release Date', 'Count', '%' | Sort-Object count -Descending
                Switch ($output)
                {
                    'Summary'
                    {
                        $ReturnAPI | Export-Csv -Path $CSVPath -NoTypeInformation
                    }
                    'Detailed'
                    {
                        $return | Export-Csv -Path $CSVPath -NoTypeInformation
                    }
                }

                # output summary to host

                Write-Host "Top 10 Chrome Versions" -ForegroundColor DarkBlue -BackgroundColor White

                $ReturnAPI | Sort-Object -Property Count -Descending | Select-Object -First 10 | Format-Table -AutoSize 

                Write-Output "Report path: $CSVPath"
            }
            Else
            {
                Write-Error 'Requirements not met.  Make sure you can connect to Google API and have all necessary cmdlets loaded.'
            } 
        }
        catch [System.UnauthorizedAccessException]
        {
            Write-Warning -Message "Access denied" ; break
        }
        catch [System.Exception]
        {
            Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
        }
        

        $timeEnd = Get-Date
    }
    End
    {
        $timeElapsed = $timeEnd - $timeStart 
        Write-Output "Total runtime: $($timeElapsed.ToString())"
    }
}