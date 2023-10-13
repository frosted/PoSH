Function Get-DMEAppVersionReport
{
    [CmdletBinding()]
    [Alias()]
    [OutputType([int])]
    Param
    (
        [string]$ApplicationVendor,
        [Parameter(Mandatory)]
        [String]$ApplicationName,
        [String]$CollectionName,
        [ValidateSet('Summary', 'Detailed')] # Output summary = group by version, detailed = full report
        [string]$Output = 'Summary',
        [System.IO.DirectoryInfo]$OutputPath,
        [System.IO.FileInfo]$OutputFileName = 'Get-DMEAppVersionReport.csv'
    )

    Begin
    {
        $timeStart = Get-Date

        Switch ($Output)
        {
            'Detailed'
            {
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
                FCM.CollectionID,
                COL.Name as 'Collection Name'
                FROM 
                v_R_System VRS
                INNER JOIN v_Add_Remove_Programs VADR on VADR.ResourceID = VRS.ResourceID
                INNER JOIN v_FullCollectionMembership FCM ON FCM.ResourceID = VRS.ResourceID
                INNER JOIN v_Collection COL ON FCM.CollectionID = COL.CollectionID 
                WHERE
                VADR.DisplayName0 IS NOT null 
                AND VADR.Publisher0 LIKE '%$($ApplicationVendor)%'
                AND VADR.DisplayName0 LIKE '%$($ApplicationName)%' 
                AND VADR.DisplayName0 NOT LIKE '%chat'
                AND VADR.DisplayName0 NOT LIKE '%legacy'
                AND VADR.DisplayName0 NOT LIKE '%updater'
                AND VADR.DisplayName0 NOT LIKE '%(DV)'
                AND VADR.DisplayName0 NOT LIKE '%(USB)'
                AND VADR.DisplayName0 NOT LIKE '%(SSON)'
                AND VADR.DisplayName0 NOT LIKE '%(AERO)'
                AND VADR.DisplayName0 NOT LIKE '%(HDX%)'
                AND VADR.DisplayName0 NOT LIKE '%(Redirection%)'
                AND VADR.DisplayName0 NOT LIKE '%Language%'
                AND VRS.Decommissioned0 = 0
                AND VRS.Active0 = 1
                "
                If ($CollectionName)
                {
                    $query += "AND COL.Name = '$($CollectionName)'"
                }
            }
            'Summary'
            {
                $query = "Select 
                COUNT(*) as Total,
                VADR.Publisher0 as 'Publisher',
                VADR.DisplayName0 as 'Software Name',
                VADR.Version0 'Software version'
                FROM 
                v_R_System VRS
                INNER JOIN v_Add_Remove_Programs VADR on VADR.ResourceID = VRS.ResourceID
                INNER JOIN v_FullCollectionMembership FCM ON FCM.ResourceID = VRS.ResourceID
                INNER JOIN v_Collection COL ON FCM.CollectionID = COL.CollectionID 
                WHERE
                VADR.DisplayName0 IS NOT null 
                AND VADR.Publisher0 LIKE '%$($ApplicationVendor)%'
                AND VADR.DisplayName0 LIKE '%$($ApplicationName)%' 
                AND VADR.DisplayName0 NOT LIKE '%chat'
                AND VADR.DisplayName0 NOT LIKE '%legacy'
                AND VADR.DisplayName0 NOT LIKE '%updater'
                AND VADR.DisplayName0 NOT LIKE '%(DV)'
                AND VADR.DisplayName0 NOT LIKE '%(USB)'
                AND VADR.DisplayName0 NOT LIKE '%(SSON)'
                AND VADR.DisplayName0 NOT LIKE '%(AERO)'
                AND VADR.DisplayName0 NOT LIKE '%(HDX%)'
                AND VADR.DisplayName0 NOT LIKE '%(Redirection%)'
                AND VADR.DisplayName0 NOT LIKE '%Language%'
                AND VRS.Decommissioned0 = 0
                AND VRS.Active0 = 1"
                
                If ($CollectionName)
                {
                    $query += [char]10 + "AND COL.Name = '$($CollectionName)'"
                }

                $query += [char]10 + "GROUP BY VADR.Publisher0,VADR.DisplayName0,VADR.Version0
                ORDER BY TOTAL DESC"
            }
        }
        
        If ($null -eq $OutputPath) { [System.IO.DirectoryInfo]$OutputPath = $env:TEMP }

        $CSVPath = $OutputPath.FullName + "\" + $Output + "_" + $OutputFileName.Name
    }
    Process
    {
        try
        {
            $return = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query

            $Return |  Export-Csv -Path $CSVPath -NoTypeInformation -Force

            # output summary to host

            Write-Host "Top 10 $($ApplicationVendor) $($ApplicationName) Versions" -ForegroundColor DarkBlue -BackgroundColor White

            Switch ($Output)
            {
                'Detailed'
                {
                    $Return | Group-Object -Property Publisher, 'Software Name', 'Software Version' | Sort-Object Count -Descending | Select-Object Count, Name -First 10 | Format-Table -AutoSize
                }
                'Summary'
                {
                    $Return | Select-Object -First 10 | Format-Table -AutoSize 
                }
            }
            

            Write-Output "Report path: $CSVPath"
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
