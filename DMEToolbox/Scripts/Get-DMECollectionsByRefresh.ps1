<#
.SYNOPSIS
Query for MECM collections by refresh type

.DESCRIPTION
Returns list of collections based on criteria, to assist with collection optimization

.PARAMETER RefreshType
Both, Continuous, Manual, None

.PARAMETER CollectionType
Device, User

.PARAMETER MemberCountMinimum
To allow or prevent the return of empty collections

.PARAMETER MemberCountMaximum
If you choose to target smaller collections or only empty collections

.PARAMETER NoQueryMembershipRules
If you are targeting collections with 0 query membership rules

.PARAMETER NoDirectMembershipRules
If you are targeting collections with 0 direct membership rules

.PARAMETER Preset
'Incremental & No Query Rules', 'Incremental & Scheduled', 'Scheduled & No Query Rules'

.PARAMETER Limit
Limit the amount of collections returned.  Handy if automating a phase of your maintenance where you need to limit the target collections (e.g. pilot phase)

.PARAMETER LogFolder
by default, the temp environment variable will be referenced unless this paramater is defined

.EXAMPLE
Get-DMECollectionsByRefresh -RefreshType Both -CollectionType Device

.NOTES
2023.07.10 - added to DMEToobox module
#>
Function Get-DMECollectionsByRefresh {
    [CmdletBinding(DefaultParameterSetName = 'Parameter Set 1', 
        SupportsShouldProcess = $true, 
        PositionalBinding = $false,
        HelpUri = 'http://www.microsoft.com/',
        ConfirmImpact = 'Medium')]
    Param
    (
        # Specify collection membership refresh type
        [Parameter(Mandatory = $True, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False, 
            ParameterSetName = 'Parameter Set 1')]
        [ValidateSet('None', 'Manual', 'Periodic', 'Continuous', 'Both')]
        [string]
        $RefreshType,
        
        # Filter by collection type. Null will result in no filter
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False, 
            ParameterSetName = 'Parameter Set 1')]
        [ValidateSet('Device', 'User', 'Other')]
        [string]
        $CollectionType = 'Device',

        # Minimum member count
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False,
            ParameterSetName = 'Parameter Set 1')]
        [int]
        $MemberCountMinimum,

        # Maximum member count
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False,
            ParameterSetName = 'Parameter Set 1')]
        [int]
        $MemberCountMaximum,

        # Change to 1 if you're looking for dynamic collections.  Default will look for collections with 0 query rules.
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False,
            ParameterSetName = 'Parameter Set 1')]
        [switch]
        $NoQueryMembershipRules,

        # Change to 1 if you're looking for static collections rules. Default will look for collections with 0 direct rules.
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False,
            ParameterSetName = 'Parameter Set 1')]
        [switch]
        $NoDirectMembershipRules,

        # Param2 help description
        [Parameter(Mandatory = $True, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False, 
            ParameterSetName = 'Parameter Set 2')]
        [ValidateSet('Incremental & No Query Rules', 'Scheduled & No Query Rules', 'Incremental & Scheduled')]
        [string]
        $Preset,

        # Limit results returned
        [Parameter(Mandatory = $False, 
            ValueFromPipeline = $False,
            ValueFromPipelineByPropertyName = $False, 
            ValueFromRemainingArguments = $False,
            ParameterSetName = 'Parameter Set 1')]
        [int]
        $Limit,

        [string]
        $LogFolder = $env:TEMP
    )

    Begin {
        Switch ($Preset) {
            'Incremental & No Query Rules' {
                $NoQueryMembershipRules = $true
                $RefreshType = 'Continuous'
            }
            'Scheduled & No Query Rules' {
                $NoQueryMembershipRules = $true
                $RefreshType = 'Periodic'
            }
            'Incremental & Scheduled' {
                $RefreshType = 'Both'
            }
        }

        # Build query
        $query = "SELECT `n"       

        $query += "CG.CollectionName,
                CG.SITEID AS [CollectionID],
                CASE VC.CollectionType
                    WHEN 0 THEN 'Other'
                    WHEN 1 THEN 'User'
                    WHEN 2 THEN 'Device'
                    ELSE 'Unknown' 
                END AS CollectionType,
                CG.LastChangeTime,
                CG.schedule, 
                CASE
                    WHEN CG.Schedule like '%0102000' THEN 'Every 1 min'
                    WHEN CG.Schedule like '%010A000' THEN 'Every 5 mins'
                    WHEN CG.Schedule like '%0114000' THEN 'Every 10 mins'
                    WHEN CG.Schedule like '%011E000' THEN 'Every 15 mins'
                    WHEN CG.Schedule like '%0128000' THEN 'Every 20 mins'
                    WHEN CG.Schedule like '%0132000' THEN 'Every 25 mins'
                    WHEN CG.Schedule like '%013C000' THEN 'Every 30 mins'
                    WHEN CG.Schedule like '%0150000' THEN 'Every 40 mins'
                    WHEN CG.Schedule like '%015A000' THEN 'Every 45 mins'
                    WHEN CG.Schedule like '%0100100' THEN 'Every 1 hour'
                    WHEN CG.Schedule like '%0100200' THEN 'Every 2 hours'
                    WHEN CG.Schedule like '%0100300' THEN 'Every 3 hours'
                    WHEN CG.Schedule like '%0100400' THEN 'Every 4 hours'
                    WHEN CG.Schedule like '%0100500' THEN 'Every 5 hours'
                    WHEN CG.Schedule like '%0100600' THEN 'Every 6 hours'
                    WHEN CG.Schedule like '%0100700' THEN 'Every 7 hours'
                    WHEN CG.Schedule like '%0100800' THEN 'Every 8 hours'
                    WHEN CG.Schedule like '%0100B00' THEN 'Every 11 Hours'
                    WHEN CG.Schedule like '%0100C00' THEN 'Every 12 Hours'
                    WHEN CG.Schedule like '%0101000' THEN 'Every 16 Hours'
                    WHEN CG.Schedule like '%0100008' THEN 'Every 1 days'
                    WHEN CG.Schedule like '%0100010' THEN 'Every 2 days'
                    WHEN CG.Schedule like '%0100018' THEN 'Every 3 days'
                    WHEN CG.Schedule like '%0100028' THEN 'Every 5 days'
                    WHEN CG.Schedule like '%0100038' THEN 'Every 7 Days'
                    WHEN CG.Schedule like '%0100C38' THEN 'Every 7 Days'
                    WHEN CG.Schedule like '%0192000' THEN '1 week'
                    WHEN CG.Schedule like '%01F2000' THEN '1 week'
                    WHEN CG.Schedule like '%01E2000' THEN '1 week'
                    WHEN CG.Schedule like '%0080000' THEN 'Update Once'
                    WHEN CG.SChedule = ''THEN 'Manual'
                END AS [Update Schedule],
                VC.RefreshType as RefreshTypeID,
                CASE VC.RefreshType
                    WHEN 0 THEN 'None'
                    WHEN 1 THEN 'Manual'
                    WHEN 2 THEN 'Periodic'
                    WHEN 4 THEN 'Continuous'
                    WHEN 6 THEN 'Both'
                    ELSE 'Unknown'
                END as RefreshType,
                VC.MemberCount,
                (SELECT COUNT(CollectionID) FROM v_CollectionRuleQuery CRQ WHERE CRQ.CollectionID = VC.SiteID) AS 'RuleQueryCount',
                (SELECT COUNT(CollectionID) FROM v_CollectionRuleDirect CRD WHERE CRD.CollectionID = VC.SiteID) AS 'RuleDirectCount',
                VC.LimitToCollectionID, VC.ObjectPath
                FROM
                dbo.collections_g CG
                LEFT JOIN v_collections VC on VC.SiteID = CG.SiteID `n"

        If ($null -ne $RefreshType) {
            Switch ($RefreshType) {
                'None' { $query += "WHERE RefreshType = '0'" }
                'Manual' { $query += "WHERE RefreshType = '1'" }
                'Periodic' { $query += "WHERE RefreshType = '2'" }
                'Continuous' { $query += "WHERE RefreshType = '4'" }
                'Both' { $query += "WHERE RefreshType = '6'" }
            }

            $query += "`n"        
        }

        If ($null -ne $CollectionType) {
            Switch ($CollectionType) {
                'Other' { $query += "AND VC.CollectionType = '0'" }
                'User' { $query += "AND VC.CollectionType = '1'" }
                'Device' { $query += "AND VC.CollectionType = '2'" }
            }

            $query += "`n"
        }

        $query += "GROUP BY CG.SiteID,VC.SiteID,CG.CollectionName,VC.CollectionType,CG.schedule,VC.RefreshType,VC.MemberCount,VC.LimitToCollectionID,VC.ObjectPath,CG.LastChangeTime `n"

        If ($NoQueryMembershipRules) {
            $query += "HAVING (SELECT COUNT(CollectionID) FROM v_CollectionRuleQuery CRQ WHERE CRQ.CollectionID = VC.SiteID) = 0 `n"
        }

        If ($NoDirectMembershipRules) {
            $query += "AND (SELECT COUNT(CollectionID) FROM v_CollectionRuleDirect CRD WHERE CRD.CollectionID = VC.SiteID) = 0 `n"
        }

        $query += "ORDER BY CG.CollectionName ASC`n"

        # END BUILD QUERY

        # Check for log folder 

        If (-not (Test-Path -Path $logFolder)) {
            $LogFolder = $env:TEMP
        }

        $LogFullPath = "$($LogFolder)\Get-DMECollectionsByRefresh_$($CollectionType)-$($RefreshType)_$(Get-Date -Format 'yyyyMMdd').csv"
    }

    Process {
        try {
            # Gather collections
            $return = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query

            $return | Select-Object *, @{ Name = 'Action'; Expression = { 'Review' } } | Export-Csv -Path $LogFullPath -NoTypeInformation

            Write-Host "Output Path: $LogFullPath" -ForegroundColor Green
        } catch [System.UnauthorizedAccessException] {
            Write-Warning -Message "Access Denied" ; break
        } catch [System.Exception] {
            Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
        }
    }

    End {
        If ($Limit -gt 0) {
            Write-Verbose -Message "Returning first $limit results"
            Return $return | Select-Object -first $Limit
        } Else {
            Write-Verbose -Message "Returning $($return.count) results"
            Return $return
        }
    }
}