<#
.SYNOPSIS
Reports Windows patch compliance 

.DESCRIPTION
Compliance based on systems that have installed the selected Cumulative Update (or later)

.PARAMETER currentRevision
Parameter description: Patch build number.  If blank, you may choose from a list of deployed cululative updates

.PARAMETER minimumMajor
Parameter description: Remove this.  Minimum Windows major version (e.g. 10 for Windows 10)

.PARAMETER minimumBuild
Parameter description: Minimum feature update version

.PARAMETER DeploymentDate
Parameter description: Systems not active after the deployment date will be added to the exclusion total

.PARAMETER DeploymentAge
Parameter description: Instead of deployment date, you can specify age in days

.PARAMETER TargetCollection
Parameter description: Return compliance data for one collection only

.PARAMETER FeatureUpdateDeployment
Parameter description: Compliance focusing only on feature update version

.PARAMETER ReportAllCollections
Parameter description: move this to a config file

.PARAMETER Config
Parameter description: need to use this for TargetCollection(s)

.PARAMETER ScriptDir
Parameter description: rename to $OutputFolder, Output folder for logs, html, csv files.  If blank, output folder will be the temp folder

.PARAMETER ExternalSourceFlag
Parameter description: this is maybe too client specific.  consider removing

.PARAMETER ExternalSource
Parameter description: again, too client specific

.PARAMETER exclude
Parameter description: not sure if this is needed.  review

.PARAMETER excludeUnavailable
Parameter description: review

.PARAMETER ExportHTML
Parameter description: Switch, used if you want to export a summary to HTML

.PARAMETER ExportCSV
Parameter description: Switch, used if you want to export all systems to CSV

.PARAMETER ReturnMembers
Parameter description: Switch, used if you want to return all systems

.EXAMPLE
$returnCompliance = Get-DMEComplianceReport -currentRevision $CurrentRevision -minimumBuild 19044 -ReturnMembers -ExportCSV -ExportHTMLAn example

.NOTES
Adding to DMEToolbox Module
#>
Function Get-DMEComplianceReport {
    [CmdLetBinding(DefaultParameterSetName = 'Default')]
    Param
    (
        [Parameter(ParameterSetName = 'Default')]
        [int]
        $currentRevision, # 2006 = July: 1826, # KB5015807 (OS Builds 19042.1826, 19043.1826, and 19044.1826)
    
        [Parameter(ParameterSetName = 'Default')]
        [int]
        $minimumMajor = 10, # 10 = Windows 10
    
        [Parameter(ParameterSetName = 'Default')]
        [int]
        $minimumBuild = 19044, # 19042 = 20H2
    
        [Parameter(ParameterSetName = 'Default')]
        [datetime]
        $DeploymentDate,
    
        [Parameter(ParameterSetName = 'Default')]
        [int]
        $DeploymentAge = 30,
    
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'focused')]
        [ValidateSet('Report - TELUS (excludes executive and stores)',
            'Report - Executive_Reports',
            'Report - All Store',
            'Deploy 99 - Windows Update Pre-Pilot',
            'Deploy 00 - Windows Update Pilot',
            'Deploy 11 - Support exec Ous',
            'Deploy 01 - Windows Update Group 1',
            'Deploy 21 - Call Center Group 1',
            'Deploy 02 - Windows Update Group 2',
            'Deploy 22 - Call Center Group 2',
            'Deploy 23 - Call Center Group 3',
            'Deploy 03 - Windows Update Group 3',
            'Deploy 04 - Windows Update Group 4',
            'Deploy 05 - Windows Update Group 5',
            'Deploy 06 - Windows Update Group 6',
            'Deploy 07 - Windows Update Group 7',
            'Deploy 08 - Windows Update Group 8',
            'Deploy 09 - Windows Update Group 9',
            'Deploy 10 - Windows Update Group 10',
            'Deploy 30 - Stores Pilot',
            'Deploy 12 - Executive Ous',
            'Deploy 31 - Stores Group 1',
            'Deploy 32 - Stores Group 2',
            'Deploy 33 - Stores Group 3',
            'Deploy 34 - Stores Group 4',
            'Deploy 35 - Stores Group 5')]
        [string]$TargetCollection,

        [Parameter(ParameterSetName = 'Default')]
        [ValidateSet(19042, 19043, 19044, 19045, 22000, 22621)]
        [int]$FeatureUpdateDeployment,
    
        [Parameter(ParameterSetName = 'Default')]
        [switch]$ReportAllCollections,
    
    
        [Parameter(ParameterSetName = 'Default')]
        [String]$Config,

        [Parameter(ParameterSetName = 'Default')]
        [String]$ScriptDir = $env:TEMP,

        #[Parameter(ParameterSetName = 'Default')]
        #[Parameter(ParameterSetName = 'externalbaseline')]
        #[switch]$ExternalSourceFlag,
    
        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'externalbaseline')]
        [ValidateSet('Report - TELUS (excludes executive and stores)', 'Report - Executive_Reports', 'Report - All Store')]
        [string]$ExternalSource,

        [Parameter(ParameterSetName = 'Default')]
        [Parameter(ParameterSetName = 'externalbaseline')]
        [string[]]$exclude,

        [Parameter(ParameterSetName = 'Default')]
        [switch]$excludeUnavailable,

        [Parameter(ParameterSetName = 'Default')]
        [switch]
        $ExportHTML,

        [Parameter()]
        [switch]
        $ExportCSV,

        [Parameter()]
        [switch]
        $ReturnMembers
    )

    Begin {
        

        #If (-NOT(Test-Path -Path $ScriptDir)) { $ScriptDir = Split-Path -Path $MyInvocation.MyCommand.Definition }

        #. $ScriptDir\Invoke-DMESQLQuery.ps1

        Function Export-toHTML {
            Param
            (
                [PSObject]$Report,
                [string]$ExportPath
            )

        
        
            $report | ConvertTo-Html -Property Report, Total, NumberCompleted, NumberExcluded, NumberUnsupported, PercentCompliant, PercentSupported, PercentAvailable | Out-File -FilePath $ExportPath -Append

            <#
        $Style = Get-Content -Path $ScriptDir\bin\style.css
        $Header = '<div class="fixed-header"><div class="container"><h1 style=font-weight:700>CGI</h1><h2>Patch Compliance</h2></div></div>'
        $Footer = '<div class="fixed-footer"><div class="container"><p>Report run 08/24/2022 02:08:46<br>Custom report - developed by Ed Frost</p><h1 style=font-weight:700>CGI</h1>'
        
        If (!(Test-Path -Path $ExportPath)) 
        {
            ConvertTo-Html –Head $Style –Body $Header –CssUri "http://www.w3schools.com/lib/w3.css" | Out-File -FilePath $ExportPath 
        }
       
        $Report | ConvertTo-Html -Head $Style | Out-File -FilePath $ExportPath -Append

        If (!(Get-Content -Path $ExportPath | Select-String '<div class="fixed-footer"><div class="container"><p>Report run 08/24/2022 02:08:46<br>Custom report - developed by Ed Frost</p><h1 style=font-weight:700>CGI</h1>' -SimpleMatch))
        {
            ConvertTo-Html –Head $Style –Body $Footer –CssUri "http://www.w3schools.com/lib/w3.css" | Out-File -FilePath $ExportPath -Append
        }
        #>
            Write-Verbose "HTML Export: $ExportPath"
        }
    
        #$supportedBuilds = @(19044, 19043, 19045) # 20H2, 21H1, 21H2

        If ($targetCollection) {
            $targetCollections = @($TargetCollection)
        } Elseif ($ReportAllCollections) {
            $targetCollections = @('Deploy 99 - Windows Update Pre-Pilot',
                'Deploy 00 - Windows Update Pilot',
                'Deploy 01 - Windows Update Group 1',
                'Deploy 02 - Windows Update Group 2',
                'Deploy 03 - Windows Update Group 3',
                'Deploy 04 - Windows Update Group 4',
                'Deploy 05 - Windows Update Group 5',
                'Deploy 06 - Windows Update Group 6',
                'Deploy 07 - Windows Update Group 7',
                'Deploy 08 - Windows Update Group 8',
                'Deploy 09 - Windows Update Group 9',
                'Deploy 10 - Windows Update Group 10',
                'Deploy 11 - Support exec Ous',
                'Deploy 12 - Executive Ous',
                'Deploy 21 - Call Center Group 1',
                'Deploy 22 - Call Center Group 2',
                'Deploy 23 - Call Center Group 3',
                'Deploy 30 - Stores Pilot',
                'Deploy 31 - Stores Group 1',
                'Deploy 32 - Stores Group 2',
                'Deploy 33 - Stores Group 3',
                'Deploy 34 - Stores Group 4',
                'Deploy 35 - Stores Group 5')
        } Else {
            $targetCollections = @('Report - TELUS (excludes executive and stores)', 'Report - Executive_Reports', 'Report - All Store')
        }

        #$targetCollection = 'Report - TELUS (excludes executive and stores)'
        #$targetCollection = 'Report - Executive_Reports'
        #$targetCollection = 'Report - All Store'
        #$targetCollection = $null
        $members = $null

        <#If ($Config -ne $null)
    {
        If (Test-Path $Config)
        {
            
        }
        $XMLSettings = Select-Xml -
        '$ScriptDir\Automate Process\_PRODUCTION\inbox\Get-MECMComplianceReport.xml'
        Import-Clixml -Path '$ScriptDir\Automate Process\_PRODUCTION\inbox\Get-MECMComplianceReport.xml'
    }#>

        # If not passed via CurrentRevision parameter, get value from system
        If (-Not($CurrentRevision)) {
            try {
                $queryLatestUpdate = "SELECT TOP(10) * FROM v_UpdateInfo WHERE IsDeployed = '1' AND Title LIKE '%Cumulative Update for Windows%' ORDER BY DatePosted DESC"
                $LatestUpdates = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $queryLatestUpdate
                #$CurrentRevision = ([version](Get-ComputerInfo | Select-Object -ExpandProperty OsHardwareAbstractionLayer)).Revision
                $SelectedUpdate = $LatestUpdates | Select-Object -Property Title, MinSourceVersion, DatePosted, InfoURL  | Out-GridView -Title 'Last 10 Updates' -OutputMode Single 
                $RegExPattern = [regex] '^[0-9]{5}[.]{1}[0-9]{4}'

                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 <# using TLS 1.2 is vitally important #>
                If ($DMEProxy) {
                    $Request = Invoke-Webrequest -URI $SelectedUpdate.InfoURL -Proxy $DMEProxy
                } else {
                    $Request = Invoke-Webrequest -URI $SelectedUpdate.InfoURL
                }
            } catch {
                If ($DMEProxy) {
                    Write-Warning -Message "Unable to connect to $($SelectedUpdate.InfoURL). Consider using proxy."; break
                } else {
                    Write-Warning -Message "Unable to connect to $($SelectedUpdate.InfoURL).  Check your proxy settings."; break
                }
            }
            
            
            $currentRevision = [string]$($Request.ParsedHtml.Title.replace(")", "").replace("(", "") -split " " | Where-Object { $_ -match $RegExPattern } | Select-Object -first 1).split(".")[-1]
            Write-Verbose "Current Revision: $($currentRevision)"
        }

        # Calculate days since deployment
        If ($DeploymentDate) { 
            $DeploymentAge = ($(Get-Date) - $DeploymentDate).Days 
        }

        $query = "DECLARE @CollectionName As VarChar(255)
            DECLARE @PatchCycleDays AS INT
            DECLARE @MinimumBuild AS INT
            DECLARE @__timezoneoffset INT SELECT @__timezoneoffset = DateDiff(ss,getutcdate(),getdate());
            SET @CollectionName = '<COLLECTIONNAME>'
            SET @PatchCycleDays = $($DeploymentAge)
            SET @MinimumBuild = $($minimumBuild)

            SELECT 
                SYS.Name0 DeviceName,
                FCM.CollectionID, 
                COL.Name AS CollectionName, 
                OS.BuildNumber0 AS BuildNumber,
                SYS.BuildExt, 
                RIGHT(SYS.BuildExt,4) AS BuildRev,
                SYS.Active0 AS Active,
                SYS.AD_Site_Name0 AS ADSite,
                SYS.Client0 AS Client,
                SYS.Client_Version0 AS ClientVersion,
                Decommissioned0 AS Decommissioned,
                SYS.Distinguished_Name0 AS DistinguishedName,
                SYS.Is_Virtual_Machine0 AS ISVirtualMachine,
                SYS.Last_Logon_Timestamp0 AS LastLogon,
                DateAdd(ss,@__timezoneoffset ,USS.LastScanTime) LastScanTime,
                DateDiff(D, USS.LastScanTime, GetDate()) 'LastScanTime (Days)',
                SYS.Obsolete0 AS Obsolete,
                'Patchable'= 
                    CASE 
                        WHEN DateDiff(D, USS.LastScanTime, GetDate()) >= @PatchCycleDays THEN 'FALSE'
                        ELSE 'TRUE'
                    END,
                'Windows 10 Version'=
                    CASE OS.BuildNumber0
                        WHEN '22000' THEN 'Windows 11'
                        WHEN '19045' THEN 'Windows 10 22H2'
                        WHEN '19044' THEN 'Windows 10 21H2'
                        WHEN '19043' THEN 'Windows 10 21H1'
                        WHEN '19042' THEN 'Windows 10 20H2'
                        WHEN '19041' THEN 'Windows 10 2004'
                        WHEN '18363' THEN 'Windows 10 1909'
                        WHEN '18362' THEN 'Windows 10 1903'
                        WHEN '17763' THEN 'Windows 10 1809'
                        WHEN '17134' THEN 'Windows 10 1803'
                        WHEN '16299' THEN 'Windows 10 1709'
                        WHEN '15063' THEN 'Windows 10 1703'
                        WHEN '14393' THEN 'Windows 10 1607'
                        WHEN '10586' THEN 'Windows 10 1511'
                        WHEN '10240' THEN 'Windows 10 1507'
                    End,
                'Supported'=
                    CASE 
                        WHEN OS.BuildNumber0 >= @MinimumBuild THEN 'TRUE'
                        ELSE 'FALSE'
                    End
            FROM V_R_SYSTEM SYS 
            INNER JOIN v_FullCollectionMembership FCM ON FCM.ResourceID = SYS.ResourceID
            LEFT JOIN v_Collection COL ON FCM.CollectionID = COL.CollectionID 
            LEFT JOIN v_GS_OPERATING_SYSTEM OS ON FCM.ResourceID = OS.ResourceID
            LEFT Join v_UpdateScanStatus USS ON SYS.ResourceID = USS.ResourceID
            WHERE COL.Name LIKE @CollectionName
            --AND OS.Caption0 like 'Microsoft Windows 10%'"

        ($targetCollections | ForEach-Object { $strCollections += "'" + $($_) + "'," })
        $strCollections = $strCollections.TrimEnd(',')

        $queryExt = "DECLARE @CollectionName As VarChar(255)
            DECLARE @PatchCycleDays AS INT
            DECLARE @MinimumBuild AS INT
            DECLARE @__timezoneoffset INT SELECT @__timezoneoffset = DateDiff(ss,getutcdate(),getdate());
            SET @CollectionName = '<COLLECTIONNAME>'
            SET @PatchCycleDays = $($DeploymentAge)
            SET @MinimumBuild = $($minimumBuild)

            SELECT 
                SYS.Name0 DeviceName,
                FCM.CollectionID, 
                COL.Name AS CollectionName, 
                OS.BuildNumber0 AS BuildNumber,
                SYS.BuildExt, 
                RIGHT(SYS.BuildExt,4) AS BuildRev,
                SYS.Active0 AS Active,
                SYS.AD_Site_Name0 AS ADSite,
                SYS.Client0 AS Client,
                SYS.Client_Version0 AS ClientVersion,
                Decommissioned0 AS Decommissioned,
                SYS.Distinguished_Name0 AS DistinguishedName,
                SYS.Is_Virtual_Machine0 AS ISVirtualMachine,
                SYS.Last_Logon_Timestamp0 AS LastLogon,
                DateAdd(ss,@__timezoneoffset ,USS.LastScanTime) LastScanTime,
                DateDiff(D, USS.LastScanTime, GetDate()) 'LastScanTime (Days)',
                SYS.Obsolete0 AS Obsolete,
                'Patchable'= 
                    CASE 
                        WHEN DateDiff(D, USS.LastScanTime, GetDate()) >= @PatchCycleDays THEN 'FALSE'
                        ELSE 'TRUE'
                    END,
                'Windows 10 Version'=
                    CASE OS.BuildNumber0
                        WHEN '22000' THEN 'Windows 11'
                        WHEN '19045' THEN 'Windows 10 22H2'
                        WHEN '19044' THEN 'Windows 10 21H2'
                        WHEN '19043' THEN 'Windows 10 21H1'
                        WHEN '19042' THEN 'Windows 10 20H2'
                        WHEN '19041' THEN 'Windows 10 2004'
                        WHEN '18363' THEN 'Windows 10 1909'
                        WHEN '18362' THEN 'Windows 10 1903'
                        WHEN '17763' THEN 'Windows 10 1809'
                        WHEN '17134' THEN 'Windows 10 1803'
                        WHEN '16299' THEN 'Windows 10 1709'
                        WHEN '15063' THEN 'Windows 10 1703'
                        WHEN '14393' THEN 'Windows 10 1607'
                        WHEN '10586' THEN 'Windows 10 1511'
                        WHEN '10240' THEN 'Windows 10 1507'
                    End,
                'Supported'=
                    CASE 
                        WHEN OS.BuildNumber0 >= @MinimumBuild THEN 'TRUE'
                        ELSE 'FALSE'
                    End
            FROM V_R_SYSTEM SYS 
            INNER JOIN v_FullCollectionMembership FCM ON FCM.ResourceID = SYS.ResourceID
            INNER JOIN v_Collection COL ON FCM.CollectionID = COL.CollectionID 
            INNER JOIN v_GS_OPERATING_SYSTEM OS ON FCM.ResourceID = OS.ResourceID
            INNER Join v_UpdateScanStatus USS ON SYS.ResourceID = USS.ResourceID
            WHERE COL.Name LIKE @CollectionName
            AND SYS.Name0 in (<SYSTEMLIST>)
            --AND OS.Caption0 like 'Microsoft Windows 10%'"
    }

    Process {
        try {
            $members = $null
            If ($targetCollection) { 
                If ($ExternalSourceFlag) {
                    Write-Verbose -Message "Retrieving members of: $($targetCollection) (external data source)"
                    $content = Import-Csv -Path $ExternalSourceFlagFile
                    $systemList = $content | Where-Object { $_.Status -eq 'Deployed' -and $_.'Asset Number' -ne '' } | Select-Object -Property 'Asset Number' | ForEach-Object { "'" + $($_.'Asset Number') + "'," }
                    $systemList[$systemList.Length - 1] = $systemList[$systemList.Length - 1].Replace(",", "")
                    $members = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $queryExt.Replace('<SYSTEMLIST>', $SystemList).Replace('<COLLECTIONNAME>', $targetCollection)

                } Else {
                    Write-Verbose -Message "Retrieving members of: $($targetCollection)"
                    #$members = Get-CMCollectionMember -CollectionName $targetCollection | Select-Object -Property *, @{ Name='BuildM'; Expression={([Version]$_.DeviceOSBuild).Major} }, @{ Name='BuildMR'; Expression={([Version]$_.DeviceOSBuild).MajorRevision} }, @{ Name='BuildB'; Expression={([Version]$_.DeviceOSBuild).Build} }, @{ Name='BuildR'; Expression={([Version]$_.DeviceOSBuild).Revision} }, @{ Name='ReportGroup'; Expression={$targetCollection} } 
                    $members = Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query.Replace('<COLLECTIONNAME>', $targetCollection)
                }
            } Else { 
                foreach ($targetCollection in $targetCollections) `
                { 
                    If ($ExternalSourceFlag -eq $true -and $ExternalSource -eq $targetCollection) {
                        Write-Verbose -Message "Retrieving members of: $($targetCollection) (external data source)"
                        $content = Import-Csv -Path $ExternalSourceFlagFile
                        $systemList = $content | Where-Object Status -eq 'Deployed' | Select-Object -Property 'Asset Number' | ForEach-Object { "'" + $($_.'Asset Number') + "'," }
                        $systemList[$systemList.Length - 1] = $systemList[$systemList.Length - 1].Replace(",", "")
                        $members += Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $queryExt.Replace('<SYSTEMLIST>', $SystemList).Replace('<COLLECTIONNAME>', $targetCollection)
                    } Else {
                        Write-Verbose -Message "Retrieving members of: $($targetCollection)" 
                        $members += Invoke-DMESQLQuery -SQLServer $DMESiteDBServer -SQLDatabase $DMESiteDB -SQLQuery $query.Replace('<COLLECTIONNAME>', $targetCollection)
                    }
                } 
                #$targetCollection = $null
            }
        } catch [System.UnauthorizedAccessException] {
            Write-Warning -Message "Access Denied"; break;
        } catch [System.Exception] {
            Write-Warning -Message "Unable to retrieve data.  Please ensure you're able to connect to $($DMESiteDBServer) or check your values in your config file."; break
        }
        

        #try
        #{
        Write-Verbose -Message "Returned $($members.count) systems"

        Write-Verbose -Message "Removing $($members | Where-Object { $_.DeviceName -notlike 'L*' -and $_.DeviceName -notlike 'D*' -and $_.DeviceName -notlike 'TB*' } | Measure-Object | Select-Object -ExpandProperty count) systems due to invalid naming convention."
        $members = $members | Where-Object { $_.DeviceName -like 'L*' -or $_.DeviceName -like 'D*' -or $_.DeviceName -like 'TB*' }

        Write-Verbose -Message "Removing $($members | Where-Object { $_.Active -ne 1 -or $_.Client -ne 1 -or $_.Decommissioned -ne 0 -or $_.Obsolete -ne 0 -or $_.IsVirtualMachine -ne $false} | Measure-Object | Select-Object -ExpandProperty Count) inactive systems"
        $members = $members | Where-Object { $_.Active -eq 1 -and $_.Client -eq 1 -and $_.Decommissioned -eq 0 -and $_.Obsolete -eq 0 -and $_.IsVirtualMachine -eq $false }

        Write-Verbose -Message "Compiling report..."
        $compliance = New-Object PSObject
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name Report 
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name Total 
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name NumberCompleted
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name NumberExcluded
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name NumberUnsupported
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name PercentCompliant
        Add-Member -InputObject $compliance -MemberType NoteProperty -Value '' -Name PercentSupported

        $ComplianceView = @()

        $ErrorActionPreference = 'SilentlyContinue'
        If ($FeatureUpdateDeployment) {
            Write-Verbose -Message "Compliance based on feature update [$($FeatureUpdateDeployment)] deployment"
            foreach ($targetCollection in $targetCollections) {
                $Report = $targetCollection
                $Total = $members | Where-Object { $_.CollectionName -eq $targetCollection } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberCompleted = $members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.BuildNumber -ge $FeatureUpdateDeployment } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberExcluded = $members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.Patchable -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberUnsupported = $members | Where-Object { $_.CollectionName -eq $targetCollection -and ($_.BuildNumber -lt $minimumBuild) } | Measure-Object | Select-Object -ExpandProperty Count
                $PercentCompliant = (($NumberCompleted + $NumberUnsupported) / $Total).ToString("P")
                $PercentSupported = ($NumberCompleted / ($Total - $NumberUnsupported)).ToString("P")
                $PercentSupportedAvailable = ($NumberCompleted / ($Total - $NumberUnsupported - $NumberExcluded)).ToString("P")

                $compliance = [PSCustomObject]@{
                    Report            = $report
                    Total             = $Total
                    NumberCompleted   = $NumberCompleted
                    NumberExcluded    = $NumberExcluded
                    NumberUnsupported = $NumberUnsupported
                    PercentCompliant  = $PercentCompliant
                    PercentSupported  = $PercentSupported
                    PercentAvailable  = $PercentSupportedAvailable
                }

                $ComplianceView += $compliance 
            }
            $TargetCollection = $null
        } Else {
            foreach ($targetCollection in $targetCollections) {
                $Report = $targetCollection
                $Total = $members | Where-Object { $_.CollectionName -eq $targetCollection } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberCompleted = $members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.BuildNumber -ge $minimumBuild -and $_.BuildRev -ge $currentRevision } | Measure-Object | Select-Object -ExpandProperty Count
                $ManualExclusions = $members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.BuildNumber -ge $minimumBuild -and $_.BuildRev -lt $currentRevision -and $_.DeviceName -in $exclude } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberExcluded = $ManualExclusions + ($members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.Patchable -eq $false } | Measure-Object | Select-Object -ExpandProperty Count)
                $NumberUnsupported = $members | Where-Object { $_.CollectionName -eq $targetCollection -and ($_.BuildNumber -lt $minimumBuild) } | Measure-Object | Select-Object -ExpandProperty Count
                $PercentCompliant = (($NumberCompleted + $NumberUnsupported) / $Total).ToString("P")
                $PercentSupported = ($NumberCompleted / ($Total - $NumberUnsupported)).ToString("P")
                $PercentSupportedAvailable = ($NumberCompleted / ($Total - $NumberUnsupported - $NumberExcluded - $manualExclusions)).ToString("P")

                $compliance = [PSCustomObject]@{
                    Report            = $report
                    Total             = $Total
                    NumberCompleted   = $NumberCompleted
                    NumberExcluded    = $NumberExcluded
                    NumberUnsupported = $NumberUnsupported
                    PercentCompliant  = $PercentCompliant
                    PercentSupported  = $PercentSupported
                    PercentAvailable  = $PercentSupportedAvailable
                }

                $ComplianceView += $compliance 
            }
            $TargetCollection = $null
        }

        If ($complianceView.Length -gt 1) {
            If ($FeatureUpdateDeployment) {
                $Report = 'All Groups'
                $Total = $members | Measure-Object | Select-Object -ExpandProperty Count
                $NumberCompleted = $members | Where-Object { $_.BuildNumber -ge $FeatureUpdateDeployment } | Measure-Object | Select-Object -ExpandProperty Count
                $ManualExclusions = $members | Where-Object { $_.CollectionName -eq $targetCollection -and $_.BuildNumber -ge $minimumBuild -and $_.BuildRev -lt $currentRevision -and $_.DeviceName -in $exclude } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberExcluded = $members | Where-Object { $_.Patchable -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberUnsupported = $members | Where-Object { ($_.BuildNumber -lt $minimumBuild) } | Measure-Object | Select-Object -ExpandProperty Count
                $PercentCompliant = (($NumberCompleted + $NumberUnsupported) / $Total).ToString("P")
                $PercentSupported = ($NumberCompleted / ($Total - $NumberUnsupported)).ToString("P")
                $PercentSupportedAvailable = ($NumberCompleted / ($Total - $NumberUnsupported - $NumberExcluded)).ToString("P")
        
                $compliance = [PSCustomObject]@{
                    Report            = $report
                    Total             = $Total
                    NumberCompleted   = $NumberCompleted
                    NumberExcluded    = $NumberExcluded
                    NumberUnsupported = $NumberUnsupported
                    PercentCompliant  = $PercentCompliant
                    PercentSupported  = $PercentSupported
                    PercentAvailable  = $PercentSupportedAvailable
                }

                $ComplianceView += $compliance 
            } Else {
                $Report = 'All Groups'
                $Total = $members | Measure-Object | Select-Object -ExpandProperty Count
                $NumberCompleted = $members | Where-Object { $_.BuildNumber -ge $minimumBuild -and $_.BuildRev -ge $currentRevision } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberExcluded = $members | Where-Object { $_.Patchable -eq $false } | Measure-Object | Select-Object -ExpandProperty Count
                $NumberUnsupported = $members | Where-Object { ($_.BuildNumber -lt $minimumBuild) } | Measure-Object | Select-Object -ExpandProperty Count
                $PercentCompliant = (($NumberCompleted + $NumberUnsupported) / $Total).ToString("P")
                $PercentSupported = ($NumberCompleted / $Total).ToString("P")
                $PercentSupportedAvailable = ($NumberCompleted / ($Total - $NumberExcluded)).ToString("P")
        
                $compliance = [PSCustomObject]@{
                    Report            = $report
                    Total             = $Total
                    NumberCompleted   = $NumberCompleted
                    NumberExcluded    = $NumberExcluded
                    NumberUnsupported = $NumberUnsupported
                    PercentCompliant  = $PercentCompliant
                    PercentSupported  = $PercentSupported
                    PercentAvailable  = $PercentSupportedAvailable
                }

                $ComplianceView += $compliance 
            }
            $ErrorActionPreference = 'Continue'
        }
        foreach ($item in $ComplianceView) { Write-Verbose "$($item.report): $($item.NumberCompleted)/ $($item.Total)" }
        #}
        #catch [System.UnauthorizedAccessException]
        #{
        #    Write-Warning -Message "Access Denied"; break;
        #}
        #catch [System.Exception]
        #{
        #    Write-Warning -Message "Unable display data."; break
        #}
    }

    End {
    
        If ($ExportHTML) {
            Export-toHTML -Report $ComplianceView -ExportPath "$ScriptDir\ComplianceReport-complianceview_$(Get-Date -Format yyyy-MM-dd).html"
        }
    
        If ($ExportCSV) {
            $members | Export-Csv -Path "$ScriptDir\ComplianceReport-complianceview_$(Get-Date -Format yyyy-MM-dd).csv"
            Write-Verbose "CSV Output: $ScriptDir\ComplianceReport-complianceview_$(Get-Date -Format yyyy-MM-dd).csv"
        }

        $ComplianceView | Format-Table -AutoSize

        If ($ReturnMembers) {
            Return $members
        }
    }
}